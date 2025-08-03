VERSION 5.00
Begin VB.Form Merge 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   1035
   ClientTop       =   2040
   ClientWidth     =   9285
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4590
   ScaleWidth      =   9285
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
      Left            =   2685
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   525
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Merge.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   24
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
            TabIndex        =   25
            Top             =   405
            Visible         =   0   'False
            Width           =   300
         End
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
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
         TabIndex        =   26
         Top             =   60
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcDnMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   4260
      ScaleHeight     =   90
      ScaleWidth      =   90
      TabIndex        =   27
      Top             =   2565
      Width           =   90
   End
   Begin VB.CommandButton cmcDropDown 
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
      Left            =   3885
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   300
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
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
      HelpContextID   =   9
      Left            =   2700
      MaxLength       =   10
      TabIndex        =   2
      Top             =   300
      Width           =   1170
   End
   Begin VB.PictureBox pbcUpMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   5295
      ScaleHeight     =   90
      ScaleWidth      =   105
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2565
      Width           =   105
   End
   Begin VB.PictureBox plcMerge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   195
      ScaleHeight     =   1155
      ScaleWidth      =   8760
      TabIndex        =   16
      Top             =   2880
      Width           =   8820
      Begin VB.ListBox lbcMerge 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   8685
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5385
      Width           =   75
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4995
      TabIndex        =   19
      Top             =   4245
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Merge"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   18
      Top             =   4245
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      ScaleHeight     =   240
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   15
      Width           =   3030
   End
   Begin VB.CommandButton cmcUpMove 
      Appearance      =   0  'Flat
      Caption         =   "Mo&ve   "
      Enabled         =   0   'False
      Height          =   285
      Left            =   4620
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2475
      Width           =   945
   End
   Begin VB.PictureBox plcFrom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   210
      ScaleHeight     =   1155
      ScaleWidth      =   4140
      TabIndex        =   8
      Top             =   1065
      Width           =   4200
      Begin VB.ListBox lbcFrom 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   30
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   30
         Width           =   4065
      End
   End
   Begin VB.PictureBox plcTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4740
      ScaleHeight     =   1155
      ScaleWidth      =   4230
      TabIndex        =   11
      Top             =   1065
      Width           =   4290
      Begin VB.ListBox lbcTo 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   30
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   45
         Width           =   4155
      End
   End
   Begin VB.PictureBox plcVersions 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   225
      ScaleHeight     =   240
      ScaleWidth      =   3885
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   570
      Width           =   3885
      Begin VB.OptionButton rbcVersions 
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   3255
         TabIndex        =   6
         Top             =   0
         Width           =   555
      End
      Begin VB.OptionButton rbcVersions 
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   2580
         TabIndex        =   5
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.CommandButton cmcDnMove 
      Appearance      =   0  'Flat
      Caption         =   "M&ove   "
      Enabled         =   0   'False
      Height          =   285
      Left            =   3615
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2475
      Width           =   945
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Effective Merge Start Date"
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
      Left            =   210
      TabIndex        =   1
      Top             =   300
      Width           =   2370
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      Caption         =   "Replacement Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   6015
      TabIndex        =   10
      Top             =   855
      Width           =   1815
   End
   Begin VB.Label lacFrom 
      Appearance      =   0  'Flat
      Caption         =   "Name to Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1515
      TabIndex        =   7
      Top             =   855
      Width           =   1560
   End
End
Attribute VB_Name = "Merge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Merge.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Merge.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Merge screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBSMode As Integer     'Backspace flag
Dim imProcess As Integer
Dim imBypassFocus As Integer
Dim lmMergeStartDate As Long
Dim smSyncDate As String
Dim smSyncTime As String
Dim hmMsg As Integer   'From file hanle
Dim lmNowDate As Long   'Todays date

Dim tmVehicle() As SORTCODE
Dim smVehicleTag As String

Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim tmAdfSrchKey As INTKEY0    'ADF key record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer 'Agency file handle
Dim tmAgf As AGF        'AGF record image
Dim tmAgfSrchKey As INTKEY0    'AGF key record image
Dim imAgfRecLen As Integer        'AGF record length
Dim hmSlf As Integer 'Salesperson file handle
Dim tmSlf As SLF        'SLF record image
Dim tmSlfSrchKey As INTKEY0    'SLF key record image
Dim imSlfRecLen As Integer        'SLF record length
Dim tmBof As BOF        'CHF record image
Dim tmBvf As BVF        'BVF record image
Dim tmCdf As CDF        'CDF record image
Dim tmChf As CHF        'CHF record image
Dim tmChfSrchKey1 As CHFKEY1
Dim tmChf1 As CHF
Dim tmCntrInfo() As CNTRINFO
Dim tmCif As CIF        'CIF record image
Dim tmCuf As CUF        'CUF record image
Dim hmCrf As Integer 'Copy Rotation file handle
Dim tmCrf As CRF        'CRF record image
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim imCrfRecLen As Integer        'CRF record length
'Dim hmCsf As Integer 'Copy Script file handle
Dim tmCsf As CSF        'CSF record image
'Dim tmCsfSrchKey As LONGKEY0    'CSF key record image
'Dim imCsfRecLen As Integer        'CSF record length
'Dim hmCyf As Integer 'Copy Feed file handle
Dim tmCyf As CYF        'CYF record image
'Dim imCyfRecLen As Integer        'CYF record length
'Dim tmCtf As CTF
'Dim hmOdf As Integer 'One day Log file handle
'Dim tmOdf As ODF        'ODF record image
'Dim tmOdfSrchKey As ODFKEY0    'ODF key record image
'Dim imOdfRecLen As Integer        'ODF record length
'Same structure as RVF
'Dim hmClf As Integer 'Contract Line file handle
Dim tmClf As CLF        'CLF record image
'Dim imClfRecLen As Integer        'CLF record length
'Dim hmDrf As Integer 'Research file handle
Dim tmDrf As DRF        'DRF record image
'Dim imDrfRecLen As Integer        'DRF record length
'Dim hmPhf As Integer 'Receivable History file handle
'Dim tmPhf As RVF        'RVF record image
'Dim tmPhfSrchKey As RVFKEY1    'RVF key record image
'Dim imPhfRecLen As Integer        'RVF record length
Dim tmPjf As PJF
Dim tmPnf As PNF
Dim hmPrf As Integer 'Product Name file handle
Dim tmPrf As PRF        'PRF record image
Dim tmPrfSrchKey0 As LONGKEY0    'PRF key record image
Dim tmPrfSrchKey1 As PRFKEY1    'PRF key record image
Dim imPrfRecLen As Integer        'PRF record length
Dim hmAxf As Integer 'Product file handle
Dim tmAxf As AXF
Dim imAxfRecLen As Integer
Dim tmAxfSrchKey0 As LONGKEY0
Dim tmAxfSrchKey1 As INTKEY0

Dim tmRcf As RCF
'Dim hmRif As Integer 'Rate Card Item file handle
Dim tmRif As RIF        'RIF record image
'Dim imRifRecLen As Integer        'RIF record length
'Dim hmRvf As Integer 'Receivable file handle
Dim tmRvf As RVF        'RVF record image
'Dim tmRvfSrchKey As RVFKEY1    'RVF key record image
'Dim imRvfRecLen As Integer        'RVF record length
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
'Dim hmSbf As Integer 'Special Bill Item file handle
Dim tmSbf As SBF        'SBF record image
'Dim imSbfRecLen As Integer        'SBF record length
'Dim hmScf As Integer 'Sales Commission Item file handle
Dim tmScf As SCF        'SCF record image
'Dim imScfRecLen As Integer        'SCF record length
Dim hmSdf As Integer 'Spot detail file handle
Dim tmSdf As SDF        'SDF record image
Dim tmSdfSrchKey3 As LONGKEY0    'SDF key record image
Dim imSdfRecLen As Integer        'SDF record length
'Dim hmPsf As Integer 'Spot detail file handle
'Dim tmPsf As SDF        'SDF record image
'Dim tmPsfSrchKey As SDFKEY2    'SDF key record image
'Dim imPsfRecLen As Integer        'SDF record length
'Dim hmShf As Integer 'Spot History file handle
'Dim tmShf As SHF        'SHF record image
'Dim tmShfSrchKey As SHFKEY0    'SHF key record image
'Dim imShfRecLen As Integer        'SHF record length
Dim tmSif As SIF
Dim hmSsf As Integer 'Spot summary file handle
Dim tmSsf As SSF        'SSF record image
Dim imSsfRecLen As Integer        'SSF record length
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmRec As LPOPREC
Dim lmRecPos() As Long
Dim hmVef As Integer 'Vehicle file handle
Dim tmVef As VEF        'ADF record image
Dim tmVefSrchKey As INTKEY0    'ADF key record image
Dim imVefRecLen As Integer        'ADF record length
Dim hmVpf As Integer 'Advertiser file handle
Dim tmVpf As VPF        'ADF record image
Dim tmVpfSrchKey As VPFKEY0    'ADF key record image
Dim imVpfRecLen As Integer        'ADF record length
Dim hmVof As Integer 'Advertiser file handle
Dim tmVof As VOF        'ADF record image
Dim tmVofSrchKey As VOFKEY0    'ADF key record image
Dim imVofRecLen As Integer        'ADF record length
Dim hmVsf As Integer 'Contract vehicles file handle
Dim tmVsf As VSF        'VSF record image

Dim tmFsf As FSF
Dim tmLst As LST
Dim tmRbt As RBT
Dim tmRaf As RAF
Dim tmBsf As BSF

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Dim tmMergeInfo() As MERGEINFO


'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeAdvtProd                  *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Combine agencies               *
'*                                                     *
'*******************************************************
Private Sub mMergeAdvtProd()
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llOldCode As Long
    Dim llNewCode As Long
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slName As String
    Dim tlPrf As PRF
    Screen.MousePointer = vbHourglass
    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Advertiser Product File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)
    ReDim tmMergeInfo(0 To 0) As MERGEINFO
    For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
        slStr = lbcMerge.List(ilLoop)
        Print #hmMsg, "  " & slStr
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        llOldCode = -1
        llNewCode = -1
        For ilTest = LBound(tgProdCode) To UBound(tgProdCode) - 1 Step 1
            slNameCode = tgProdCode(ilTest).sKey
            If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Advertiser Product code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                llOldCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        For ilTest = LBound(tgProdCode) To UBound(tgProdCode) - 1 Step 1
            slNameCode = tgProdCode(ilTest).sKey
            If InStr(1, slTo, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slTo, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Advertiser Product code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                llNewCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        tmMergeInfo(UBound(tmMergeInfo)).iType = 1
        tmMergeInfo(UBound(tmMergeInfo)).iOldCode = 0
        tmMergeInfo(UBound(tmMergeInfo)).iNewCode = 0
        tmMergeInfo(UBound(tmMergeInfo)).lOldCode = llOldCode
        tmMergeInfo(UBound(tmMergeInfo)).lNewCode = llNewCode
        ReDim Preserve tmMergeInfo(0 To UBound(tmMergeInfo) + 1) As MERGEINFO
    Next ilLoop

        'If (llOldCode > 0) And (llNewCode > 0) Then
        If UBound(tmMergeInfo) > LBound(tmMergeInfo) Then
            'Phf
            ilOffSet = gFieldOffset("Phf", "PhfPrfCode")
            If Not mConvert("PHF.BTR", Len(tmRvf), ilOffSet, 4) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivables History File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Pnf
            ilOffSet = gFieldOffset("Pjf", "PjfPrfCode")
            If Not mConvert("PJF.BTR", Len(tmPjf), ilOffSet, 4) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Projection File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Rvf
            ilOffSet = gFieldOffset("Rvf", "RvfPrfCode")
            If Not mConvert("RVF.BTR", Len(tmRvf), ilOffSet, 4) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivables File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Fsf
            ilOffSet = gFieldOffset("Fsf", "FsfPrfCode")
            If Not mConvert("FSF.BTR", Len(tmFsf), ilOffSet, 4) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Feed Spot File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            For ilLoop = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                llOldCode = tmMergeInfo(ilLoop).lOldCode
                llNewCode = tmMergeInfo(ilLoop).lNewCode
                If (tgSpf.sRemoteUsers <> "Y") And (lmMergeStartDate = 0) And (rbcVersions(0).Value) Then
                    'Move A/R amount
                    Do
                        tmPrfSrchKey0.lCode = llOldCode
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Advertiser Product File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        tmPrfSrchKey0.lCode = llNewCode
                        ilRet = btrGetEqual(hmPrf, tlPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Advertiser Product File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        ilRet = btrUpdate(hmPrf, tlPrf, imPrfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    'Remove Prf
                    tmPrfSrchKey0.lCode = llOldCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when reading Advertiser Product File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                        Exit Sub
                    End If
                    ilRet = btrDelete(hmPrf)
                Else
                    'Set state to Dormant in Prf
                    Do
                        tmPrfSrchKey0.lCode = llOldCode
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Advertiser Product File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        tmPrf.sState = "D"
                        tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                        gPackDate smSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                        gPackTime smSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                        ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
            Next ilLoop
        End If
    'Next ilLoop
    Screen.MousePointer = vbDefault
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVsfVefCode                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain matching records and    *
'*                      replace values                 *
'*                                                     *
'*******************************************************
Private Function mVsfVefCode(ilOldValue As Integer, ilNewValue As Integer) As Integer

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlVsfSrchKey As LONGKEY0
    Dim ilRecLen As Integer     'Vsf record length
    Dim llLkVsfCode As Long
    Dim ilUpdate As Integer


    ilUpdate = False
    ilRecLen = Len(tmVsf)  'btrRecordLength(hlVpf)  'Get and save record length
    If tmChf.lVefCode < 0 Then
        tlVsfSrchKey.lCode = -tmChf.lVefCode
        ilRet = btrGetEqual(hmVsf, tmVsf, ilRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) = ilOldValue Then
                    tmVsf.iFSCode(ilLoop) = ilNewValue
                    ilUpdate = True
                End If
            Next ilLoop
            If ilUpdate Then
                ilRet = btrUpdate(hmVsf, tmVsf, ilRecLen)
                If ilRet <> BTRV_ERR_NONE Then
                    mVsfVefCode = False
                    Exit Function
                End If
            End If
            llLkVsfCode = tmVsf.lLkVsfCode
            Do While llLkVsfCode > 0
                tlVsfSrchKey.lCode = llLkVsfCode
                ilRet = btrGetEqual(hmVsf, tmVsf, ilRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilLoop) = ilOldValue Then
                        tmVsf.iFSCode(ilLoop) = ilNewValue
                    End If
                Next ilLoop
                ilRet = btrUpdate(hmVsf, tmVsf, ilRecLen)
                If ilRet <> BTRV_ERR_NONE Then
                    mVsfVefCode = False
                    Exit Function
                End If
                llLkVsfCode = tmVsf.lLkVsfCode
            Loop
        End If
    End If
    mVsfVefCode = True
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeVehicle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Combine Vehicles               *
'*                                                     *
'*******************************************************
Private Sub mMergeVehicle()
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilOldCode As Integer
    Dim ilNewCode As Integer
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slName As String
    Screen.MousePointer = vbHourglass
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Vehicle File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Vehicle Preferences File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imVpfRecLen = Len(tmVpf)

    hmVof = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVof, "", sgDBPath & "Vof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        ilRet = btrClose(hmVpf)
        btrDestroy hmVpf
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Vehicle Log File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imVofRecLen = Len(tmVof)

    ReDim tmMergeInfo(0 To 0) As MERGEINFO

    For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
        slStr = lbcMerge.List(ilLoop)
        Print #hmMsg, "  " & slStr
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        ilOldCode = -1
        ilNewCode = -1
        For ilTest = LBound(tmVehicle) To UBound(tmVehicle) - 1 Step 1
            slNameCode = tmVehicle(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilRet = gParseItem(slName, 3, "|", slName)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                ilRet = gParseItem(slName, 3, "|", slName)
            End If
            If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    mCloseMergeVehicle "Error when getting Salesperson code"
                    Exit Sub
                End If
                ilOldCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        For ilTest = LBound(tmVehicle) To UBound(tmVehicle) - 1 Step 1
            slNameCode = tmVehicle(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slTo, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilRet = gParseItem(slName, 3, "|", slName)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                ilRet = gParseItem(slName, 3, "|", slName)
            End If
            If StrComp(slTo, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    mCloseMergeVehicle "Error when getting Salesperson code"
                    Exit Sub
                End If
                If ilOldCode <> Val(slCode) Then
                    ilNewCode = Val(slCode)
                    Exit For
                End If
            End If
        Next ilTest
        tmMergeInfo(UBound(tmMergeInfo)).iType = 0
        tmMergeInfo(UBound(tmMergeInfo)).iOldCode = ilOldCode
        tmMergeInfo(UBound(tmMergeInfo)).iNewCode = ilNewCode
        tmMergeInfo(UBound(tmMergeInfo)).lOldCode = 0
        tmMergeInfo(UBound(tmMergeInfo)).lNewCode = 0
        ReDim Preserve tmMergeInfo(0 To UBound(tmMergeInfo) + 1) As MERGEINFO
    Next ilLoop

        'If (ilOldCode > 0) And (ilNewCode > 0) Then
        If UBound(tmMergeInfo) > LBound(tmMergeInfo) Then
            'Bvf
            ilOffSet = gFieldOffset("Bvf", "BvfVefCode")
            If Not mConvert("BVF.BTR", Len(tmBvf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Budget by Vehicle File"
                Exit Sub
            End If
            'Chf- must be done prior to rvf; phf
            ilOffSet = gFieldOffset("Chf", "ChfVefCode")
            If Not mConvert("CHF.BTR", Len(tmChf), ilOffSet, 4) Then
                mCloseMergeVehicle "Error when converting Contract Header File"
                Exit Sub
            End If
            'Clf
            ilOffSet = gFieldOffset("Clf", "ClfVefCode")
            If Not mConvert("CLF.BTR", Len(tmClf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Contract Line File"
                Exit Sub
            End If
'            'VSF related to CHF
'            If Not mVsfVefConvert(ilOldCode, ilNewCode) Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Error when converting Contract File", vbOkOnly + vbCritical, "Error"
'                Exit Sub
'            End If
            'Crf
            ilOffSet = gFieldOffset("Crf", "CrfVefCode")
            If Not mConvert("CRF.BTR", Len(tmCrf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Rotation File"
                Exit Sub
            End If
            'Cyf
            ilOffSet = gFieldOffset("Cyf", "CyfVefCode")
            If Not mConvert("CYF.BTR", Len(tmCyf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Copy Feed File"
                Exit Sub
            End If
            'Drf
            ilOffSet = gFieldOffset("Drf", "DrfVefCode")
            If Not mConvert("DRF.BTR", Len(tmDrf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Research File"
                Exit Sub
            End If
            'Phf
            ilOffSet = gFieldOffset("Phf", "PhfAirVefCode")
            If Not mConvert("PHF.BTR", Len(tmRvf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Revenue History File"
                Exit Sub
            End If
            'Phf
            ilOffSet = gFieldOffset("Phf", "PhfBillVefCode")
            If Not mConvert("PHF.BTR", Len(tmRvf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Revenue History File"
                Exit Sub
            End If
            'Pjf
            ilOffSet = gFieldOffset("Pjf", "PjfVefCode")
            If Not mConvert("PJF.BTR", Len(tmPjf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Projection File"
                Exit Sub
            End If
            'Psf
            ilOffSet = gFieldOffset("Psf", "PsfVefCode")
            If Not mConvert("PSF.BTR", Len(tmSdf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Package Spots File"
                Exit Sub
            End If
            'Rcf
            ilOffSet = gFieldOffset("Rcf", "RcfVefCode")
            If Not mConvert("RCF.BTR", Len(tmRcf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Rate Card File"
                Exit Sub
            End If
            'Rif
            ilOffSet = gFieldOffset("Rif", "RifVefCode")
            If Not mConvert("RIF.BTR", Len(tmRif), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Rate Card Item File"
                Exit Sub
            End If
            'Rvf
            ilOffSet = gFieldOffset("Rvf", "RvfAirVefCode")
            If Not mConvert("RVF.BTR", Len(tmRvf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Receivable File"
                Exit Sub
            End If
            'Rvf
            ilOffSet = gFieldOffset("Rvf", "RvfBillVefCode")
            If Not mConvert("RVF.BTR", Len(tmRvf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Receivable File"
                Exit Sub
            End If
            'Sbf
            ilOffSet = gFieldOffset("Sbf", "SbfAirVefCode")
            If Not mConvert("SBF.BTR", Len(tmSbf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Special Bill File"
                Exit Sub
            End If
            'Sbf
            ilOffSet = gFieldOffset("Sbf", "SbfBillVefCode")
            If Not mConvert("SBF.BTR", Len(tmSbf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Special Bill File"
                Exit Sub
            End If
            'Scf
            ilOffSet = gFieldOffset("Scf", "ScfVefCode")
            If Not mConvert("SCF.BTR", Len(tmScf), ilOffSet, 2) Then
                mCloseMergeVehicle "Error when converting Sales Commission File"
                Exit Sub
            End If

            For ilLoop = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                ilOldCode = tmMergeInfo(ilLoop).iOldCode
                ilNewCode = tmMergeInfo(ilLoop).iNewCode
                If (tgSpf.sRemoteUsers <> "Y") And (lmMergeStartDate = 0) And (rbcVersions(0).Value) Then
                    Do
                        tmVefSrchKey.iCode = ilOldCode
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            mCloseMergeVehicle "Error when reading Vehicle File" & ", Error #" & str$(ilRet)
                            Exit Sub
                        End If
                        ilRet = btrDelete(hmVef)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        tmVpfSrchKey.iVefKCode = ilOldCode
                        ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            mCloseMergeVehicle "Error when reading Vehicle Option File" & ", Error #" & str$(ilRet)
                            Exit Sub
                        End If
                        ilRet = btrDelete(hmVpf)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        tmVofSrchKey.iVefCode = ilOldCode
                        tmVofSrchKey.sType = "L"
                        ilRet = btrGetEqual(hmVof, tmVof, imVofRecLen, tmVofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
    '                        Screen.MousePointer = vbDefault
    '                        MsgBox "Error when reading Vehicle Log Option File", vbOkOnly + vbCritical, "Error"
                            Exit Do
                        End If
                        ilRet = btrDelete(hmVof)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        tmVofSrchKey.iVefCode = ilOldCode
                        tmVofSrchKey.sType = "L"
                        ilRet = btrGetEqual(hmVof, tmVof, imVofRecLen, tmVofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
    '                        Screen.MousePointer = vbDefault
    '                        MsgBox "Error when reading Vehicle Log Option File", vbOkOnly + vbCritical, "Error"
                            Exit Do
                        End If
                        ilRet = btrDelete(hmVof)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    Do
                        tmVofSrchKey.iVefCode = ilOldCode
                        tmVofSrchKey.sType = "O"
                        ilRet = btrGetEqual(hmVof, tmVof, imVofRecLen, tmVofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
    '                        Screen.MousePointer = vbDefault
    '                        MsgBox "Error when reading Vehicle Log Option File", vbOkOnly + vbCritical, "Error"
                            Exit Do
                        End If
                        ilRet = btrDelete(hmVof)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Else
                    'Set state to Dormant in Vef
                    Do
                        tmVefSrchKey.iCode = ilOldCode
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            mCloseMergeVehicle "Error when reading Salesperson File" & ", Error #" & str$(ilRet)
                            Exit Sub
                        End If
                        tmVef.sState = "D"
                        'gPackDate smSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
                        'gPackTime smSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
                        ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
            Next ilLoop
        End If
    'Next ilLoop
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVof)
    btrDestroy hmVof
    
    '11/26/17
    gFileChgdUpdate "vef.btr", False
    gFileChgdUpdate "vpf.btr", False
    
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If imProcess Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Erase lmRecPos
    Erase tmCntrInfo
    Erase tmVehicle
    Unload Merge
    Set Merge = Nothing   'Remove data segment
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcDnMove_Click()
    Dim ilFrom As Integer
    Dim ilTo As Integer
    Dim slFrom As String
    Dim slTo As String
    Dim ilLoop As Integer

    If (lbcFrom.ListIndex >= 0) And (lbcTo.ListIndex >= 0) Then
        If StrComp(lbcFrom.List(lbcFrom.ListIndex), lbcTo.List(lbcTo.ListIndex), 1) <> 0 Then
            ilFrom = lbcFrom.ListIndex
            ilTo = lbcTo.ListIndex
            slFrom = lbcFrom.List(lbcFrom.ListIndex)
            slTo = lbcTo.List(lbcTo.ListIndex)
            lbcMerge.AddItem "Replace: " & lbcFrom.List(lbcFrom.ListIndex) & " with: " & lbcTo.List(lbcTo.ListIndex)
            lbcFrom.RemoveItem lbcFrom.ListIndex
'            lbcTo.RemoveItem lbcTo.ListIndex
'            If ilFrom < ilTo Then
'                lbcTo.RemoveItem ilFrom
'                lbcFrom.RemoveItem ilTo - 1
'            Else
'                lbcTo.RemoveItem ilFrom - 1
'                lbcFrom.RemoveItem ilTo
'            End If
            For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
                If StrComp(lbcFrom.List(ilLoop), slTo, 1) = 0 Then
                    lbcFrom.RemoveItem ilLoop
                    Exit For
                End If
            Next ilLoop
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                If StrComp(lbcTo.List(ilLoop), slFrom, 1) = 0 Then
                    lbcTo.RemoveItem ilLoop
                    Exit For
                End If
            Next ilLoop
        Else
            Beep
        End If
    End If
    mSetCommands
End Sub
Private Sub cmcDnMove_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcDropDown_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
    Dim ilRet As Integer

    If Not mMerge() Then
        Exit Sub
    End If
    lbcMerge.Clear
    If igMergeCallSource = ADVERTISERSLIST Then
        'Traffic!lbcAdvertiser.Tag = ""
        sgCommAdfStamp = "~"
        ilRet = gObtainAdvt()
        sgAdvertiserTag = ""
        mAdvtPop
    ElseIf igMergeCallSource = AGENCIESLIST Then
        'Traffic!lbcAgency.Tag = ""
        sgCommAgfStamp = "~"
        ilRet = gObtainAgency()
        sgAgencyTag = ""
        mAgencyPop
    ElseIf igMergeCallSource = SALESPEOPLELIST Then
        'Traffic!lbcAgency.Tag = ""
        sgMSlfStamp = ""
        ilRet = gObtainSalesperson()
        sgSalespersonTag = ""
        mSPersonPop
    ElseIf igMergeCallSource = VEHICLESLIST Then
        'Traffic!lbcAgency.Tag = ""
        sgMVefStamp = ""
        ilRet = gObtainVef()
        smVehicleTag = ""
        mVehiclePop
    ElseIf igMergeCallSource = 100 Then     'Advertiser Product
        sgProdCodeTag = ""
        mAdvtProdPop
    End If
    mSetCommands
    cmcCancel.Caption = "&Done"
End Sub
Private Sub cmcUpdate_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcUpMove_Click()
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim slTstFrom As String
    Dim slTstTo As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer

    If lbcMerge.ListIndex >= 0 Then
        slStr = lbcMerge.List(lbcMerge.ListIndex)
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        lbcMerge.RemoveItem lbcMerge.ListIndex
        lbcFrom.AddItem slFrom
        'lbcFrom.AddItem slTo
        ilFound = False
        For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
            slStr = lbcMerge.List(ilLoop)
            ilPos = InStr(1, slStr, "With:", 1)
            slTstFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
            slTstTo = Trim$(Mid$(slStr, ilPos + 6))
            If StrComp(slTstTo, slTo, 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lbcFrom.AddItem slTo
        End If
        lbcTo.AddItem slFrom
        'lbcTo.AddItem slTo
    End If
    mSetCommands
End Sub
Private Sub cmcUpMove_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcStartDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
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
    Me.KeyPreview = True
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase tmMergeInfo
End Sub

Private Sub lbcFrom_Click()
    mSetCommands
End Sub
Private Sub lbcFrom_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub lbcMerge_Click()
    mSetCommands
End Sub
Private Sub lbcMerge_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub lbcTo_Click()
    mSetCommands
End Sub
Private Sub lbcTo_GotFocus()
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilDuplName As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String

    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(Merge, lbcFrom, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(Merge, lbcFrom, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", Merge
        On Error GoTo 0
        ilDuplName = False
        For ilLoop = 0 To UBound(tgAdvertiser) - 2 Step 1
            slNameCode = tgAdvertiser(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slNameCode = tgAdvertiser(ilLoop + 1).sKey  'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            If StrComp(slName, slStr, vbTextCompare) = 0 Then
                ilDuplName = True
                Exit For
            End If
        Next ilLoop
'        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
'            lbcTo.AddItem lbcFrom.List(ilLoop)
'        Next ilLoop
        If ilDuplName Then
            lbcFrom.Clear
            For ilLoop = 0 To UBound(tgAdvertiser) - 1 Step 1
                slNameCode = tgAdvertiser(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                lbcFrom.AddItem slName & "\" & slCode
            Next ilLoop
        End If
        lbcTo.Clear
        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            lbcTo.AddItem lbcFrom.List(ilLoop)
        Next ilLoop
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
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilDuplName As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String

    'ilRet = gPopAgyBox(Merge, lbcFrom, Traffic!lbcAgency)
    ilRet = gPopAgyBox(Merge, lbcFrom, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gIMoveListBox)", Merge
        On Error GoTo 0
        ilDuplName = False
        For ilLoop = 0 To UBound(tgAgency) - 2 Step 1
            slNameCode = tgAgency(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slNameCode = tgAgency(ilLoop + 1).sKey  'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            If StrComp(slName, slStr, vbTextCompare) = 0 Then
                ilDuplName = True
                Exit For
            End If
        Next ilLoop
        If ilDuplName Then
            lbcFrom.Clear
            For ilLoop = 0 To UBound(tgAgency) - 1 Step 1
                slNameCode = tgAgency(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                lbcFrom.AddItem slName & "\" & slCode
            Next ilLoop
        End If
        lbcTo.Clear
        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            lbcTo.AddItem lbcFrom.List(ilLoop)
        Next ilLoop
    End If
    Exit Sub
mAgencyPopErr:
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
    slStr = edcStartDate.Text
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
'*      Procedure Name:mChfSlfConvert                  *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain matching records and    *
'*                      replace values                 *
'*                                                     *
'*******************************************************
Private Function mChfSlfConvert()
    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim tlExtRec As POPLCODE
    Dim llRecPos As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilRecLen As Integer
    Dim ilOffSet As Integer
    Dim ilDataLen As Integer
    Dim llCount As Long
    Dim ilLoop As Integer
    Dim llLoop As Long
    Dim llDate As Long
    Dim ilIndex0 As Integer
    Dim ilIndex1 As Integer
    Dim ilUpdate As Integer
    Dim ilFound As Integer
    Dim ilCheck As Integer
    Dim ilOldValue As Integer
    Dim ilNewValue As Integer
    ReDim lmRecPos(0 To 32000, 0 To 0) As Long
    ilIndex0 = 0
    ilIndex1 = 0
    llCount = 0
    hlFile = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlFile, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mChfSlfConvert = False
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    ilRecLen = Len(tmChf)
    ilDataLen = 2
    ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
    ilExtLen = ilDataLen  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hlFile, tmChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        mChfSlfConvert = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            mChfSlfConvert = False
            Exit Function
        End If
    End If
    'Should be I not L because ilDataLen = 2 used
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "POPICODEPK", POPICODEPK)  'Set extract limits (all records)
    If UBound(tmMergeInfo) <= LBound(tmMergeInfo) + 1 Then
        ilOldValue = tmMergeInfo(0).iOldCode
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 2, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 4, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 6, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 8, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 10, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 12, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 14, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 16, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_OR, ilOldValue, ilDataLen)
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet + 18, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, ilOldValue, ilDataLen)
    End If
    ilRet = btrExtAddField(hlFile, ilOffSet, ilDataLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        mChfSlfConvert = False
        Exit Function
    End If
    ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            mChfSlfConvert = False
            Exit Function
        End If
        ilExtLen = ilDataLen
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            For ilCheck = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                If ilDataLen = 2 Then
                    If tlExtRec.lCode = CLng(tmMergeInfo(ilCheck).iOldCode) Then
                        ilFound = True
                        Exit For
                    End If
                Else
                    If tlExtRec.lCode = tmMergeInfo(ilCheck).lOldCode Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilCheck
            If ilFound Then
                lmRecPos(ilIndex0, ilIndex1) = llRecPos
                ilIndex0 = ilIndex0 + 1
                llCount = llCount + 1
                If ilIndex0 > 32000 Then
                    ilIndex0 = 0
                    ilIndex1 = ilIndex1 + 1
                    ReDim Preserve lmRecPos(0 To 32000, 0 To ilIndex1) As Long
                End If
            End If
            ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilIndex0 = 0
    ilIndex1 = 0
    For llLoop = 0 To llCount - 1 Step 1
        Do
            llRecPos = lmRecPos(ilIndex0, ilIndex1)
            ilRet = btrGetDirect(hlFile, tmChf, ilRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If (ilRet <> BTRV_ERR_NONE) Then
                ilRet = btrClose(hlFile)
                btrDestroy hlFile
                mChfSlfConvert = False
                Exit Function
            End If
            If (tmChf.iCntRevNo <= 0) Then  'Proposal
                tmChf1 = tmChf
                ilRet = BTRV_ERR_NONE
            Else
                tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000 'tlChf.iCntRevNo
                tmChfSrchKey1.iPropVer = 32000  'tlChf.iPropVer
                ilRet = btrGetGreaterOrEqual(hlFile, tmChf1, ilRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                llRecPos = lmRecPos(ilIndex0, ilIndex1)     '1-27-03 re-establish the original record
                ilRet = btrGetDirect(hlFile, tmChf, ilRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)

            End If
            gUnpackDateLong tmChf1.iEndDate(0), tmChf1.iEndDate(1), llDate
            ilUpdate = False
            If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmChf1.lCntrNo) And ((llDate >= lmMergeStartDate) Or (lmMergeStartDate = 0)) Then
                If rbcVersions(0).Value Then
                    ilUpdate = True
                Else
                    If tmChf.lCode = tmChf1.lCode Then
                        ilUpdate = True
                    End If
                End If
            End If
            If ilUpdate Then
                'tmRec = tmChf
                'ilRet = gGetByKeyForUpdate("CHF", hlFile, tmRec)
                'tmChf = tmRec
                'If (ilRet <> BTRV_ERR_NONE) Then
                '    ilRet = btrClose(hlFile)
                '    btrDestroy hlFile
                '    mChfSlfConvert = False
                '    Exit Function
                'End If
                For ilUpdate = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                    ilOldValue = tmMergeInfo(ilUpdate).iOldCode
                    ilNewValue = tmMergeInfo(ilUpdate).iNewCode
                    For ilLoop = LBound(tmChf.iSlfCode) To UBound(tmChf.iSlfCode) Step 1
                        If tmChf.iSlfCode(ilLoop) = ilOldValue Then
                            tmChf.iSlfCode(ilLoop) = ilNewValue
                        End If
                    Next ilLoop
                Next ilUpdate
                If tmChf.lCode = tmChf1.lCode Then
                    'tmChf.iSourceID = tgUrf(0).iRemoteUserID
                    'gPackDate smSyncDate, tmChf.iSyncDate(0), tmChf.iSyncDate(1)
                    'gPackTime smSyncTime, tmChf.iSyncTime(0), tmChf.iSyncTime(1)
                    tmCntrInfo(UBound(tmCntrInfo)).lCode = tmChf.lCode
                    tmCntrInfo(UBound(tmCntrInfo)).lCntrNo = tmChf.lCntrNo
                    ReDim Preserve tmCntrInfo(0 To UBound(tmCntrInfo) + 1) As CNTRINFO
                End If
                ilRet = btrUpdate(hlFile, tmChf, ilRecLen)
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_CONFLICT) Then
                    ilRet = btrClose(hlFile)
                    btrDestroy hlFile
                    mChfSlfConvert = False
                    Exit Function
                End If
                Print #hmMsg, "    Contract:" & str$(tmChf.lCntrNo) & " R" & str$(tmChf.iCntRevNo) & " changed"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilIndex0 = ilIndex0 + 1
        If ilIndex0 > 32000 Then
            ilIndex0 = 0
            ilIndex1 = ilIndex1 + 1
        End If
    Next llLoop
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    mChfSlfConvert = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvert                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain matching records and    *
'*                      replace values                 *
'*                                                     *
'*******************************************************
Private Function mConvert(slFileName As String, ilRecLen As Integer, ilOffSet As Integer, ilDataLen As Integer)

'  the code must be changed to set each field instead of using HMemCpy
'  the file image names are: bof; cdf; chf; cif; cuf; csf; pjf; pnf;
'                            prf; sdf; psf; shf; sif; adf; phf; rvf and agf.
'  replace offset with flag to indicate which item is to be replaced:
'          Ad for advertiser
'          Ag for agency
'          Sa for salesperson
'
'  only integer (2 byte) values changed.  remove the long lloldvalue and llnewvalue.
'
'  example
'    case bof
'      if slFlag = "ad" then
'          bofadfcode = ilnewvalue
'      end If
'

    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim tlExtRec As POPLCODE
    Dim llRecPos As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilReadRecLen As Integer
    Dim llCount As Long
    Dim llLoop As Long
    Dim ilIndex0 As Integer
    Dim ilIndex1 As Integer
    Dim ilUpdate As Integer
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilCheck As Integer
    Dim ilTstOffset As Integer
    Dim ilOldValue As Integer
    Dim ilNewValue As Integer
    Dim llOldValue As Long
    Dim llNewValue As Long

    ReDim lmRecPos(0 To 32000, 0 To 0) As Long
    ilIndex0 = 0
    ilIndex1 = 0
    llCount = 0
    If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
        hmVsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mConvert = False
            ilRet = btrClose(hmVsf)
            btrDestroy hmVsf
            Exit Function
        End If
    End If
    If StrComp(slFileName, "ODF.BTR", 1) = 0 Then
        hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    Else
        hlFile = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    End If
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mConvert = False
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
            ilRet = btrClose(hmVsf)
            btrDestroy hmVsf
        End If
        Exit Function
    End If
    ilExtLen = ilDataLen  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    ilReadRecLen = ilRecLen
    'ilRet = btrGetFirst(hlFile, tmRec, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (StrComp(slFileName, "SDF.BTR", 1) = 0) Or (StrComp(slFileName, "PSF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmSdf, ilReadRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "DRF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmDrf, ilReadRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "RVF.BTR", 1) = 0) Or (StrComp(slFileName, "PHF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmRvf, ilReadRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "BVF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmBvf, ilReadRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "PJF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmPjf, ilReadRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "SBF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmSbf, ilReadRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "RIF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmRif, ilReadRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "CLF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmClf, ilReadRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ElseIf (StrComp(slFileName, "CDF.BTR", 1) = 0) Then
        ilRet = btrGetFirst(hlFile, tmCdf, ilReadRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Else
        ilRet = btrGetFirst(hlFile, tmRec, ilReadRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    End If
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        mConvert = True
        If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
            ilRet = btrClose(hmVsf)
            btrDestroy hmVsf
        End If
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
                ilRet = btrClose(hmVsf)
                btrDestroy hmVsf
            End If
            mConvert = False
            Exit Function
        End If
    End If
    'Call btrExtSetBounds(hlFile, llNoRec, -1, "UC") 'Set extract limits (all records)
    If ilDataLen = 2 Then
        Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "POPICODEPK", POPICODEPK) 'Set extract limits (all records)
        If UBound(tmMergeInfo) <= LBound(tmMergeInfo) + 1 Then
            ilOldValue = tmMergeInfo(0).iOldCode
            ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, ilOldValue, ilDataLen)
        End If
    Else
        Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "POPLCODEPK", POPLCODEPK) 'Set extract limits (all records)
        If UBound(tmMergeInfo) <= LBound(tmMergeInfo) + 1 Then
            llOldValue = tmMergeInfo(0).lOldCode
            ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, llOldValue, ilDataLen)
        End If
    End If
    ilRet = btrExtAddField(hlFile, ilOffSet, ilDataLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
            ilRet = btrClose(hmVsf)
            btrDestroy hmVsf
        End If
        mConvert = False
        Exit Function
    End If
    ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
                ilRet = btrClose(hmVsf)
                btrDestroy hmVsf
            End If
            mConvert = False
            Exit Function
        End If
        ilExtLen = ilDataLen
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            For ilCheck = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                If ilDataLen = 2 Then
                    If tlExtRec.lCode = CLng(tmMergeInfo(ilCheck).iOldCode) Then
                        ilFound = True
                        Exit For
                    End If
                Else
                    If tlExtRec.lCode = tmMergeInfo(ilCheck).lOldCode Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilCheck
            If ilFound Then
                lmRecPos(ilIndex0, ilIndex1) = llRecPos
                ilIndex0 = ilIndex0 + 1
                llCount = llCount + 1
                If ilIndex0 > 32000 Then
                    ilIndex0 = 0
                    ilIndex1 = ilIndex1 + 1
                    ReDim Preserve lmRecPos(0 To 32000, 0 To ilIndex1) As Long
                End If
            End If
            ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilIndex0 = 0
    ilIndex1 = 0
    For llLoop = 0 To llCount - 1 Step 1
        Do
            ilReadRecLen = ilRecLen
            llRecPos = lmRecPos(ilIndex0, ilIndex1)
            'ilRet = btrGetDirect(hlFile, tmRec, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If StrComp(slFileName, "ADF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmAdf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmAdf.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmAdf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "AGF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmAgf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmAgf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "BOF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmBof, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmBof.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "BSF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmBsf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmBsf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "BVF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmBvf, ilReadRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmBvf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CDF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmCdf, ilReadRecLen, llRecPos, INDEXKEY2, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmCdf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmCdf.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CHF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmChf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmChf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmChf.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = CInt(tmChf.lVefCode)
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CLF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmClf, ilReadRecLen, llRecPos, INDEXKEY2, BTRV_LOCK_NONE)
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        ilOldValue = tmClf.iVefCode
                    ElseIf igMergeCallSource = 100 Then
                    End If
            ElseIf StrComp(slFileName, "CIF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmCif, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmCif.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CUF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmCuf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmCuf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CRF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmCrf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmCrf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CSF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmCsf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmCsf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "CYF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmCyf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmCyf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "DRF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmDrf, ilReadRecLen, llRecPos, INDEXKEY2, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmDrf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "FSF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmFsf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmFsf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                ElseIf igMergeCallSource = 100 Then
                    llOldValue = tmPjf.lPrfCode
                End If
            ElseIf StrComp(slFileName, "LST.MKD", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmSif, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmLst.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmLst.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "PHF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmRvf, ilReadRecLen, llRecPos, INDEXKEY2, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmRvf.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmRvf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilTstOffset = gFieldOffset("Phf", "PhfAirVefCode")
                    If ilOffSet = ilTstOffset Then
                        ilOldValue = tmRvf.iAirVefCode
                    Else
                        ilOldValue = tmRvf.iBillVefCode
                    End If
                ElseIf igMergeCallSource = 100 Then
                    llNewValue = tmRvf.lPrfCode
                End If
            ElseIf StrComp(slFileName, "PJF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmPjf, ilReadRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmPjf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmPjf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmPjf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                    llOldValue = tmPjf.lPrfCode
                End If
            ElseIf StrComp(slFileName, "PNF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmPnf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmPnf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmPnf.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "PRF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmPrf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmPrf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "PSF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmSdf, ilReadRecLen, llRecPos, INDEXKEY3, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmSdf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmSdf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
           ElseIf StrComp(slFileName, "RAF.BTR", 1) = 0 Then
                 ilRet = btrGetDirect(hlFile, tmRaf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmRaf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "RBT.MKD", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmRbt, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmRbt.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "RCF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmRcf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmRcf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "RIF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmRif, ilReadRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmRif.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "RVF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmRvf, ilReadRecLen, llRecPos, INDEXKEY2, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                    ilOldValue = tmRvf.iAgfCode
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmRvf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilTstOffset = gFieldOffset("Phf", "PhfAirVefCode")
                    If ilOffSet = ilTstOffset Then
                        ilOldValue = tmRvf.iAirVefCode
                    Else
                        ilOldValue = tmRvf.iBillVefCode
                    End If
                ElseIf igMergeCallSource = 100 Then
                    llOldValue = tmRvf.lPrfCode
                End If
            ElseIf StrComp(slFileName, "SBF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmSbf, ilReadRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilTstOffset = gFieldOffset("Sbf", "SbfAirVefCode")
                    If ilOffSet = ilTstOffset Then
                        ilOldValue = tmSbf.iAirVefCode
                    Else
                        ilOldValue = tmSbf.iBillVefCode
                    End If
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "SCF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmScf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ilOldValue = tmScf.iSlfCode
                ElseIf igMergeCallSource = VEHICLESLIST Then
                    ilOldValue = tmScf.iVefCode
                ElseIf igMergeCallSource = 100 Then
                End If
            ElseIf StrComp(slFileName, "SDF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmSdf, ilReadRecLen, llRecPos, INDEXKEY3, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmSdf.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
           ElseIf StrComp(slFileName, "SIF.BTR", 1) = 0 Then
                ilRet = btrGetDirect(hlFile, tmSif, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If igMergeCallSource = ADVERTISERSLIST Then
                    ilOldValue = tmSif.iAdfCode
                ElseIf igMergeCallSource = AGENCIESLIST Then
                ElseIf igMergeCallSource = SALESPEOPLELIST Then
                ElseIf igMergeCallSource = VEHICLESLIST Then
                ElseIf igMergeCallSource = 100 Then
                End If
            End If
            If (ilRet <> BTRV_ERR_NONE) Then
                ilRet = btrClose(hlFile)
                btrDestroy hlFile
                If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
                    ilRet = btrClose(hmVsf)
                    btrDestroy hmVsf
                End If
                mConvert = False
                Exit Function
            End If
            ilFound = False
            For ilCheck = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                If ilDataLen = 2 Then
                    If ilOldValue = tmMergeInfo(ilCheck).iOldCode Then
                        ilFound = True
                        ilNewValue = tmMergeInfo(ilCheck).iNewCode
                        Exit For
                    End If
                Else
                    If llOldValue = tmMergeInfo(ilCheck).lOldCode Then
                        ilFound = True
                        llNewValue = tmMergeInfo(ilCheck).lNewCode
                        Exit For
                    End If
                End If
            Next ilCheck
            If ilFound Then
                ilUpdate = True
                If StrComp(slFileName, "CHF.BTR", 1) = 0 Then
                    'tmChf = tmRec
                    'If (tmChf.iCntRevNo <= 0) Then  'Proposal
                    If (tmChf.iCntRevNo <= 0) And ((tmChf.sStatus <> "H") And (tmChf.sStatus <> "O") And (tmChf.sStatus <> "G") And (tmChf.sStatus <> "N")) Then 'Proposal
                        tmChf1 = tmChf
                        ilRet = BTRV_ERR_NONE
                    Else
                        tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
                        tmChfSrchKey1.iCntRevNo = 32000 'tlChf.iCntRevNo
                        tmChfSrchKey1.iPropVer = 32000  'tlChf.iPropVer
                        ilRet = btrGetGreaterOrEqual(hlFile, tmChf1, ilRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        llRecPos = lmRecPos(ilIndex0, ilIndex1)
                        ilRet = btrGetDirect(hlFile, tmChf, ilReadRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    End If
                    gUnpackDateLong tmChf1.iEndDate(0), tmChf1.iEndDate(1), llDate
                    ilUpdate = False
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmChf1.lCntrNo) And ((llDate >= lmMergeStartDate) Or (lmMergeStartDate = 0)) Then
                        If rbcVersions(0).Value Then
                            ilUpdate = True
                        Else
                            If tmChf.lCode = tmChf1.lCode Then
                                ilUpdate = True
                            End If
                        End If
                    End If
                    'If revised contract, merge. see ttp 2962
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmChf1.lCntrNo) And (Not ilUpdate) Then
                        If (tmChf.sStatus = "W") Or (tmChf.sStatus = "C") Or (tmChf.sStatus = "I") Or (tmChf.sStatus = "G") Or (tmChf.sStatus = "O") Then
                            ilUpdate = True
                        End If
                    End If
                ElseIf (StrComp(slFileName, "BSF.BTR", 1) = 0) Then
                    gUnpackDateLong tmBsf.iStartDate(0), tmBsf.iStartDate(1), llDate
                    llDate = llDate + 52    'Obtain end date
                    If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "BVF.BTR", 1) = 0) Then
                    gUnpackDateLong tmBvf.iStartDate(0), tmBvf.iStartDate(1), llDate
                    llDate = llDate + 52    'Obtain end date
                    If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "CRF.BTR", 1) = 0) Then
    '                'tmRvf = tmRec
    '                ilFound = False
    '                For ilLoop = 0 To UBound(tmCntrInfo) - 1 Step 1
    '                    If tmRvf.lCntrNo = tmCntrInfo(ilLoop).lCntrNo Then
    '                        ilFound = True
    '                        Exit For
    '                    End If
    '                Next ilLoop
    '                If Not ilFound Then
                        gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), llDate
                        If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                            ilUpdate = False
                        End If
    '                End If
                ElseIf (StrComp(slFileName, "CYF.BTR", 1) = 0) Then
                    gUnpackDateLong tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), llDate
                    If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "RVF.BTR", 1) = 0) Or (StrComp(slFileName, "PHF.BTR", 1) = 0) Then
                    'tmRvf = tmRec
                    ilFound = False
                    For ilLoop = 0 To UBound(tmCntrInfo) - 1 Step 1
                        If tmRvf.lCntrNo = tmCntrInfo(ilLoop).lCntrNo Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
                        If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                            ilUpdate = False
                        End If
                    End If
                ElseIf (StrComp(slFileName, "SDF.BTR", 1) = 0) Or (StrComp(slFileName, "PSF.BTR", 1) = 0) Then
                    'tmSdf = tmRec
                    ilFound = False
                    For ilLoop = 0 To UBound(tmCntrInfo) - 1 Step 1
                        If tmSdf.lChfCode = tmCntrInfo(ilLoop).lCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                        If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                            ilUpdate = False
                        End If
                    End If
                ElseIf (StrComp(slFileName, "BOF.BTR", 1) = 0) Then
                    'tmBof = tmRec
                    gUnpackDateLong tmBof.iEndDate(0), tmBof.iEndDate(1), llDate
                    If (llDate <> 0) And (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "CDF.BTR", 1) = 0) Then
                    'tmCdf = tmRec
                    gUnpackDateLong tmCdf.iDateEntrd(0), tmCdf.iDateEntrd(1), llDate
                    If (llDate <> 0) And (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "PJF.BTR", 1) = 0) Then
                    'tmPjf = tmRec
                    gUnpackDateLong tmPjf.iEffDate(0), tmPjf.iEffDate(1), llDate
                    If (llDate <> 0) And (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "RIF.BTR", 1) = 0) Then
                    'tmPjf = tmRec
                    llDate = gDateValue("12/31/" & Trim$(str$(tmRif.iYear)))
                    If (llDate <> 0) And (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "RCF.BTR", 1) = 0) Then
                    'tmPjf = tmRec
                    llDate = gDateValue("12/31/" & Trim$(str$(tmRcf.iYear)))
                    If (llDate <> 0) And (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                ElseIf (StrComp(slFileName, "SBF.BTR", 1) = 0) Then
                    'tmSdf = tmRec
                    ilFound = False
                    For ilLoop = 0 To UBound(tmCntrInfo) - 1 Step 1
                        If tmSbf.lChfCode = tmCntrInfo(ilLoop).lCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
                        If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                            ilUpdate = False
                        End If
                    End If
                ElseIf (StrComp(slFileName, "SCF.BTR", 1) = 0) Then
                    'tmPjf = tmRec
                    gUnpackDateLong tmScf.iEndDate(0), tmScf.iEndDate(1), llDate
                    If (llDate <> 0) And (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                        ilUpdate = False
                    End If
                End If
            Else
                ilUpdate = False
            End If
            If ilUpdate Then
                'If StrComp(slFileName, "ODF.BTR", 1) <> 0 Then
                '    If StrComp(slFileName, "SHF.BTR", 1) <> 0 Then
                '        ilRet = gGetByKeyForUpdate(slFileName, hlFile, tmRec)
                '    Else
                '        ilRet = gGetByKeyForUpdateXF(slFileName, hlFile, tmRec)
                '    End If
                '    If (ilRet <> BTRV_ERR_NONE) Then
                '        ilRet = btrClose(hlFile)
                '        btrDestroy hlFile
                '        mConvert = False
                '        Exit Function
                '    End If
                'End If
'                If ilDataLen = 2 Then
'                    'ReplaceInt tmRec, ilOffset, ilNewValue
'                    ReDim bgByteArray(LenB(tmRec))
'                    HMemCpy bgByteArray(0), tmRec, LenB(tmRec)
'                    HMemCpy bgByteArray(0), ilNewValue, 2
'                    HMemCpy tmRec, bgByteArray(0), LenB(tmRec)
'
'                Else
'                    'ReplaceLong tmRec, ilOffset, llNewValue
'                    ReDim bgByteArray(LenB(tmRec))
'                    HMemCpy bgByteArray(0), tmRec, LenB(tmRec)
'                    HMemCpy bgByteArray(0), llNewValue, 4
'                    HMemCpy tmRec, bgByteArray(0), LenB(tmRec)
'                End If
                If StrComp(slFileName, "ADF.BTR", 1) = 0 Then
                    'tmAdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                        tmAdf.iAgfCode = ilNewValue
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmAdf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmAdf
                ElseIf StrComp(slFileName, "AGF.BTR", 1) = 0 Then
                    'tmAgf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmAgf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    '10148 removed fields
                   ' tmAgf.iSourceID = tgUrf(0).iRemoteUserID
                  '  gPackDate smSyncDate, tmAgf.iSyncDate(0), tmAgf.iSyncDate(1)
                  '  gPackTime smSyncTime, tmAgf.iSyncTime(0), tmAgf.iSyncTime(1)
                    'tmRec = tmAgf
                ElseIf StrComp(slFileName, "BOF.BTR", 1) = 0 Then
                    'tmBof = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmBof.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmBof
                ElseIf StrComp(slFileName, "BSF.BTR", 1) = 0 Then
                    'tmBof = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmBsf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "BVF.BTR", 1) = 0 Then
                    'tmBof = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmBvf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "CDF.BTR", 1) = 0 Then
                    'tmCdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmCdf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                        tmCdf.iAgfCode = ilNewValue
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmCdf
                ElseIf StrComp(slFileName, "CHF.BTR", 1) = 0 Then
                    'tmChf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmChf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                        tmChf.iAgfCode = ilNewValue
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmChf.lVefCode = llNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    If tmChf.lCode = tmChf1.lCode Then
                        'tmChf.iSourceID = tgUrf(0).iRemoteUserID
                        'gPackDate smSyncDate, tmChf.iSyncDate(0), tmChf.iSyncDate(1)
                        'gPackTime smSyncTime, tmChf.iSyncTime(0), tmChf.iSyncTime(1)
                        tmCntrInfo(UBound(tmCntrInfo)).lCode = tmChf.lCode
                        tmCntrInfo(UBound(tmCntrInfo)).lCntrNo = tmChf.lCntrNo
                        ReDim Preserve tmCntrInfo(0 To UBound(tmCntrInfo) + 1) As CNTRINFO
                    End If
                    'tmRec = tmChf
                    Print #hmMsg, "    Contract:" & str$(tmChf.lCntrNo) & " R" & str$(tmChf.iCntRevNo) & " changed"
                ElseIf StrComp(slFileName, "CIF.BTR", 1) = 0 Then
                    'tmCif = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmCif.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmCif
                ElseIf StrComp(slFileName, "CUF.BTR", 1) = 0 Then
                    'tmCuf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmCuf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "CLF.BTR", 1) = 0 Then
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmClf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "CRF.BTR", 1) = 0 Then
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmCrf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "CSF.BTR", 1) = 0 Then
                    'tmCsf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmCsf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "CYF.BTR", 1) = 0 Then
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmCyf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "DRF.BTR", 1) = 0 Then
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmDrf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "PHF.BTR", 1) = 0 Then
                    'tmPhf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                        tmRvf.iAgfCode = ilNewValue
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmRvf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        ilTstOffset = gFieldOffset("Phf", "PhfAirVefCode")
                        If ilOffSet = ilTstOffset Then
                            tmRvf.iAirVefCode = ilNewValue
                        Else
                            tmRvf.iBillVefCode = ilNewValue
                        End If
                    ElseIf igMergeCallSource = 100 Then
                        tmRvf.lPrfCode = llNewValue
                    End If
                    'tmRec = tmPhf
                ElseIf StrComp(slFileName, "PJF.BTR", 1) = 0 Then
                    'tmPjf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmPjf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmPjf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmPjf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                        tmPjf.lPrfCode = llNewValue
                    End If
                    tmPjf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate smSyncDate, tmPjf.iSyncDate(0), tmPjf.iSyncDate(1)
                    gPackTime smSyncTime, tmPjf.iSyncTime(0), tmPjf.iSyncTime(1)
                    'tmRec = tmPjf
                ElseIf StrComp(slFileName, "PNF.BTR", 1) = 0 Then
                    'tmPnf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmPnf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                        tmPnf.iAgfCode = ilNewValue
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    tmPnf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate smSyncDate, tmPnf.iSyncDate(0), tmPnf.iSyncDate(1)
                    gPackTime smSyncTime, tmPnf.iSyncTime(0), tmPnf.iSyncTime(1)
                    'tmRec = tmPnf
                ElseIf StrComp(slFileName, "PRF.BTR", 1) = 0 Then
                    'tmPrf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmPrf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate smSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                    gPackTime smSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                    'tmRec = tmPrf
                ElseIf StrComp(slFileName, "PSF.BTR", 1) = 0 Then
                    'tmPsf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmSdf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmSdf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmPsf
                ElseIf StrComp(slFileName, "RCF.BTR", 1) = 0 Then
                    'tmSdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmRcf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "RIF.BTR", 1) = 0 Then
                    'tmSdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmRif.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "RVF.BTR", 1) = 0 Then
                    'tmRvf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                        tmRvf.iAgfCode = ilNewValue
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmRvf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        ilTstOffset = gFieldOffset("Phf", "PhfAirVefCode")
                        If ilOffSet = ilTstOffset Then
                            tmRvf.iAirVefCode = ilNewValue
                        Else
                            tmRvf.iBillVefCode = ilNewValue
                        End If
                    ElseIf igMergeCallSource = 100 Then
                        tmRvf.lPrfCode = llNewValue
                    End If
                    'tmRec = tmRvf
                ElseIf StrComp(slFileName, "SBF.BTR", 1) = 0 Then
                    'tmSdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        ilTstOffset = gFieldOffset("Sbf", "SbfAirVefCode")
                        If ilOffSet = ilTstOffset Then
                            tmSbf.iAirVefCode = ilNewValue
                        Else
                            tmSbf.iBillVefCode = ilNewValue
                        End If
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "SCF.BTR", 1) = 0 Then
                    'tmSdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                        tmScf.iSlfCode = ilNewValue
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                        tmScf.iVefCode = ilNewValue
                    ElseIf igMergeCallSource = 100 Then
                    End If
                ElseIf StrComp(slFileName, "SDF.BTR", 1) = 0 Then
                    'tmSdf = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmSdf.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmSdf
'                ElseIf StrComp(slFileName, "SHF.BTR", 1) = 0 Then
'                    tmShf = tmRec
'                    If igMergeCallSource = ADVERTISERSLIST Then
'                        tmShf.iAdfCode = ilNewValue
'                    ElseIf igMergeCallSource = AGENCIESLIST Then
'                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
'                    End If
'                    tmRec = tmShf
                ElseIf StrComp(slFileName, "SIF.BTR", 1) = 0 Then
                    'tmSif = tmRec
                    If igMergeCallSource = ADVERTISERSLIST Then
                        tmSif.iAdfCode = ilNewValue
                    ElseIf igMergeCallSource = AGENCIESLIST Then
                    ElseIf igMergeCallSource = SALESPEOPLELIST Then
                    ElseIf igMergeCallSource = VEHICLESLIST Then
                    ElseIf igMergeCallSource = 100 Then
                    End If
                    'tmRec = tmSif
                End If
                'Delete record incase field is part of key
                ilRet = btrDelete(hlFile)
                'ilRet = btrUpdate(hlFile, tmRec, ilReadRecLen)
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_CONFLICT) Then
                    ilRet = btrClose(hlFile)
                    btrDestroy hlFile
                    If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
                        ilRet = btrClose(hmVsf)
                        btrDestroy hmVsf
                    End If
                    mConvert = False
                    Exit Function
                End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilUpdate Then
            'ilRet = btrInsert(hlFile, tmRec, ilReadRecLen, INDEXKEY0)
            If StrComp(slFileName, "ADF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmAdf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "AGF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmAgf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "BOF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmBof, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "BSF.BTR", 1) = 0 Then
                'Don't add, just delete old record
                'ilRet = btrInsert(hlFile, tmBsf, ilReadRecLen, INDEXKEY1)
            ElseIf StrComp(slFileName, "BVF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmBvf, ilReadRecLen, INDEXKEY1)
            ElseIf StrComp(slFileName, "CDF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmCdf, ilReadRecLen, INDEXKEY2)
            ElseIf StrComp(slFileName, "CHF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmChf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "CIF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmCif, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "CUF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmCuf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "CLF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmClf, ilReadRecLen, INDEXKEY2)
            ElseIf StrComp(slFileName, "CRF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmCrf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "CSF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmCsf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "CYF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmCyf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "DRF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmDrf, ilReadRecLen, INDEXKEY2)
            ElseIf StrComp(slFileName, "FSF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmFsf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "LST.MKD", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmLst, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "PHF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmRvf, ilReadRecLen, INDEXKEY2)
            ElseIf StrComp(slFileName, "PJF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmPjf, ilReadRecLen, INDEXKEY1)
            ElseIf StrComp(slFileName, "PNF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmPnf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "PRF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmPrf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "PSF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmSdf, ilReadRecLen, INDEXKEY3)
            ElseIf StrComp(slFileName, "RAF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmRaf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "RBT.MKD", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmRbt, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "RCF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmRcf, ilReadRecLen, INDEXKEY0)
            ElseIf StrComp(slFileName, "RIF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmRif, ilReadRecLen, INDEXKEY1)
            ElseIf StrComp(slFileName, "RVF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmRvf, ilReadRecLen, INDEXKEY2)
            ElseIf StrComp(slFileName, "SBF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmSbf, ilReadRecLen, INDEXKEY1)
            ElseIf StrComp(slFileName, "SCF.BTR", 1) = 0 Then
                If igMergeCallSource = VEHICLESLIST Then
                    ilRet = btrInsert(hlFile, tmScf, ilReadRecLen, INDEXKEY0)
                End If
            ElseIf StrComp(slFileName, "SDF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmSdf, ilReadRecLen, INDEXKEY3)
            ElseIf StrComp(slFileName, "SIF.BTR", 1) = 0 Then
                ilRet = btrInsert(hlFile, tmSif, ilReadRecLen, INDEXKEY0)
            End If
            If (ilRet <> BTRV_ERR_NONE) Then
                ilRet = btrClose(hlFile)
                btrDestroy hlFile
                If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
                    ilRet = btrClose(hmVsf)
                    btrDestroy hmVsf
                End If
                mConvert = False
                Exit Function
            End If
            If StrComp(slFileName, "CHF.BTR", 1) = 0 Then
                ilRet = mVsfVefCode(ilOldValue, ilNewValue)
            End If
        End If
        ilIndex0 = ilIndex0 + 1
        If ilIndex0 > 32000 Then
            ilIndex0 = 0
            ilIndex1 = ilIndex1 + 1
        End If
    Next llLoop
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    If (StrComp(slFileName, "CHF.BTR", 1) = 0) Then
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
    End If
    mConvert = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCrfAdvtMerge                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge Crf for advertisers      *
'*                                                     *
'*******************************************************
Private Function mCrfAdvtMerge(ilOldValue As Integer, ilNewValue As Integer) As Integer
'
'   ilRet = mCrfAdvtMerge(ilOldValue, ilNewValue)
'   Where:
'       ilOldValue(I)- Old values to be checked for
'       ilNewValue(I)- New value to be inserted if old value found
'
    Dim ilRet As Integer
    Dim ilHighestRotNo As Integer
    Dim ilDone As Integer
    Dim llRecPos As Long
    Dim ilCRet As Integer
    Dim ilMergeInfo As Integer

    hmCrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "CRF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCrfAdvtMerge = False
        ilRet = btrClose(hmCrf)
        btrDestroy hmCrf
        Exit Function
    End If
    'Find the current Hihest Rotation Number for the new advertiser
    For ilMergeInfo = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
        ilOldValue = tmMergeInfo(ilMergeInfo).iOldCode
        ilNewValue = tmMergeInfo(ilMergeInfo).iNewCode
        ilHighestRotNo = 0
        imCrfRecLen = Len(tmCrf)
        tmCrfSrchKey1.sRotType = "A"
        tmCrfSrchKey1.iEtfCode = 0
        tmCrfSrchKey1.iEnfCode = 0
        tmCrfSrchKey1.iAdfCode = ilNewValue
        tmCrfSrchKey1.lChfCode = 0
        tmCrfSrchKey1.lFsfCode = 0
        tmCrfSrchKey1.iVefCode = 0
        tmCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.iAdfCode = ilNewValue)
            If tmCrf.iRotNo > ilHighestRotNo Then
                ilHighestRotNo = tmCrf.iRotNo
            End If
            ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        ilDone = False
        Do
            ilHighestRotNo = ilHighestRotNo + 1
            'Process in reverse order (lowest Ror # to Highest Rot #)
            tmCrfSrchKey1.sRotType = "A"
            tmCrfSrchKey1.iEtfCode = 0
            tmCrfSrchKey1.iEnfCode = 0
            tmCrfSrchKey1.iAdfCode = ilOldValue
            tmCrfSrchKey1.lChfCode = 99999999
            tmCrfSrchKey1.iVefCode = 32000
            tmCrfSrchKey1.iRotNo = 0    '32000
            ilRet = btrGetLessOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tmCrf.iAdfCode = ilOldValue) Then
                ilRet = btrGetPosition(hmCrf, llRecPos)
                Do
                    'tmRec = tmCrf
                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                    ilRet = btrDelete(hmCrf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                'tmCrf.lCode not changed so all references are
                'still valid like: Cnf.lCifCode
                tmCrf.iAdfCode = ilNewValue
                tmCrf.iRotNo = ilHighestRotNo
                ilRet = btrInsert(hmCrf, tmCrf, imCrfRecLen, INDEXKEY0)
            Else
                ilDone = True
            End If
        Loop While Not ilDone
    Next ilMergeInfo
    mCrfAdvtMerge = True
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim slStr As String
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imProcess = False
    imTerminate = False
    imBypassFocus = False
    If igMergeCallSource = ADVERTISERSLIST Then
        mAdvtPop
        If imTerminate Then
            Exit Sub
        End If
    ElseIf igMergeCallSource = AGENCIESLIST Then
        mAgencyPop
        If imTerminate Then
            Exit Sub
        End If
    ElseIf igMergeCallSource = SALESPEOPLELIST Then
        mSPersonPop
        If imTerminate Then
            Exit Sub
        End If
    ElseIf igMergeCallSource = VEHICLESLIST Then
        mVehiclePop
        If imTerminate Then
            Exit Sub
        End If
    ElseIf igMergeCallSource = 100 Then
        mAdvtProdPop
        If imTerminate Then
            Exit Sub
        End If
        lacStartDate.Visible = False
        edcStartDate.Visible = False
        cmcDropDown.Visible = False
        plcVersions.Visible = False
        rbcVersions(0).Value = True
    End If
    Screen.MousePointer = vbHourglass
    mInitBox
    Merge.height = cmcCancel.Top + 5 * cmcCancel.height / 3
    gCenterStdAlone Merge
    slStr = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
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
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.height
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMerge                          *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge                          *
'*                                                     *
'*******************************************************
Private Function mMerge() As Integer
'
'   mMerge
'   Where:
'
    Dim slDate As String
    mMerge = False
    If mTestRef() Then
        Exit Function
    End If
    slDate = Trim$(edcStartDate.Text)
    If slDate <> "" Then
        If Not gValidDate(slDate) Then
            MsgBox "Invalid Start Date", vbOKOnly + vbCritical, "Error"
            edcStartDate.SetFocus
            Exit Function
        Else
            If gDateValue(slDate) > lmNowDate Then
                MsgBox "Start Date must be prior to " & Format$(lmNowDate, "m/d/yy"), vbOKOnly + vbCritical, "Error"
                edcStartDate.SetFocus
                Exit Function
            End If
            lmMergeStartDate = gDateValue(slDate)
        End If
    Else
        lmMergeStartDate = 0
    End If
    If (Not rbcVersions(0).Value) And (Not rbcVersions(1).Value) Then
        MsgBox "'Change All Contract Revisions' must be set", vbOKOnly + vbCritical, "Error"
        rbcVersions(0).SetFocus
        Exit Function
    End If
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Function
    End If
    If lmMergeStartDate <> 0 Then
        Print #hmMsg, "Merge Start Date: " & slDate
    Else
        Print #hmMsg, "Merge Start Date: " & "All Dates"
    End If
    If rbcVersions(0).Value Then
        Print #hmMsg, "All Contract Revisions"
    Else
        Print #hmMsg, "Only Current Contract Revision"
    End If
    mMerge = True
    ReDim tmCntrInfo(0 To 0) As CNTRINFO
    gGetSyncDateTime smSyncDate, smSyncTime
    mUpdate
    Print #hmMsg, "** Completed Merge: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    Close #hmMsg
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeAdvt                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Combine advertisers            *
'*                                                     *
'*******************************************************
Private Sub mMergeAdvt()
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilOldCode As Integer
    Dim ilNewCode As Integer
    Dim slOldDate As String
    Dim slNewDate As String
    Dim slOldStr As String
    Dim slNewStr As String
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slName As String
    Dim ilSdfExist As Integer
    Dim ilPsfExist As Integer
    Dim tlAdf As ADF
    Screen.MousePointer = vbHourglass
    hmAdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Advertiser File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "PRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Product File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)
    
    hmAxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAxf, "", sgDBPath & "AXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Audio Vault to X-Digital Cue File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imAxfRecLen = Len(tmAxf)

'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "DSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Error when opening Delete Stamp File", vbOkOnly + vbCritical, "Error"
'        Exit Sub
'    End If
'    imDsfRecLen = Len(tmDsf)
    ReDim tmMergeInfo(0 To 0) As MERGEINFO
    For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
        slStr = lbcMerge.List(ilLoop)
        Print #hmMsg, "  " & slStr
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        ilOldCode = -1
        ilNewCode = -1
        For ilTest = LBound(tgAdvertiser) To UBound(tgAdvertiser) - 1 Step 1
            slNameCode = tgAdvertiser(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Advertiser code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                ilOldCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        For ilTest = LBound(tgAdvertiser) To UBound(tgAdvertiser) - 1 Step 1
            slNameCode = tgAdvertiser(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slTo, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slTo, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Advertiser code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                ilNewCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        tmMergeInfo(UBound(tmMergeInfo)).iType = 0
        tmMergeInfo(UBound(tmMergeInfo)).iOldCode = ilOldCode
        tmMergeInfo(UBound(tmMergeInfo)).iNewCode = ilNewCode
        tmMergeInfo(UBound(tmMergeInfo)).lOldCode = 0
        tmMergeInfo(UBound(tmMergeInfo)).lNewCode = 0
        ReDim Preserve tmMergeInfo(0 To UBound(tmMergeInfo) + 1) As MERGEINFO
    Next ilLoop

        'If (ilOldCode > 0) And (ilNewCode > 0) Then
        If UBound(tmMergeInfo) > LBound(tmMergeInfo) Then
            'Bof
            ilOffSet = gFieldOffset("Bof", "BofAdfCode")
            If Not mConvert("BOF.BTR", Len(tmBof), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Blackout File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Cdf
            ilOffSet = gFieldOffset("Cdf", "CdfAdfCode")
            If Not mConvert("CDF.BTR", Len(tmCdf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Comment File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Chf- must be done prior to rvf; phf; sdf and ssf
            ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
            If Not mConvert("CHF.BTR", Len(tmChf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Contract File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Cif
            ilOffSet = gFieldOffset("Cif", "CifAdfCode")
            If Not mConvert("CIF.BTR", Len(tmCif), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Inventory File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Cuf
            ilOffSet = gFieldOffset("Cuf", "CufAdfCode")
            If Not mConvert("CUF.BTR", Len(tmCuf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Inventory File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Crf
            'Add rotation from old advt to new advt with highest rotation numbers
            'since these instructions can't be superseding an of the instruction
            'in the new advertiser (different contracts)
            ilOffSet = gFieldOffset("Crf", "CrfAdfCode")
            If Not mCrfAdvtMerge(ilOldCode, ilNewCode) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Rotation File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Csf
            ilOffSet = gFieldOffset("Csf", "CsfAdfCode")
            If Not mConvert("CSF.BTR", Len(tmCsf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Script File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ''Ctf
            'ilOffset = gFieldOffset("Ctf", "CtfAdfCode")
            'If Not mConvert("CTF.BTR", Len(tmCtf), ilOffset, 2, ilOldCode, ilNewCode, 0, 0) Then
            '    Screen.MousePointer = vbDefault
            '    MsgBox "Error when converting Contract Summary File", vbOkOnly + vbCritical, "Error"
            '    Exit Sub
            'End If
            'Odf
            'ilOffset = gFieldOffset("Odf", "OdfAdfCode")
            'If Not mConvert("ODF.BTR", Len(tmOdf), ilOffset, 2, ilOldCode, ilNewCode, 0, 0) Then
            '    Screen.MousePointer = vbDefault
            '    MsgBox "Error when converting Log File", vbOkOnly + vbCritical, "Error"
            '    Exit Sub
            'End If
            'Phf- reveivable history
            If Not mPhfOrRvfConvert("PHF.BTR") Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivables History File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Pjf
            ilOffSet = gFieldOffset("Pjf", "PjfAdfCode")
            If Not mConvert("PJF.BTR", Len(tmPjf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Contract Projection File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Pjf
            ilOffSet = gFieldOffset("Pnf", "PnfAdfCode")
            If Not mConvert("PNF.BTR", Len(tmPnf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Contract Projection File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Prf
            ilOffSet = gFieldOffset("Prf", "PrfAdfCode")
            If Not mConvert("PRF.BTR", Len(tmPrf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Product File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Rvf
            If Not mPhfOrRvfConvert("RVF.BTR") Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivables File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Sdf
            If UBound(tmMergeInfo) <= LBound(tmMergeInfo) + 1 Then
                ilOldCode = tmMergeInfo(0).iOldCode
                ilSdfExist = gIICodeRefExist(Merge, ilOldCode, "Sdf.Btr", "SdfAdfCode")
            Else
                ilSdfExist = True
            End If
            If ilSdfExist Then
                ilOffSet = gFieldOffset("Sdf", "SdfAdfCode")
                If Not mConvert("SDF.BTR", Len(tmSdf), ilOffSet, 2) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when converting Spot File", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
            End If
            'Psf (same image as sdf)
            If UBound(tmMergeInfo) <= LBound(tmMergeInfo) + 1 Then
                ilOldCode = tmMergeInfo(0).iOldCode
                ilPsfExist = gIICodeRefExist(Merge, ilOldCode, "Psf.Btr", "PsfAdfCode")
            Else
                ilPsfExist = True
            End If
            If ilPsfExist Then
                ilOffSet = gFieldOffset("Psf", "PsfAdfCode")
                If Not mConvert("PSF.BTR", Len(tmSdf), ilOffSet, 2) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when converting Package Spot File", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
            End If
            'Shf
'            'ilOffset = gFieldOffset("Shf", "ShfAdfCode")
'            'ilOffset = GetOffSetForInt(tmShf, tmShf.iAdfCode) '6
'            ilOffset = gFieldOffset("Shf", "ShfAdfCode")
'            If Not mConvert("SHF.BTR", Len(tmShf), ilOffset, 2, ilOldCode, ilNewCode, 0, 0) Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Error when converting Spot History File", vbOkOnly + vbCritical, "Error"
'                Exit Sub
'            End If
            'Sif
            ilOffSet = gFieldOffset("Sif", "SifAdfCode")
            If Not mConvert("SIF.BTR", Len(tmSif), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Short Title File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Ssf
            If ilSdfExist Then
                If Not mSsfConvert(0, ilOldCode, ilNewCode) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when converting Spot Summary File", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
            End If
            ilOffSet = gFieldOffset("Fsf", "FsfAdfCode")
            If Not mConvert("FSF.BTR", Len(tmFsf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Feed Spot File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ilOffSet = gFieldOffset("Raf", "RafAdfCode")
            If Not mConvert("RAF.BTR", Len(tmRaf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Region Definition File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ilOffSet = gFieldOffset("Lst", "LstAdfCode")
            If Not mConvert("LST.MKD", Len(tmLst), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Affiliate Log Spots File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ilOffSet = gFieldOffset("Rbt", "RbtAdfCode")
            If Not mConvert("RBT.MKD", Len(tmRbt), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Affiliate Region Blackout Info File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            For ilLoop = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                ilOldCode = tmMergeInfo(ilLoop).iOldCode
                ilNewCode = tmMergeInfo(ilLoop).iNewCode
                'Move A/R amount
                Do
                    tmAdfSrchKey.iCode = ilOldCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when reading Advertiser File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                        Exit Sub
                    End If
                    tmAdfSrchKey.iCode = ilNewCode
                    ilRet = btrGetEqual(hmAdf, tlAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when reading Advertiser File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                        Exit Sub
                    End If
                    gUnpackDate tmAdf.iDateLstInv(0), tmAdf.iDateLstInv(1), slOldDate
                    gUnpackDate tlAdf.iDateLstInv(0), tlAdf.iDateLstInv(1), slNewDate
                    If Trim$(slOldDate) <> "" Then
                        If gDateValue(slOldDate) > gDateValue(slNewDate) Then
                            tlAdf.iDateLstInv(0) = tmAdf.iDateLstInv(0)
                            tlAdf.iDateLstInv(1) = tmAdf.iDateLstInv(1)
                        End If
                        gPDNToStr tmAdf.sCurrAR, 2, slOldStr
                        gPDNToStr tlAdf.sCurrAR, 2, slNewStr
                        slStr = gAddStr(slOldStr, slNewStr)
                        gStrToPDN slStr, 2, 6, tlAdf.sCurrAR
                        gPDNToStr tmAdf.sTotalGross, 2, slOldStr
                        gPDNToStr tlAdf.sTotalGross, 2, slNewStr
                        slStr = gAddStr(slOldStr, slNewStr)
                        gStrToPDN slStr, 2, 6, tlAdf.sTotalGross
                        gPDNToStr tlAdf.sCurrAR, 2, slStr
                        gPDNToStr tlAdf.sHiCredit, 2, slNewStr
                        If gCompNumberStr(slStr, slNewStr) > 0 Then
                            gStrToPDN slStr, 2, 6, tmAdf.sHiCredit
                        End If
                        tlAdf.iNSFChks = tlAdf.iNSFChks + tmAdf.iNSFChks
                        gUnpackDate tmAdf.iDateLstPaym(0), tmAdf.iDateLstPaym(1), slOldDate
                        gUnpackDate tlAdf.iDateLstPaym(0), tlAdf.iDateLstPaym(1), slNewDate
                        If Trim$(slOldDate) <> "" Then
                            If gDateValue(slOldDate) > gDateValue(slNewDate) Then
                                tlAdf.iDateLstPaym(0) = tmAdf.iDateLstPaym(0)
                                tlAdf.iDateLstPaym(1) = tmAdf.iDateLstPaym(1)
                            End If
                        End If
                        tlAdf.iNoInvPd = tlAdf.iNoInvPd + tmAdf.iNoInvPd
                        ilRet = btrUpdate(hmAdf, tlAdf, imAdfRecLen)
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT

                If (tgSpf.sRemoteUsers <> "Y") And (lmMergeStartDate = 0) And (rbcVersions(0).Value) Then
                    'Remove Adf
                    tmPrfSrchKey1.iAdfCode = ilOldCode
                    ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                    Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = ilOldCode)
                        tmPrfSrchKey0.lCode = tmPrf.lCode
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        ilRet = btrDelete(hmPrf)
                        tmPrfSrchKey1.iAdfCode = ilOldCode
                        ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                    Loop
                    tmAdfSrchKey.iCode = ilOldCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when reading Advertiser File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                        Exit Sub
                    End If
                    
                    tmAxfSrchKey1.iCode = ilOldCode
                    ilRet = btrGetEqual(hmAxf, tmAxf, imAxfRecLen, tmAxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrDelete(hmAxf)
                    End If
                    
                    ilRet = btrDelete(hmAdf)
                Else
                    'Set state to Dormant in Adf
                    Do
                        tmAdfSrchKey.iCode = ilOldCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Advertiser File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        tmAdf.sState = "D"
                        ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
            Next ilLoop
        End If
    'Next ilLoop
    Screen.MousePointer = vbDefault
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    On Error Resume Next
    ilRet = btrClose(hmAxf)
    btrDestroy hmAxf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeAgency                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Combine agencies               *
'*                                                     *
'*******************************************************
Private Sub mMergeAgency()
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilOldCode As Integer
    Dim ilNewCode As Integer
    Dim slOldDate As String
    Dim slNewDate As String
    Dim slOldStr As String
    Dim slNewStr As String
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slName As String
    Dim tlagf As AGF
    Screen.MousePointer = vbHourglass
    hmAgf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Agency File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Error when opening Delete Stamp File", vbOkOnly + vbCritical, "Error"
'        Exit Sub
'    End If
'    imDsfRecLen = Len(tmDsf)
    'slNameCode = tgAgency(lbcFrom.ListIndex).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
    'ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
    'If ilRet <> CP_MSG_NONE Then
    '    Screen.MousePointer = vbDefault
    '    MsgBox "Error when getting Agency code", vbOkOnly + vbCritical, "Error"
    '    Exit Sub
    'End If
    'ilOldCode = Val(slCode)
    'slNameCode = tgAgency(lbcTo.ListIndex).sKey    'Traffic!lbcAgency.List(lbcTo.ListIndex)
    'ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
    'If ilRet <> CP_MSG_NONE Then
    '    Screen.MousePointer = vbDefault
    '    MsgBox "Error when getting Agency code", vbOkOnly + vbCritical, "Error"
    '    Exit Sub
    'End If
    'ilNewCode = Val(slCode)
    ReDim tmMergeInfo(0 To 0) As MERGEINFO
    For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
        slStr = lbcMerge.List(ilLoop)
        Print #hmMsg, "  " & slStr
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        ilOldCode = -1
        ilNewCode = -1
        For ilTest = LBound(tgAgency) To UBound(tgAgency) - 1 Step 1
            slNameCode = tgAgency(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Agency code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                ilOldCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        For ilTest = LBound(tgAgency) To UBound(tgAgency) - 1 Step 1
            slNameCode = tgAgency(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slTo, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slTo, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Agency code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                ilNewCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        tmMergeInfo(UBound(tmMergeInfo)).iType = 0
        tmMergeInfo(UBound(tmMergeInfo)).iOldCode = ilOldCode
        tmMergeInfo(UBound(tmMergeInfo)).iNewCode = ilNewCode
        tmMergeInfo(UBound(tmMergeInfo)).lOldCode = 0
        tmMergeInfo(UBound(tmMergeInfo)).lNewCode = 0
        ReDim Preserve tmMergeInfo(0 To UBound(tmMergeInfo) + 1) As MERGEINFO
    Next ilLoop

        'If (ilOldCode > 0) And (ilNewCode > 0) Then
        If UBound(tmMergeInfo) > LBound(tmMergeInfo) Then
            'Adf
            ilOffSet = gFieldOffset("Adf", "AdfAgfCode")
            If Not mConvert("ADF.BTR", Len(tmAdf), ilOffSet, 2) Then
                MsgBox "Error when converting Advertiser File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Cdf
            ilOffSet = gFieldOffset("Cdf", "CdfAgfCode")
            If Not mConvert("CDF.BTR", Len(tmCdf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Comment File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Chf- must be done prior to rvf; phf
            ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
            If Not mConvert("CHF.BTR", Len(tmChf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Contract File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ''Ctf
            'ilOffset = gFieldOffset("Ctf", "CtfAgfCode")
            'If Not mConvert("CTF.BTR", Len(tmCtf), ilOffset, 2, ilOldCode, ilNewCode, 0, 0) Then
            '    Screen.MousePointer = vbDefault
            '    MsgBox "Error when converting Contract Summary File", vbOkOnly + vbCritical, "Error"
            '    Exit Sub
            'End If
            'Phf
            ilOffSet = gFieldOffset("Phf", "PhfAgfCode")
            If Not mConvert("PHF.BTR", Len(tmRvf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivables History File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Pnf
            ilOffSet = gFieldOffset("Pnf", "PnfAgfCode")
            If Not mConvert("PNF.BTR", Len(tmPnf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Personnel File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Rvf
            ilOffSet = gFieldOffset("Rvf", "RvfAgfCode")
            If Not mConvert("RVF.BTR", Len(tmRvf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivables File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ilOffSet = gFieldOffset("Lst", "LstAgfCode")
            If Not mConvert("LST.MKD", Len(tmLst), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Affiliate Log Spots File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            For ilLoop = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                ilOldCode = tmMergeInfo(ilLoop).iOldCode
                ilNewCode = tmMergeInfo(ilLoop).iNewCode
                If (tgSpf.sRemoteUsers <> "Y") And (lmMergeStartDate = 0) And (rbcVersions(0).Value) Then
                    'Move A/R amount
                    Do
                        tmAgfSrchKey.iCode = ilOldCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Agency File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        tmAgfSrchKey.iCode = ilNewCode
                        ilRet = btrGetEqual(hmAgf, tlagf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Agency File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        gUnpackDate tmAgf.iDateLstInv(0), tmAgf.iDateLstInv(1), slOldDate
                        gUnpackDate tlagf.iDateLstInv(0), tlagf.iDateLstInv(1), slNewDate
                        If Trim$(slOldDate) <> "" Then
                            If gDateValue(slOldDate) > gDateValue(slNewDate) Then
                                tlagf.iDateLstInv(0) = tmAgf.iDateLstInv(0)
                                tlagf.iDateLstInv(1) = tmAgf.iDateLstInv(1)
                            End If
                            gPDNToStr tmAgf.sCurrAR, 2, slOldStr
                            gPDNToStr tlagf.sCurrAR, 2, slNewStr
                            slStr = gAddStr(slOldStr, slNewStr)
                            gStrToPDN slStr, 2, 6, tlagf.sCurrAR
                            gPDNToStr tmAgf.sTotalGross, 2, slOldStr
                            gPDNToStr tlagf.sTotalGross, 2, slNewStr
                            slStr = gAddStr(slOldStr, slNewStr)
                            gStrToPDN slStr, 2, 6, tlagf.sTotalGross
                            gPDNToStr tlagf.sCurrAR, 2, slStr
                            gPDNToStr tlagf.sHiCredit, 2, slNewStr
                            If gCompNumberStr(slStr, slNewStr) > 0 Then
                                gStrToPDN slStr, 2, 6, tmAgf.sHiCredit
                            End If
                            tlagf.iNSFChks = tlagf.iNSFChks + tmAgf.iNSFChks
                            gUnpackDate tmAgf.iDateLstPaym(0), tmAgf.iDateLstPaym(1), slOldDate
                            gUnpackDate tlagf.iDateLstPaym(0), tlagf.iDateLstPaym(1), slNewDate
                            If Trim$(slOldDate) <> "" Then
                                If gDateValue(slOldDate) > gDateValue(slNewDate) Then
                                    tlagf.iDateLstPaym(0) = tmAgf.iDateLstPaym(0)
                                    tlagf.iDateLstPaym(1) = tmAgf.iDateLstPaym(1)
                                End If
                            End If
                            tlagf.iNoInvPd = tlagf.iNoInvPd + tmAgf.iNoInvPd
                            ilRet = btrUpdate(hmAgf, tlagf, imAgfRecLen)
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    'Remove Agf
                    tmAgfSrchKey.iCode = ilOldCode
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when reading Agency File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                        Exit Sub
                    End If
                    ilRet = btrDelete(hmAgf)
                Else
                    'Set state to Dormant in Agf
                    Do
                        tmAgfSrchKey.iCode = ilOldCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Agency File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        tmAgf.sState = "D"
                        '10148 removed fields
                       ' tmAgf.iSourceID = tgUrf(0).iRemoteUserID
                       ' gPackDate smSyncDate, tmAgf.iSyncDate(0), tmAgf.iSyncDate(1)
                      '  gPackTime smSyncTime, tmAgf.iSyncTime(0), tmAgf.iSyncTime(1)
                        ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
            Next ilLoop
        End If
    'Next ilLoop
    Screen.MousePointer = vbDefault
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMergeSPerson                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Combine salesperson            *
'*                                                     *
'*******************************************************
Private Sub mMergeSPerson()
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilOldCode As Integer
    Dim ilNewCode As Integer
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slName As String
    Screen.MousePointer = vbHourglass
    hmSlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        MsgBox "Error when opening Salesperson File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    ReDim tmMergeInfo(0 To 0) As MERGEINFO
    For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
        slStr = lbcMerge.List(ilLoop)
        Print #hmMsg, "  " & slStr
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        ilOldCode = -1
        ilNewCode = -1
        For ilTest = LBound(tgSalesperson) To UBound(tgSalesperson) - 1 Step 1
            slNameCode = tgSalesperson(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Salesperson code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                ilOldCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        For ilTest = LBound(tgSalesperson) To UBound(tgSalesperson) - 1 Step 1
            slNameCode = tgSalesperson(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            'slNameCode = tgAdvertiser(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
            If InStr(1, slTo, "\", vbTextCompare) > 0 Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slName = Trim$(slName) & "\" & slCode
            Else
                ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
            End If
            If StrComp(slTo, Trim$(slName), 1) = 0 Then
                ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                If ilRet <> CP_MSG_NONE Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Error when getting Salesperson code", vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
                ilNewCode = Val(slCode)
                Exit For
            End If
        Next ilTest
        tmMergeInfo(UBound(tmMergeInfo)).iType = 0
        tmMergeInfo(UBound(tmMergeInfo)).iOldCode = ilOldCode
        tmMergeInfo(UBound(tmMergeInfo)).iNewCode = ilNewCode
        tmMergeInfo(UBound(tmMergeInfo)).lOldCode = 0
        tmMergeInfo(UBound(tmMergeInfo)).lNewCode = 0
        ReDim Preserve tmMergeInfo(0 To UBound(tmMergeInfo) + 1) As MERGEINFO
    Next ilLoop

        If UBound(tmMergeInfo) > LBound(tmMergeInfo) Then
            'Adf
            ilOffSet = gFieldOffset("Adf", "AdfSlfCode")
            If Not mConvert("ADF.BTR", Len(tmAdf), ilOffSet, 2) Then
                MsgBox "Error when converting Advertiser File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Agf
            ilOffSet = gFieldOffset("Agf", "AgfSlfCode")
            If Not mConvert("AGF.BTR", Len(tmAgf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Agency File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Chf- must be done prior to rvf; phf
            If Not mChfSlfConvert() Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Contract File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Pjf
            ilOffSet = gFieldOffset("Pjf", "PjfSlfCode")
            If Not mConvert("PJF.BTR", Len(tmPjf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Projection File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Rvf
            ilOffSet = gFieldOffset("Rvf", "RvfSlfCode")
            If Not mConvert("RVF.BTR", Len(tmRvf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Receivable File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            'Phf
            ilOffSet = gFieldOffset("Phf", "PhfSlfCode")
            If Not mConvert("PHF.BTR", Len(tmRvf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Payment History File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ilOffSet = gFieldOffset("Bsf", "BsfSlfCode")
            If Not mConvert("BSF.BTR", Len(tmBsf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Budget by Salesperson File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            ilOffSet = gFieldOffset("Scf", "ScfSlfCode")
            If Not mConvert("SCF.BTR", Len(tmScf), ilOffSet, 2) Then
                Screen.MousePointer = vbDefault
                MsgBox "Error when converting Sales commission File", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
            For ilLoop = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                ilOldCode = tmMergeInfo(ilLoop).iOldCode
                ilNewCode = tmMergeInfo(ilLoop).iNewCode
                'Retain salesperson as dormant as URF can be referenced in contract or other files
                'If (tgSpf.sRemoteUsers <> "Y") And (lmMergeStartDate = 0) And (rbcVersions(0).Value) Then
                '    Do
                '        tmSlfSrchKey.iCode = ilOldCode
                '        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                '        If ilRet <> BTRV_ERR_NONE Then
                '            Screen.MousePointer = vbDefault
                '            MsgBox "Error when reading Salesperson File" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical, "Error"
                '            Exit Sub
                '        End If
                '        ilRet = btrDelete(hmSlf)
                '    Loop While ilRet = BTRV_ERR_CONFLICT
                'Else
                    'Set state to Dormant in Slf
                    Do
                        tmSlfSrchKey.iCode = ilOldCode
                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            MsgBox "Error when reading Salesperson File" & ", Error #" & str$(ilRet), vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        tmSlf.sState = "D"
                        gPackDate smSyncDate, tmSlf.iSyncDate(0), tmSlf.iSyncDate(1)
                        gPackTime smSyncTime, tmSlf.iSyncTime(0), tmSlf.iSyncTime(1)
                        ilRet = btrUpdate(hmSlf, tmSlf, imSlfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                'End If
            Next ilLoop
        End If
    'Next ilLoop
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
End Sub
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
    'On Error GoTo mOpenMsgFileErr:
    If igMergeCallSource = ADVERTISERSLIST Then
        'slToFile = sgExportPath & "MergeAdf.Txt"
        slToFile = sgDBPath & "Messages\" & "MergeAdf.Txt"
    ElseIf igMergeCallSource = AGENCIESLIST Then
        'slToFile = sgExportPath & "MergeAgf.Txt"
        slToFile = sgDBPath & "Messages\" & "MergeAgf.Txt"
    ElseIf igMergeCallSource = SALESPEOPLELIST Then
        'slToFile = sgExportPath & "MergeSlf.Txt"
        slToFile = sgDBPath & "Messages\" & "MergeSlf.Txt"
    ElseIf igMergeCallSource = VEHICLESLIST Then
        'slToFile = sgExportPath & "MergeSlf.Txt"
        slToFile = sgDBPath & "Messages\" & "MergeVef.Txt"
    ElseIf igMergeCallSource = 100 Then
        'slToFile = sgExportPath & "MergeSlf.Txt"
        slToFile = sgDBPath & "Messages\" & "MergeProd.Txt"
    End If
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        '5/19/16: Always Append
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        'If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        'Else
        '    Kill slToFile
        '    On Error GoTo 0
        '    ilRet = 0
        '    On Error GoTo mOpenMsgFileErr:
        '    hmMsg = FreeFile
        '    Open slToFile For Output As hmMsg
        '    If ilRet <> 0 Then
        '        Screen.MousePointer = vbDefault
        '        MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        '        mOpenMsgFile = False
        '        Exit Function
        '    End If
        'End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    If igMergeCallSource = ADVERTISERSLIST Then
        Print #hmMsg, "** Merge Advertisers: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    ElseIf igMergeCallSource = AGENCIESLIST Then
        Print #hmMsg, "** Merge Agencies: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    ElseIf igMergeCallSource = SALESPEOPLELIST Then
        Print #hmMsg, "** Merge Salespersons: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    ElseIf igMergeCallSource = VEHICLESLIST Then
        Print #hmMsg, "** Merge Vehicles: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    ElseIf igMergeCallSource = 100 Then
        Print #hmMsg, "** Merge Advertiser/Product: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    End If
    '5/19/16: Add user name to report
    If Trim$(tgUrf(0).sRept) <> "" Then
        Print #hmMsg, "Merge performed by: " & Trim$(tgUrf(0).sRept)
    Else
        Print #hmMsg, "Merge performed by: " & Trim$(tgUrf(0).sName)
    End If
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPhfOrRvfConvert                *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Combine advertisers for PHF or *
'*                      RVF                            *
'*                                                     *
'*******************************************************
Private Function mPhfOrRvfConvert(slFileName As String) As Integer
'
'   ilRet = mPhfOrRvfConvert(slFileName, ilOldValue, ilNewValue)
'   Where:
'       slFileName(I)- File name to be converted (Rvf.btr or Phf.btr)
'       ilOldValue(I)- Old AdfCode to be replaced
'       ilNewValue(I)- New AdfCode to be used to replace oldvalue
'
'       hmPrf(I)- Opened to Prf.btr
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim tlExtRec As POPLCODE
    Dim llRecPos As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilRecLen As Integer
    Dim ilOffSet As Integer
    Dim ilFound As Integer
    Dim ilDataLen As Integer
    Dim ilIndex0 As Integer
    Dim ilIndex1 As Integer
    Dim llCount As Long
    Dim llLoop As Long
    Dim llDate As Long
    Dim ilUpdate As Integer
    Dim ilLoop As Integer
    Dim ilCheck As Integer
    Dim tlPrf As PRF
    Dim ilOldValue As Integer
    Dim ilNewValue As Integer
    ReDim lmRecPos(0 To 32000, 0 To 0) As Long
    ilIndex0 = 0
    ilIndex1 = 0
    llCount = 0
    hlFile = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mPhfOrRvfConvert = False
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    ilRecLen = Len(tmRvf)
    ilDataLen = 2
    imPrfRecLen = Len(tmPrf)
    ilOffSet = gFieldOffset("Rvf", "RvfAdfCode")
    ilExtLen = ilDataLen  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hlFile, tmRvf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        mPhfOrRvfConvert = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            mPhfOrRvfConvert = False
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "POPICODEPK", POPICODEPK) 'Set extract limits (all records)
    If UBound(tmMergeInfo) <= LBound(tmMergeInfo) + 1 Then
        ilOldValue = tmMergeInfo(0).iOldCode
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, ilDataLen, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, ilOldValue, ilDataLen)
    End If
    ilRet = btrExtAddField(hlFile, ilOffSet, ilDataLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        mPhfOrRvfConvert = False
        Exit Function
    End If
    ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            mPhfOrRvfConvert = False
            Exit Function
        End If
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            For ilCheck = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                If ilDataLen = 2 Then
                    If tlExtRec.lCode = CLng(tmMergeInfo(ilCheck).iOldCode) Then
                        ilFound = True
                        Exit For
                    End If
                Else
                    If tlExtRec.lCode = tmMergeInfo(ilCheck).lOldCode Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilCheck
            If ilFound Then
                lmRecPos(ilIndex0, ilIndex1) = llRecPos
                ilIndex0 = ilIndex0 + 1
                llCount = llCount + 1
                If ilIndex0 > 32000 Then
                    ilIndex0 = 0
                    ilIndex1 = ilIndex1 + 1
                    ReDim Preserve lmRecPos(0 To 32000, 0 To ilIndex1) As Long
                End If
            End If
            ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFile, tlExtRec, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilIndex0 = 0
    ilIndex1 = 0
    For llLoop = 0 To llCount - 1 Step 1
        Do
            llRecPos = lmRecPos(ilIndex0, ilIndex1)
            ilRet = btrGetDirect(hlFile, tmRvf, ilRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If (ilRet <> BTRV_ERR_NONE) Then
                ilRet = btrClose(hlFile)
                btrDestroy hlFile
                mPhfOrRvfConvert = False
                Exit Function
            End If
            'gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
            'If (llDate >= lmMergeStartDate) Or (lmMergeStartDate = 0) Then
            ilUpdate = True
            ilFound = False
            For ilLoop = 0 To UBound(tmCntrInfo) - 1 Step 1
                If tmRvf.lCntrNo = tmCntrInfo(ilLoop).lCntrNo Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
                If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                    ilUpdate = False
                End If
            End If
            If ilUpdate Then
                ilUpdate = False
                For ilCheck = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                    If tmRvf.iAdfCode = tmMergeInfo(ilCheck).iOldCode Then
                        ilUpdate = True
                        ilNewValue = tmMergeInfo(ilCheck).iNewCode
                        Exit For
                    End If
                Next ilCheck
            End If
            If ilUpdate Then
                'tmRec = tmRvf
                'ilRet = gGetByKeyForUpdate("RVF", hlFile, tmRec)
                'tmRvf = tmRec
                If tmRvf.lPrfCode > 0 Then
                    tmPrfSrchKey0.lCode = tmRvf.lPrfCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
                    If ilRet = BTRV_ERR_NONE Then
                        'Determine if matching record exist within newcode
                        ilFound = False
                        tmPrfSrchKey1.iAdfCode = ilNewValue
                        ilRet = btrGetGreaterOrEqual(hmPrf, tlPrf, imPrfRecLen, tmPrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                        Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = ilNewValue)
                            If StrComp(Trim$(tlPrf.sName), Trim$(tmPrf.sName), 1) = 0 Then
                                ilFound = True
                                tmRvf.lPrfCode = tlPrf.lCode
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmPrf, tlPrf, imPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If Not ilFound Then
                            Do  'Loop until record updated or added
                                tlPrf.lCode = 0
                                tlPrf.iAdfCode = ilNewValue
                                tlPrf.sName = tmPrf.sName
                                tlPrf.iMnfComp(0) = tmPrf.iMnfComp(0)
                                tlPrf.iMnfComp(1) = tmPrf.iMnfComp(1)
                                tlPrf.iMnfExcl(0) = tmPrf.iMnfExcl(0)
                                tlPrf.iMnfExcl(1) = tmPrf.iMnfExcl(1)
                                tlPrf.iPnfBuyer = 0
                                tlPrf.sCppCpm = ""
                                For ilLoop = 0 To 3
                                    tlPrf.iMnfDemo(ilLoop) = 0
                                    tlPrf.lTarget(ilLoop) = 0
                                    tlPrf.lLastCPP(ilLoop) = 0
                                    tlPrf.lLastCPM(ilLoop) = 0
                                Next ilLoop
                                tlPrf.sState = "A"
                                tlPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                            '    ilRet = btrInsert(hmPrf, tlPrf, imPrfRecLen, INDEXKEY0)
                            'Loop While ilRet = BTRV_ERR_CONFLICT
                                tlPrf.iRemoteID = tgUrf(0).iRemoteUserID
                                tlPrf.lAutoCode = 0
                                ilRet = btrInsert(hmPrf, tlPrf, imPrfRecLen, INDEXKEY0)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If (ilRet <> BTRV_ERR_NONE) Then
                                ilRet = btrClose(hlFile)
                                btrDestroy hlFile
                                mPhfOrRvfConvert = False
                                Exit Function
                            End If
                            Do
                                'tmPrfSrchKey0.lCode = tmPrf.lCode
                                'ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                'slMsg = "mSaveRec (btrGetEqual:Product)"
                                'On Error GoTo mSaveRecErr
                                'gBtrvErrorMsg ilRet, slMsg, Advt
                                'On Error GoTo 0
                                tlPrf.iRemoteID = tgUrf(0).iRemoteUserID
                                tlPrf.lAutoCode = tlPrf.lCode
                                tlPrf.iSourceID = tgUrf(0).iRemoteUserID
                                gPackDate smSyncDate, tlPrf.iSyncDate(0), tlPrf.iSyncDate(1)
                                gPackTime smSyncTime, tlPrf.iSyncTime(0), tlPrf.iSyncTime(1)
                                ilRet = btrUpdate(hmPrf, tlPrf, imPrfRecLen)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If (ilRet <> BTRV_ERR_NONE) Then
                                ilRet = btrClose(hlFile)
                                btrDestroy hlFile
                                mPhfOrRvfConvert = False
                                Exit Function
                            End If
                            tmRvf.lPrfCode = tlPrf.lCode
                        End If
                    Else
                        tmRvf.lPrfCode = 0
                    End If
                End If
                tmRvf.iAdfCode = ilNewValue
                ilRet = btrUpdate(hlFile, tmRvf, ilRecLen)
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_CONFLICT) Then
                    ilRet = btrClose(hlFile)
                    btrDestroy hlFile
                    mPhfOrRvfConvert = False
                    Exit Function
                End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilIndex0 = ilIndex0 + 1
        If ilIndex0 > 32000 Then
            ilIndex0 = 0
            ilIndex1 = ilIndex1 + 1
        End If
    Next llLoop
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    mPhfOrRvfConvert = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    'If (lbcFrom.ListIndex >= 0) And (lbcTo.ListIndex >= 0) Then
    '    cmcUpdate.Enabled = True
    'Else
    '    cmcUpdate.Enabled = False
    'End If
    If (lbcFrom.ListIndex >= 0) And (lbcTo.ListIndex >= 0) Then
        cmcDnMove.Enabled = True
    Else
        cmcDnMove.Enabled = False
    End If
    If lbcMerge.ListIndex >= 0 Then
        cmcUpMove.Enabled = True
    Else
        cmcUpMove.Enabled = False
    End If
    If lbcMerge.ListCount > 0 Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSPersonPop                     *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSPersonPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilDuplName As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String

    'ilRet = gPopSalespersonBox(SPerson, 0, True, True, cbcSelect, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(Merge, 0, True, True, lbcFrom, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", Merge
        On Error GoTo 0
        ilDuplName = False
        For ilLoop = 0 To UBound(tgSalesperson) - 2 Step 1
            slNameCode = tgSalesperson(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slNameCode = tgSalesperson(ilLoop + 1).sKey  'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            If StrComp(slName, slStr, vbTextCompare) = 0 Then
                ilDuplName = True
                Exit For
            End If
        Next ilLoop
        If ilDuplName Then
            lbcFrom.Clear
            For ilLoop = 0 To UBound(tgSalesperson) - 1 Step 1
                slNameCode = tgSalesperson(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                lbcFrom.AddItem slName & "\" & slCode
            Next ilLoop
        End If
        lbcTo.Clear
        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            lbcTo.AddItem lbcFrom.List(ilLoop)
        Next ilLoop
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSsffConvert                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert subrecords within Ssf  *
'*                                                     *
'*******************************************************
Private Function mSsfConvert(ilType As Integer, ilOldValue As Integer, ilNewValue As Integer) As Integer
'
'   ilRet = mSsfConvert(ilType, ilOldValue, ilNewValue)
'   Where:
'       ilType(I)- 0=Convert Advertiser field
'       ilOldValue(I)- Old values to be checked for
'       ilNewValue(I)- New value to be inserted if old value found
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilSpotIndex As Integer
    Dim llRecPos As Long
    Dim ilUpdate As Integer
    Dim ilOk As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim ilMergeInfo As Integer

    hmSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "SSF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mSsfConvert = False
        ilRet = btrClose(hmSsf)
        btrDestroy hmSsf
        Exit Function
    End If
    imSsfRecLen = Len(tmSsf)
    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "SDF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mSsfConvert = False
        ilRet = btrClose(hmSsf)
        btrDestroy hmSsf
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)
    'ilRet = btrGetFirst(hmSsf, tmSsf, imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    ilRet = gSSFGetFirst(hmSsf, tmSsf, imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
        ilRet = btrGetPosition(hmSsf, llRecPos)
        ilUpdate = False
        ilLoop = 1
        Do While ilLoop <= tmSsf.iCount
           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                For ilSpotIndex = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                    If ilType = 0 Then  'Test advertiser
                        For ilMergeInfo = LBound(tmMergeInfo) To UBound(tmMergeInfo) - 1 Step 1
                            ilOldValue = tmMergeInfo(ilMergeInfo).iOldCode
                            ilNewValue = tmMergeInfo(ilMergeInfo).iNewCode
                            If tmSpot.iAdfCode = ilOldValue Then
                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                ilOk = True
                                ilFound = False
                                For ilLoop = 0 To UBound(tmCntrInfo) - 1 Step 1
                                    If tmSdf.lChfCode = tmCntrInfo(ilLoop).lCode Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    If (llDate < lmMergeStartDate) And (lmMergeStartDate <> 0) Then
                                        ilOk = False
                                    End If
                                End If
                                If ilOk Then
                                    ilUpdate = True
                                    tmSpot.iAdfCode = ilNewValue
                                    LSet tmSsf.tPas(ADJSSFPASBZ + ilSpotIndex) = tmSpot
                                End If
                                Exit For
                            End If
                        Next ilMergeInfo
                    End If
                Next ilSpotIndex
                ilLoop = ilLoop + 1 + tmAvail.iNoSpotsThis
            Else
                ilLoop = ilLoop + 1
            End If
        Loop
        If ilUpdate Then
            imSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
            'ilRet = btrUpdate(hmSsf, tmSsf, imSsfRecLen)
            ilRet = gSSFUpdate(hmSsf, tmSsf, imSsfRecLen)
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_CONFLICT) Then
                mSsfConvert = False
                ilRet = btrClose(hmSsf)
                btrDestroy hmSsf
                ilRet = btrClose(hmSdf)
                btrDestroy hmSdf
                Exit Function
            End If
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet = BTRV_ERR_CONFLICT Then
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            'ilRet = btrGetDirect(hmSsf, tmSsf, imSsfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            ilRet = gSSFGetDirect(hmSsf, tmSsf, imSsfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf)
        Else
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            'ilRet = btrGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        End If
    Loop
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mSsfConvert = True
    Else
        mSsfConvert = False
    End If
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestRef                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if Advertiser or Agency   *
'*                      is referenced by a contract    *
'*                                                     *
'*            This is temporary code                   *
'*                                                     *
'*******************************************************
Private Function mTestRef() As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim ilOldCode As Integer
    Dim slName As String
    Dim slCode As String
    Dim slMsg As String
    Dim slStr As String
    Dim slFrom As String
    Dim slTo As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilTest As Integer
    'Ignore Test as of 3/24/00 jim
    mTestRef = False
    Exit Function
    Screen.MousePointer = vbHourglass
    For ilLoop = 0 To lbcMerge.ListCount - 1 Step 1
        slStr = lbcMerge.List(ilLoop)
        ilPos = InStr(1, slStr, "With:", 1)
        slFrom = Trim$(Mid$(slStr, 9, ilPos - 10))
        slTo = Trim$(Mid$(slStr, ilPos + 6))
        If igMergeCallSource = ADVERTISERSLIST Then
            For ilTest = LBound(tgAdvertiser) To UBound(tgAdvertiser) - 1 Step 1
                slNameCode = tgAdvertiser(ilTest).sKey  'Traffic!lbcAdvertiser.List(lbcFrom.ListIndex)
                If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    slName = Trim$(slName) & "\" & slCode
                Else
                    ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                End If
                If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                    If ilRet <> CP_MSG_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when getting Advertiser code", vbOKOnly + vbCritical, "Error"
                        mTestRef = True
                        Exit Function
                    End If
                    ilOldCode = Val(slCode)
                    ilRet = gIICodeRefExist(Merge, ilOldCode, "Chf.Btr", "CHFADFCODE") 'chfagfCode
                    If ilRet Then
                        Screen.MousePointer = vbDefault
                        ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                        slMsg = slName & " used by Contracts. " & " Ok to Proceed"
                        ilRet = MsgBox(slMsg, vbYesNo + vbQuestion, "Merge")
                        If ilRet = vbNo Then
                            mTestRef = True
                            Exit Function
                        Else
                            mTestRef = False
                            Exit Function
                        End If
                    End If
                    Exit For
                End If
            Next ilTest
        ElseIf igMergeCallSource = AGENCIESLIST Then
            For ilTest = LBound(tgAgency) To UBound(tgAgency) - 1 Step 1
                slNameCode = tgAgency(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
                If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    slName = Trim$(slName) & "\" & slCode
                Else
                    ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                End If
                If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                    If ilRet <> CP_MSG_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when getting Agency code", vbOKOnly + vbCritical, "Error"
                        mTestRef = True
                        Exit Function
                    End If
                    ilOldCode = Val(slCode)
                    ilRet = gIICodeRefExist(Merge, ilOldCode, "Chf.Btr", "CHFAGFCODE") 'chfagfCode
                    If ilRet Then
                        Screen.MousePointer = vbDefault
                        ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                        slMsg = slName & " used by Contracts. " & " Ok to Proceed"
                        ilRet = MsgBox(slMsg, vbYesNo + vbQuestion, "Merge")
                        If ilRet = vbNo Then
                            mTestRef = True
                            Exit Function
                        Else
                            mTestRef = False
                            Exit Function
                        End If
                    End If
                    Exit For
                End If
            Next ilTest
        ElseIf igMergeCallSource = SALESPEOPLELIST Then
            For ilTest = LBound(tgSalesperson) To UBound(tgSalesperson) - 1 Step 1
                slNameCode = tgSalesperson(ilTest).sKey  'Traffic!lbcAgency.List(lbcFrom.ListIndex)
                If InStr(1, slFrom, "\", vbTextCompare) > 0 Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    slName = Trim$(slName) & "\" & slCode
                Else
                    ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                End If
                If StrComp(slFrom, Trim$(slName), 1) = 0 Then
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)  'Obtain Index and code number
                    If ilRet <> CP_MSG_NONE Then
                        Screen.MousePointer = vbDefault
                        MsgBox "Error when getting Salesperson code", vbOKOnly + vbCritical, "Error"
                        mTestRef = True
                        Exit Function
                    End If
                    ilOldCode = Val(slCode)
                    ilRet = gSlfCodeExistInChf(Merge, ilOldCode) 'chfagfCode
                    If ilRet Then
                        Screen.MousePointer = vbDefault
                        ilRet = gParseItem(slNameCode, 1, "\", slName)  'Obtain Index and code number
                        slMsg = slName & " used by Contracts. " & " Ok to Proceed"
                        ilRet = MsgBox(slMsg, vbYesNo + vbQuestion, "Merge")
                        If ilRet = vbNo Then
                            mTestRef = True
                            Exit Function
                        Else
                            mTestRef = False
                            Exit Function
                        End If
                    End If
                    Exit For
                End If
            Next ilTest
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    mTestRef = False
    Exit Function
End Function
Private Sub mUpdate()
    Dim slMsg As String
    Dim ilRet As Integer
    'slMsg = "Replace " & lbcFrom.List(lbcFrom.ListIndex) & " with " & lbcTo.List(lbcTo.ListIndex)
    'slMsg = "Merge names selected"

    If igMergeCallSource = 100 Then
        slMsg = "Warning!!! This process can not be undone.  The name conversion will be permanent in Receivables, Projections."
    ElseIf igMergeCallSource = VEHICLESLIST Then
        slMsg = "Warning!!! This process can not be undone.  The name conversion will be permanent in Contracts, Budgets, Projections, Rate Cards, Receivables."
    Else
        slMsg = "Warning!!! This process can not be undone.  The name conversion will be permanent in Contracts, Advertisers, Agencies, Receivables."
    End If
    ilRet = MsgBox(slMsg, vbOKCancel + vbQuestion, "Merge")
    If ilRet = vbCancel Then
        Exit Sub
    End If
    If lbcMerge.ListCount = 1 Then
        slMsg = lbcMerge.List(0)
        ilRet = MsgBox(slMsg, vbOKCancel + vbQuestion, "Merge")
        If ilRet = vbCancel Then
            Exit Sub
        End If
    End If
    imProcess = True
    If igMergeCallSource = ADVERTISERSLIST Then
        mMergeAdvt
    ElseIf igMergeCallSource = AGENCIESLIST Then
        mMergeAgency
    ElseIf igMergeCallSource = SALESPEOPLELIST Then
        mMergeSPerson
    ElseIf igMergeCallSource = VEHICLESLIST Then
        mMergeVehicle
    ElseIf igMergeCallSource = 100 Then
        mMergeAdvtProd
    End If
    imProcess = False
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
                edcStartDate.Text = Format$(llDate, "m/d/yy")
                edcStartDate.SelStart = 0
                edcStartDate.SelLength = Len(edcStartDate.Text)
                imBypassFocus = True
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcDnMove_Paint()
    pbcDnMove.CurrentX = 0
    pbcDnMove.CurrentY = 0
    pbcDnMove.Print "t"
End Sub

Private Sub pbcUpMove_Paint()
    pbcUpMove.CurrentX = 0
    pbcUpMove.CurrentY = 0
    pbcUpMove.Print "s"
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcVersions_Paint()
    plcVersions.CurrentX = 0
    plcVersions.CurrentY = 0
    plcVersions.Print "Change All Contract Versions"
End Sub

Private Sub rbcVersions_GotFocus(Index As Integer)
    plcCalendar.Visible = False
End Sub

Private Sub mVehiclePop()
    Dim ilRet As Integer 'btrieve status
    Dim llFilter As Long
    Dim ilLoop As Integer
    Dim ilRemove As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilDuplName As Integer
    Dim slStr As String

    If igVpfType = 0 Then
        'ilFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHLOGVEHICLE + VEHSIMUL + ACTIVEVEH + DORMANTVEH
        'Don't include virtual vehicles 4/2/02- Jim/Mary
        llFilter = VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH
    Else
        llFilter = VEHPACKAGE + ACTIVEVEH + DORMANTVEH
        hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        imVefRecLen = Len(tmVef)
    End If
    'ilRet = gPopUserVehicleBox(Vehicle, ilFilter, cbcSelect, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(Merge, llFilter, lbcFrom, tmVehicle(), smVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehiclePopErr
        gCPErrorMsg ilRet, "mVehiclePop (gPopUserVehicleBox)", Merge
        On Error GoTo 0
        If igVpfType <> 0 Then
            'Filter out standard packages
            lbcFrom.Clear
            For ilLoop = UBound(tmVehicle) - 1 To LBound(tmVehicle) Step -1
                slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    tmVefSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                Else
                    ilRet = Not BTRV_ERR_NONE
                End If
                If (ilRet <> BTRV_ERR_NONE) Or (tmVef.lPvfCode > 0) Then
                    'Remove item
                    For ilRemove = ilLoop To UBound(tmVehicle) - 1 Step 1
                        tmVehicle(ilRemove) = tmVehicle(ilRemove + 1)
                    Next ilRemove
                    ReDim Preserve tmVehicle(0 To UBound(tmVehicle) - 1) As SORTCODE
                End If
            Next ilLoop
            ilDuplName = False
            For ilLoop = 0 To UBound(tmVehicle) - 2 Step 1
                slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                slNameCode = tmVehicle(ilLoop + 1).sKey  'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                ilRet = gParseItem(slStr, 3, "|", slStr)
                If StrComp(slName, slStr, vbTextCompare) = 0 Then
                    ilDuplName = True
                    Exit For
                End If
            Next ilLoop

            For ilLoop = 0 To UBound(tmVehicle) - 1 Step 1
                slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '            If ilRet <> CP_MSG_NONE Then
    '                gPopUserVehicleBox = CP_MSG_PARSE
    '                Exit Function
    '            End If
                ilRet = gParseItem(slName, 3, "|", slName)
    '            If ilRet <> CP_MSG_NONE Then
    '                gPopUserVehicleBox = CP_MSG_PARSE
    '                Exit Function
    '            End If
                slName = Trim$(slName)
                If ilDuplName Then
                    slName = slName & "\" & slCode
                End If
    '            If Not gOkAddStrToListBox(slName, llLen, True) Then
    '                Exit For
    '            End If
                lbcFrom.AddItem slName  'Add ID to list box
            Next ilLoop
        End If
        lbcTo.Clear
        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            lbcTo.AddItem lbcFrom.List(ilLoop)
        Next ilLoop
    End If
    If igVpfType <> 0 Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
    End If
    Exit Sub
mVehiclePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mAdvtProdPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilDuplName As Integer
    Dim slStr As String
    Dim slName As String

    'ilRet = gPopAdvtProdBox(AdvtProd, ilCode, cbcSelect, lbcAdvtProdCode)
    ilRet = gPopAdvtProdBox(Merge, igMergeAdfCode, lbcFrom, tgProdCode(), sgProdCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtProdPopErr
        gCPErrorMsg ilRet, "mAdvtProdPop (gPopAdvtProdBox)", Merge
        On Error GoTo 0
        ilDuplName = False
        For ilLoop = 0 To UBound(tgProdCode) - 2 Step 1
            slNameCode = tgProdCode(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            slNameCode = tgProdCode(ilLoop + 1).sKey  'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            If StrComp(slName, slStr, vbTextCompare) = 0 Then
                ilDuplName = True
                Exit For
            End If
        Next ilLoop
'        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
'            lbcTo.AddItem lbcFrom.List(ilLoop)
'        Next ilLoop
        If ilDuplName Then
            lbcFrom.Clear
            For ilLoop = 0 To UBound(tgProdCode) - 1 Step 1
                slNameCode = tgProdCode(ilLoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                lbcFrom.AddItem slName & "\" & slCode
            Next ilLoop
        End If
        lbcTo.Clear
        For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            lbcTo.AddItem lbcFrom.List(ilLoop)
        Next ilLoop
    End If
    Exit Sub
mAdvtProdPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mCloseMergeVehicle(slMsg As String)
    Dim ilRet As Integer
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVof)
    btrDestroy hmVof
    Screen.MousePointer = vbDefault
    MsgBox slMsg, vbOKOnly + vbCritical, "Error"
End Sub
