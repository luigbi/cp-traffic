VERSION 5.00
Begin VB.Form RCImpact 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6075
   ClientLeft      =   570
   ClientTop       =   1470
   ClientWidth     =   9360
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6075
   ScaleWidth      =   9360
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1155
      Top             =   5535
   End
   Begin VB.PictureBox plcType 
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
      Height          =   225
      Left            =   6525
      ScaleHeight     =   225
      ScaleWidth      =   2775
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5385
      Width           =   2775
      Begin VB.OptionButton rbcType 
         Caption         =   "Week"
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
         Height          =   195
         Index           =   2
         Left            =   1785
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   870
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Month"
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
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   915
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Quarter"
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
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.PictureBox plcShow 
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2880
      ScaleHeight     =   225
      ScaleWidth      =   3120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5385
      Width           =   3120
      Begin VB.OptionButton rbcShow 
         Caption         =   "Standard"
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
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton rbcShow 
         Caption         =   "Corporate"
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
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.PictureBox pbcImpact 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   2955
      Picture         =   "Rcimpact.frx":0000
      ScaleHeight     =   4680
      ScaleWidth      =   6030
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   6030
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   15
         TabIndex        =   9
         Top             =   750
         Visible         =   0   'False
         Width           =   6000
      End
   End
   Begin VB.PictureBox plcImpact 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   2895
      ScaleHeight     =   4755
      ScaleWidth      =   6315
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   6375
      Begin VB.VScrollBar vbcImpact 
         Height          =   4515
         LargeChange     =   21
         Left            =   6090
         Max             =   1
         Min             =   1
         TabIndex        =   7
         Top             =   240
         Value           =   1
         Width           =   240
      End
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   1440
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   4215
      TabIndex        =   4
      Top             =   5730
      Width           =   945
   End
   Begin VB.PictureBox plcProposal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4680
      Left            =   150
      ScaleHeight     =   4620
      ScaleWidth      =   2535
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   2595
      Begin VB.ListBox lbcBudget 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   45
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   60
         Width           =   2475
      End
      Begin VB.ListBox lbcProposal 
         Appearance      =   0  'Flat
         Height          =   2340
         Left            =   45
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   2250
         Width           =   2475
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   210
      Top             =   5595
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RCImpact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rcimpact.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCImpact.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
Dim tmICtrls(0 To 5)  As FIELDAREA
Dim imLBICtrls As Integer
Dim tmWKCtrls(0 To 4)  As FIELDAREA
Dim imLBWKCtrls As Integer
Dim smIShow() As String
'Vehicle
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim imVefCode() As Integer
'Virtual Vehicle
Dim hmVsf As Integer        'Virtual Vehicle file handle
Dim tmVsf As VSF            'VSF record image
Dim imVsfRecLen As Integer  'VSF record length
'Contract
Dim hmCHF As Integer        'Contract file handle
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
'Line
Dim hmClf As Integer        'Line file handle
Dim tmClf As CLF            'CLF record image
Dim imClfRecLen As Integer  'CLF record length
'Flight
Dim hmCff As Integer        'Flight file handle
Dim tmCff As CFF            'CFF record image
Dim imCffRecLen As Integer  'CFF record length
'Rate Card
Dim hmRcf As Integer
Dim tmRcf() As RCF
Dim imRcfRecLen As Integer
Dim hmRif As Integer
Dim tmRif() As RIF
Dim imRifRecLen As Integer
Dim tmDPBudgetInfo() As DPBUDGETINFO
'Daypart
Dim hmRdf As Integer
Dim tmRdf() As RDF
Dim imRdfRecLen As Integer
'Library Calendar
Dim hmLcf As Integer
'Spot
Dim hmSdf As Integer    'file handle
Dim imSdfRecLen As Integer  'Record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0    'Key 3
Dim tmSdfSrchKey2 As SDFKEY2
'MG File
Dim hmSmf As Integer    'file handle
Dim imSmfRecLen As Integer  'Record length
Dim tmSmf As SMF
'Budget by Office
Dim hmBvf As Integer    'Rate Card file handle
Dim imBvfRecLen As Integer        'Rcf record length
Dim tmBvfVeh() As BVF   'Budget by vehicle
Dim tmBvfVeh2() As BVF   'Budget by vehicle
Dim tmBvf As BVF
Dim imBdMnf() As Integer
Dim imBdYr() As Integer
Dim lmStartDateBd(0 To 1) As Long
Dim lmEndDateBd(0 To 1) As Long
'Spot Summary
Dim hmSsf As Integer    'file handle
'Replaced with tgRCSsf
'Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imSettingValue As Integer   'True=Don't enable any box with change
Dim smCntrStartDate As String
Dim lmCntrStartDate As Long
Dim smCntrEndDate As String
Dim lmCntrEndDate As Long
Dim imNoWks As Integer
Dim imShowIndex As Integer
Dim imTypeIndex As Integer
Dim imBSelectedIndex As Integer
Dim imChgMode As Integer
Dim imBSMode As Integer
Dim lmSplitDate As Long
Dim lmCntrCode() As Long
Dim tmSortCode() As SORTCODE
Dim smSortCodeTag As String
Dim tmRCSpotInfo() As RCSPOTINFO
'Period (column) Information
'Period (column) Information
Dim imPdYear As Integer
Dim imPdStartWk As Integer 'start week number
Dim imIStartYear As Integer
Dim imIStartWk As Integer
Dim imINoYears As Integer
Dim tmPdGroups(0 To 4) As PDGROUPS          'Index zero ignored
Dim imHotSpot(0 To 4, 0 To 4) As Integer    'Index zero ignored
Dim imInHotSpot As Integer
Dim imUpdateAllowed As Integer

Const LBONE = 1

Const WK1INDEX = 1
Const WK2INDEX = 2
Const WK3INDEX = 3
Const WK4INDEX = 4
Const NAMEINDEX = 1
Const DOLLAR1INDEX = 2
Const DOLLAR2INDEX = 3
Const DOLLAR3INDEX = 4
Const DOLLAR4INDEX = 5
Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    RCImpact.Refresh
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
    mInit
    If imTerminate Then
        cmcDone_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgNameCode

    Erase tmSortCode
    Erase smIShow
    Erase lmCntrCode
    Erase tgImpactRec
    Erase tgDollarRec
    Erase tmRCSpotInfo
    Erase imVefCode
    Erase tmDPBudgetInfo
    Erase tmBvfVeh
    Erase tmBvfVeh2
    Erase imBdMnf
    Erase imBdYr
    Erase tgClfRC
    Erase tgCffRC

    btrDestroy hmVef
    btrDestroy hmVsf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmRcf
    btrDestroy hmRif
    btrDestroy hmRdf
    btrDestroy hmLcf
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmBvf
    btrDestroy hmSsf
    
    Set RCImpact = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcBudget_Click()
    tmcDelay.Enabled = False
    If tgSpf.sRUseCorpCal <> "Y" Then
        If lbcBudget.SelCount = 1 Then
            If lbcProposal.SelCount > 0 Then
                tmcDelay.Enabled = True
            End If
        End If
    Else
        If lbcBudget.SelCount > 0 Then
            If lbcProposal.SelCount > 0 Then
                tmcDelay.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub lbcBudget_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'mInitDDE
    End If
    gCtrlGotFocus lbcBudget
End Sub
Private Sub lbcBudget_TopIndexChange(TopIndex As Integer)
    tmcDelay.Enabled = False
    If tgSpf.sRUseCorpCal <> "Y" Then
        If lbcBudget.SelCount = 1 Then
            If lbcProposal.SelCount > 0 Then
                tmcDelay.Enabled = True
            End If
        End If
    Else
        If lbcBudget.SelCount > 0 Then
            If lbcProposal.SelCount > 0 Then
                tmcDelay.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub lbcBudget_Scroll()
    tmcDelay.Enabled = False
    If tgSpf.sRUseCorpCal <> "Y" Then
        If lbcBudget.SelCount = 1 Then
            If lbcProposal.SelCount > 0 Then
                tmcDelay.Enabled = True
            End If
        End If
    Else
        If lbcBudget.SelCount > 0 Then
            If lbcProposal.SelCount > 0 Then
                tmcDelay.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub lbcProposal_Click()
    tmcDelay.Enabled = False
    If lbcProposal.SelCount > 0 Then
        tmcDelay.Enabled = True
    End If
End Sub
Private Sub lbcProposal_TopIndexChange(TopIndex As Integer)
    tmcDelay.Enabled = False
    If lbcProposal.SelCount > 0 Then
        tmcDelay.Enabled = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCompMonths                     *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Months for Year        *
'*                                                     *
'*******************************************************
Private Sub mCompMonths(ilYear As Integer, ilStartWk() As Integer, ilNoWks() As Integer)
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim ilLoop As Integer
    If rbcShow(0).Value Then    'Corporate
        slDate = "1/15/" & Trim$(str$(ilYear))
        slStart = gObtainStartCorp(slDate, True)
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndCorp(slStart, True)
            ilNoWks(ilLoop) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
            If ilLoop = 1 Then
                ilStartWk(1) = 1      'Set 1 0 below
            Else
                ilStartWk(ilLoop) = ilStartWk(ilLoop - 1) + ilNoWks(ilLoop - 1)
            End If
            slStart = gIncOneDay(slEnd)
        Next ilLoop
    Else                        'Standard
        'Compute start week number for each month
        slDate = "1/15/" & Trim$(str$(ilYear))
        slStart = gObtainStartStd(slDate)
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndStd(slStart)
            ilNoWks(ilLoop) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
            If ilLoop = 1 Then
                ilStartWk(1) = 1      'Set 1 0 below
            Else
                ilStartWk(ilLoop) = ilStartWk(ilLoop - 1) + ilNoWks(ilLoop - 1)
            End If
            slStart = gIncOneDay(slEnd)
        Next ilLoop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenSold                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Make Sold from SSF             *
'*                                                     *
'*******************************************************
Private Function mGenSold(llStartDate As Long, llEndDate As Long) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llTime As Long
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilAvailOk As Integer
    Dim ilTime As Integer
    Dim ilSpot As Integer
    Dim ilVpfIndex As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilIndex As Integer
    Dim ilRdf As Integer
    Dim ilWkIndex As Integer
    Dim ilVef As Integer
    Dim ilDone As Integer
    Dim ilVefCode As Integer
    Dim llLatestDate As Long
    Dim llTstDate As Long
    Dim ilVefIndex As Integer
    Dim ilType As Integer

    For ilVef = LBONE To UBound(tgImpactRec) - 1 Step 1
        ilDone = False
        For ilLoop = LBONE To ilVef - 1 Step 1
            If tgImpactRec(ilLoop).iVefCode = tgImpactRec(ilVef).iVefCode Then
                ilDone = True
                Exit For
            End If
        Next ilLoop
        'Exclude sports for now
        ilVefIndex = gBinarySearchVef(tgImpactRec(ilVef).iVefCode)
        If ilVefIndex <> -1 Then
            If tgMVef(ilVefIndex).sType = "G" Then
                'ilDone = True
            End If
        Else
            ilDone = True
        End If
        If Not ilDone Then
            ilVefCode = tgImpactRec(ilVef).iVefCode
            llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
            For llDate = llStartDate To llEndDate Step 1
                ilVefIndex = gBinarySearchVef(ilVefCode)
                If ilVefIndex <> -1 Then
                    If tgMVef(ilVefIndex).sType <> "G" Then
                        ilType = 0
                        tmSsfSrchKey.iType = 0
                        tmSsfSrchKey.iVefCode = ilVefCode
                        slDate = Format$(llDate, "m/d/yy")
                        gPackDate slDate, tmSsfSrchKey.iDate(0), tmSsfSrchKey.iDate(1)
                        tmSsfSrchKey.iStartTime(0) = 0
                        tmSsfSrchKey.iStartTime(1) = 0
                        imSsfRecLen = Len(tgRCSsf)
                        ilRet = gSSFGetGreaterOrEqual(hmSsf, tgRCSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Else
                        tmSsfSrchKey2.iVefCode = ilVefCode
                        slDate = Format$(llDate, "m/d/yy")
                        gPackDate slDate, tmSsfSrchKey2.iDate(0), tmSsfSrchKey2.iDate(1)
                        imSsfRecLen = Len(tgRCSsf)
                        ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tgRCSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                        ilType = tgRCSsf.iType
                    End If

                    ilRet = gBuildAvails(ilRet, ilVefCode, llDate, llLatestDate, ilType)
                    Do While (ilRet = BTRV_ERR_NONE) And (tgRCSsf.iType = ilType) And (tgRCSsf.iVefCode = ilVefCode)
                        gUnpackDateLong tgRCSsf.iDate(0), tgRCSsf.iDate(1), llTstDate
                        If llTstDate <> llDate Then
                            Exit Do
                        End If
                        ilVpfIndex = gVpfFind(RCImpact, tgRCSsf.iVefCode)
                        ilWkIndex = (llDate - lmCntrStartDate) \ 7 + 1
                        gUnpackDate tgRCSsf.iDate(0), tgRCSsf.iDate(1), slDate
                        ilDay = gWeekDayStr(slDate)
                        For ilLoop = 1 To tgRCSsf.iCount Step 1
                           LSet tmAvail = tgRCSsf.tPas(ADJSSFPASBZ + ilLoop)
                            'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                            If tmAvail.iRecType = 2 Then    'Cmml Avails only
                                ReDim tmRCSpotInfo(0 To tmAvail.iNoSpotsThis) As RCSPOTINFO
                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tgRCSsf.tPas(ADJSSFPASBZ + ilLoop + ilSpot)
                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    tmRCSpotInfo(ilSpot - 1).iLen = tmSpot.iPosLen And &HFFF
                                    tmRCSpotInfo(ilSpot - 1).lPrice = mGetCost(tmSdf, hmClf, hmCff, hmSmf, hmVef, hmVsf)
                                    tmRCSpotInfo(ilSpot - 1).iRank = (tmSpot.iRank And RANKMASK)
                                    tmRCSpotInfo(ilSpot - 1).iRecType = tmSpot.iRecType
                                Next ilSpot
                                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                'Increment inventory
                                For ilIndex = LBONE To UBound(tgImpactRec) - 1 Step 1
                                    If tgImpactRec(ilIndex).iVefCode = ilVefCode Then
                                        For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                                            If tmRdf(ilRdf).iCode = tgImpactRec(ilIndex).iRdfCode Then
                                                tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).iAvailDefined = 1
                                                ilAvailOk = True
                                                If (tmRdf(ilRdf).sInOut = "I") Then
                                                    If (tmRdf(ilRdf).ianfCode <> tmAvail.ianfCode) Then
                                                        ilAvailOk = False
                                                    End If
                                                End If
                                                If (tmRdf(ilRdf).sInOut = "O") Then
                                                    If tmAvail.ianfCode = tmRdf(ilRdf).ianfCode Then
                                                        ilAvailOk = False
                                                    End If
                                                End If
                                                If ilAvailOk Then
                                                    ilAvailOk = False
                                                    For ilTime = LBound(tmRdf(ilRdf).iStartTime, 2) To UBound(tmRdf(ilRdf).iStartTime, 2) Step 1
                                                        If (tmRdf(ilRdf).iStartTime(0, ilTime) <> 1) Or (tmRdf(ilRdf).iStartTime(1, ilTime) <> 0) Then
                                                            'If tmRdf(ilRdf).sWkDays(ilTime, ilDay + 1) = "Y" Then
                                                            If tmRdf(ilRdf).sWkDays(ilTime, ilDay) = "Y" Then
                                                                gUnpackTimeLong tmRdf(ilRdf).iStartTime(0, ilTime), tmRdf(ilRdf).iStartTime(1, ilTime), False, llRdfStartTime
                                                                gUnpackTimeLong tmRdf(ilRdf).iEndTime(0, ilTime), tmRdf(ilRdf).iEndTime(1, ilTime), True, llRdfEndTime
                                                                If (llTime >= llRdfStartTime) And (llTime < llRdfEndTime) Then
                                                                    ilAvailOk = True
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilTime
                                                End If
                                                If ilAvailOk Then
                                                    ilLen = tmAvail.iLen
                                                    ilUnits = tmAvail.iAvInfo And &H1F
                                                    ilNo30 = 0
                                                    ilNo60 = 0
                                                    If ilLen >= 30 Then
                                                        If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                                            If (ilLen Mod 30) = 0 Then
                                                                Do While ilLen >= 30
                                                                    ilNo30 = ilNo30 + 1
                                                                    ilLen = ilLen - 30
                                                                Loop
                                                            End If
                                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                            If (ilLen Mod 30) = 0 Then
                                                                Do While ilLen >= 30
                                                                    ilNo30 = ilNo30 + 1
                                                                    ilLen = ilLen - 30
                                                                Loop
                                                            End If
                                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                            If (ilLen Mod 30) = 0 Then
                                                                Do While ilLen >= 30
                                                                    ilNo30 = ilNo30 + 1
                                                                    ilLen = ilLen - 30
                                                                Loop
                                                            End If
                                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                        End If
                                                    Else
                                                        ilNo30 = 1
                                                    End If
                                                    tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Inv = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Inv + ilNo30
                                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                                        ilNo30 = 0
                                                        ilNo60 = 0
                                                        'ilLen = tmSpot.iPosLen And &HFFF
                                                        ilLen = tmRCSpotInfo(ilSpot - 1).iLen
                                                        If ilLen >= 30 Then
                                                            If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                                                If (ilLen Mod 30) = 0 Then
                                                                    Do While ilLen >= 30
                                                                        ilNo30 = ilNo30 + 1
                                                                        ilLen = ilLen - 30
                                                                    Loop
                                                                End If
                                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                                If (ilLen Mod 30) = 0 Then
                                                                    Do While ilLen >= 30
                                                                        ilNo30 = ilNo30 + 1
                                                                        ilLen = ilLen - 30
                                                                    Loop
                                                                End If
                                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                                If (ilLen Mod 30) = 0 Then
                                                                    Do While ilLen >= 30
                                                                        ilNo30 = ilNo30 + 1
                                                                        ilLen = ilLen - 30
                                                                    Loop
                                                                End If
                                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                            End If
                                                        Else
                                                            ilNo30 = 1
                                                        End If
                                                        If (tmRCSpotInfo(ilSpot - 1).iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE Then
                                                        Else
                                                        End If
                                                        If tmRCSpotInfo(ilSpot - 1).iRank <= 1000 Then
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + tmRCSpotInfo(ilSpot - 1).lPrice
                                                        ElseIf (tmRCSpotInfo(ilSpot - 1).iRank = REMNANTRANK) Then 'Remnant
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + tmRCSpotInfo(ilSpot - 1).lPrice
                                                        ElseIf (tmRCSpotInfo(ilSpot - 1).iRank = DIRECTRESPONSERANK) Or (tmRCSpotInfo(ilSpot - 1).iRank = 1030) Then 'Direct Response or per Inquiry
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + tmRCSpotInfo(ilSpot - 1).lPrice
                                                        ElseIf (tmRCSpotInfo(ilSpot - 1).iRank = TRADERANK) Then 'Trade
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                                            tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + tmRCSpotInfo(ilSpot - 1).lPrice
                                                        ElseIf (tmRCSpotInfo(ilSpot - 1).iRank = PROMORANK) Then 'Promo
                                                        ElseIf (tmRCSpotInfo(ilSpot - 1).iRank = PSARANK) Then  'PSA
                                                        End If
                                                    Next ilSpot
                                                End If
                                            End If
                                        Next ilRdf
                                    End If
                                Next ilIndex
                            End If
                        Next ilLoop
                        imSsfRecLen = Len(tgRCSsf)
                        ilRet = gSSFGetNext(hmSsf, tgRCSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If tgMVef(ilVefIndex).sType = "G" Then
                            ilType = tgRCSsf.iType
                        End If
                    Loop
                End If
            Next llDate
            mGetMissed ilVefCode, llStartDate, llEndDate
        End If
    Next ilVef
    mGenSold = True
    Exit Function

    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetBudgetDollars               *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get budget dollars             *
'*                                                     *
'*            Note: Similar code in RateCard.Frm       *
'*                                                     *
'*******************************************************
Private Sub mGetBudgetDollars()
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim ilFound As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilWkIndex As Integer
    Dim ilRCWkNo As Integer
    Dim ilBdWkNo As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilFirstLastWk As Integer
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilBvf As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilAvailOk As Integer
    Dim llSsfDate As Long
    Dim llTime As Long
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilRet As Integer
    Dim ilTime As Integer
    Dim ilDay As Integer
    Dim ilVpfIndex As Integer
    Dim ilDP As Integer
    Dim llPrice As Long
    Dim llBudget As Long
    Dim llDPInv As Long
    Dim ilDPNotZero As Integer
    If lbcBudget.SelCount <= 0 Then
        Exit Sub
    End If
    If lbcProposal.SelCount <= 0 Then
        Exit Sub
    End If
    If UBound(tmBvfVeh) = LBound(tmBvfVeh) Then
        Exit Sub
    End If
    'Build array of vehicles
    ReDim imVefCode(0 To 0) As Integer
    For ilVef = LBONE To UBound(tgImpactRec) - 1 Step 1
        ilFound = False
        For ilLoop = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
            If imVefCode(ilLoop) = tgImpactRec(ilVef).iVefCode Then
                ilFound = True
            End If
        Next ilLoop
        If Not ilFound Then
            imVefCode(UBound(imVefCode)) = tgImpactRec(ilVef).iVefCode
            ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
        End If
    Next ilVef
    'Count number of Base dayparts within the vehicle
    For ilVef = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
        ReDim tmDPBudgetInfo(0 To 0) As DPBUDGETINFO
        For llRif = LBound(tmRif) To UBound(tmRif) - 1 Step 1
            'If (tmRif(ilRif).iRcfCode = tgImpactRec(1).iRcfCode) And (tmRif(ilRif).iVefCode = imVefCode(ilVef)) And (tmRif(ilRif).iYear = imYear) Then
            If (tmRif(llRif).iRcfCode = tgImpactRec(1).iRcfCode) And (tmRif(llRif).iVefCode = imVefCode(ilVef)) Then
                'Test if daypart is base daypart
                'For ilRdf = 1 To UBound(tmRdf) - 1 Step 1
                For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                    If tmRif(llRif).iRdfCode = tmRdf(ilRdf).iCode Then
                        If tmRdf(ilRdf).sBase = "Y" Then
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lRifIndex = llRif
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).iRdfIndex = ilRdf
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).iInv = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lRCPrice = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lPrice = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lBudget = 0
                            ReDim Preserve tmDPBudgetInfo(0 To UBound(tmDPBudgetInfo) + 1) As DPBUDGETINFO
                        End If
                        Exit For
                    End If
                Next ilRdf
            End If
        Next llRif
        If UBound(tmDPBudgetInfo) > LBound(tmDPBudgetInfo) Then
            For ilBvf = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
                If tmBvfVeh(ilBvf).iVefCode = imVefCode(ilVef) Then
                    ilVpfIndex = gVpfFind(RCImpact, imVefCode(ilVef))
                    For llDate = lmCntrStartDate To lmCntrEndDate Step 7
                        slDate = Format$(llDate, "m/d/yy")
                        gObtainMonthYear 0, slDate, ilMonth, ilYear
                        'If rbcShow(0).Value Then
                        '    ilType = 4
                        'Else
                        '    ilType = 0
                        'End If
                        'gObtainWkNo ilType, slDate, ilWkNo, ilFirstLastWk
                        gObtainWkNo 0, slDate, ilRCWkNo, ilFirstLastWk
                        gObtainWkNo 5, slDate, ilBdWkNo, ilFirstLastWk
                        If LBound(tmDPBudgetInfo) + 1 = UBound(tmDPBudgetInfo) Then
                            tmDPBudgetInfo(LBound(tmDPBudgetInfo)).iInv = 1
                        Else
                            tmSsfSrchKey.iType = 0
                            tmSsfSrchKey.iVefCode = imVefCode(ilVef)
                            gPackDate slDate, tmSsfSrchKey.iDate(0), tmSsfSrchKey.iDate(1)
                            tmSsfSrchKey.iStartTime(0) = 0
                            tmSsfSrchKey.iStartTime(1) = 0
                            imSsfRecLen = Len(tgRCSsf)
                            ilRet = gSSFGetGreaterOrEqual(hmSsf, tgRCSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                            Do While (ilRet = BTRV_ERR_NONE) And (tgRCSsf.iType = 0) And (tgRCSsf.iVefCode = imVefCode(ilVef))
                                gUnpackDateLong tgRCSsf.iDate(0), tgRCSsf.iDate(1), llSsfDate
                                If llSsfDate > llDate + 6 Then
                                    Exit Do
                                End If
                                ilDay = gWeekDayLong(llSsfDate)
                                For ilLoop = 1 To tgRCSsf.iCount Step 1
                                   LSet tmAvail = tgRCSsf.tPas(ADJSSFPASBZ + ilLoop)
                                    'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                                    If tmAvail.iRecType = 2 Then    'Cmml Avails only
                                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                            If ilYear = tmRif(tmDPBudgetInfo(ilDP).lRifIndex).iYear Then    'ilBdYear Then
                                                ilRdf = tmDPBudgetInfo(ilDP).iRdfIndex
                                                ilAvailOk = True
                                                If (tmRdf(ilRdf).sInOut = "I") Then
                                                    If (tmRdf(ilRdf).ianfCode <> tmAvail.ianfCode) Then
                                                        ilAvailOk = False
                                                    End If
                                                End If
                                                If (tmRdf(ilRdf).sInOut = "O") Then
                                                    If tmAvail.ianfCode = tmRdf(ilRdf).ianfCode Then
                                                        ilAvailOk = False
                                                    End If
                                                End If
                                                If ilAvailOk Then
                                                    ilAvailOk = False
                                                    For ilTime = LBound(tmRdf(ilRdf).iStartTime, 2) To UBound(tmRdf(ilRdf).iStartTime, 2) Step 1
                                                        If (tmRdf(ilRdf).iStartTime(0, ilTime) <> 1) Or (tmRdf(ilRdf).iStartTime(1, ilTime) <> 0) Then
                                                            'If tmRdf(ilRdf).sWkDays(ilTime, ilDay + 1) = "Y" Then
                                                            If tmRdf(ilRdf).sWkDays(ilTime, ilDay) = "Y" Then
                                                                gUnpackTimeLong tmRdf(ilRdf).iStartTime(0, ilTime), tmRdf(ilRdf).iStartTime(1, ilTime), False, llRdfStartTime
                                                                gUnpackTimeLong tmRdf(ilRdf).iEndTime(0, ilTime), tmRdf(ilRdf).iEndTime(1, ilTime), True, llRdfEndTime
                                                                If (llTime >= llRdfStartTime) And (llTime < llRdfEndTime) Then
                                                                    ilAvailOk = True
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilTime
                                                End If
                                                If ilAvailOk Then
                                                    ilLen = tmAvail.iLen
                                                    ilUnits = tmAvail.iAvInfo And &H1F
                                                    ilNo30 = 0
                                                    ilNo60 = 0
                                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                                        If (ilLen Mod 30) = 0 Then
                                                            Do While ilLen >= 30
                                                                ilNo30 = ilNo30 + 1
                                                                ilLen = ilLen - 30
                                                            Loop
                                                        End If
                                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                        If (ilLen Mod 30) = 0 Then
                                                            Do While ilLen >= 30
                                                                ilNo30 = ilNo30 + 1
                                                                ilLen = ilLen - 30
                                                            Loop
                                                        End If
                                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                        If (ilLen Mod 30) = 0 Then
                                                            Do While ilLen >= 30
                                                                ilNo30 = ilNo30 + 1
                                                                ilLen = ilLen - 30
                                                            Loop
                                                        End If
                                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                    End If
                                                    'Add
                                                    tmDPBudgetInfo(ilDP).iInv = tmDPBudgetInfo(ilDP).iInv + ilNo30
                                                End If
                                            End If
                                        Next ilDP
                                    End If
                                Next ilLoop
                                imSsfRecLen = Len(tgRCSsf)
                                ilRet = gSSFGetNext(hmSsf, tgRCSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            Loop
                        End If
                        'Compute Budget by daypart
                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                            'If ilYear = imYear Then
                            If ilYear = tmRif(tmDPBudgetInfo(ilDP).lRifIndex).iYear Then    'ilBdYear Then
                                tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tmRif(tmDPBudgetInfo(ilDP).lRifIndex).lRate(ilRCWkNo)
                            End If
                                'If ilWkNo = 1 Then
                                '    If (ilFirstLastWk) Or rbcShow(1).Value Then
                                '        tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(0) + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(ilWkNo)
                                '    Else
                                '        tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(ilWkNo)
                                '    End If
                                'ElseIf ilWkNo = 52 Then
                                '    If (ilFirstLastWk) Or rbcShow(1).Value Then
                                '        tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(53) + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(ilWkNo)
                                '    Else
                                '        tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(ilWkNo)
                                '    End If
                                'Else
                                '    tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tmRif(tmDPBudgetInfo(ilDP).iRifIndex).lRate(ilWkNo)
                                'End If
                            'End If
                        Next ilDP
                        'If ilWkNo = 1 Then
                        '    If (ilFirstLastWk) Or rbcShow(1).Value Then
                        '        llBudget = tmBvfVeh(ilBvf).lGross(ilWkNo) + tmBvfVeh(ilBvf).lGross(0)
                        '    Else
                        '        llBudget = tmBvfVeh(ilBvf).lGross(ilWkNo)
                        '    End If
                        'ElseIf ilWkNo = 52 Then
                         '   If (ilFirstLastWk) Or rbcShow(1).Value Then
                        '        llBudget = tmBvfVeh(ilBvf).lGross(ilWkNo) + tmBvfVeh(ilBvf).lGross(53)
                        '    Else
                        '        llBudget = tmBvfVeh(ilBvf).lGross(ilWkNo)
                        '    End If
                        'Else
                        '    llBudget = tmBvfVeh(ilBvf).lGross(ilWkNo)
                        'End If
                        llBudget = 0
                        If llDate < lmSplitDate Then
                            If (llDate >= lmStartDateBd(0)) And (llDate <= lmEndDateBd(0)) Then
                                llBudget = tmBvfVeh(ilBvf).lGross(ilBdWkNo)
                            End If
                        Else
                            'Scan tmBvfVeh2 for matching tmBvfVeh
                            For ilLoop = LBound(tmBvfVeh2) To UBound(tmBvfVeh2) - 1 Step 1
                                If tmBvfVeh(ilBvf).iVefCode = tmBvfVeh2(ilLoop).iVefCode Then
                                    If (llDate >= lmStartDateBd(1)) And (llDate <= lmEndDateBd(1)) Then
                                        llBudget = tmBvfVeh2(ilBvf).lGross(ilBdWkNo)
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                        ilDPNotZero = -1
                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                            If tmDPBudgetInfo(ilDP).lRCPrice > 0 Then
                                ilDPNotZero = ilDP
                                Exit For
                            End If
                        Next ilDP
                        If ilDPNotZero >= LBound(tmDPBudgetInfo) Then
                            llDPInv = 0
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                '4-12-11 maintain better accuracy
                                llDPInv = llDPInv + (100 * CSng(tmDPBudgetInfo(ilDP).lRCPrice) * tmDPBudgetInfo(ilDP).iInv) / tmDPBudgetInfo(ilDPNotZero).lRCPrice
                            Next ilDP
                            If llDPInv > 0 Then
                                llPrice = (10000 * CSng(llBudget)) / llDPInv
                            Else
                                llPrice = 0
                            End If
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                tmDPBudgetInfo(ilDP).lPrice = (CSng(tmDPBudgetInfo(ilDP).lRCPrice) * llPrice) / tmDPBudgetInfo(ilDPNotZero).lRCPrice
                            Next ilDP
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                tmDPBudgetInfo(ilDP).lBudget = (tmDPBudgetInfo(ilDP).lPrice * tmDPBudgetInfo(ilDP).iInv) / 100
                                tmDPBudgetInfo(ilDP).lPrice = tmDPBudgetInfo(ilDP).lPrice / 100
                            Next ilDP
                        Else
                            For ilDP = LBound(tmDPBudgetInfo) + 1 To UBound(tmDPBudgetInfo) - 1 Step 1
                                tmDPBudgetInfo(ilDP).lBudget = 0
                            Next ilDP
                        End If
                        ilWkIndex = (llDate - lmCntrStartDate) \ 7 + 1
                        For ilLoop = LBONE To UBound(tgImpactRec) - 1 Step 1
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                ilRdf = tmDPBudgetInfo(ilDP).iRdfIndex
                                If (tgImpactRec(ilLoop).iVefCode = imVefCode(ilVef)) And (tgImpactRec(ilLoop).iRdfCode = tmRdf(ilRdf).iCode) Then
                                    tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).lBudget = tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).lBudget + tmDPBudgetInfo(ilDP).lBudget * 100 'Get Pennies
                                    'This line is ok- removed to match the code in RCImpact.Bas
                                    'Exit For
                                End If
                            Next ilDP
                        Next ilLoop
                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                            tmDPBudgetInfo(ilDP).iInv = 0
                            tmDPBudgetInfo(ilDP).lRCPrice = 0
                            tmDPBudgetInfo(ilDP).lPrice = 0
                            tmDPBudgetInfo(ilDP).lBudget = 0
                        Next ilDP
                    Next llDate
                    Exit For
                End If
            Next ilBvf
        End If
    Next ilVef
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetMissed                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Missed Information         *
'*                                                     *
'*******************************************************
Private Sub mGetMissed(ilVefCode As Integer, llStartDate As Long, llEndDate As Long)
    Dim ilPass As Integer
    Dim slType As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilIndex As Integer
    Dim ilRdf As Integer
    Dim ilAvailOk As Integer
    Dim ilVpfIndex As Integer
    Dim ilWkIndex As Integer
    Dim ilDay As Integer
    Dim llTime As Long
    Dim ilTime As Integer
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilLen As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim llPrice As Long
    Dim ilLnRdf As Integer
    Dim ilRdfAnfCode As Integer

    'Key 2: VefCode; SchStatus; AdfCode; Date, Time
    For ilPass = 0 To 2 Step 1
        tmSdfSrchKey2.iVefCode = ilVefCode
        If ilPass = 0 Then
            slType = "M"
        ElseIf ilPass = 1 Then
            slType = "R"
        ElseIf ilPass = 2 Then
            slType = "U"
        End If
        tmSdfSrchKey2.sSchStatus = slType
        tmSdfSrchKey2.iAdfCode = 0
        slDate = Format$(llStartDate, "m/d/yy")
        gPackDate slDate, tmSdfSrchKey2.iDate(0), tmSdfSrchKey2.iDate(1)
        tmSdfSrchKey2.iTime(0) = 0
        tmSdfSrchKey2.iTime(1) = 0
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
        'This code added as replacement for Ext operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                ilVpfIndex = gVpfFind(RCImpact, tmSdf.iVefCode)
                ilWkIndex = (llDate - llStartDate) \ 7 + 1
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                ilDay = gWeekDayStr(slDate)
                gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                'Increment inventory
                For ilIndex = LBONE To UBound(tgImpactRec) - 1 Step 1
                    If tgImpactRec(ilIndex).iVefCode = ilVefCode Then
                        For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                            If tmRdf(ilRdf).iCode = tgImpactRec(ilIndex).iRdfCode Then
                                ilAvailOk = False
                                For ilTime = LBound(tmRdf(ilRdf).iStartTime, 2) To UBound(tmRdf(ilRdf).iStartTime, 2) Step 1
                                    If (tmRdf(ilRdf).iStartTime(0, ilTime) <> 1) Or (tmRdf(ilRdf).iStartTime(1, ilTime) <> 0) Then
                                        'If tmRdf(ilRdf).sWkDays(ilTime, ilDay + 1) = "Y" Then
                                        If tmRdf(ilRdf).sWkDays(ilTime, ilDay) = "Y" Then
                                            gUnpackTimeLong tmRdf(ilRdf).iStartTime(0, ilTime), tmRdf(ilRdf).iStartTime(1, ilTime), False, llRdfStartTime
                                            gUnpackTimeLong tmRdf(ilRdf).iEndTime(0, ilTime), tmRdf(ilRdf).iEndTime(1, ilTime), True, llRdfEndTime
                                            If (llTime >= llRdfStartTime) And (llTime < llRdfEndTime) Then
                                                ilAvailOk = True
                                            End If
                                        End If
                                    End If
                                Next ilTime

                                If ilAvailOk Then
                                    llPrice = mGetCost(tmSdf, hmClf, hmCff, hmSmf, hmVef, hmVsf)
                                    ilRdfAnfCode = 0
                                    ilLnRdf = gBinarySearchRdf(tmClf.iRdfCode)
                                    If ilLnRdf <> -1 Then
                                        ilRdfAnfCode = tgMRdf(ilLnRdf).ianfCode
                                    End If
                                    If (tmRdf(ilRdf).sInOut = "I") Then
                                        If (tmRdf(ilRdf).ianfCode <> ilRdfAnfCode) Then
                                            ilAvailOk = False
                                        End If
                                    End If
                                    If (tmRdf(ilRdf).sInOut = "O") Then
                                        If ilRdfAnfCode = tmRdf(ilRdf).ianfCode Then
                                            ilAvailOk = False
                                        End If
                                    End If
                                End If

                                If ilAvailOk Then
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    'ilLen = tmSpot.iPosLen And &HFFF
                                    ilLen = tmSdf.iLen
                                    If ilLen >= 30 Then
                                        If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                            If (ilLen Mod 30) = 0 Then
                                                Do While ilLen >= 30
                                                    ilNo30 = ilNo30 + 1
                                                    ilLen = ilLen - 30
                                                Loop
                                            End If
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                            If (ilLen Mod 30) = 0 Then
                                                Do While ilLen >= 30
                                                    ilNo30 = ilNo30 + 1
                                                    ilLen = ilLen - 30
                                                Loop
                                            End If
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                            If (ilLen Mod 30) = 0 Then
                                                Do While ilLen >= 30
                                                    ilNo30 = ilNo30 + 1
                                                    ilLen = ilLen - 30
                                                Loop
                                            End If
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                        End If
                                    Else
                                        ilNo30 = 1
                                    End If
                                    tmChfSrchKey.lCode = tmSdf.lChfCode
                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) And (tmChf.sType <> "M") And (tmChf.sType <> "S") Then
                                        tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                        tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + llPrice
                                    End If
                                    'If (tmSpot.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE Then
                                    'Else
                                    'End If
                                    'If tmSpot.iRank <= 1000 Then
                                    '    tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                    '    tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + mGetCost()
                                    'ElseIf (tmSpot.iRank = 1020) Then   'Remnant
                                    'ElseIf (tmSpot.iRank = 1010) Or (tmSpot.iRank = 1030) Then   'Direct Response or per Inquiry
                                    'ElseIf (tmSpot.iRank = 1040) Then   'Trade
                                    'ElseIf (tmSpot.iRank = 1050) Then 'Promo
                                    'ElseIf (tmSpot.iRank = 1060) Then    'PSA
                                    'End If
                                End If
                            End If
                        Next ilRdf
                    End If
                Next ilIndex
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilPass
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowDates                   *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mGetShowDates()
'
'   mGetShowDates
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim slWkEnd As String
    Dim ilFound As Integer
    Dim ilYearOk As Integer
    Dim ilWkNo As Integer
    Dim ilWkCount As Integer
    ReDim ilStartWk(0 To 12) As Integer 'Index zero ignored
    ReDim ilNoWks(0 To 12) As Integer   'Index zero ignored
    Dim slFontName As String
    Dim flFontSize As Single
    'If UBound(tgImpactRec) <= 1 Then
    '    For ilIndex = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
    '        tmPdGroups(ilIndex).iStartWkNo = -1
    '        tmPdGroups(ilIndex).iNoWks = 0
    '        tmPdGroups(ilIndex).iTrueNoWks = 0
    '        tmPdGroups(ilIndex).iFltNo = 0
    '        tmPdGroups(ilIndex).sStartDate = ""
    '        tmPdGroups(ilIndex).sEndDate = ""
    '        gSetShow pbcImpact, "", tmWKCtrls(ilIndex)
    '    Next ilIndex
    '    Exit Sub
    'End If
    If imPdYear = 0 Then
        Exit Sub
    End If
    slFontName = pbcImpact.FontName
    flFontSize = pbcImpact.FontSize
    pbcImpact.FontBold = False
    pbcImpact.FontSize = 7
    pbcImpact.FontName = "Arial"
    pbcImpact.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    tmPdGroups(1).iYear = imPdYear
    tmPdGroups(1).iStartWkNo = imPdStartWk

    ilIndex = 1
    Do
        ilFound = False
        If ilIndex > 1 Then
            If tmPdGroups(ilIndex).iYear <> tmPdGroups(ilIndex - 1).iYear Then
                ilYearOk = False
            Else
                ilYearOk = True
            End If
        Else
            ilYearOk = False
        End If
        If Not ilYearOk Then
            mCompMonths tmPdGroups(ilIndex).iYear, ilStartWk(), ilNoWks()
        End If
        If rbcType(0).Value Then        'Quarter
            For ilLoop = 1 To 12 Step 3
                If (tmPdGroups(ilIndex).iStartWkNo >= ilStartWk(ilLoop)) And (tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(ilLoop) + ilNoWks(ilLoop) + ilNoWks(ilLoop + 1) + ilNoWks(ilLoop + 2) - 1) Then
                    tmPdGroups(ilIndex).iNoWks = ilNoWks(ilLoop) + ilNoWks(ilLoop + 1) + ilNoWks(ilLoop + 2) - (tmPdGroups(ilIndex).iStartWkNo - ilStartWk(ilLoop))
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        ElseIf rbcType(1).Value Then    'Month
            For ilLoop = 1 To 12 Step 1
                If (tmPdGroups(ilIndex).iStartWkNo >= ilStartWk(ilLoop)) And (tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(ilLoop) + ilNoWks(ilLoop) - 1) Then
                    tmPdGroups(ilIndex).iNoWks = ilNoWks(ilLoop) - (tmPdGroups(ilIndex).iStartWkNo - ilStartWk(ilLoop))
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        ElseIf rbcType(2).Value Then    'Week
            If tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(12) + ilNoWks(12) - 1 Then
                tmPdGroups(ilIndex).iNoWks = 1
                ilFound = True
            End If
        End If
        If ilFound Then
            If ilIndex <> 4 Then
                tmPdGroups(ilIndex + 1).iStartWkNo = tmPdGroups(ilIndex).iStartWkNo + tmPdGroups(ilIndex).iNoWks
                tmPdGroups(ilIndex + 1).iYear = tmPdGroups(ilIndex).iYear   'imPdYear
            End If
            ilIndex = ilIndex + 1
        Else
            tmPdGroups(ilIndex).iYear = tmPdGroups(ilIndex).iYear + 1
            tmPdGroups(ilIndex).iStartWkNo = 1
            'Test if year exist
            If tmPdGroups(ilIndex).iYear > imIStartYear + imINoYears - 1 Then
                For ilLoop = ilIndex To 4 Step 1
                    tmPdGroups(ilLoop).iStartWkNo = -1
                    tmPdGroups(ilLoop).iTrueNoWks = 0
                    tmPdGroups(ilLoop).iNoWks = 0
                Next ilLoop
                Exit Do
            End If
        End If
    Loop Until ilIndex > 4
    'Compute Start/End Date if groups
    For ilIndex = LBONE To UBound(tmPdGroups) Step 1
        If tmPdGroups(ilIndex).iStartWkNo > 0 Then
            If rbcShow(0).Value Then    'Corporate
                slDate = "1/15/" & Trim$(str$(tmPdGroups(ilIndex).iYear))
                slStart = gObtainStartCorp(slDate, True)
                slDate = "12/15/" & Trim$(str$(tmPdGroups(ilIndex).iYear))
                slEnd = gObtainEndCorp(slDate, True)
            Else
                slDate = "1/15/" & Trim$(str$(tmPdGroups(ilIndex).iYear))
                slStart = gObtainStartStd(slDate)
                slDate = "12/15/" & Trim$(str$(tmPdGroups(ilIndex).iYear))
                slEnd = gObtainEndStd(slDate)
            End If
            ilWkNo = 1
            Do
                If ilWkNo = tmPdGroups(ilIndex).iStartWkNo Then
                    tmPdGroups(ilIndex).sStartDate = slStart
                    slWkEnd = gObtainNextSunday(slStart)
                    ilWkCount = 1
                    Do
                        If ilWkNo = tmPdGroups(ilIndex).iStartWkNo + tmPdGroups(ilIndex).iNoWks - 1 Then
                            tmPdGroups(ilIndex).sEndDate = slWkEnd
                            tmPdGroups(ilIndex).iTrueNoWks = ilWkCount
                            'slDate = tmPdGroups(ilIndex).sStartDate & "-" & tmPdGroups(ilIndex).sEndDate
                            slDate = tmPdGroups(ilIndex).sStartDate
                            slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                            gSetShow pbcImpact, slDate, tmWKCtrls(ilIndex)
                            Exit Do
                        Else
                            ilWkNo = ilWkNo + 1
                            ilWkCount = ilWkCount + 1
                            slWkEnd = gIncOneWeek(slWkEnd)
                            If gDateValue(slWkEnd) > gDateValue(slEnd) Then
                                tmPdGroups(ilIndex).sEndDate = slEnd
                                tmPdGroups(ilIndex).iTrueNoWks = ilWkCount - 1
                                'slDate = Left$(tmPdGroups(ilIndex).sStartDate, Len(tmPdGroups(ilIndex).sStartDate) - 3)
                                'slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                                slDate = tmPdGroups(ilIndex).sStartDate
                                slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                                gSetShow pbcImpact, slDate, tmWKCtrls(ilIndex)
                                Exit Do
                            End If
                        End If
                    Loop
                    Exit Do
                Else
                    ilWkNo = ilWkNo + 1
                    slStart = gIncOneWeek(slStart)
                End If
            Loop
        Else
            tmPdGroups(ilIndex).sStartDate = ""
            tmPdGroups(ilIndex).sEndDate = ""
            gSetShow pbcImpact, "", tmWKCtrls(ilIndex)
        End If
    Next ilIndex
    For ilIndex = LBONE To UBound(tmPdGroups) Step 1
        If tmPdGroups(ilIndex).sStartDate <> "" Then
            If (lmCntrStartDate >= gDateValue(tmPdGroups(ilIndex).sStartDate)) And (lmCntrStartDate <= gDateValue(tmPdGroups(ilIndex).sEndDate)) Then
                tmPdGroups(ilIndex).iStartWkNo = tmPdGroups(ilIndex).iStartWkNo + (lmCntrStartDate - gDateValue(tmPdGroups(ilIndex).sStartDate)) \ 7
                tmPdGroups(ilIndex).sStartDate = smCntrStartDate
                slDate = tmPdGroups(ilIndex).sStartDate
                slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                gSetShow pbcImpact, slDate, tmWKCtrls(ilIndex)
            End If
            If (lmCntrEndDate >= gDateValue(tmPdGroups(ilIndex).sStartDate)) And (lmCntrEndDate <= gDateValue(tmPdGroups(ilIndex).sEndDate)) Then
                tmPdGroups(ilIndex).sEndDate = smCntrEndDate
                slDate = tmPdGroups(ilIndex).sStartDate
                slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                gSetShow pbcImpact, slDate, tmWKCtrls(ilIndex)
            End If
            If gDateValue(tmPdGroups(ilIndex).sStartDate) > lmCntrEndDate Then
                tmPdGroups(ilIndex).iStartWkNo = 0
                tmPdGroups(ilIndex).sStartDate = ""
                tmPdGroups(ilIndex).sEndDate = ""
                gSetShow pbcImpact, "", tmWKCtrls(ilIndex)
            End If
        End If
    Next ilIndex
    pbcImpact.FontSize = flFontSize
    pbcImpact.FontName = slFontName
    pbcImpact.FontSize = flFontSize
    pbcImpact.FontBold = True
    mGetShowPrices
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowPrices                  *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mGetShowPrices()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRdf As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim ilStartWk As Integer
    Dim ilEndWk As Integer
    Dim slStr As String
    Dim ilVefDp As Integer
    Dim llRCPrice As Long   'Current Rate Card Price (From flight-Proposal price)
    Dim llSPrice As Long    'Spot Price
    Dim llTSPrice As Long   'Total spot price
    Dim llTSSold As Long
    Dim llCBudget As Long   'Current budget
    Dim llDollarSold As Long      'Sold
    Dim llAvail As Long
    Dim llBudget1 As Long
    Dim llBudget2 As Long
    Dim ilRet As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilWksUndefined As Integer
    Dim slColor As String
    ReDim smIShow(0 To 5, 0 To 7 * (UBound(tgImpactRec) - 1) + 1) As String
    For ilLoop = LBound(smIShow, 1) To UBound(smIShow, 1) Step 1
        For ilIndex = LBound(smIShow, 2) To UBound(smIShow, 2) Step 1
            smIShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    For ilVefDp = LBONE To UBound(tgImpactRec) - 1 Step 1
        tmVefSrchKey.iCode = tgImpactRec(ilVefDp).iVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slStr = Trim$(tmVef.sName)
            ilRdf = -1
            For ilLoop = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                If tmRdf(ilLoop).iCode = tgImpactRec(ilVefDp).iRdfCode Then
                    ilRdf = ilLoop
                    Exit For
                End If
            Next ilLoop
            If ilRdf >= 0 Then
                llColor = pbcImpact.ForeColor
                slFontName = pbcImpact.FontName
                flFontSize = pbcImpact.FontSize
                pbcImpact.ForeColor = BLUE
                pbcImpact.FontBold = False
                pbcImpact.FontSize = 7
                pbcImpact.FontName = "Arial"
                pbcImpact.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
                slStr = slStr & "/" & Trim$(tmRdf(ilRdf).sName)
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 1) = tmICtrls(NAMEINDEX).sShow
                slStr = "MCurrent Rate Card Price"
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 2) = "  " & Mid$(tmICtrls(NAMEINDEX).sShow, 2)
                slStr = "MProposal Spot Price"
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 3) = "  " & Mid$(tmICtrls(NAMEINDEX).sShow, 2)
                slStr = "MDifference"
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 4) = "  " & Mid$(tmICtrls(NAMEINDEX).sShow, 2)
                slStr = "MCurrent Spot Price To Make Budget"
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 5) = "  " & Mid$(tmICtrls(NAMEINDEX).sShow, 2)
                slStr = "MResulting Spot Price To Make Budget"
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 6) = "  " & Mid$(tmICtrls(NAMEINDEX).sShow, 2)
                slStr = "MDifference"
                gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX)
                smIShow(NAMEINDEX, 7 * (ilVefDp - 1) + 7) = "  " & Mid$(tmICtrls(NAMEINDEX).sShow, 2)
                pbcImpact.FontSize = flFontSize
                pbcImpact.FontName = slFontName
                pbcImpact.FontSize = flFontSize
                pbcImpact.ForeColor = llColor
                pbcImpact.FontBold = True
                For ilGroup = LBONE To UBound(tmPdGroups) Step 1
                    If tmPdGroups(ilGroup).sStartDate <> "" Then
                        ilStartWk = (gDateValue(tmPdGroups(ilGroup).sStartDate) - lmCntrStartDate) \ 7 + 1
                        ilEndWk = (gDateValue(tmPdGroups(ilGroup).sEndDate) - lmCntrStartDate) \ 7 + 1
                        llRCPrice = 0
                        llSPrice = 0
                        llTSPrice = 0
                        llTSSold = 0
                        llCBudget = 0
                        llDollarSold = 0
                        llAvail = 0
                        ilWksUndefined = 0
                        For ilWk = ilStartWk To ilEndWk Step 1
                            llRCPrice = llRCPrice + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lRCPrice
                            llSPrice = llSPrice + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lSPrice
                            llTSPrice = llTSPrice + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lTSPrice
                            llTSSold = llTSSold + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).iTSpots
                            llAvail = llAvail + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).l30Inv - tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).l30Sold
                            If lbcBudget.SelCount > 0 Then   'imBSelectedIndex >= 0 Then
                                llCBudget = llCBudget + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lBudget    'lDollarSold
                                llDollarSold = llDollarSold + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lDollarSold
                            End If
                            If tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).iAvailDefined = 0 Then
                                ilWksUndefined = ilWksUndefined + 1
                            End If
                        Next ilWk
                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 2) = tmICtrls(NAMEINDEX + ilGroup).sShow
                        llRCPrice = (llRCPrice / (ilEndWk - ilStartWk + 1)) / 100
                        slStr = gLongToStrDec(llRCPrice, 0)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX + ilGroup)
                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 2) = tmICtrls(NAMEINDEX + ilGroup).sShow
                        llSPrice = (llSPrice / (ilEndWk - ilStartWk + 1)) / 100
                        slStr = gLongToStrDec(llSPrice, 0)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX + ilGroup)
                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 3) = tmICtrls(NAMEINDEX + ilGroup).sShow
                        slStr = gLongToStrDec(llSPrice - llRCPrice, 0)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX + ilGroup)
                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 4) = tmICtrls(NAMEINDEX + ilGroup).sShow
                        If lbcBudget.SelCount > 0 Then   'imBSelectedIndex >= 0 Then
                            If ilWksUndefined <> ilEndWk - ilStartWk + 1 Then
                                If llAvail > 0 Then
                                    slColor = ""
                                    llBudget1 = (llCBudget - llDollarSold) / llAvail
                                    If llBudget1 / 100 <= llRCPrice Then
                                        slColor = "G"
                                    Else
                                        slColor = "R"
                                    End If
                                    slStr = gLongToStrDec(llBudget1 / 100, 0)
                                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                    gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX + ilGroup)
                                    smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 5) = slColor & tmICtrls(NAMEINDEX + ilGroup).sShow
                                    If llAvail - llTSSold > 0 Then
                                        llBudget2 = (llCBudget - llDollarSold - llTSPrice) / (llAvail - llTSSold)
                                        If llBudget2 / 100 <= llRCPrice Then
                                            slColor = "G"
                                        Else
                                            slColor = "R"
                                        End If
                                        slStr = gLongToStrDec(llBudget2 / 100, 0)
                                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                        gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX + ilGroup)
                                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 6) = slColor & tmICtrls(NAMEINDEX + ilGroup).sShow
                                        llBudget2 = llBudget2 - llBudget1
                                        slStr = gLongToStrDec(llBudget2 / 100, 0)
                                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                        gSetShow pbcImpact, slStr, tmICtrls(NAMEINDEX + ilGroup)
                                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 7) = tmICtrls(NAMEINDEX + ilGroup).sShow
                                    Else
                                        gSetShow pbcImpact, "Sold Out", tmICtrls(NAMEINDEX + ilGroup)
                                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 6) = tmICtrls(NAMEINDEX + ilGroup).sShow
                                        smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 7) = ""
                                    End If
                                Else
                                    gSetShow pbcImpact, "Sold Out", tmICtrls(NAMEINDEX + ilGroup)
                                    smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 5) = tmICtrls(NAMEINDEX + ilGroup).sShow
                                    smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 6) = tmICtrls(NAMEINDEX + ilGroup).sShow
                                    smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 7) = ""
                                End If
                            Else
                                gSetShow pbcImpact, "Undefined", tmICtrls(NAMEINDEX + ilGroup)
                                smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 5) = tmICtrls(NAMEINDEX + ilGroup).sShow
                                smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 6) = tmICtrls(NAMEINDEX + ilGroup).sShow
                                smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 7) = ""
                            End If
                        Else
                            smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 5) = ""
                            smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 6) = ""
                            smIShow(NAMEINDEX + ilGroup, 7 * (ilVefDp - 1) + 7) = ""
                        End If
                    End If
                Next ilGroup
            End If
        End If
    Next ilVefDp
    imSettingValue = True
    vbcImpact.Min = LBONE   'LBound(smIShow, 2)
    imSettingValue = True
    If UBound(smIShow, 2) - 1 <= vbcImpact.LargeChange + 1 Then ' + 1 Then
        vbcImpact.Max = LBONE   'LBound(smIShow, 2)
    Else
        vbcImpact.Max = UBound(smIShow, 2) - vbcImpact.LargeChange
    End If
    imSettingValue = True
    vbcImpact.Value = vbcImpact.Min
    imSettingValue = False
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
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    imLBICtrls = 1
    imLBWKCtrls = 1
    imTerminate = False
    imFirstActivate = True
    imPdYear = 0
    imFirstFocus = True
    Screen.MousePointer = vbHourglass
    RCImpact.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm RCImpact
    'RCImpact.Show
    Screen.MousePointer = vbHourglass
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", RCImpact
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", RCImpact
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", RCImpact
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", RCImpact
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", RCImpact
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmRcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rcf.Btr)", RCImpact
    On Error GoTo 0
    ReDim tmRcf(0 To 0) As RCF
    imRcfRecLen = Len(tmRcf(0))
    hmRif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rif.Btr)", RCImpact
    On Error GoTo 0
    ReDim tmRif(0 To 0) As RIF
    imRifRecLen = Len(tmRif(0))
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", RCImpact
    On Error GoTo 0
    'ReDim tmRdf(1 To 1) As RDF
    ReDim tmRdf(0 To 0) As RDF
    'imRdfRecLen = Len(tmRdf(1))
    imRdfRecLen = Len(tmRdf(0))
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", RCImpact
    On Error GoTo 0
    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", RCImpact
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", RCImpact
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    hmBvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bvf.Btr)", RCImpact
    On Error GoTo 0
    ReDim tmBvfVeh(0 To 0) As BVF
    imBvfRecLen = Len(tmBvf)
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", RCImpact
    On Error GoTo 0
    smSortCodeTag = ""
    imBSelectedIndex = -1
    imChgMode = False
    imBSMode = False
    mInitBox
    sgNameCodeTag = ""
    ReDim tgNameCode(0 To 0) As SORTCODE
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    mObtainRcf
    If imTerminate Then
        Exit Sub
    End If
    mObtainRif
    If imTerminate Then
        Exit Sub
    End If
    mObtainRdf
    If imTerminate Then
        Exit Sub
    End If
    ReDim tgImpactRec(0 To 1) As IMPACTREC
    ReDim smIShow(0 To 5, 0 To 1) As String
    For ilLoop = LBound(smIShow, 1) To UBound(smIShow, 1) Step 1
        For ilIndex = LBound(smIShow, 2) To UBound(smIShow, 2) Step 1
            smIShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    imPdYear = 0
    lmCntrStartDate = 0
    lmCntrEndDate = 0
    imNoWks = 0
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm RCImpact
    imInHotSpot = False
    imHotSpot(1, 1) = 2415  'Left
    imHotSpot(1, 2) = 15    'Top
    imHotSpot(1, 3) = 2415 + 150 'Right
    imHotSpot(1, 4) = 15 + 180  'Bottom
    imHotSpot(2, 1) = 2565  'Left
    imHotSpot(2, 2) = 15    'Top
    imHotSpot(2, 3) = 2565 + 150 'Right
    imHotSpot(2, 4) = 15 + 180  'Bottom
    imHotSpot(3, 1) = 5760  'Left
    imHotSpot(3, 2) = 15    'Top
    imHotSpot(3, 3) = 5760 + 150 'Right
    imHotSpot(3, 4) = 15 + 180  'Bottom
    imHotSpot(4, 1) = 5910  'Left
    imHotSpot(4, 2) = 15    'Top
    imHotSpot(4, 3) = 5910 + 150 'Right
    imShowIndex = 1 'Std Month
    imTypeIndex = 1 'Quarter
    If tgSpf.sRUseCorpCal <> "Y" Then
        rbcShow(0).Enabled = False
    End If
    imSettingValue = False
    Screen.MousePointer = vbDefault
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
    Dim flTextHeight As Single  'Standard text height
    flTextHeight = pbcImpact.TextHeight("1") - 35
    'plcSelect.Move 5340, 60
    plcProposal.Move 135, 570
    plcImpact.Move 2895, plcProposal.Top - 30, pbcImpact.Width + vbcImpact.Width + fgPanelAdj, pbcImpact.Height + fgPanelAdj
    pbcImpact.Move plcImpact.Left + fgBevelX, plcImpact.Top + fgBevelY
    vbcImpact.Move pbcImpact.Width + fgBevelX + 15, plcImpact.Height - vbcImpact.Height - fgBevelY

    'Week 1
    gSetCtrl tmWKCtrls(WK1INDEX), 2430, 195, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmWKCtrls(WK2INDEX), 3330, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmWKCtrls(WK3INDEX), 4230, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmWKCtrls(WK4INDEX), 5130, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH

    'Name
    gSetCtrl tmICtrls(NAMEINDEX), 30, 390, 2385, fgBoxGridH
    'Dollar 1
    gSetCtrl tmICtrls(DOLLAR1INDEX), 2430, tmICtrls(NAMEINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 2
    gSetCtrl tmICtrls(DOLLAR2INDEX), 3330, tmICtrls(DOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Dollar 3
    gSetCtrl tmICtrls(DOLLAR3INDEX), 4230, tmICtrls(DOLLAR1INDEX).fBoxY, 885, fgBoxGridH
    'Dollar 4
    gSetCtrl tmICtrls(DOLLAR4INDEX), 5130, tmICtrls(DOLLAR1INDEX).fBoxY, 885, fgBoxGridH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainRcf                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Rate Definitions        *
'*                                                     *
'*******************************************************
Private Sub mObtainRcf()
    Dim ilRecLen As Integer
    Dim ilUpperBound As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilRet As Integer
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    ReDim tmRcf(0 To 0) As RCF
    ilRecLen = Len(tmRcf(0)) 'btrRecordLength(hlAdf)  'Get and save record length
    ilUpperBound = UBound(tmRcf)
    ilExtLen = Len(tmRcf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmRcf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmRcf, tmRcf(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            imTerminate = True
            Exit Sub
        End If
        Call btrExtSetBounds(hmRcf, llNoRec, -1, "UC", "RCF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hmRcf, ilOffSet, ilRecLen)  'Extract iCode field
        If ilRet <> BTRV_ERR_NONE Then
            imTerminate = True
            Exit Sub
        End If
        'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
        ilUpperBound = UBound(tmRcf)
        ilRet = btrExtGetNext(hmRcf, tmRcf(ilUpperBound), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                imTerminate = True
                Exit Sub
            End If
            ilUpperBound = UBound(tmRcf)
            ilExtLen = Len(tmRcf(ilUpperBound))  'Extract operation record size
            'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmRcf, tmRcf(ilUpperBound), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tmRcf(0 To ilUpperBound) As RCF
                ilRet = btrExtGetNext(hmRcf, tmRcf(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRcf, tmRcf(ilUpperBound), ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainRdf                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tmRdf                  *
'*                                                     *
'*******************************************************
Private Sub mObtainRdf()
'
'   mObtainRdf
'   Where:
'       tmRdf() (I)- RDF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilUpperBound As Integer
    Dim llNoRec As Long
    'ReDim tmRdf(1 To 1) As RDF
    ReDim tmRdf(0 To 0) As RDF
    ilUpperBound = UBound(tmRdf)
    ilExtLen = Len(tmRdf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmRdf) 'Obtain number of records
    btrExtClear hmRdf   'Clear any previous extend operation
    'ilRet = btrGetFirst(hmRdf, tmRdf(1), imRdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ilRet = btrGetFirst(hmRdf, tmRdf(0), imRdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            imTerminate = True
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hmRdf, llNoRec, -1, "UC", "RDF", "") 'Set extract limits (all records)
    ilRet = btrExtAddField(hmRdf, 0, ilExtLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hmRdf)    'Extract record
    ilUpperBound = UBound(tmRdf)
    ilRet = btrExtGetNext(hmRdf, tmRdf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            imTerminate = True
            Exit Sub
        End If
        ilUpperBound = UBound(tmRdf)
        ilExtLen = Len(tmRdf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hmRdf, tmRdf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmRdf, tmRdf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tmRdf(1 To ilUpperBound) As RDF
            ReDim Preserve tmRdf(0 To ilUpperBound) As RDF
            ilRet = btrExtGetNext(hmRdf, tmRdf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmRdf, tmRdf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainRif                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tmRif                  *
'*                                                     *
'*******************************************************
Private Sub mObtainRif()
'
'   mObtainRif
'   Where:
'       tmRif() (I)- RIF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim llUpperBound As Long
    Dim llNoRec As Long

    ReDim tmRif(0 To 0) As RIF
    llUpperBound = UBound(tmRif)
    ilExtLen = Len(tmRif(0))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmRif) 'Obtain number of records
    btrExtClear hmRif   'Clear any previous extend operation
    ilRet = btrGetFirst(hmRif, tmRif(0), imRifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            imTerminate = True
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hmRif, llNoRec, -1, "UC", "RIF", "") 'Set extract limits (all records)
    ilRet = btrExtAddField(hmRif, 0, ilExtLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hmRif)    'Extract record
    llUpperBound = UBound(tmRif)
    ilRet = btrExtGetNext(hmRif, tmRif(llUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            imTerminate = True
            Exit Sub
        End If
        llUpperBound = UBound(tmRif)
        ilExtLen = Len(tmRif(0))  'Extract operation record size
        'ilRet = btrExtGetFirst(hmRif, tmRif(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmRif, tmRif(llUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            llUpperBound = llUpperBound + 1
            ReDim Preserve tmRif(0 To llUpperBound) As RIF
            ilRet = btrExtGetNext(hmRif, tmRif(llUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmRif, tmRif(llUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilAAS As Integer
    Dim ilAASCode As Integer
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilHOType As Integer
    Dim ilShow As Integer

    imPopReqd = False
    ilRet = gPopVehBudgetBox(RCImpact, 0, 0, 1, lbcBudget, tgNameCode(), sgNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopBudgetBox)", RCImpact
        On Error GoTo 0
    End If
    ilAAS = -1
    ilAASCode = 0
    slStatus = "WIC"
    'slCntrType = "" 'All
    If (tgSpf.sSchdPromo <> "Y") And (tgSpf.sSchdPSA <> "Y") Then
        slCntrType = "" 'All
    Else
        slCntrType = "CVTRQ" 'All except Promo
        If tgSpf.sSchdPromo <> "Y" Then
            slCntrType = slCntrType & "M"
        End If
        If tgSpf.sSchdPSA <> "Y" Then
            slCntrType = slCntrType & "S"
        End If
    End If
    ilCurrent = 0
    ilHOType = -1
    ilShow = 1  '5  '#/Advertiser
    ilRet = gPopCntrForAASBox(RCImpact, ilAAS, ilAASCode, slStatus, slCntrType, ilCurrent, ilHOType, ilShow, lbcProposal, tmSortCode(), smSortCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopCntrForAASBox)", RCImpact
        On Error GoTo 0
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
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

    sgNameCodeTag = ""

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload RCImpact
    igManUnload = NO
End Sub

Private Sub lbcProposal_Scroll()
    tmcDelay.Enabled = False
    If lbcProposal.SelCount > 0 Then
        tmcDelay.Enabled = True
    End If
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcImpact_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    Dim ilWk As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer
    If Button = 2 Then  'Right Mouse
        Exit Sub
    End If
    'Check if hot spot
    If imInHotSpot Then
        Exit Sub
    End If
    If lbcProposal.ListCount <= 0 Then
        Exit Sub
    End If
    If UBound(tgImpactRec) > 1 Then
        For ilLoop = LBONE To UBound(imHotSpot, 1) Step 1
            If (X >= imHotSpot(ilLoop, 1)) And (X <= imHotSpot(ilLoop, 3)) And (Y >= imHotSpot(ilLoop, 2)) And (Y <= imHotSpot(ilLoop, 4)) Then
                Screen.MousePointer = vbHourglass
                imInHotSpot = True
                Select Case ilLoop
                    Case 1  'Goto Start
                        imPdYear = imIStartYear
                        imPdStartWk = imIStartWk
                    Case 2  'Reduce by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(1).iYear = imIStartYear) And (tmPdGroups(1).iStartWkNo <= imIStartWk) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo <= 1) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 9 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                If (tmPdGroups(1).iStartWkNo >= 12) And (tmPdGroups(1).iStartWkNo <= 14) Then
                                    imPdStartWk = ilStartWk(1)
                                ElseIf (tmPdGroups(1).iStartWkNo >= 25) And (tmPdGroups(1).iStartWkNo <= 27) Then
                                    imPdStartWk = ilStartWk(1)  'Compute start of second quarter
                                    For ilWk = 1 To 3 Step 1
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                    Next ilWk
                                Else
                                    imPdStartWk = ilStartWk(1)  'Compute start of third quarter
                                    For ilWk = 1 To 6 Step 1
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                    Next ilWk
                                End If
                                If (imPdYear = imIStartYear) And (imPdStartWk < imIStartWk) Then
                                    imPdStartWk = imIStartWk
                                End If
                            End If
                        ElseIf rbcType(1).Value Then    'Month
                            If (tmPdGroups(1).iYear = imIStartYear) And (tmPdGroups(1).iStartWkNo <= imIStartWk) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo <= 1) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 11 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 2 To 12 Step 1
                                    If tmPdGroups(1).iStartWkNo = ilStartWk(ilWk) Then
                                        Exit For
                                    End If
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk - 1)
                                Next ilWk
                            End If
                            If (imPdYear = imIStartYear) And (imPdStartWk < imIStartWk) Then
                                imPdStartWk = imIStartWk
                            End If
                        ElseIf rbcType(2).Value Then    'Week
                            If (tmPdGroups(1).iYear = imIStartYear) And (tmPdGroups(1).iStartWkNo <= imIStartWk) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo <= 1) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 12 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                                imPdStartWk = imPdStartWk - 1
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                imPdStartWk = tmPdGroups(1).iStartWkNo - 1
                            End If
                            If (imPdYear = imIStartYear) And (imPdStartWk < imIStartWk) Then
                                imPdStartWk = imIStartWk
                            End If
                        End If
                    Case 3  'Increase by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(4).iYear = imIStartYear + imINoYears - 1) And (tmPdGroups(4).iStartWkNo > 39) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths tmPdGroups(4).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(4).iYear = imIStartYear + imINoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12)) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths tmPdGroups(4).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(4).iYear = imIStartYear + imINoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        End If
                    Case 4  'GoTo End
                        imPdYear = imIStartYear + imINoYears - 1
                        If rbcType(0).Value Then    'Quarter
                            imPdStartWk = 1
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(9)  'At end
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(12) + ilNoWks(12) - 4
                        End If
                End Select
                pbcImpact.Cls
                mGetShowDates
                pbcImpact_Paint
                Screen.MousePointer = vbDefault
                imInHotSpot = False
                Exit Sub
            End If
        Next ilLoop
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub pbcImpact_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    llColor = pbcImpact.ForeColor
    slFontName = pbcImpact.FontName
    flFontSize = pbcImpact.FontSize
    pbcImpact.ForeColor = BLUE
    pbcImpact.FontBold = False
    pbcImpact.FontSize = 7
    pbcImpact.FontName = "Arial"
    pbcImpact.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = imLBWKCtrls To UBound(tmWKCtrls) Step 1
        'gPaintArea pbcImpact, tmWKCtrls(ilBox).fBoxX, tmWKCtrls(ilBox).fBoxY, tmWKCtrls(ilBox).fBoxW - 15, tmWKCtrls(ilBox).fBoxH - 15, WHITE
        pbcImpact.CurrentX = tmWKCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcImpact.CurrentY = tmWKCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcImpact.Print tmWKCtrls(ilBox).sShow
    Next ilBox
    pbcImpact.FontSize = flFontSize
    pbcImpact.FontName = slFontName
    pbcImpact.FontSize = flFontSize
    pbcImpact.ForeColor = llColor
    pbcImpact.FontBold = True
    ilStartRow = vbcImpact.Value '+ 1  'Top location
    ilEndRow = vbcImpact.Value + vbcImpact.LargeChange ' + 1
    If ilEndRow > UBound(smIShow, 2) Then
        ilEndRow = UBound(smIShow, 2)
    End If
    llColor = pbcImpact.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBICtrls To UBound(tmICtrls) Step 1
            If ilBox = imLBICtrls Then
                pbcImpact.CurrentX = tmICtrls(ilBox).fBoxX + fgBoxInsetX
                pbcImpact.CurrentY = tmICtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) ' - 30'+ fgBoxInsetY
                slFontName = pbcImpact.FontName
                flFontSize = pbcImpact.FontSize
                pbcImpact.ForeColor = BLUE
                pbcImpact.FontBold = False
                pbcImpact.FontSize = 7
                pbcImpact.FontName = "Arial"
                pbcImpact.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
                slStr = smIShow(ilBox, ilRow)
                pbcImpact.Print slStr
                pbcImpact.ForeColor = llColor
                pbcImpact.FontSize = flFontSize
                pbcImpact.FontName = slFontName
                pbcImpact.FontSize = flFontSize
                pbcImpact.ForeColor = llColor
                pbcImpact.FontBold = True
            Else
                slStr = smIShow(ilBox, ilRow)
                'If InStr(1, slStr, "-") <> 0 Then
                '    pbcImpact.ForeColor = RED
                'End If
                If Left$(slStr, 1) = "R" Then
                    pbcImpact.ForeColor = RED
                    slStr = right$(slStr, Len(slStr) - 1)
                ElseIf Left$(slStr, 1) = "G" Then
                    pbcImpact.ForeColor = DARKGREEN
                    slStr = right$(slStr, Len(slStr) - 1)
                Else
                    pbcImpact.ForeColor = llColor
                End If
                If InStr(1, slStr, "Sold", 1) = 0 Then
                    pbcImpact.CurrentX = gRightJustifyShowStr(pbcImpact, slStr, tmICtrls(ilBox)) 'tmICtrls(ilBox).fBoxX + fgBoxInsetX
                Else
                    pbcImpact.CurrentX = tmICtrls(ilBox).fBoxX + fgBoxInsetX
                End If
                pbcImpact.CurrentY = tmICtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 15 ' - 30'+ fgBoxInsetY
                pbcImpact.Print slStr
                pbcImpact.ForeColor = llColor
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcShow_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShow(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer

    If Value Then
        Screen.MousePointer = vbHourglass
        pbcImpact.Cls
        If imIStartYear <> 0 Then
            imPdYear = tmPdGroups(1).iYear
            If imTypeIndex = 1 Then 'By month
                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                imPdStartWk = tmPdGroups(1).iStartWkNo
                ilFound = False
                Do
                    For ilLoop = 1 To 12 Step 1
                        If imPdStartWk = ilStartWk(ilLoop) Then
                            ilFound = True
                            Exit Do
                        End If
                    Next ilLoop
                    If imPdStartWk <= 1 Then
                        imPdStartWk = 1
                        ilFound = True
                        Exit Do
                    End If
                    imPdStartWk = imPdStartWk - 1
                Loop Until ilFound
            ElseIf imTypeIndex = 2 Then 'Weeks- make sure not pass end
                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                imPdStartWk = tmPdGroups(4).iStartWkNo
                ilFound = False
                Do
                    If (imPdStartWk > ilStartWk(12) + ilNoWks(12) - 6) Then
                        If imPdStartWk <= 1 Then
                            imPdStartWk = 1
                            ilFound = True
                            Exit Do
                        End If
                        imPdStartWk = imPdStartWk - 1
                    Else
                        ilFound = True
                        Exit Do
                    End If
                Loop Until ilFound
            End If
        End If
        imShowIndex = Index
        If lbcBudget.SelCount > 0 Then
            mGetBudgetDollars
        End If
        mGetShowDates
        pbcImpact_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    ReDim ilStartWk(0 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer

    If Value Then
        Screen.MousePointer = vbHourglass
        pbcImpact.Cls
        If imIStartYear <> 0 Then
            imPdYear = tmPdGroups(1).iYear
            If Index = 0 Then   'Change to Quarter
                If (imPdYear = imIStartYear) Then
                    imPdStartWk = imIStartWk
                Else
                    imPdStartWk = 1
                End If
            ElseIf Index = 1 Then   'Month
                If (imTypeIndex = 0) Or (imTypeIndex = 3) Then
                    If (imPdYear = imIStartYear) Then
                        imPdStartWk = imIStartWk
                    Else
                        imPdStartWk = 1
                    End If
                Else    'by week- back up to start of month
                    mCompMonths imPdYear, ilStartWk(), ilNoWks()
                    imPdStartWk = tmPdGroups(1).iStartWkNo
                    ilFound = False
                    Do
                        For ilLoop = 1 To 12 Step 1
                            If imPdStartWk = ilStartWk(ilLoop) Then
                                ilFound = True
                                Exit Do
                            End If
                        Next ilLoop
                        If imPdStartWk <= 1 Then
                            imPdStartWk = 1
                            ilFound = True
                            Exit Do
                        End If
                        If (imPdYear = imIStartYear) And (imPdStartWk = imIStartWk) Then
                            ilFound = True
                            Exit Do
                        End If
                        imPdStartWk = imPdStartWk - 1
                    Loop Until ilFound
                End If
            ElseIf Index = 2 Then   'Week
                If (imTypeIndex = 0) Or (imTypeIndex = 3) Then
                    If (imPdYear = imIStartYear) Then
                        imPdStartWk = imIStartWk
                    Else
                        imPdStartWk = 1
                    End If
                Else  'Month
                    imPdStartWk = tmPdGroups(1).iStartWkNo
                End If
            End If
        End If
        imTypeIndex = Index
        mGetShowDates
        pbcImpact_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub tmcDelay_Timer()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim ilImpact As Integer
    Dim ilFound As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilWkIndex As Integer
    Dim ilDay As Integer
    Dim ilUnits As Integer
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slDate As String
    Dim slStart As String
    Dim llPriceAdj As Long
    Dim ilRdf As Integer
    Dim ilBaseDP As Integer
    Dim ilBvf As Integer
    Dim slNameYear As String
    Dim slBdMnfName As String
    Dim slYear As String
    Dim ilCount As Integer
    tmcDelay.Enabled = False
    ReDim tgImpactRec(0 To 1) As IMPACTREC
    ReDim smIShow(0 To 5, 0 To 1) As String
    For ilLoop = LBound(smIShow, 1) To UBound(smIShow, 1) Step 1
        For ilIndex = LBound(smIShow, 2) To UBound(smIShow, 2) Step 1
            smIShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    imPdYear = 0
    lmCntrStartDate = 0
    lmCntrEndDate = 0
    imNoWks = 0
    pbcImpact.Cls
    If lbcProposal.SelCount <= 0 Then
        Exit Sub
    End If
    ReDim lmCntrCode(0 To lbcProposal.SelCount) As Long
    Screen.MousePointer = vbHourglass
    'Compute Overall Dates- Used to determine # week Buckets
    ilIndex = 1
    For ilLoop = 0 To lbcProposal.ListCount - 1 Step 1
        If lbcProposal.Selected(ilLoop) Then
            slNameCode = tmSortCode(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            lmCntrCode(ilIndex) = Val(slCode)
            tmChfSrchKey.lCode = lmCntrCode(ilIndex)
            ilIndex = ilIndex + 1
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmChf.iStartDate(0), tmChf.iStartDate(1), llStartDate
                gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), llEndDate
                If llStartDate <= llEndDate Then
                    If lmCntrStartDate = 0 Then
                        lmCntrStartDate = llStartDate
                        lmCntrEndDate = llEndDate
                    Else
                        If llStartDate < lmCntrStartDate Then
                            lmCntrStartDate = llStartDate
                        End If
                        If llEndDate > lmCntrEndDate Then
                            lmCntrEndDate = llEndDate
                        End If
                    End If
                End If
            End If
        End If
    Next ilLoop
    If lmCntrStartDate = 0 Then
        Exit Sub
    End If
    smCntrStartDate = Format$(lmCntrStartDate, "m/d/yy")
    smCntrEndDate = Format$(lmCntrEndDate, "m/d/yy")
    smCntrStartDate = gObtainPrevMonday(smCntrStartDate)
    smCntrEndDate = gObtainNextSunday(smCntrEndDate)
    lmCntrStartDate = gDateValue(smCntrStartDate)
    lmCntrEndDate = gDateValue(smCntrEndDate)
    imNoWks = (lmCntrEndDate - lmCntrStartDate) \ 7 + 1
    gObtainMonthYear 0, smCntrStartDate, ilMonth, ilYear
    imPdYear = ilYear
    imPdStartWk = 1
    imIStartYear = imPdYear
    If rbcShow(0).Value Then    'Corporate
        slDate = "1/15/" & Trim$(str$(ilYear))
        slStart = gObtainStartCorp(slDate, True)
    Else                        'Standard
        slDate = "1/15/" & Trim$(str$(ilYear))
        slStart = gObtainStartStd(slDate)
    End If
    imIStartWk = (lmCntrStartDate - gDateValue(slStart)) \ 7 + 1
    imPdStartWk = imIStartWk
    gObtainMonthYear 0, smCntrEndDate, ilMonth, ilYear
    imINoYears = ilYear - imPdYear + 1
    ReDim tgDollarRec(0 To imNoWks, 0 To 1) As DOLLARREC
    'Build Buckets for each contract selected
    For ilLoop = LBONE To UBound(lmCntrCode) Step 1
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, lmCntrCode(ilLoop), False, tgChfRC, tgClfRC(), tgCffRC())
        If ilRet Then
            For ilClf = LBound(tgClfRC) To UBound(tgClfRC) - 1 Step 1
                gUnpackDateLong tgClfRC(ilClf).ClfRec.iStartDate(0), tgClfRC(ilClf).ClfRec.iStartDate(1), llStartDate
                gUnpackDateLong tgClfRC(ilClf).ClfRec.iEndDate(0), tgClfRC(ilClf).ClfRec.iEndDate(1), llEndDate
                If (llStartDate <= llEndDate) And (tgClfRC(ilClf).ClfRec.sType <> "O") And (tgClfRC(ilClf).ClfRec.sType <> "A") Then
                    llPriceAdj = 100  'mAdjPrice(tgChfRC.iRcfCode, tgClfRC(ilClf).ClfRec.iLen)
                    ilFound = False
                    For ilImpact = LBONE To UBound(tgImpactRec) - 1 Step 1
                        If (tgImpactRec(ilImpact).iVefCode = tgClfRC(ilClf).ClfRec.iVefCode) And (tgImpactRec(ilImpact).iRdfCode = tgClfRC(ilClf).ClfRec.iRdfCode) And (tgImpactRec(ilImpact).iRcfCode = tgChfRC.iRcfCode) Then
                            ilFound = True
                            ilIndex = ilImpact
                            Exit For
                        End If
                    Next ilImpact
                    'Only include base dayparts
                    If Not ilFound Then
                        ilBaseDP = False
                        'For ilRdf = 1 To UBound(tmRdf) - 1 Step 1
                        For ilRdf = LBound(tmRdf) To UBound(tmRdf) - 1 Step 1
                            If tgClfRC(ilClf).ClfRec.iRdfCode = tmRdf(ilRdf).iCode Then
                                If tmRdf(ilRdf).sBase = "Y" Then
                                    ilBaseDP = True
                                End If
                                Exit For
                            End If
                        Next ilRdf
                    Else
                        ilBaseDP = True
                    End If
                    If ilBaseDP Then
                        If Not ilFound Then
                            ilIndex = UBound(tgImpactRec)
                            tgImpactRec(ilIndex).iVefCode = tgClfRC(ilClf).ClfRec.iVefCode
                            tgImpactRec(ilIndex).iRdfCode = tgClfRC(ilClf).ClfRec.iRdfCode
                            tgImpactRec(ilIndex).iRcfCode = tgChfRC.iRcfCode
                            tgImpactRec(ilIndex).iPtDollarRec = ilIndex
                            ReDim Preserve tgImpactRec(0 To ilIndex + 1) As IMPACTREC
                            ReDim Preserve tgDollarRec(0 To imNoWks, 0 To ilIndex + 1) As DOLLARREC
                        End If
                        ilUnits = (tgClfRC(ilClf).ClfRec.iLen - 1) \ 30 + 1
                        ilCff = tgClfRC(ilClf).iFirstCff
                        Do While ilCff >= 0
                            gUnpackDateLong tgCffRC(ilCff).CffRec.iStartDate(0), tgCffRC(ilCff).CffRec.iStartDate(1), llStartDate
                            gUnpackDateLong tgCffRC(ilCff).CffRec.iEndDate(0), tgCffRC(ilCff).CffRec.iEndDate(1), llEndDate
                            If llEndDate < llStartDate Then
                                Exit Do
                            End If
                            Do While gWeekDayLong(llStartDate) <> 0
                                llStartDate = llStartDate - 1
                            Loop
                            For llDate = llStartDate To llEndDate Step 7
                                ilWkIndex = (llDate - lmCntrStartDate) \ 7 + 1
                                tgDollarRec(ilWkIndex, ilIndex).iNoHits = tgDollarRec(ilWkIndex, ilIndex).iNoHits + 1
                                tgDollarRec(ilWkIndex, ilIndex).lRCPrice = tgDollarRec(ilWkIndex, ilIndex).lRCPrice + (llPriceAdj * tgCffRC(ilCff).CffRec.lPropPrice) '/ 100
                                If tgCffRC(ilCff).CffRec.sPriceType = "T" Then
                                    If tgCffRC(ilCff).CffRec.sDyWk = "D" Then
                                        For ilDay = 0 To 6 Step 1
                                            If llDate + ilDay <= llEndDate Then
                                                tgDollarRec(ilWkIndex, ilIndex).lSPrice = tgDollarRec(ilWkIndex, ilIndex).lSPrice + (llPriceAdj * tgCffRC(ilCff).CffRec.lActPrice) / 100
                                                tgDollarRec(ilWkIndex, ilIndex).lTSPrice = tgDollarRec(ilWkIndex, ilIndex).lTSPrice + tgCffRC(ilCff).CffRec.iDay(ilDay) * tgCffRC(ilCff).CffRec.lActPrice
                                                tgDollarRec(ilWkIndex, ilIndex).iTSpots = tgDollarRec(ilWkIndex, ilIndex).iTSpots + ilUnits * tgCffRC(ilCff).CffRec.iDay(ilDay)
                                            End If
                                        Next ilDay
                                    Else
                                        tgDollarRec(ilWkIndex, ilIndex).lSPrice = tgDollarRec(ilWkIndex, ilIndex).lSPrice + (llPriceAdj * tgCffRC(ilCff).CffRec.lActPrice) / 100
                                        tgDollarRec(ilWkIndex, ilIndex).lTSPrice = tgDollarRec(ilWkIndex, ilIndex).lTSPrice + (tgCffRC(ilCff).CffRec.iSpotsWk + tgCffRC(ilCff).CffRec.iXSpotsWk) * tgCffRC(ilCff).CffRec.lActPrice
                                        tgDollarRec(ilWkIndex, ilIndex).iTSpots = tgDollarRec(ilWkIndex, ilIndex).iTSpots + ilUnits * (tgCffRC(ilCff).CffRec.iSpotsWk + tgCffRC(ilCff).CffRec.iXSpotsWk)
                                    End If
                                End If
                            Next llDate
                            ilCff = tgCffRC(ilCff).iNextCff
                        Loop
                    End If
                End If
            Next ilClf
        End If
    Next ilLoop
    'Compute Average values for Rate Card Price and Proposal Price
    For ilWkIndex = LBONE To UBound(tgDollarRec, 1) Step 1
        For ilIndex = LBONE To UBound(tgDollarRec, 2) Step 1
            If tgDollarRec(ilWkIndex, ilIndex).iNoHits > 0 Then
                tgDollarRec(ilWkIndex, ilIndex).lRCPrice = tgDollarRec(ilWkIndex, ilIndex).lRCPrice / tgDollarRec(ilWkIndex, ilIndex).iNoHits
                tgDollarRec(ilWkIndex, ilIndex).lSPrice = tgDollarRec(ilWkIndex, ilIndex).lSPrice / tgDollarRec(ilWkIndex, ilIndex).iNoHits
            End If
        Next ilIndex
    Next ilWkIndex
    'Compute $ Sold and # 30 Units Available for each vehicle and week
    ilRet = mGenSold(lmCntrStartDate, lmCntrEndDate)
    If lbcBudget.SelCount = 1 Then
        'Get the budgets
        ReDim imBdMnf(0 To 0) As Integer
        ReDim imBdYr(0 To 0) As Integer
        For ilLoop = 0 To lbcBudget.ListCount - 1 Step 1
            If lbcBudget.Selected(ilLoop) Then
                slNameCode = tgNameCode(ilLoop).sKey  'lbcBudget.List(ilIndex - 1)
                ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
                ilRet = gParseItem(slNameYear, 2, "/", slBdMnfName)
                ilRet = gParseItem(slNameYear, 1, "/", slYear)
                slYear = gSubStr("9999", slYear)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imBdMnf(0) = Val(slCode)
                imBdYr(0) = Val(slYear)
                Exit For
            End If
        Next ilLoop
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If imBdYr(0) = tgMCof(ilLoop).iYear Then
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), lmSplitDate
                lmSplitDate = lmSplitDate + 1
                Exit For
            End If
        Next ilLoop
        lmStartDateBd(1) = -1
        lmEndDateBd(1) = -1
        For ilLoop = 0 To UBound(imBdYr) Step 1
            If tgSpf.sRUseCorpCal <> "Y" Then
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearStartDate(0, slDate)
                lmStartDateBd(ilLoop) = gDateValue(slDate)
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearEndDate(0, slDate)
                lmEndDateBd(ilLoop) = gDateValue(slDate)
            Else
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearStartDate(5, slDate)
                lmStartDateBd(ilLoop) = gDateValue(slDate)
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearEndDate(5, slDate)
                lmEndDateBd(ilLoop) = gDateValue(slDate)
            End If
        Next ilLoop
        If mReadBvfRec(hmBvf, imBdMnf(0), imBdYr(0), tmBvfVeh()) Then
            ReDim tmBvfVeh2(0 To 0) As BVF
            mGetBudgetDollars
        End If
    ElseIf lbcBudget.SelCount = 2 Then
        'Get the budgets
        ReDim imBdMnf(0 To 1) As Integer
        ReDim imBdYr(0 To 1) As Integer
        ilCount = 0
        For ilLoop = 0 To lbcBudget.ListCount - 1 Step 1
            If lbcBudget.Selected(ilLoop) Then
                slNameCode = tgNameCode(ilLoop).sKey  'lbcBudget.List(ilIndex - 1)
                ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
                ilRet = gParseItem(slNameYear, 2, "/", slBdMnfName)
                ilRet = gParseItem(slNameYear, 1, "/", slYear)
                slYear = gSubStr("9999", slYear)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilCount = 0 Then
                    imBdMnf(0) = Val(slCode)
                    imBdYr(0) = Val(slYear)
                    ilCount = 1
                Else
                    If imBdYr(0) < Val(slYear) Then
                        imBdMnf(1) = Val(slCode)
                        imBdYr(1) = Val(slYear)
                    Else
                        imBdMnf(1) = imBdMnf(0)
                        imBdYr(1) = imBdYr(0)
                        imBdMnf(0) = Val(slCode)
                        imBdYr(0) = Val(slYear)
                    End If
                    Exit For
                End If
            End If
        Next ilLoop
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If imBdYr(0) = tgMCof(ilLoop).iYear Then
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), lmSplitDate
                lmSplitDate = lmSplitDate + 1
                Exit For
            End If
        Next ilLoop
        lmStartDateBd(1) = -1
        lmEndDateBd(1) = -1
        For ilLoop = 0 To UBound(imBdYr) Step 1
            If tgSpf.sRUseCorpCal <> "Y" Then
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearStartDate(0, slDate)
                lmStartDateBd(ilLoop) = gDateValue(slDate)
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearEndDate(0, slDate)
                lmEndDateBd(ilLoop) = gDateValue(slDate)
            Else
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearStartDate(5, slDate)
                lmStartDateBd(ilLoop) = gDateValue(slDate)
                slDate = "1/15/" & Trim$(str$(imBdYr(ilLoop)))
                slDate = gObtainYearEndDate(5, slDate)
                lmEndDateBd(ilLoop) = gDateValue(slDate)
            End If
        Next ilLoop
        If mReadBvfRec(hmBvf, imBdMnf(0), imBdYr(0), tmBvfVeh()) Then
            If mReadBvfRec(hmBvf, imBdMnf(1), imBdYr(1), tmBvfVeh2()) Then
                For ilLoop = LBound(tmBvfVeh2) To UBound(tmBvfVeh2) - 1 Step 1
                    ilFound = False
                    For ilBvf = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
                        If tmBvfVeh2(ilLoop).iVefCode = tmBvfVeh(ilBvf).iVefCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilBvf
                    If Not ilFound Then
                        tmBvfVeh(UBound(tmBvfVeh)) = tmBvfVeh2(ilLoop)
                        ReDim Preserve tmBvfVeh(LBound(tmBvfVeh) To UBound(tmBvfVeh) + 1) As BVF
                    End If
                Next ilLoop
                mGetBudgetDollars
            End If
        End If
    Else
        ReDim tmBvfVeh(0 To 0) As BVF
        ReDim tmBvfVeh2(0 To 0) As BVF
    End If
    mGetShowDates
    pbcImpact_Paint
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcImpact_Change()
    If imSettingValue Then
        pbcImpact.Cls
        pbcImpact_Paint
        imSettingValue = False
    Else
        pbcImpact.Cls
        pbcImpact_Paint
    End If
End Sub
Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show by"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Proposal Impact"
End Sub
