VERSION 5.00
Begin VB.Form LogChk 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4620
   ClientLeft      =   495
   ClientTop       =   2070
   ClientWidth     =   7845
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
   ScaleHeight     =   4620
   ScaleWidth      =   7845
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
      Left            =   1995
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   14
      Top             =   1785
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4290
      TabIndex        =   2
      Top             =   4215
      Width           =   945
   End
   Begin VB.Frame frcLog 
      Caption         =   "Check for"
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
      Height          =   780
      Left            =   90
      TabIndex        =   3
      Top             =   270
      Width           =   6840
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Inconsistent"
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
         Index           =   6
         Left            =   5355
         TabIndex        =   10
         Top             =   495
         Width           =   1380
      End
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Missed Spots"
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
         Index           =   5
         Left            =   3555
         TabIndex        =   9
         Top             =   495
         Width           =   1455
      End
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Hold Contracts"
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
         Index           =   4
         Left            =   1575
         TabIndex        =   8
         Top             =   495
         Width           =   1830
      End
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Reservation"
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
         Index           =   3
         Left            =   75
         TabIndex        =   7
         Top             =   495
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Missing Rotation"
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
         Left            =   3555
         TabIndex        =   6
         Top             =   255
         Value           =   1  'Checked
         Width           =   1740
      End
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Copy Not Assigned"
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
         Left            =   1575
         TabIndex        =   5
         Top             =   255
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox ckcCheck 
         Caption         =   "Unsold Avails"
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
         Left            =   75
         TabIndex        =   4
         Top             =   255
         Value           =   1  'Checked
         Width           =   1485
      End
   End
   Begin VB.PictureBox plcStatus 
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
      Height          =   210
      Left            =   195
      ScaleHeight     =   210
      ScaleWidth      =   7515
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3975
      Width           =   7515
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1140
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   2490
      TabIndex        =   1
      Top             =   4215
      Width           =   945
   End
   Begin VB.PictureBox plcChk 
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
      Height          =   2730
      Left            =   90
      ScaleHeight     =   2670
      ScaleWidth      =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7620
      Begin VB.ListBox lbcUnsold 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   30
         TabIndex        =   12
         Top             =   60
         Width           =   7485
      End
   End
   Begin VB.Timer tmcPrt 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7335
      Top             =   4200
   End
   Begin VB.Image imcPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7230
      Picture         =   "Logchk.frx":0000
      Top             =   285
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   285
      Top             =   4200
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "LogChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Logchk.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmGhfSrchKey0                 tmGsfSrchKey0                                           *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LogChk.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Log Check screen code
Option Explicit
Option Compare Text
'Btrieve files
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0     'VEF key 0 image
Dim imVefRecLen As Integer      'VEF record length
Dim hmVef As Integer            'Vehicle file handle
Dim tmVpf As VPF                'VPF record image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle preference file handle
Dim tmVlf As VLF                'VLF record image
Dim imVlfRecLen As Integer      'VLF record length
Dim hmVLF As Integer            'Vehicle Link file handle
'Unsold test
'Advertiser Name
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0     'ADF key 0 image
Dim imAdfRecLen As Integer      'ADF record length
Dim hmAdf As Integer            'Advertiser Name file handle
'Avail Name
Dim tmAnf As ANF                'ANF record image
Dim tmAnfSrchKey As INTKEY0     'ANF key 0 image
Dim imAnfRecLen As Integer      'ANF record length
Dim hmAnf As Integer            'Avail Name file handle
'Copy rotation
Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrf As CRF            'CRF record image
'Dim tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Dim tmCrfSrchKey4 As CRFKEY4 'CRF key record image
Dim imCrfRecLen As Integer     'CRF record length
' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer     'TZF record length
'Spot Detail
Dim tmSdf As SDF                'SDF record image
Dim tmSdfSrchKey3 As LONGKEY0     'SDF key 0 image
Dim imSdfRecLen As Integer      'SDF record length
'Region assign copy
Dim hmRsf As Integer              'Regional copy assignment
Dim tmRsf As RSF
Dim tmRsfSrchKey1 As LONGKEY0
Dim tmRsfSrchKey3 As RSFKEY3
Dim imRsfRecLen As Integer        'RAF record length

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf() As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

'Copy Air Game
Dim tmCaf As CAF            'CAF record image
Dim tmCafSrchKey As LONGKEY0  'CAF key record image
Dim tmCafSrchKey1 As CAFKEY1  'CAF key record image
Dim hmCaf As Integer        'CAF Handle
Dim imCafRecLen As Integer      'CAF record length

'Copy Vehicles
Dim tmCvf As CVF            'CVF record image
Dim tmCvfSrchKey As LONGKEY0  'CVF key record image
Dim tmCvfSrchKey1 As LONGKEY0  'CVF key record image
Dim hmCvf As Integer        'CVF Handle
Dim imCvfRecLen As Integer      'CVF record length

Dim hmCHF As Integer            'Spot file handle
Dim tmChf As CHF                'SDF record image
Dim tmChfSrchKey0 As LONGKEY0     'SDF key 0 image
Dim imCHFRecLen As Integer      'SDF record length
'Line
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmSdf As Integer            'Spot file handle
Dim hmSmf As Integer
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imNoZones As Integer
Dim tmRotNo(0 To 6) As COPYROTNO    'Index zero ignored
Dim tmLogTst() As LOGSEL
Dim smStatusCaption As String

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    tmcPrt.Enabled = False
    pbcPrinting.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Screen.MousePointer = vbHourglass
    mUnsold
    DoEvents
    Screen.MousePointer = vbDefault
    'If cmcCancel.Enabled Then
    '    cmcCancel.SetFocus
    'End If
End Sub
Private Sub cmcDone_GotFocus()
    tmcPrt.Enabled = False
    pbcPrinting.Visible = False
    gCtrlGotFocus ActiveControl
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
    LogChk.Refresh
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmLogTst
    Erase igSVefCode
    'Close btrieve files
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmCaf)
    btrDestroy hmCaf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmTzf)
    btrDestroy hmTzf
    ilRet = btrClose(hmRsf)
    btrDestroy hmRsf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    
    Set LogChk = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcPrt_Click()
    Dim ilCurrentLineNo As Integer
    Dim ilLinesPerPage As Integer
    Dim slRecord As String
    Dim slHeading As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    pbcPrinting.Visible = True
    DoEvents
    ilCurrentLineNo = 0
    ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    ilRet = 0
    On Error GoTo imcPrtErr:
    slHeading = "Counterpoint Log Check Information for " & Trim$(tgUrf(0).sRept) & " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    '6/6/16: Replaced GoSub
    'GoSub mHeading1
    mHeading1 ilRet, slHeading, ilCurrentLineNo
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    'Output Information
    For ilLoop = 0 To lbcUnsold.ListCount - 1 Step 1
        slRecord = "    " & lbcUnsold.List(ilLoop)
        '6/6/16: Replaced GoSub
        'GoSub mLineOutput
        mLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage
        If ilRet <> 0 Then
            Printer.EndDoc
            On Error GoTo 0
            pbcPrinting.Visible = False
            Exit Sub
        End If
    Next ilLoop
    Printer.EndDoc
    On Error GoTo 0
    'pbcPrinting.Visible = False
    tmcPrt.Enabled = True
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
imcPrtErr:
    ilRet = Err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gAssignCopyTest                 *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Assign copy test (taken from   *
'*                      gAssignCopyToSpots)            *
'*                                                     *
'*******************************************************
Private Function mAssignCopyTest(ilSSFType As Integer, ilVpfIndex As Integer, ilAsgnVefCode As Integer, ilRotNo As Integer, ilLnVefCode As Integer, slLive As String, ilGsf As Integer) As Integer
'
'   ilRet = gAssignCopyTest( slSsfType, ilVpfIndex)
'
'   Where:
'       slSsfType(I) "O"=On Air; "A" = Alternate
'       ilVpfIndex(I): Vehicle index into tgVpf
'       ilAsgnVefCode(O): Vehicle code that requires copy
'       ilRet(O)- 0=None Defined; 1=Copy defined but not assigned;
'                 2=Copy Assigned but superseded; 3=Zone copy Missing;
'                 4=Copy assignment Ok
'
'       tmSdf(I)
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llSpotTime As Long
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim ilAsgnDate0 As Integer
    Dim ilAsgnDate1 As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slType As String
    Dim ilDay As Integer
    Dim llAvailTime As Long
    Dim ilAvailOk As Integer
    Dim ilEvtIndex As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    Dim ilMatch As Integer
    Dim ilCrfVefCode As Integer
    Dim ilBypassCrf As Integer
    Dim ilVpf As Integer
    Dim ilVef As Integer
    Dim blVefFound As Boolean

    'Dim imNoZones As Integer
    'ReDim slZone(1 To 6) As String * 3
    imNoZones = 0
    For ilLoop = 1 To 6 Step 1
        tmRotNo(ilLoop).iRotNo = 0
        tmRotNo(ilLoop).sZone = ""
    Next ilLoop
    ilRotNo = -1
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
    llDate = gDateValue(slDate)
    ilAsgnDate0 = tmSdf.iDate(0)
    ilAsgnDate1 = tmSdf.iDate(1)
    ilDay = gWeekDayStr(slDate)
    'If (tmSsf.sType <> slSsfType) Or (tmSsf.iVefCode <> tmSdf.iVefCode) Or (tmSsf.iDate(0) <> ilAsgnDate0) Or (tmSsf.iDate(1) <> ilAsgnDate1) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
    '    tmSsfSrchKey.sType = slSsfType
    '    tmSsfSrchKey.iVefCode = tmSdf.iVefCode
    '    tmSsfSrchKey.iDate(0) = ilAsgnDate0
    '    tmSsfSrchKey.iDate(1) = ilAsgnDate1
    '    tmSsfSrchKey.iStartTime(0) = 0
    '    tmSsfSrchKey.iStartTime(1) = 0
    '    imSsfRecLen = Len(tmSsf)
    '    ilRet = btrGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
    'Else
        ilRet = BTRV_ERR_NONE
    'End If
    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
        ilEvtIndex = 1
        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
        llSpotTime = CLng(gTimeToCurrency(slTime, False)) ' - 1
        'Find rotation to assign
        'Code later- test spot type to determine which rotation type
        'ilCrfVefCode = gGetCrfVefCode(hmClf, tmSdf)
        ilCrfVefCode = ilAsgnVefCode
        ilAvailOk = True
        slType = "A"
        'tmCrfSrchKey1.sRotType = slType
        'tmCrfSrchKey1.iEtfCode = 0
        'tmCrfSrchKey1.iEnfCode = 0
        'tmCrfSrchKey1.iadfCode = tmSdf.iadfCode
        'tmCrfSrchKey1.lChfCode = tmSdf.lChfCode
        'tmCrfSrchKey1.lFsfCode = 0
        'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'tmSdf.iVefCode
        'tmCrfSrchKey1.iRotNo = 32000
        'ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get last current record to obtain date
        'Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iadfCode = tmSdf.iadfCode) And (tmCrf.lChfCode = tmSdf.lChfCode)
        tmCrfSrchKey4.sRotType = slType
        tmCrfSrchKey4.iEtfCode = 0
        tmCrfSrchKey4.iEnfCode = 0
        tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
        tmCrfSrchKey4.lChfCode = tmSdf.lChfCode
        tmCrfSrchKey4.lFsfCode = 0
        tmCrfSrchKey4.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) 'And (tmCrf.iVefCode = ilCrfVefCode)
            blVefFound = gCheckCrfVehicle(ilCrfVefCode, tmCrf, hmCvf)
            If blVefFound Then
                'Test date, time, day and zone
                ilBypassCrf = False
                'Test if looking for Live or Recorded rotations
                If tmCrf.sState <> "D" Then
                    If slLive = "L" Then
                        If tmCrf.sLiveCopy <> "L" Then
                            ilBypassCrf = True
                        End If
                    ElseIf slLive = "M" Then
                        If tmCrf.sLiveCopy <> "M" Then
                            ilBypassCrf = True
                        End If
                    ElseIf slLive = "S" Then
                        If tmCrf.sLiveCopy <> "S" Then
                            ilBypassCrf = True
                        End If
                    ElseIf slLive = "P" Then
                        If tmCrf.sLiveCopy <> "P" Then
                            ilBypassCrf = True
                        End If
                    ElseIf slLive = "Q" Then
                        If tmCrf.sLiveCopy <> "Q" Then
                            ilBypassCrf = True
                        End If
                    Else
                        If (tmCrf.sLiveCopy = "L") Or (tmCrf.sLiveCopy = "M") Or (tmCrf.sLiveCopy = "S") Or (tmCrf.sLiveCopy = "P") Or (tmCrf.sLiveCopy = "Q") Then
                            ilBypassCrf = True
                        End If
                    End If
                    If Not ilBypassCrf Then
                        If (tmCrf.lRafCode > 0) And (Trim$(tmCrf.sZone) = "R") Then
                            If (Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY Then
                                'ilVpf = gBinarySearchVpf(tmCrf.iVefCode)
                                'If ilVpf <> -1 Then
                                '    If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                '        ilBypassCrf = True
                                '    End If
                                'Else
                                '    ilBypassCrf = True
                                'End If
                                ilVef = gBinarySearchVef(tmCrf.iVefCode)
                                If ilVef <> -1 Then
                                    '5/11/11: Allow Selling to be set to No
                                    'If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Then
                                    If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "S") Then
                                        ilVpf = gBinarySearchVpf(tmCrf.iVefCode)
                                        If ilVpf <> -1 Then
                                            If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                                ilBypassCrf = True
                                            End If
                                        Else
                                            ilBypassCrf = True
                                        End If
                                    'ElseIf (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "P") Then
                                    ElseIf (tgMVef(ilVef).sType = "P") Then
                                        ilBypassCrf = False
                                    Else
                                        ilBypassCrf = True
                                    End If
                                Else
                                    ilBypassCrf = True
                                End If
                            End If
                        End If
                    End If
                Else
                    ilBypassCrf = True
                End If
                If (tmCrf.sDay(ilDay) = "Y") And (tmSdf.iLen = tmCrf.iLen) And (Not ilBypassCrf) Then
                    gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
                    llSAsgnDate = gDateValue(slDate)
                    gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                    llEAsgnDate = gDateValue(slDate)
                    If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
                        gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
                        llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
                        gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
                        llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
                        If tmCrf.sAirGameType = "G" Then
                            llSAsgnTime = 999999
                            llEAsgnTime = -1
                            If tmSsf.iType > 0 Then
                                tmCafSrchKey1.lCrfCode = tmCrf.lCode
                                ilRet = btrGetEqual(hmCaf, tmCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmCaf.lCrfCode = tmCrf.lCode)
                                    If tmCaf.sType = "G" Then
                                        If tmCaf.iGameNo = tmSsf.iType Then
                                            llSAsgnTime = llSpotTime
                                            llEAsgnTime = llSpotTime
                                            Exit Do
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmCaf, tmCaf, imCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                        ElseIf tmCrf.sAirGameType = "T" Then
                            llSAsgnTime = 999999
                            llEAsgnTime = -1
                            If tmSsf.iType > 0 Then
                                tmCafSrchKey1.lCrfCode = tmCrf.lCode
                                ilRet = btrGetEqual(hmCaf, tmCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmCaf.lCrfCode = tmCrf.lCode)
                                    If tmCaf.sType = "T" Then
                                        If (tmCaf.iTeamMnfCode = tmGsf(ilGsf).iHomeMnfCode) Or (tmCaf.iTeamMnfCode = tmGsf(ilGsf).iVisitMnfCode) Then
                                            llSAsgnTime = llSpotTime
                                            llEAsgnTime = llSpotTime
                                            Exit Do
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmCaf, tmCaf, imCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                        End If
                        If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
                            ilAvailOk = True    'Ok to book into
                            If (tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O") Then
                                ilEvtIndex = 1
                                Do
                                    If ilEvtIndex > tmSsf.iCount Then
                                        imSsfRecLen = Len(tmSsf)
                                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
                                            ilEvtIndex = 1
                                        Else
                                            mAssignCopyTest = 0
                                            Exit Function
                                        End If
                                    End If
                                    'Scan for avail that matches time of spot- then test avail name
                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvtIndex)
                                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                                        'Test time-
                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                        llAvailTime = CLng(gTimeToCurrency(slTime, False))
                                        If llSpotTime = llAvailTime Then
                                            If tmCrf.sInOut = "I" Then
                                                If tmCrf.ianfCode <> tmAvail.ianfCode Then
                                                    ilAvailOk = False   'No
                                                End If
                                            ElseIf tmCrf.sInOut = "O" Then
                                                If tmCrf.ianfCode = tmAvail.ianfCode Then
                                                    ilAvailOk = False   'No
                                                End If
                                            End If
                                            Exit Do
                                        ElseIf llSpotTime < llAvailTime Then
                                            'Spot missing from Ssf
                                            mAssignCopyTest = 0
                                            Exit Function
                                        End If
                                    End If
                                    ilEvtIndex = ilEvtIndex + 1
                                Loop
                            End If
                            If ilAvailOk Then
                                If ilRotNo = -1 Then
                                    ilRotNo = tmCrf.iRotNo
                                End If
                                If Trim$(tmCrf.sZone) <> "R" Then
                                    If Trim$(tmCrf.sZone) = "" Then
                                        tmRsfSrchKey1.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If tmRsf.iRotNo < tmCrf.iRotNo Then
                                                mAssignCopyTest = 2
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Else
                                    '7/15/14
                                    tmRsfSrchKey3.lSdfCode = tmSdf.lCode
                                    tmRsfSrchKey3.lRafCode = tmCrf.lRafCode
                                    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
                                        If tmRsf.lRafCode = tmCrf.lRafCode Then
                                            If tmRsf.iRotNo < tmCrf.iRotNo Then
                                                mAssignCopyTest = 2
                                                Exit Function
                                            Else
                                                ilAvailOk = False
                                            End If
                                        Else
                                            Exit Do
                                        End If
                                        ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                    If ilAvailOk Then
                                        mAssignCopyTest = 1
                                        Exit Function
                                    End If
                                End If
                                If (Trim$(tmCrf.sZone) = "") And ilAvailOk Then 'All zones
                                    imNoZones = imNoZones + 1
                                    tmRotNo(imNoZones).iRotNo = tmCrf.iRotNo
                                    tmRotNo(imNoZones).sZone = "Oth"
                                    'Add supersede test
                                    If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3") Then
                                        mAssignCopyTest = 4
                                        'Test if superseded
                                        If tmSdf.sPtType = "1" Then
                                            If imNoZones = 1 Then
                                                If tmRotNo(1).iRotNo > tmSdf.iRotNo Then
                                                    mAssignCopyTest = 2
                                                End If
                                            Else
                                                mAssignCopyTest = 3
                                            End If
                                        ElseIf tmSdf.sPtType = "2" Then
                                        Else    'Zones defined
                                            If imNoZones = 1 Then
                                                mAssignCopyTest = 2
                                            Else
                                                tmTzfSrchKey.lCode = tmSdf.lCopyCode
                                                ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    For ilLoop = 1 To imNoZones Step 1
                                                        ilMatch = False
                                                        For ilIndex = 1 To 6 Step 1
                                                            If tmTzf.lCifZone(ilIndex - 1) > 0 Then
                                                                If StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), Trim$(tmRotNo(ilLoop).sZone), 1) = 0 Then
                                                                    ilMatch = True
                                                                    If tmRotNo(ilLoop).iRotNo > tmTzf.iRotNo(ilIndex - 1) Then
                                                                        mAssignCopyTest = 2
                                                                        Exit Function
                                                                    End If
                                                                End If
                                                            End If
                                                        Next ilIndex
                                                        If Not ilMatch Then
                                                            mAssignCopyTest = 3
                                                        End If
                                                    Next ilLoop
                                                Else
                                                    mAssignCopyTest = 2
                                                End If
                                            End If
                                        End If
                                    Else
                                        mAssignCopyTest = 1
                                    End If
                                    Exit Function
                                End If
                                If ilAvailOk Then
                                    For ilLoop = 1 To 6 Step 1
                                        If Trim$(tmRotNo(ilLoop).sZone) = "" Then
                                            tmRotNo(ilLoop).iRotNo = tmCrf.iRotNo
                                            tmRotNo(ilLoop).sZone = tmCrf.sZone
                                            imNoZones = imNoZones + 1
                                            Exit For
                                        End If
                                        If StrComp(tmCrf.sZone, tmRotNo(ilLoop).sZone, 1) = 0 Then
                                            Exit For
                                        End If
                                    Next ilLoop
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If (imNoZones = 0) Or (ilVpfIndex = -1) Then
        mAssignCopyTest = 0
    Else
        'Test if all zones specified
        For ilIndex = LBound(tgVpf(ilVpfIndex).sGZone) To LBound(tgVpf(ilVpfIndex).sGZone) Step 1
            If Trim$(tgVpf(ilVpfIndex).sGZone(ilIndex)) <> "" Then
                ilFound = False
                For ilLoop = 1 To imNoZones Step 1
                    If StrComp(Trim$(tgVpf(ilVpfIndex).sGZone(ilIndex)), Trim$(tmRotNo(ilLoop).sZone), 1) = 0 Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3") Then
                        mAssignCopyTest = 3
                    Else
                        mAssignCopyTest = 1
                    End If
                    Exit Function
                End If
            End If
        Next ilIndex
        'Test if superseded
        If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3") Then
            mAssignCopyTest = 4
            If tmSdf.sPtType = "1" Then
                mAssignCopyTest = 2
            ElseIf tmSdf.sPtType = "2" Then
            Else
                tmTzfSrchKey.lCode = tmSdf.lCopyCode
                ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    For ilLoop = 1 To imNoZones Step 1
                        ilMatch = False
                        For ilIndex = 1 To 6 Step 1
                            If tmTzf.lCifZone(ilIndex - 1) > 0 Then
                                If StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), Trim$(tmRotNo(ilLoop).sZone), 1) = 0 Then
                                    ilMatch = True
                                    If tmRotNo(ilLoop).iRotNo > tmTzf.iRotNo(ilIndex - 1) Then
                                        mAssignCopyTest = 2
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next ilIndex
                        If Not ilMatch Then
                            mAssignCopyTest = 3
                        End If
                    Next ilLoop
                Else
                    mAssignCopyTest = 2
                End If
            End If
        Else
            mAssignCopyTest = 1
        End If
    End If
    Exit Function
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
    Dim ilLoop As Integer
    Dim slStr As String
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imcPrt.Picture = IconTraf!imcPrinter.Picture    'IconTraf!imcCamera.Picture
    LogChk.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone LogChk
    'Open btrieve files
    hmVef = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VEF.BTR)", LogChk
    On Error GoTo 0
    imVefRecLen = Len(tmVef)    'Save VEF record length
    hmVpf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VPF.BTR)", LogChk
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)    'Save VEF record length
    hmVLF = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VLF.BTR)", LogChk
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)    'Save VEF record length
    hmSdf = CBtrvTable(TWOHANDLES)          'Save VEF handle
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SDF.BTR)", LogChk
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)    'Save VEF record length
    hmCHF = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CHF.BTR)", LogChk
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)    'Save VEF record length
    hmClf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CLF.BTR)", LogChk
    On Error GoTo 0
    imClfRecLen = Len(tmClf)    'Save VEF record length
    hmAdf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.BTR)", LogChk
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)    'Save VEF record length
    hmAnf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ANF.BTR)", LogChk
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)    'Save VEF record length
    hmCrf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CRF.BTR)", LogChk
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)    'Save VEF record length
    hmTzf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: TZF.BTR)", LogChk
    On Error GoTo 0
    imTzfRecLen = Len(tmTzf)    'Save VEF record length
    hmRsf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RSF.BTR)", LogChk
    On Error GoTo 0
    imRsfRecLen = Len(tmRsf)    'Save VEF record length
    hmSsf = CBtrvTable(ONEHANDLE)          'Save STF handle
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SSF.BTR)", LogChk
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf)    'Save STF record length
    hmSmf = CBtrvTable(ONEHANDLE)          'Save STF handle
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SMF.BTR)", LogChk
    On Error GoTo 0
    hmGhf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: GHF.BTR)", LogChk
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)    'Save VEF record length
    hmGsf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: GSF.BTR)", LogChk
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))    'Save VEF record length
    hmCaf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmCaf, "", sgDBPath & "Caf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CAF.BTR)", LogChk
    On Error GoTo 0
    imCafRecLen = Len(tmCaf)    'Save VEF record length
    
    hmCvf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CVF.BTR)", LogChk
    On Error GoTo 0
    imCvfRecLen = Len(tmCvf)    'Save VEF record length
    
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    'LogChk.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    If ckcCheck(6).Value = vbChecked Then
        For ilLoop = LBound(tgChkMsg) To UBound(tgChkMsg) - 1 Step 1
            slStr = Trim$(tgChkMsg(ilLoop).sVehicle)
            Select Case tgChkMsg(ilLoop).iStatus
                Case 1
                    slStr = slStr & ", No Selling to Airing Links"
                Case 2
                    slStr = slStr & ", No Dates in Future"
                Case 3
                    slStr = slStr & "(Log), No Conventional"
                Case 4
                    slStr = slStr & "(Log), Conventional No Dates in Future"
                Case 5
                    slStr = slStr & ", No Dates in Future"
            End Select
            lbcUnsold.AddItem slStr
        Next ilLoop
    End If
    If tgSpf.sCDefLogCopy = "Y" Then
        ckcCheck(1).Value = vbUnchecked
    Else
        ckcCheck(1).Value = vbChecked
    End If
    pbcPrinting.Move (LogChk.Width - pbcPrinting.Width) / 2, (LogChk.Height - pbcPrinting.Height) / 2
    ' Dan M 9-23-09 adjust look of 'wait' message
    gAdjustScreenMessage Me, pbcPrinting

'    gCenterModalForm LogChk
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMissedTest                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain missed spots for         *
'*                     specified dates                 *
'*                                                     *
'*******************************************************
Private Function mMissedTest(slType As String, ilVefCode As Integer, llStartDate As Long, llEndDate As Long) As Integer
'
'   ilRet = mMissedTest (slType, ilVefCode, ilAdfCode, slStartDate, slEndDate, ilSortOrder, lbcSortCtrl, tlSdfMdExt())
'   Where:
'       slType(I)- "M"=Missed; "C"=Cancelled; "H"=Hidden
'       ilVefCode(I)- Vehicle code
'       slStartDate(I)- Start date
'       slEndDate(I)- End date
'       ilRet (O)- True if missed obtained OK; False if error
'
    Dim hlSdf As Integer        'Sdf handle
    Dim ilSdfRecLen As Integer     'Record length
    Dim tlSdf As SDF
    Dim tlSdfSrchKey2 As SDFKEY2
    Dim ilRet As Integer
    Dim llDate As Long
    mMissedTest = False
    hlSdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlSdf)
        btrDestroy hlSdf
        Exit Function
    End If
    ilSdfRecLen = Len(tlSdf) 'btrRecordLength(hlSdf)  'Get and save record length
    tlSdfSrchKey2.iVefCode = ilVefCode
    tlSdfSrchKey2.sSchStatus = slType
    tlSdfSrchKey2.iAdfCode = 0
    gPackDateLong llStartDate, tlSdfSrchKey2.iDate(0), tlSdfSrchKey2.iDate(1)
    tlSdfSrchKey2.iTime(0) = 0
    tlSdfSrchKey2.iTime(1) = 0
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
    Do While (ilRet = BTRV_ERR_NONE) And (tlSdf.iVefCode = ilVefCode) And (tlSdf.sSchStatus = slType)
        gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
            ilRet = btrClose(hlSdf)
            btrDestroy hlSdf
            mMissedTest = True
            Exit Function
        End If
        ilRet = btrGetNext(hlSdf, tlSdf, ilSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlSdf)
    btrDestroy hlSdf
    Exit Function
End Function
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload LogChk
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnsold                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:4/27/94       By:               *
'*                                                     *
'*            Comments:Test for unsold avails          *
'*                                                     *
'*      1-14-05 Incorrect avail time showing for message
'*          "No rotation defined"                      *
'*******************************************************
Private Sub mUnsold()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

    Dim ilAvLen As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilEvt As Integer
    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilAssign As Integer
    Dim slTime As String
    Dim llDate As Long
    Dim ilSpot As Integer
    Dim ilVpfIndex As Integer
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim llLen As Long
    Dim ilAsgnVefCode As Integer
    Dim ilPkgRet As Integer
    Dim ilPkgVefCode As Integer
    Dim ilSchPkgVefCode As Integer
    Dim ilRotNo As Integer
    Dim ilPkgRotNo As Integer
    Dim ilLnVefCode As Integer
    Dim ilLnRotNo As Integer
    Dim ilLnRet As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    Dim tlVef As VEF
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilTst As Integer
    Dim ilLoop As Integer
    Dim ilLink As Integer
    Dim slAdvtName As String
    Dim slStartDate As String
    Dim ilGameNo As Integer
    Dim ilGsf As Integer
    Dim llGsfDate As Long
    ReDim tlBBSdf(0 To 0) As SDF

    llLen = 0
    ReDim tmLogTst(0 To 0) As LOGSEL
    lbcUnsold.Clear
    If ckcCheck(6).Value = vbChecked Then
        For ilLoop = LBound(tgChkMsg) To UBound(tgChkMsg) - 1 Step 1
            slStr = Trim$(tgChkMsg(ilLoop).sVehicle)
            Select Case tgChkMsg(ilLoop).iStatus
                Case 1
                    slStr = slStr & ", No Selling to Airing Links"
                Case 2
                    slStr = slStr & ", No Dates in Future"
                Case 3
                    slStr = slStr & "(Log), No Conventional"
                Case 4
                    slStr = slStr & "(Log), Conventional No Dates in Future"
                Case 5
                    slStr = slStr & ", No Dates in Future"
            End Select
            lbcUnsold.AddItem slStr
        Next ilLoop
    End If
    For ilLoop = LBound(tgSel) To UBound(tgSel) - 1 Step 1
        If (tgSel(ilLoop).iChk = 1) And (tgSel(ilLoop).iStatus = 0) Then
            ilVefCode = tgSel(ilLoop).iVefCode
            tmVefSrchKey.iCode = ilVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If tmVef.sType = "L" Then
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                        ilFound = False
                        For ilTst = 0 To UBound(tmLogTst) - 1 Step 1
                            If tmLogTst(ilTst).iVefCode = tgMVef(ilVef).iCode Then
                                If tgSel(ilLoop).lStartDate < tmLogTst(ilTst).lStartDate Then
                                    tmLogTst(ilTst).lStartDate = tgSel(ilLoop).lStartDate
                                End If
                                If tgSel(ilLoop).lEndDate > tmLogTst(ilTst).lEndDate Then
                                    tmLogTst(ilTst).lEndDate = tgSel(ilLoop).lEndDate
                                End If
                                ilFound = True
                                Exit For
                            End If
                        Next ilTst
                        If Not ilFound Then
                            tmLogTst(UBound(tmLogTst)) = tgSel(ilLoop)
                            tmLogTst(UBound(tmLogTst)).iVefCode = tgMVef(ilVef).iCode
                            If bgLogFirstCallToVpfFind Then
                                tmLogTst(UBound(tmLogTst)).iVpfIndex = gVpfFind(LogChk, tmLogTst(UBound(tmLogTst)).iVefCode)
                                bgLogFirstCallToVpfFind = False
                            Else
                                tmLogTst(UBound(tmLogTst)).iVpfIndex = gVpfFindIndex(tmLogTst(UBound(tmLogTst)).iVefCode)
                            End If
                            ReDim Preserve tmLogTst(0 To UBound(tmLogTst) + 1) As LOGSEL
                        End If
                    End If
                Next ilVef
            ElseIf tmVef.sType = "A" Then
                slStartDate = Format$(tgSel(ilLoop).lStartDate, "m/d/yy")
                gBuildLinkArray hmVLF, tmVef, slStartDate, igSVefCode()
                'For ilLink = LBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) To UBound(tgVpf(tgSel(ilLoop).iVpfIndex).iGLink) Step 1
                '    If tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) > 0 Then
                For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                        ilFound = False
                        For ilTst = 0 To UBound(tmLogTst) - 1 Step 1
                            'If tmLogTst(ilTst).iVefCode = tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink) Then
                            If tmLogTst(ilTst).iVefCode = igSVefCode(ilLink) Then
                                If tgSel(ilLoop).lStartDate < tmLogTst(ilTst).lStartDate Then
                                    tmLogTst(ilTst).lStartDate = tgSel(ilLoop).lStartDate
                                End If
                                If tgSel(ilLoop).lEndDate > tmLogTst(ilTst).lEndDate Then
                                    tmLogTst(ilTst).lEndDate = tgSel(ilLoop).lEndDate
                                End If
                                ilFound = True
                                Exit For
                            End If
                        Next ilTst
                        If Not ilFound Then
                            tmLogTst(UBound(tmLogTst)) = tgSel(ilLoop)
                            tmLogTst(UBound(tmLogTst)).iVefCode = igSVefCode(ilLink)    'tgVpf(tgSel(ilLoop).iVpfIndex).iGLink(ilLink)
                            If bgLogFirstCallToVpfFind Then
                                tmLogTst(UBound(tmLogTst)).iVpfIndex = gVpfFind(LogChk, tmLogTst(UBound(tmLogTst)).iVefCode)
                                bgLogFirstCallToVpfFind = False
                            Else
                                tmLogTst(UBound(tmLogTst)).iVpfIndex = gVpfFindIndex(tmLogTst(UBound(tmLogTst)).iVefCode)
                            End If
                            ReDim Preserve tmLogTst(0 To UBound(tmLogTst) + 1) As LOGSEL
                        End If
                    'End If
                Next ilLink
            Else
                ilFound = False
                For ilTst = 0 To UBound(tmLogTst) - 1 Step 1
                    If tmLogTst(ilTst).iVefCode = tgSel(ilLoop).iVefCode Then
                        If tgSel(ilLoop).lStartDate < tmLogTst(ilTst).lStartDate Then
                            tmLogTst(ilTst).lStartDate = tgSel(ilLoop).lStartDate
                        End If
                        If tgSel(ilLoop).lEndDate > tmLogTst(ilTst).lEndDate Then
                            tmLogTst(ilTst).lEndDate = tgSel(ilLoop).lEndDate
                        End If
                        ilFound = True
                        Exit For
                    End If
                Next ilTst
                If Not ilFound Then
                    tmLogTst(UBound(tmLogTst)) = tgSel(ilLoop)
                    ReDim Preserve tmLogTst(0 To UBound(tmLogTst) + 1) As LOGSEL
                End If
            End If
        End If
    Next ilLoop
    For ilVef = LBound(tmLogTst) To UBound(tmLogTst) - 1 Step 1
        ilVefCode = tmLogTst(ilVef).iVefCode
        tmVefSrchKey.iCode = ilVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        smStatusCaption = "Checking " & Trim$(tmVef.sName)
        plcStatus.Cls
        plcStatus_Paint
        ilVpfIndex = tmLogTst(ilVef).iVpfIndex  'gVpfFind(Logs, ilVefCode)
        ReDim tmGsf(0 To 1) As GSF
        tmGsf(0).iGameNo = 0
        If tmVef.sType = "G" Then
            ilRet = mGhfGsfReadRec(ilVefCode, tmLogTst(ilVef).lStartDate, tmLogTst(ilVef).lEndDate)
        End If
        ilType = 0
        If (tgSpf.sUsingBBs = "Y") And ((ckcCheck(1).Value = vbChecked) Or (ckcCheck(2).Value = vbChecked)) Then
            ilRet = gMakeBBAndAssignCopy(hmSdf, hmVLF, ilVefCode, tmLogTst(ilVef).lStartDate, tmLogTst(ilVef).lEndDate)
        End If
        For llDate = tmLogTst(ilVef).lStartDate To tmLogTst(ilVef).lEndDate Step 1
            For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
                ilType = tmGsf(ilGsf).iGameNo
                If ilType <> 0 Then
                    gUnpackDateLong tmGsf(ilGsf).iAirDate(0), tmGsf(ilGsf).iAirDate(1), llGsfDate
                Else
                    llGsfDate = llDate
                End If
                If llDate = llGsfDate Then
                    ilGameNo = ilType
                    slDate = Format$(llDate, "m/d/yy")
                    gPackDate slDate, ilDate0, ilDate1

                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    tmSsfSrchKey.iType = ilType
                    tmSsfSrchKey.iVefCode = ilVefCode
                    tmSsfSrchKey.iDate(0) = ilDate0
                    tmSsfSrchKey.iDate(1) = ilDate1
                    tmSsfSrchKey.iStartTime(0) = 0
                    tmSsfSrchKey.iStartTime(1) = 0
                    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1)
                        ilEvt = 1
                        Do While ilEvt <= tmSsf.iCount
                           LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                            ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                ilAvLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime    '1-14-05

                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                        'If ckcCheck(1).Value Then
                                        '    If (tmSdf.sPtType <> "1") And (tmSdf.sPtType <> "2") And (tmSdf.sPtType <> "3") Then
                                        '        tmChfSrchKey0.lCode = tmSdf.lchfCode
                                        '        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        '        If ilRet = BTRV_ERR_NONE Then
                                        '            If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Copy Missing: " & slDate & " at " & slTime & " " & Trim$(Str$(tmChf.lCntrNo)), llLen, True) Then
                                        '                Exit Sub
                                        '            End If
                                        '            lbcUnsold.AddItem Trim$(tmVef.sName) & " Copy Missing: " & slDate & " at " & slTime & " Contract #" & Str$(tmChf.lCntrNo)
                                        '        Else
                                        '            If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Copy Missing: " & slDate & " at " & slTime, llLen, True) Then
                                        '                Exit Sub
                                        '            End If
                                        '            lbcUnsold.AddItem Trim$(tmVef.sName) & " Copy Missing: " & slDate & " at " & slTime
                                        '        End If
                                        '    End If
                                        'End If
                                        If (ckcCheck(1).Value = vbChecked) Or (ckcCheck(2).Value = vbChecked) Then
                                            ilSchPkgVefCode = 0
                                            ilRet = gGetCrfVefCode(hmClf, tmSdf, ilAsgnVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                            If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSpotType = "X") Then
                                                slStr = gGetMGCopyAssign(tmSdf, ilPkgVefCode, ilLnVefCode, slLive, hmSmf, hmCrf)
                                                If (slStr = "S") Or (slStr = "B") Then
                                                    ilSchPkgVefCode = gGetMGPkgVefCode(hmClf, tmSdf)
                                                End If
                                                If slStr = "O" Then
                                                    ilAsgnVefCode = ilLnVefCode
                                                    ilLnVefCode = 0
                                                ElseIf slStr = "S" Then
                                                    ilPkgVefCode = ilSchPkgVefCode
                                                    ilSchPkgVefCode = 0
                                                    ilLnVefCode = 0
                                                Else
                                                    If ilPkgVefCode = ilSchPkgVefCode Then
                                                        ilSchPkgVefCode = 0
                                                    End If
                                                End If
                                            Else
                                                ilLnVefCode = 0
                                            End If
                                            ilAssign = mAssignCopyTest(ilType, ilVpfIndex, ilAsgnVefCode, ilRotNo, ilLnVefCode, slLive, ilGsf)
                                            If ilPkgVefCode > 0 Then
                                                ilPkgRet = mAssignCopyTest(ilType, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilLnVefCode, slLive, ilGsf)
                                                If (ilAssign <> 0) And (ilPkgRet <> 0) Then
                                                    If ilPkgRotNo > ilRotNo Then
                                                        ilRotNo = ilPkgRotNo
                                                        ilAssign = ilPkgRet
                                                        ilAsgnVefCode = ilPkgVefCode
                                                    End If
                                                ElseIf (ilAssign = 0) And (ilPkgRet = 0) Then
                                                    ilRotNo = ilPkgRotNo
                                                    ilAssign = ilPkgRet
                                                    ilAsgnVefCode = ilPkgVefCode
                                                ElseIf (ilAssign = 0) And (ilPkgRet <> 0) Then
                                                    ilRotNo = ilPkgRotNo
                                                    ilAssign = ilPkgRet
                                                    ilAsgnVefCode = ilPkgVefCode
                                                End If
                                            End If
                                            If ilSchPkgVefCode > 0 Then
                                                ilPkgVefCode = ilSchPkgVefCode
                                                ilPkgRet = mAssignCopyTest(ilType, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilLnVefCode, slLive, ilGsf)
                                                If (ilAssign <> 0) And (ilPkgRet <> 0) Then
                                                    If ilPkgRotNo > ilRotNo Then
                                                        ilRotNo = ilPkgRotNo
                                                        ilAssign = ilPkgRet
                                                        ilAsgnVefCode = ilPkgVefCode
                                                    End If
                                                ElseIf (ilAssign = 0) And (ilPkgRet = 0) Then
                                                    ilRotNo = ilPkgRotNo
                                                    ilAssign = ilPkgRet
                                                    ilAsgnVefCode = ilPkgVefCode
                                                ElseIf (ilAssign = 0) And (ilPkgRet <> 0) Then
                                                    ilRotNo = ilPkgRotNo
                                                    ilAssign = ilPkgRet
                                                    ilAsgnVefCode = ilPkgVefCode
                                                End If
                                            End If
                                            If (ilAsgnVefCode <> ilLnVefCode) And (ilLnVefCode > 0) Then
                                                ilLnRet = mAssignCopyTest(ilType, ilVpfIndex, ilLnVefCode, ilLnRotNo, ilLnVefCode, slLive, ilGsf)
                                                If (ilAssign <> 0) And (ilLnRet <> 0) Then
                                                    If ilLnRotNo > ilRotNo Then
                                                        ilRotNo = ilLnRotNo
                                                        ilAssign = ilLnRet
                                                        ilAsgnVefCode = ilLnVefCode
                                                    End If
                                                ElseIf (ilAssign = 0) And (ilLnRet <> 0) Then
                                                    ilRotNo = ilLnRotNo
                                                    ilAssign = ilLnRet
                                                    ilAsgnVefCode = ilLnVefCode
                                                End If
                                            End If
                                            If (ilAssign <> 4) Then
                                                If (ilAssign = 0) Or (ilAssign = 3) Or (ckcCheck(1).Value = vbChecked) Then
                                                    If ilAsgnVefCode <> tmVef.iCode Then
                                                        tmVefSrchKey.iCode = ilAsgnVefCode
                                                        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        slStr = Trim$(tlVef.sName)
                                                    Else
                                                        slStr = Trim$(tmVef.sName)
                                                    End If
                                                    Select Case ilAssign
                                                        Case 0
                                                            slStr = slStr & " No Rotation Defined: "
                                                        Case 1
                                                            slStr = slStr & " Copy Not Assigned: "
                                                        Case 2
                                                            slStr = slStr & " Copy Superseded: "
                                                        Case 3
                                                            slStr = slStr & " Time Zone Missing: "
                                                    End Select
                                                    tmChfSrchKey0.lCode = tmSdf.lChfCode
                                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        If tmAdf.iCode <> tmChf.iAdfCode Then
                                                            tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                tmAdf.iCode = 0
                                                                tmAdf.sName = ""
                                                            End If
                                                        End If
                                                        'slAdvtName = Trim$(tmAdf.sName)
                                                        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                                            slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                                        Else
                                                            slAdvtName = Trim$(tmAdf.sName)
                                                        End If
                                                        If Not gOkAddStrToListBox(slStr & slDate & " at " & slTime & " #" & Trim$(str$(tmChf.lCntrNo)) & " " & slAdvtName, llLen, True) Then
                                                            Exit Sub
                                                        End If
                                                        lbcUnsold.AddItem slStr & slDate & " at " & slTime & " #" & str$(tmChf.lCntrNo) & " " & slAdvtName
                                                    Else
                                                        If Not gOkAddStrToListBox(slStr & slDate & " at " & slTime, llLen, True) Then
                                                            Exit Sub
                                                        End If
                                                        lbcUnsold.AddItem slStr & slDate & " at " & slTime
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If (ckcCheck(3).Value = vbChecked) Then
                                            If (tmSpot.iRank And RANKMASK) = RESERVATIONRANK Then
                                                tmChfSrchKey0.lCode = tmSdf.lChfCode
                                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    If tmAdf.iCode <> tmChf.iAdfCode Then
                                                        tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            tmAdf.iCode = 0
                                                            tmAdf.sName = ""
                                                        End If
                                                    End If
                                                    'slAdvtName = Trim$(tmAdf.sName)
                                                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                                        slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                                    Else
                                                        slAdvtName = Trim$(tmAdf.sName)
                                                    End If
                                                    If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Reservation: " & slDate & " at " & slTime & " #" & Trim$(str$(tmChf.lCntrNo)) & " " & slAdvtName, llLen, True) Then
                                                        Exit Sub
                                                    End If
                                                    lbcUnsold.AddItem Trim$(tmVef.sName) & " Reservation: " & slDate & " at " & slTime & " #" & str$(tmChf.lCntrNo) & " " & slAdvtName
                                                Else
                                                    If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Reservation: " & slDate & " at " & slTime, llLen, True) Then
                                                        Exit Sub
                                                    End If
                                                    lbcUnsold.AddItem Trim$(tmVef.sName) & " Reservation: " & slDate & " at " & slTime
                                                End If
                                            End If
                                        End If
                                        If (ckcCheck(4).Value = vbChecked) Then
                                            tmChfSrchKey0.lCode = tmSdf.lChfCode
                                            If tmChf.lCode <> tmSdf.lChfCode Then
                                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Else
                                                ilRet = BTRV_ERR_NONE
                                            End If
                                            If (ilRet = BTRV_ERR_NONE) And (tmChf.sStatus = "H") Then
                                                If tmAdf.iCode <> tmChf.iAdfCode Then
                                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet <> BTRV_ERR_NONE Then
                                                        tmAdf.iCode = 0
                                                        tmAdf.sName = ""
                                                    End If
                                                End If
                                                'slAdvtName = Trim$(tmAdf.sName)
                                                If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                                    slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                                Else
                                                    slAdvtName = Trim$(tmAdf.sName)
                                                End If
                                                If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Hold: " & slDate & " at " & slTime & " #" & Trim$(str$(tmChf.lCntrNo)) & " " & slAdvtName, llLen, True) Then
                                                    Exit Sub
                                                End If
                                                lbcUnsold.AddItem Trim$(tmVef.sName) & " Hold: " & slDate & " at " & slTime & " #" & str$(tmChf.lCntrNo) & " " & slAdvtName
                                            End If
                                        End If
                                    End If
                                    ilLen = tmSpot.iPosLen And &HFFF
                                    If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                        If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                            ilUnits = ilUnits - 1
                                            ilAvLen = ilAvLen - ilLen
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                            ilUnits = ilUnits - 1
                                            ilAvLen = ilAvLen - ilLen
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                            ilUnits = ilUnits - 1
                                            ilAvLen = ilAvLen - ilLen
                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                        End If
                                    End If
                                Next ilSpot
                                If (ckcCheck(0).Value = vbChecked) Then
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                        If (ilUnits > 0) And (ilAvLen > 0) Then
                                            tmAnfSrchKey.iCode = tmAvail.ianfCode
                                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            'gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime       get the time of avail when first accessing it, not here
                                            If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Unsold Avail: " & slDate & " at " & slTime & " " & Trim$(tmAnf.sName), llLen, True) Then
                                                Exit Sub
                                            End If
                                            lbcUnsold.AddItem Trim$(tmVef.sName) & " Unsold Avail: " & slDate & " at " & slTime & " " & Trim$(tmAnf.sName)
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                        If (ilUnits > 0) And (ilAvLen > 0) Then
                                            tmAnfSrchKey.iCode = tmAvail.ianfCode
                                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            'gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime        get the time of avail when first accessing it, not here
                                            If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Unsold Avail: " & slDate & " at " & slTime & " " & Trim$(tmAnf.sName), llLen, True) Then
                                                Exit Sub
                                            End If
                                            lbcUnsold.AddItem Trim$(tmVef.sName) & " Unsold Avail: " & slDate & " at " & slTime & " " & Trim$(tmAnf.sName)
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                        If (ilUnits > 0) And (ilAvLen > 0) Then
                                            tmAnfSrchKey.iCode = tmAvail.ianfCode
                                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            'gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime        get the time of avail when first accessing it, not here
                                            If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Unsold Avail: " & slDate & " at " & slTime & " " & Trim$(tmAnf.sName), llLen, True) Then
                                                Exit Sub
                                            End If
                                            lbcUnsold.AddItem Trim$(tmVef.sName) & " Unsold Avail: " & slDate & " at " & slTime & " " & Trim$(tmAnf.sName)
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                    End If
                                End If
                                ilEvt = ilEvt + tmAvail.iNoSpotsThis    'bypass spots
                            End If
                            ilEvt = ilEvt + 1   'Increment to next event
                        Loop
                        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tgSpf.sUsingBBs = "Y") And ((ckcCheck(1).Value = vbChecked) Or (ckcCheck(2).Value = vbChecked)) Then
                        ReDim tlBBSdf(0 To 0) As SDF
                        ilRet = gGetBBSpots(hmSdf, ilVefCode, ilGameNo, slDate, tlBBSdf())
                        For ilEvt = 0 To UBound(tlBBSdf) - 1 Step 1
                            tmSdf = tlBBSdf(ilEvt)
                            If (tmSdf.sPtType <> "1") And (tmSdf.sPtType <> "2") And (tmSdf.sPtType <> "3") Then
                                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime    '1-14-05
                                slStr = Trim$(tmVef.sName)
                                slStr = slStr & " No BB Rotation Defined: "
                                tmChfSrchKey0.lCode = tmSdf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    If tmAdf.iCode <> tmChf.iAdfCode Then
                                        tmAdfSrchKey.iCode = tmChf.iAdfCode
                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tmAdf.iCode = 0
                                            tmAdf.sName = ""
                                        End If
                                    End If
                                    'slAdvtName = Trim$(tmAdf.sName)
                                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                        slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                    Else
                                        slAdvtName = Trim$(tmAdf.sName)
                                    End If
                                    If Not gOkAddStrToListBox(slStr & slDate & " at " & slTime & " #" & Trim$(str$(tmChf.lCntrNo)) & " " & slAdvtName, llLen, True) Then
                                        Exit Sub
                                    End If
                                    lbcUnsold.AddItem slStr & slDate & " at " & slTime & " #" & str$(tmChf.lCntrNo) & " " & slAdvtName
                                Else
                                    If Not gOkAddStrToListBox(slStr & slDate & " at " & slTime, llLen, True) Then
                                        Exit Sub
                                    End If
                                    lbcUnsold.AddItem slStr & slDate & " at " & slTime
                                End If
                            End If
                        Next ilEvt
                    End If
                End If
            Next ilGsf
        Next llDate
        If ckcCheck(5).Value = vbChecked Then
            If mMissedTest("M", ilVefCode, tmLogTst(ilVef).lStartDate, tmLogTst(ilVef).lEndDate) Then
                If Not gOkAddStrToListBox(Trim$(tmVef.sName) & " Missed Spots between " & Format$(tmLogTst(ilVef).lStartDate, "m/d/yy") & " - " & Format$(tmLogTst(ilVef).lEndDate, "m/d/yy"), llLen, True) Then
                    Exit Sub
                End If
                lbcUnsold.AddItem Trim$(tmVef.sName) & " Missed Spots between " & Format$(tmLogTst(ilVef).lStartDate, "m/d/yy") & " - " & Format$(tmLogTst(ilVef).lEndDate, "m/d/yy")
            End If
        End If
    Next ilVef
    Erase tlBBSdf
    smStatusCaption = "Done Checking"
    plcStatus.Cls
    plcStatus_Paint
End Sub
Private Sub pbcPrinting_Paint()
    pbcPrinting.CurrentX = (pbcPrinting.Width - pbcPrinting.TextWidth("Printing Log Check Information....")) / 2
    pbcPrinting.CurrentY = (pbcPrinting.Height - pbcPrinting.TextHeight("Printing Log Check Information....")) / 2 - 30
    pbcPrinting.Print "Printing Log Check Information...."
End Sub

Private Sub plcStatus_Paint()
    plcStatus.CurrentX = 0
    plcStatus.CurrentY = 0
    plcStatus.Print smStatusCaption
End Sub

Private Sub tmcPrt_Timer()
    tmcPrt.Enabled = False
    pbcPrinting.Visible = False
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Log Check"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGhfGsfReadRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mGhfGsfReadRec(ilVefCode As Integer, llStartDate As Long, llEndDate As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mGhfGsfReadRecErr                                                                     *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim llDate As Long

    ReDim tmGsf(0 To 0) As GSF
    ilUpper = 0
    tmGhfSrchKey1.iVefCode = ilVefCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        tmGsfSrchKey1.lGhfCode = tmGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf(ilUpper), imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf(ilUpper).lGhfCode)
            gUnpackDateLong tmGsf(ilUpper).iAirDate(0), tmGsf(ilUpper).iAirDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                ReDim Preserve tmGsf(0 To UBound(tmGsf) + 1) As GSF
                ilUpper = UBound(tmGsf)
            End If
            ilRet = btrGetNext(hmGsf, tmGsf(ilUpper), imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    Else
        mGhfGsfReadRec = False
        Exit Function
    End If
    mGhfGsfReadRec = True
    Exit Function
mGhfGsfReadRecErr: 'VBC NR
    On Error GoTo 0
    mGhfGsfReadRec = False
    Exit Function
End Function


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
