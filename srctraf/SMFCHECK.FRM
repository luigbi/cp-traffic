VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SMFCheck 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5505
   ClientLeft      =   75
   ClientTop       =   1470
   ClientWidth     =   8865
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
   ScaleHeight     =   5505
   ScaleWidth      =   8865
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   195
      Left            =   855
      TabIndex        =   28
      Top             =   1020
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame frcCreate 
      Caption         =   "Create SMF"
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   4830
      TabIndex        =   5
      Top             =   105
      Visible         =   0   'False
      Width           =   3900
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
         Height          =   1080
         Left            =   450
         TabIndex        =   25
         Top             =   2205
         Width           =   3225
      End
      Begin VB.TextBox edcTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1365
         TabIndex        =   14
         Top             =   3735
         Width           =   2085
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1365
         TabIndex        =   12
         Top             =   3390
         Width           =   2085
      End
      Begin VB.CommandButton cmcUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update"
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
         Left            =   1350
         TabIndex        =   15
         Top             =   4125
         Width           =   1050
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Original Airing Vehicle"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   26
         Top             =   1965
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Spot Vehicle:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   6
         Left            =   180
         TabIndex        =   24
         Top             =   1650
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Line Vehicle:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   23
         Top             =   1410
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Missed Time:"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   180
         TabIndex        =   13
         Top             =   3780
         Width           =   1335
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Missed Date:"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   180
         TabIndex        =   11
         Top             =   3420
         Width           =   1200
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Air Time:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   10
         Top             =   1185
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Air Date:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   945
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Line #:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Contract #:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   495
         Width           =   3600
      End
      Begin VB.Label lacCreate 
         Appearance      =   0  'Flat
         Caption         =   "Advertiser:"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   3600
      End
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
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1320
      TabIndex        =   20
      Top             =   0
      Width           =   1320
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6555
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5100
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6720
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6930
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5145
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcErrors 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1530
      Width           =   4470
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   7545
      Top             =   4995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton cmcCheck 
      Appearance      =   0  'Flat
      Caption         =   "&Check/Fix"
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
      Left            =   2865
      TabIndex        =   18
      Top             =   4965
      Width           =   1830
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
      Left            =   4965
      TabIndex        =   0
      Top             =   4965
      Width           =   1050
   End
   Begin VB.Label lacDBPath 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      TabIndex        =   27
      Top             =   5265
      Width           =   6045
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Check SMFCheck Results using Messages Menu"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   660
      TabIndex        =   16
      Top             =   4650
      Visible         =   0   'False
      Width           =   7530
   End
   Begin VB.Label lacCntr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   135
      TabIndex        =   22
      Top             =   4290
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   210
      Top             =   4740
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   855
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Label lacCntr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   1020
      TabIndex        =   3
      Top             =   660
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label lacCntr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Processing"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   420
      Width           =   3720
   End
End
Attribute VB_Name = "SMFCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SMFCHECK.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SMFCheck.Frm
'
' Release: 1.0
'
' Description:
'   This file check the Linking of SMF and SDF
Option Explicit
Option Compare Text
Dim lmTotalNo As Long
Dim lmProcessedNo As Long
Dim hmTo As Integer   'From file hanle
Dim hmSdf As Integer    'file handle
Dim imSdfRecLen As Integer  'Record length
Dim tmSdf As SDF
Dim tmSdfInfo() As SDFINFO
Dim tmSdfSrchKey3 As LONGKEY0
Dim hmSmf As Integer    'file handle
Dim imSmfRecLen As Integer  'Record length
Dim tmSmf As SMF
Dim tmSmfSrchKey1 As LONGKEY0
Dim tmSmfMake() As SMFMAKE
Dim tmDelSmf() As DELSMF
Dim hmCHF As Integer    'file handle
Dim imCHFRecLen As Integer  'Record length
Dim tmChf As CHF
Dim tmChfSrchKey As LONGKEY0
Dim hmClf As Integer    'file handle
Dim imClfRecLen As Integer  'Record length
Dim tmClf As CLF
Dim tmClfSrchKey As CLFKEY0
'Advertiser
Dim hmAdf As Integer            'Advertiser file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
'Vehicle
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
'Virtual Vehicle
Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
'Game schedule

Dim hmGsf As Integer
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length
Dim tmGsf As GSF

Dim hmGhf As Integer
Dim tmGhfSrchKey0 As LONGKEY0    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length
Dim tmGhf As GHF

'Dim tmRec As LPOPREC
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imListIndex As Integer
Dim imFromList As Integer
Dim imFromUpdate As Integer
Dim imChgFlag As Integer
Dim lmPercent As Long
Dim imFirstActivate As Integer

Private Sub cmcCancel_Click()
    If imConverting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
End Sub
Private Sub cmcCheck_Click()
    Dim ilRet As Integer
    If imConverting Then
        Exit Sub
    End If
    lbcErrors.Clear
    lbcErrors.Visible = True
    plcGauge.Visible = True
    lacMsg.Visible = True
    frcCreate.Visible = False
    imListIndex = -1
    ilRet = 0
    'On Error GoTo cmcSetErr:
    'hmTo = FreeFile
    ''Open "SMFCheck.Txt" For Output As hmTo
    'Open sgDBPath & "Messages\" & "SMFCheck.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "SMFCheck.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & "SMFCheck.Txt" & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    On Error GoTo 0
    Print #hmTo, "Checking SMF on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    imConverting = True
    lmTotalNo = mNoRecords("Smf.Btr")
    lmTotalNo = lmTotalNo + mNoRecords("Sdf.Btr")
    lmProcessedNo = 0
    lacCntr(0).Visible = True
    lacCntr(1).Visible = True
    lacCntr(2).Visible = True
    lacCntr(0).Caption = ""
    lacCntr(1).Caption = ""
    lacCount.Caption = ""
    lacCount.Visible = True
    plcGauge.Value = 0
    lmPercent = 0
    DoEvents
    If Not mSMFCheck() Then
        Print #hmTo, "Checking SMF terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close hmTo
        Screen.MousePointer = vbDefault
        imConverting = False
        If imTerminate Then
            cmcCancel_Click
        End If
        Exit Sub
    End If
    lacCount.Visible = False
    Print #hmTo, "Checking SMF successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close hmTo
    If lbcErrors.ListCount > 0 Then
        frcCreate.Visible = True
        lbcVehicle.ZOrder vbBringToFront
        DoEvents
        lbcErrors.ListIndex = 0
    Else
        cmcCancel.Caption = "Done"
    End If
    lacCntr(0).Visible = False
    lacCntr(1).Visible = False
    plcGauge.Value = 100
    imConverting = False
    Screen.MousePointer = vbDefault
    Exit Sub
'cmcSetErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
Private Sub cmcCheck_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
End Sub
Private Sub cmcUpdate_Click()
    Dim slSchDate As String
    Dim slSchTime As String
    Dim slMissedDate As String
    Dim slMissedTime As String
    Dim ilOrigVefCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilListIndex As Integer
    Dim llRecPos As Long
    Dim ilCRet As Integer
    If imListIndex < 0 Then
        Exit Sub
    End If
    slMissedDate = edcDate.Text
    If Not gValidDate(slMissedDate) Then
        Beep
        edcDate.SetFocus
        Exit Sub
    End If
    slMissedTime = edcTime.Text
    If Not gValidTime(slMissedTime) Then
        Beep
        edcTime.SetFocus
        Exit Sub
    End If
    If lbcVehicle.ListIndex < 0 Then
        Beep
        lbcVehicle.SetFocus
        Exit Sub
    End If
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slSchTime
    ilListIndex = lbcVehicle.ListIndex
    slNameCode = tgVehicle(ilListIndex).sKey   'Traffic!lbcVehicle.List(ilListIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilOrigVefCode = Val(slCode)
    mMakeSmf hmSmf, tmSmf, tmSdf.sSchStatus, tmSdf, ilOrigVefCode, slMissedDate, slMissedTime, slSchDate, slSchTime
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    ilRet = btrGetPosition(hmSmf, tmSmfMake(imListIndex).lRecPos)
    tmSdfSrchKey3.lCode = tmSmfMake(imListIndex).lSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If tmSdf.lSmfCode <> tmSmf.lCode Then
        ilRet = btrGetPosition(hmSdf, llRecPos)
        Do
            tmSdf.lSmfCode = tmSmf.lCode
            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
            If ilRet = BTRV_ERR_CONFLICT Then
                'ilCRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY3, BTRV_LOCK_NONE)
                tmSdfSrchKey3.lCode = tmSmfMake(imListIndex).lSdfCode
                ilCRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    If Not imFromList Then
        If imListIndex + 1 >= lbcErrors.ListCount Then
            cmcCancel.Caption = "Done"
            cmcCancel.SetFocus
            Exit Sub
        Else
            imFromUpdate = True
            lbcErrors.ListIndex = imListIndex + 1
            imFromUpdate = False
            edcDate.SetFocus
        End If
    End If
    imFromList = False
    imChgFlag = False
End Sub
Private Sub edcDate_Change()
    If (Trim$(edcDate.Text) <> "") And (Trim$(edcTime.Text) <> "") And (lbcVehicle.ListIndex >= 0) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    imChgFlag = True
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    'igParentRestarted = True
End Sub
Private Sub edcTime_Change()
    If (Trim$(edcDate.Text) <> "") And (Trim$(edcTime.Text) <> "") And (lbcVehicle.ListIndex >= 0) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    imChgFlag = True
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
    Me.Refresh
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
    'sgDoneMsg = CmdStr
    'igChildDone = True
    'Cancel = 0
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
    
    Erase tmSdfInfo
    Erase tmSmfMake
    Erase tmDelSmf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    
    Set SMFCheck = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcErrors_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    If (imListIndex >= 0) And (cmcUpdate.Enabled) And (Not imFromUpdate) And (imChgFlag) Then
        ilRet = MsgBox("SMF not Saved, Save", vbOKCancel + vbQuestion + vbApplicationModal, "Save SMF")
        If ilRet = vbOK Then
            imFromList = True
            cmcUpdate_Click
            imFromList = False
        End If
    End If
    imFromUpdate = False
    imListIndex = lbcErrors.ListIndex
    tmSdfSrchKey3.lCode = tmSmfMake(imListIndex).lSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
    If mGetChfClfAdf(tmSdf.lChfCode, tmSdf.iLineNo) Then
        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
            lacCreate(0).Caption = "Advertiser: " & Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
        Else
            lacCreate(0).Caption = "Advertiser: " & Trim$(tmAdf.sName)
        End If
        lacCreate(1).Caption = "Contract #:" & Str$(tmChf.lCntrNo)
        lacCreate(2).Caption = "Line #:" & Str$(tmClf.iLine)
        lacCreate(3).Caption = "Air Date: " & slDate
        lacCreate(4).Caption = "Air Time: " & slTime
        tmVefSrchKey.iCode = tmClf.iVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        lacCreate(5).Caption = "Line Vehicle: " & Trim$(tmVef.sName) & " Type: " & tmVef.sType
    Else
        lacCreate(0).Caption = "Advertiser: Missing"
        lacCreate(1).Caption = "Contract #: Missing"
        lacCreate(2).Caption = "Line #: Missing"
        lacCreate(3).Caption = "Air Date: " & slDate
        lacCreate(4).Caption = "Air Time: " & slTime
        lacCreate(5).Caption = "Line Vehicle: Missing"
    End If
    tmVefSrchKey.iCode = tmSdf.iVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    lacCreate(6).Caption = "Spot Vehicle: " & Trim$(tmVef.sName) & " Type: " & tmVef.sType
    If tmSmfMake(imListIndex).lRecPos > 0 Then
        ilRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, tmSmfMake(imListIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
        edcDate.Text = slDate
        gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slTime
        edcTime.Text = slTime
        lbcVehicle.ListIndex = -1
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tmSmf.iOrigSchVef = Val(slCode) Then
                lbcVehicle.ListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    Else
        edcDate.Text = ""
        edcTime.Text = ""
        tmSmf.lSdfCode = 0
        tmSmf.lChfCode = 0
        lbcVehicle.ListIndex = -1
        For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tmSdf.iVefCode = Val(slCode) Then
                lbcVehicle.ListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    cmcUpdate.Enabled = False
    imChgFlag = False
End Sub
Private Sub lbcVehicle_Click()
    If (Trim$(edcDate.Text) <> "") And (Trim$(edcTime.Text) <> "") And (lbcVehicle.ListIndex >= 0) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    imChgFlag = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetChfClfAdf                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get record for contract, line   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Private Function mGetChfClfAdf(llChfCode As Long, ilLineNo As Integer) As Integer
    Dim ilRet As Integer
    tmChfSrchKey.lCode = llChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        tmClfSrchKey.lChfCode = tmChf.lCode 'Can't use tmSmf.lChfCode as this only points to latest version header
                                            'but line might belong to another version
        tmClfSrchKey.iLine = ilLineNo
        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSmf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F")) 'And (tmClf.sSchStatus = "A")
            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (ilLineNo = tmClf.iLine) Then
            tmAdfSrchKey.iCode = tmChf.iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                mGetChfClfAdf = False
                Exit Function
            End If
        Else
            mGetChfClfAdf = False
            Exit Function
        End If
    Else
        mGetChfClfAdf = False
        Exit Function
    End If
    mGetChfClfAdf = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetRecLength                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the record length from   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Private Function mGetRecLength(slFileName As String) As Integer
'
'   ilRecLen = mGetRecLength(slName)
'   Where:
'       slName (I)- Name of the file
'       ilRecLen (O)- record length within the file
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mGetRecLength = -ilRet
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    mGetRecLength = btrRecordLength(hlFile)  'Get and save record length
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
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
    imFirstActivate = True
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If (Len(sgMDBPath) > 0) And (Len(sgSDBPath) > 0) Then
        lacDBPath.Caption = "D.B. Path=> " & sgMDBPath & " and " & sgSDBPath
    Else
        lacDBPath.Caption = "D.B. Path=> " & sgDBPath
    End If
    imConverting = False
    imFirstFocus = True
    lmTotalNo = 0
    lmProcessedNo = 0
    imListIndex = -1
    imFromList = False
    imFromUpdate = False
    imChgFlag = False
    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
    hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SMFCheck
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)
    'plcGauge.Move SMFCheck.Width / 2 - plcGauge.Width / 2
    'cmcSet.Move SMFCheck.Width / 2 - cmcSet.Width / 2
    'cmcCancel.Move SMFCheck.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    gCenterStdAlone SMFCheck
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        imTerminate = True
    End If
    lbcVehicle.Clear
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeSmf                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create a Spot MG Spec from     *
'*                      missed spot (copied gMakeSmf)  *
'*                                                     *
'*******************************************************
Private Sub mMakeSmf(hlSmf As Integer, tlSmf As SMF, slSchStatus As String, tlSdf As SDF, ilOrigVefCode As Integer, slMissedDate As String, slMissedTime As String, slSchDate As String, slSchTime As String)
'
'   gMakeSmf hlSmf, tlSmf, slSchStatus, tlSdf, slMissedDate, slMissedTime, slSchDate, slSchTime
'   Where:
'       hlSmf(I)- SMF open handle
'       tlSmf(I)- old SMF image if existed (tlSmf.lChfCode = 0 => didn't exist)
'                 old image must have been deleted previously
'       slSchStatus(I)- "G" for makgood, "O' for outside contract limits
'       llChfCode(I)- Contract code number
'       ilLineNo(I)- Contract line number
'       slMissedDate(I)- original spot date (missed or scheduled)
'       slMissedTime(I)- original spot time (missed or scheduled)
'       slSchDate(I)- schedule date
'       slSchTime(I)- schedule time
'
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    If (slSchStatus <> "G") And (slSchStatus <> "O") Then
        Exit Sub
    End If
    'Determine if one required to be created or updated
    'If (tlSmf.lChfCode = tlSdf.lChfCode) And (tlSmf.iLineNo = tlSdf.iLineNo) Then
    'Contract test left because old code set lChfCode = 0 prior to adding lSdfCode
    If (tlSmf.lSdfCode = tlSdf.lCode) And (tlSmf.lChfCode = tlSdf.lChfCode) And (tlSmf.iLineNo = tlSdf.iLineNo) Then
        'Test if spot is within old specification- if so, retain specifications
        gPackDate slMissedDate, tlSmf.iMissedDate(0), tlSmf.iMissedDate(1)
        gPackTime slMissedTime, tlSmf.iMissedTime(0), tlSmf.iMissedTime(1)
        gPackDate slSchDate, tlSmf.iStartDate(0), tlSmf.iStartDate(1)
        gPackDate slSchDate, tlSmf.iEndDate(0), tlSmf.iEndDate(1)
        gPackDate slSchDate, tlSmf.iActualDate(0), tlSmf.iActualDate(1)
        gPackTime slSchTime, tlSmf.iStartTime(0), tlSmf.iStartTime(1)
        gPackTime slSchTime, tlSmf.iEndTime(0), tlSmf.iEndTime(1)
        gPackTime slSchTime, tlSmf.iActualTime(0), tlSmf.iActualTime(1)
        For ilLoop = 1 To 7
            tlSmf.sWkDays(ilLoop - 1) = "N" 'Weekday flag: Y=day allowed; N=Day disallowed
        Next ilLoop
        'tlSmf.sWkDays(gWeekDayStr(slSchDate) + 1) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
        tlSmf.sWkDays(gWeekDayStr(slSchDate)) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
        If tlSdf.iVefCode <> ilOrigVefCode Then
            tlSmf.sMGSource = "A"     'A=Create by mouse move to a different vehicle
        Else
            tlSmf.sMGSource = "M"     'S=Created by MG Spec, M=Created by mouse move outside contract
        End If
        tlSmf.sSchStatus = slSchStatus    'G=Makegood; O=Outside contract limits
        tlSmf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
        tlSmf.iOrigSchVef = ilOrigVefCode
        ilRet = btrUpdate(hlSmf, tlSmf, imSmfRecLen)
        Exit Sub
    End If
    'Create SMF record
    imSmfRecLen = Len(tlSmf)  'Get and save SMF record length
    tlSmf.lCode = 0
    tlSmf.lSdfCode = tlSdf.lCode
    tlSmf.lChfCode = tlSdf.lChfCode        'Contract code
    tlSmf.iLineNo = tlSdf.iLineNo      'Line number
    tlSmf.lFsfCode = 0
    gPackDate slMissedDate, tlSmf.iMissedDate(0), tlSmf.iMissedDate(1)
    gPackTime slMissedTime, tlSmf.iMissedTime(0), tlSmf.iMissedTime(1)
    gPackDate slSchDate, tlSmf.iStartDate(0), tlSmf.iStartDate(1)
    gPackDate slSchDate, tlSmf.iEndDate(0), tlSmf.iEndDate(1)
    gPackDate slSchDate, tlSmf.iActualDate(0), tlSmf.iActualDate(1)
    gPackTime slSchTime, tlSmf.iStartTime(0), tlSmf.iStartTime(1)
    gPackTime slSchTime, tlSmf.iEndTime(0), tlSmf.iEndTime(1)
    gPackTime slSchTime, tlSmf.iActualTime(0), tlSmf.iActualTime(1)
    For ilLoop = 1 To 7
        tlSmf.sWkDays(ilLoop - 1) = "N" 'Weekday flag: Y=day allowed; N=Day disallowed
    Next ilLoop
    'tlSmf.sWkDays(gWeekDayStr(slSchDate) + 1) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
    tlSmf.sWkDays(gWeekDayStr(slSchDate)) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
    If tlSdf.iVefCode <> ilOrigVefCode Then
        tlSmf.sMGSource = "A"     'A=Create by mouse move to a different vehicle
    Else
        tlSmf.sMGSource = "M"     'S=Created by MG Spec, M=Created by mouse move outside contract
    End If
    tlSmf.sSchStatus = slSchStatus    'G=Makegood; O=Outside contract limits
    tlSmf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
    tlSmf.iOrigSchVef = ilOrigVefCode
    ilRet = btrInsert(hlSmf, tlSmf, imSmfRecLen, INDEXKEY0)
    On Error GoTo mMakeSMFErr
    gBtrvErrorMsg ilRet, "mMakeSMF (btrInsert)", SMFCheck
    On Error GoTo 0
    Exit Sub
mMakeSMFErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mNoRecords                      *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the number of records    *
'*                                                     *
'*******************************************************
Private Function mNoRecords(slFileName As String) As Long
    Dim hlFile As Integer
    Dim ilRet As Integer
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mNoRecords = 0
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    mNoRecords = btrRecords(hlFile)  'Get and save record length
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
End Function
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
    slCommand = sgCommandStr    'Command$
    'ilPos = InStr(1, slCommand, "/CS", 1)
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
    '    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    ilPos = InStr(1, slCommand, "/S:Test", 1)
    '    If ilPos > 0 Then
    '        ilTestSystem = True
    '    End If
    '    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    '    If ilPos > 0 Then
    '        ilTestSystem = False
    '    End If
    '    slCommand = "Traffic\Guide\" & Trim$(Str$(CALLNONE))
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
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
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone SMFCheck, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSMFCheck                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Check SMF                      *
'*                                                     *
'*******************************************************
Private Function mSMFCheck() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilFound As Integer
    Dim llPercent As Long
    Dim llNoRecs As Long
    Dim llNoProcs As Long
    Dim llRecPos As Long
    Dim ilCRet As Integer
    Dim ilUpper As Integer
    Dim slMGDate As String
    Dim slMGTime As String
    Dim slMissedDate As String
    Dim slMissedTime As String
    Dim ilDeleted As Integer
    Dim ilStatusChg As Integer
    Dim ilStartStatusCount As Integer
    Dim ilSmfMissing As Integer
    Dim llMissedDate As Long
    Dim llSeasonStartDate As Long
    Dim llSeasonEndDate As Long
    Dim blSeasonFd As Boolean
    ReDim tmSdfInfo(0 To 0) As SDFINFO
    ReDim tmDelSmf(0 To 0) As DELSMF
    ilUpper = UBound(tmSdfInfo)
    llNoRecs = btrRecords(hmSdf)
    llNoProcs = 0
    lacCount.Caption = Trim$(Str$(llNoProcs)) & " of" & Str$(llNoRecs)
    DoEvents
    If imTerminate Then
        mTerminate
        mSMFCheck = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    lacCntr(0).Caption = "Testing SDF Records"
    'Gather all O and G spots to compare to smf
    ilRet = btrGetFirst(hmSdf, tmSdf, imSdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imTerminate Then
            Screen.MousePointer = vbDefault
            mTerminate
            mSMFCheck = False
            Exit Function
        End If
        If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
            tmSmfSrchKey1.lCode = tmSdf.lSmfCode
            ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If (ilRet <> BTRV_ERR_NONE) Or (tmSmf.lSdfCode <> tmSdf.lCode) Then
                ilUpper = UBound(tmSdfInfo)
                tmSdfInfo(ilUpper).lCode = tmSdf.lCode
                tmSdfInfo(ilUpper).lChfCode = tmSdf.lChfCode
                tmSdfInfo(ilUpper).iLineNo = tmSdf.iLineNo
                tmSdfInfo(ilUpper).iDate(0) = tmSdf.iDate(0)
                tmSdfInfo(ilUpper).iDate(1) = tmSdf.iDate(1)
                tmSdfInfo(ilUpper).iTime(0) = tmSdf.iTime(0)
                tmSdfInfo(ilUpper).iTime(1) = tmSdf.iTime(1)
                tmSdfInfo(ilUpper).sSchStatus = tmSdf.sSchStatus
                tmSdfInfo(ilUpper).lSmfCode = tmSdf.lSmfCode
                tmSdfInfo(ilUpper).iMatchFound = False
                ilUpper = ilUpper + 1
                ReDim Preserve tmSdfInfo(0 To ilUpper) As SDFINFO
            End If
        End If
        llNoProcs = llNoProcs + 1
        lacCount.Caption = Trim$(Str$(llNoProcs)) & " of" & Str$(llNoRecs)
        lmProcessedNo = lmProcessedNo + 1
        llPercent = (lmProcessedNo * CSng(100)) / lmTotalNo
        If llPercent >= 100 Then
            llPercent = 100
        End If
        If lmPercent <> llPercent Then
            plcGauge.Value = llPercent
            lmPercent = llPercent
        End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    lacCntr(0).Caption = "Test SMF Records"
    ilDeleted = 0
    ilStatusChg = 0
    ilSmfMissing = 0
    llNoRecs = btrRecords(hmSmf)
    llNoProcs = 0
    lacCount.Caption = Trim$(Str$(llNoProcs)) & " of" & Str$(llNoRecs)
    DoEvents
    If imTerminate Then
        mTerminate
        mSMFCheck = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    ilRet = btrGetFirst(hmSmf, tmSmf, imSmfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If imTerminate Then
            Screen.MousePointer = vbDefault
            mTerminate
            mSMFCheck = False
            Exit Function
        End If
        If (tmSmf.sSchStatus = "O") Or (tmSmf.sSchStatus = "G") Then
            ilFound = False
            tmSdfSrchKey3.lCode = tmSmf.lSdfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            '6/28/12: If spot status is not MG or Outside, remove spot
            If (ilRet = BTRV_ERR_NONE) And (tmSdf.lSmfCode = tmSmf.lCode) Then
                If (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                    If tmSdf.lSmfCode > 0 Then
                        tmSdf.lSmfCode = 0
                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                    End If
                    ilRet = -1
                End If
            End If
            If (ilRet = BTRV_ERR_NONE) And (tmSdf.lSmfCode = tmSmf.lCode) Then
                ilFound = True
                'Adjust missed date and time if game buy
                If tmSmf.iGameNo > 0 Then
                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                    tmGsfSrchKey3.iVefCode = tmSmf.iOrigSchVef
                    tmGsfSrchKey3.iGameNo = tmSmf.iGameNo
                    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If (ilRet = BTRV_ERR_NONE) Then
                        ilRet = btrGetPosition(hmSmf, llRecPos)
                        On Error GoTo mSMFCheckErr
                        gBtrvErrorMsg ilRet, "mSMFCheck (btrGetPosition)", SMFCheck
                        On Error GoTo 0
                        blSeasonFd = False
                        tmGhfSrchKey0.lCode = tmGsf.lghfcode
                        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
                        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tmSmf.iOrigSchVef) And (tmGsf.iGameNo = tmSmf.iGameNo)
                            gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStartDate
                            gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEndDate
                            If (llMissedDate >= llSeasonStartDate) And (llMissedDate <= llSeasonEndDate) Then
                                blSeasonFd = True
                                Do
                                    If (tmGsf.iAirDate(0) <> tmSmf.iMissedDate(0)) Or (tmGsf.iAirDate(1) <> tmSmf.iMissedDate(1)) Then
                                        tmSmf.iMissedDate(0) = tmGsf.iAirDate(0)
                                        tmSmf.iMissedDate(1) = tmGsf.iAirDate(1)
                                    End If
                                    tmSmf.iMissedTime(0) = tmGsf.iAirTime(0)
                                    tmSmf.iMissedTime(1) = tmGsf.iAirTime(1)
                                    ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                                    If ilRet = BTRV_ERR_CONFLICT Then
                                        ilCRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                                        On Error GoTo mSMFCheckErr
                                        gBtrvErrorMsg ilCRet, "mSMFCheck (btrGetDirect)", SMFCheck
                                        On Error GoTo 0
                                    End If
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                On Error GoTo mSMFCheckErr
                                gBtrvErrorMsg ilRet, "mSMFCheck (btrUpdate)", SMFCheck
                                On Error GoTo 0
                                tmSmfSrchKey1.lCode = tmSmf.lCode
                                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                Exit Do
                            Else
                                ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    Exit Do
                                End If
                                tmGhfSrchKey0.lCode = tmGsf.lghfcode
                                ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
                            End If
                        Loop
                        If Not blSeasonFd Then
                            gUnpackDate tmSmf.iActualDate(0), tmSmf.iActualDate(1), slMGDate
                            gUnpackTime tmSmf.iActualTime(0), tmSmf.iActualTime(1), "A", "1", slMGTime
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
                            gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
                            If mGetChfClfAdf(tmSmf.lChfCode, tmSmf.iLineNo) Then
                                Print #hmTo, "Season Not Found, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & " Game # " & tmSmf.iGameNo
                            Else
                                Print #hmTo, "Season Not Found, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & " Game # " & tmSmf.iGameNo
                            End If
                        End If
                    End If
                End If
                ilStartStatusCount = ilStatusChg
                If (tmSdf.sSchStatus <> tmSmf.sSchStatus) Then
                    gUnpackDate tmSmf.iActualDate(0), tmSmf.iActualDate(1), slMGDate
                    gUnpackTime tmSmf.iActualTime(0), tmSmf.iActualTime(1), "A", "1", slMGTime
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
                    gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
                    If mGetChfClfAdf(tmSmf.lChfCode, tmSmf.iLineNo) Then
                        Print #hmTo, "Schedule Status Not Matching, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Updated"
                        'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                    Else
                        Print #hmTo, "Schedule Status Not Matching, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Updated"
                        'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                    End If
                    ilRet = btrGetPosition(hmSmf, llRecPos)
                    On Error GoTo mSMFCheckErr
                    gBtrvErrorMsg ilRet, "mSMFCheck (btrGetPosition)", SMFCheck
                    On Error GoTo 0
                    Do
                        'tmRec = tmSmf
                        'ilRet = gGetByKeyForUpdate("SMF", hmSmf, tmRec)
                        'tmSmf = tmRec
                        'On Error GoTo mSMFCheckErr
                        'gBtrvErrorMsg ilRet, "mSMFCheck (Get by Key)", SMFCheck
                        'On Error GoTo 0
                        tmSmf.sSchStatus = tmSdf.sSchStatus
                        ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            ilCRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                            On Error GoTo mSMFCheckErr
                            gBtrvErrorMsg ilCRet, "mSMFCheck (btrGetDirect)", SMFCheck
                            On Error GoTo 0
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    On Error GoTo mSMFCheckErr
                    gBtrvErrorMsg ilRet, "mSMFCheck (btrUpdate)", SMFCheck
                    On Error GoTo 0
                    ilStatusChg = ilStatusChg + 1
                    tmSmfSrchKey1.lCode = tmSmf.lCode
                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                End If
                If (tmSdf.lChfCode <> tmSmf.lChfCode) Then
                    gUnpackDate tmSmf.iActualDate(0), tmSmf.iActualDate(1), slMGDate
                    gUnpackTime tmSmf.iActualTime(0), tmSmf.iActualTime(1), "A", "1", slMGTime
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
                    gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
                    If mGetChfClfAdf(tmSdf.lChfCode, tmSdf.iLineNo) Then
                        Print #hmTo, "Contract Code Not Matching, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Updated"
                        'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                    Else
                        Print #hmTo, "Contract Code Not Matching, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Updated"
                        'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                    End If
                    ilRet = btrGetPosition(hmSmf, llRecPos)
                    On Error GoTo mSMFCheckErr
                    gBtrvErrorMsg ilRet, "mSMFCheck (btrGetPosition)", SMFCheck
                    On Error GoTo 0
                    Do
                        'tmRec = tmSmf
                        'ilRet = gGetByKeyForUpdate("SMF", hmSmf, tmRec)
                        'tmSmf = tmRec
                        'On Error GoTo mSMFCheckErr
                        'gBtrvErrorMsg ilRet, "mSMFCheck (Get by Key)", SMFCheck
                        'On Error GoTo 0
                        tmSmf.lChfCode = tmSdf.lChfCode
                        ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            ilCRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                            On Error GoTo mSMFCheckErr
                            gBtrvErrorMsg ilCRet, "mSMFCheck (btrGetDirect)", SMFCheck
                            On Error GoTo 0
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    On Error GoTo mSMFCheckErr
                    gBtrvErrorMsg ilRet, "mSMFCheck (btrUpdate)", SMFCheck
                    On Error GoTo 0
                    If ilStartStatusCount = ilStatusChg Then
                        ilStatusChg = ilStatusChg + 1
                    End If
                    tmSmfSrchKey1.lCode = tmSmf.lCode
                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                End If
                If (tmSdf.iLineNo <> tmSmf.iLineNo) Then
                    gUnpackDate tmSmf.iActualDate(0), tmSmf.iActualDate(1), slMGDate
                    gUnpackTime tmSmf.iActualTime(0), tmSmf.iActualTime(1), "A", "1", slMGTime
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
                    gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
                    If mGetChfClfAdf(tmSdf.lChfCode, tmSdf.iLineNo) Then
                        Print #hmTo, "Line #s Not Matching, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Updated"
                        'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                    Else
                        Print #hmTo, "Line #s Not Matching, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Updated"
                        'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                    End If
                    ilRet = btrGetPosition(hmSmf, llRecPos)
                    On Error GoTo mSMFCheckErr
                    gBtrvErrorMsg ilRet, "mSMFCheck (btrGetPosition)", SMFCheck
                    On Error GoTo 0
                    Do
                        'tmRec = tmSmf
                        'ilRet = gGetByKeyForUpdate("SMF", hmSmf, tmRec)
                        'tmSmf = tmRec
                        'On Error GoTo mSMFCheckErr
                        'gBtrvErrorMsg ilRet, "mSMFCheck (Get by Key)", SMFCheck
                        'On Error GoTo 0
                        tmSmf.iLineNo = tmSdf.iLineNo
                        ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            ilCRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                            On Error GoTo mSMFCheckErr
                            gBtrvErrorMsg ilCRet, "mSMFCheck (btrGetDirect)", SMFCheck
                            On Error GoTo 0
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    On Error GoTo mSMFCheckErr
                    gBtrvErrorMsg ilRet, "mSMFCheck (btrUpdate)", SMFCheck
                    On Error GoTo 0
                    If ilStartStatusCount = ilStatusChg Then
                        ilStatusChg = ilStatusChg + 1
                    End If
                    tmSmfSrchKey1.lCode = tmSmf.lCode
                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                End If
            End If
            If Not ilFound Then
                gUnpackDate tmSmf.iActualDate(0), tmSmf.iActualDate(1), slMGDate
                gUnpackTime tmSmf.iActualTime(0), tmSmf.iActualTime(1), "A", "1", slMGTime
                gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
                gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
                If mGetChfClfAdf(tmSmf.lChfCode, tmSmf.iLineNo) Then
                    Print #hmTo, "SdfSmfCode <> SmfCode Not Matching, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Removed"
                    'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                Else
                    Print #hmTo, "SdfSmfCode <> SmfCode Not Matching, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Removed"
                    'lbcErrors.AddItem "Schedule Status Not Matching, SMF Created"
                End If
                ilRet = btrGetPosition(hmSmf, llRecPos)
                tmDelSmf(UBound(tmDelSmf)).lRecPos = llRecPos
                tmDelSmf(UBound(tmDelSmf)).iReason = 2
                ReDim Preserve tmDelSmf(0 To UBound(tmDelSmf) + 1) As DELSMF
            End If
        End If
        llNoProcs = llNoProcs + 1
        lacCount.Caption = Trim$(Str$(llNoProcs)) & " of" & Str$(llNoRecs)
        lmProcessedNo = lmProcessedNo + 1
        llPercent = (lmProcessedNo * CSng(100)) / lmTotalNo
        If llPercent >= 100 Then
            llPercent = 100
        End If
        If lmPercent <> llPercent Then
            plcGauge.Value = llPercent
            lmPercent = llPercent
        End If
        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    'Remove records to be deleted
    For ilLoop = 0 To UBound(tmDelSmf) - 1 Step 1
        ilRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, tmDelSmf(ilLoop).lRecPos, INDEXKEY1, BTRV_LOCK_NONE)
        On Error GoTo mSMFCheckErr
        gBtrvErrorMsg ilRet, "mSMFCheck (btrGetDirect)", SMFCheck
        On Error GoTo 0
        gUnpackDate tmSmf.iActualDate(0), tmSmf.iActualDate(1), slMGDate
        gUnpackTime tmSmf.iActualTime(0), tmSmf.iActualTime(1), "A", "1", slMGTime
        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
        gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
        If mGetChfClfAdf(tmSmf.lChfCode, tmSmf.iLineNo) Then
            If tmDelSmf(ilLoop).iReason = 1 Then
                Print #hmTo, "Multi Smf found for Sdf, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Removed"
            Else
                Print #hmTo, "SDF Not Found For SMF, Contract #" & Str$(tmChf.lCntrNo) & " Line #" & Str$(tmSmf.iLineNo) & "MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & ", SMF Removed"
            End If
            'lbcErrors.AddItem "SDF Not Found For SMF, SMF Removed"
        Else
            If tmDelSmf(ilLoop).iReason = 1 Then
                Print #hmTo, "Multi Smf found for Sdf, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & " SMF Removed"
            Else
                Print #hmTo, "SDF Not Found For SMF, Can't find Contract #/Line/Advt" & " MG Date/Time " & slMGDate & " " & slMGTime & " Missed Date/Time " & slMissedDate & " " & slMissedTime & " SMF Removed"
            End If
            'lbcErrors.AddItem "SDF Not Found For SMF, SMF Removed"
        End If
        Do
            'tmRec = tmSmf
            'ilRet = gGetByKeyForUpdate("SMF", hmSmf, tmRec)
            'tmSmf = tmRec
            'On Error GoTo mSMFCheckErr
            'gBtrvErrorMsg ilRet, "mSMFCheck (Get by Key)", SMFCheck
            'On Error GoTo 0
            ilRet = btrDelete(hmSmf)
            If ilRet = BTRV_ERR_CONFLICT Then
                ilCRet = btrGetDirect(hmSmf, tmSmf, imSmfRecLen, tmDelSmf(ilLoop).lRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                On Error GoTo mSMFCheckErr
                gBtrvErrorMsg ilCRet, "mSMFCheck (btrGetDirect)", SMFCheck
                On Error GoTo 0
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSMFCheckErr
        gBtrvErrorMsg ilRet, "mSMFCheck (btrDelete)", SMFCheck
        On Error GoTo 0
        ilDeleted = ilDeleted + 1
    Next ilLoop
    ReDim tmSmfMake(0 To 0) As SMFMAKE
    ilUpper = UBound(tmSmfMake)
    For ilLoop = LBound(tmSdfInfo) To UBound(tmSdfInfo) - 1 Step 1
        If Not tmSdfInfo(ilLoop).iMatchFound Then
            gUnpackDate tmSdfInfo(ilLoop).iDate(0), tmSdfInfo(ilLoop).iDate(1), slDate
            gUnpackTime tmSdfInfo(ilLoop).iTime(0), tmSdfInfo(ilLoop).iTime(1), "A", "1", slTime
            If mGetChfClfAdf(tmSdfInfo(ilLoop).lChfCode, tmSdfInfo(ilLoop).iLineNo) Then
                If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                    tmSmfMake(ilUpper).sInfo = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & Str$(tmChf.lCntrNo) & Str$(tmClf.iLine) & " On " & slDate & " @ " & slTime
                    Print #hmTo, "SMF Missing for " & Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & " Contract #:" & Str$(tmChf.lCntrNo) & " Line #:" & Str$(tmClf.iLine) & " On " & slDate & " @ " & slTime
                Else
                    tmSmfMake(ilUpper).sInfo = Trim$(tmAdf.sName) & Str$(tmChf.lCntrNo) & Str$(tmClf.iLine) & " On " & slDate & " @ " & slTime
                    Print #hmTo, "SMF Missing for " & Trim$(tmAdf.sName) & " Contract #:" & Str$(tmChf.lCntrNo) & " Line #:" & Str$(tmClf.iLine) & " On " & slDate & " @ " & slTime
                End If
            Else
                tmSmfMake(ilUpper).sInfo = "Advt/Contract Missing Line" & Str$(tmClf.iLine) & " On " & slDate & " @ " & slTime
                Print #hmTo, "SMF Missing for " & "Advt/Contract Missing Line" & " Line #:" & Str$(tmClf.iLine) & " On " & slDate & " @ " & slTime
            End If
            lbcErrors.AddItem Trim$(tmSmfMake(ilUpper).sInfo)
            tmSmfMake(ilUpper).lSdfCode = tmSdfInfo(ilLoop).lCode
            tmSmfMake(ilUpper).lRecPos = 0
            ilUpper = ilUpper + 1
            ilSmfMissing = ilSmfMissing + 1
            ReDim Preserve tmSmfMake(0 To ilUpper) As SMFMAKE
        End If
    Next ilLoop
    lacCntr(2).Caption = "# Deleted" & Str$(ilDeleted) & "; # Status Chgs" & Str$(ilStatusChg) & "; # Missing SMF" & Str$(ilSmfMissing)
    mSMFCheck = True
    Exit Function
mSMFCheckErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSMFCheck = False
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
    Screen.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload SMFCheck
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestRecLengths                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if record lengths match    *
'*                                                     *
'*******************************************************
Private Function mTestRecLengths() As Integer
    Dim ilSizeError As Integer
    Dim ilSize As Integer
    ilSizeError = False
    ilSize = mGetRecLength("Sdf.Btr")
    If ilSize <> Len(tmSdf) Then
        If ilSize > 0 Then
            MsgBox "Sdf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Sdf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Smf.Btr")
    If ilSize <> Len(tmSmf) Then
        If ilSize > 0 Then
            MsgBox "Smf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSmf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Smf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Clf.Btr")
    If ilSize <> Len(tmClf) Then
        If ilSize > 0 Then
            MsgBox "Clf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmClf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Clf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Chf.Btr")
    If ilSize <> Len(tmChf) Then
        If ilSize > 0 Then
            MsgBox "Chf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmChf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Chf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Adf.Btr")
    If ilSize <> Len(tmAdf) Then
        If ilSize > 0 Then
            MsgBox "Adf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmAdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Adf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    mTestRecLengths = ilSizeError
End Function
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

    'ilRet = gPopUserVehicleBox(SMFCheck, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(SMFCheck, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcVehicle, tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SMFCheck
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "SMF Check/Fix"
End Sub
