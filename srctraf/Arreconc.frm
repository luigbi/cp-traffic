VERSION 5.00
Begin VB.Form ARReconc 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   885
   ClientTop       =   1185
   ClientWidth     =   8850
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   8850
   Begin VB.PictureBox plcBalanceMsg 
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
      Height          =   240
      Left            =   2175
      ScaleHeight     =   240
      ScaleWidth      =   4440
      TabIndex        =   2
      Top             =   4635
      Width           =   4440
   End
   Begin VB.PictureBox plcMsgDate 
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   5325
      TabIndex        =   3
      Top             =   2730
      Width           =   5385
      Begin VB.PictureBox pbcClosing 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   855
         Left            =   1110
         Picture         =   "Arreconc.frx":0000
         ScaleHeight     =   825
         ScaleWidth      =   3195
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   420
         Width           =   3225
      End
      Begin VB.Label lacMsgDate 
         Appearance      =   0  'Flat
         Caption         =   "Ending Dates of Closing Period"
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
         Height          =   210
         Left            =   1440
         TabIndex        =   16
         Top             =   105
         Width           =   2580
      End
   End
   Begin VB.PictureBox plcPrtMsg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   180
      ScaleHeight     =   615
      ScaleWidth      =   8400
      TabIndex        =   15
      Top             =   300
      Width           =   8400
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Don't Close"
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
      Left            =   4530
      TabIndex        =   4
      Top             =   4995
      Width           =   1245
   End
   Begin VB.CommandButton cmcClose 
      Appearance      =   0  'Flat
      Caption         =   "&Close Period"
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
      Left            =   3105
      TabIndex        =   5
      Top             =   4995
      Width           =   1245
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
      Height          =   240
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   15
      Width           =   915
   End
   Begin VB.PictureBox plcReconcile 
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   180
      ScaleHeight     =   1515
      ScaleWidth      =   8325
      TabIndex        =   1
      Top             =   1020
      Width           =   8385
      Begin VB.PictureBox plcARActual 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6120
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   14
         Top             =   1140
         Width           =   2085
      End
      Begin VB.PictureBox plcARPredicted 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6120
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   13
         Top             =   795
         Width           =   2085
      End
      Begin VB.PictureBox plcARCurrent 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6120
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   12
         Top             =   450
         Width           =   2085
      End
      Begin VB.PictureBox plcARPrevious 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6120
         ScaleHeight     =   255
         ScaleWidth      =   2055
         TabIndex        =   11
         Top             =   105
         Width           =   2085
      End
      Begin VB.Label lacARActual 
         Appearance      =   0  'Flat
         Caption         =   "Actual Current Receivable Balance"
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
         Height          =   210
         Left            =   75
         TabIndex        =   10
         Top             =   1185
         Width           =   5925
      End
      Begin VB.Label lacARPredicted 
         Appearance      =   0  'Flat
         Caption         =   "Predicted Current Receivable Balance"
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
         Height          =   210
         Left            =   90
         TabIndex        =   9
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label lacARCurrent 
         Appearance      =   0  'Flat
         Caption         =   "Current Period Transactions"
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
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   495
         Width           =   5910
      End
      Begin VB.Label lacARPrevious 
         Appearance      =   0  'Flat
         Caption         =   "Account Receivable Balance, Previous Closing"
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
         Height          =   210
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   5895
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   4890
      Width           =   360
   End
End
Attribute VB_Name = "ARReconc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Arreconc.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ARReconc.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Reconcile message screen code
Option Explicit
Option Compare Text
'Receivable
Dim imBalanced As Integer
Dim hmSpf As Integer        'Rvf handle
Dim imSpfRecLen As Integer     'Record length
Dim hmRhf As Integer
Dim tmRhf As RHF
Dim imRhfRecLen As Integer
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smEndPrevPeriod As String
Dim smEndCurrPeriod As String
Dim smEndNextPeriod As String
Dim smNewEndPeriod As String
Dim smPlcARPreviousP As String
Dim smPlcARCurrentP As String
Dim smPlcARPredictedP As String
Dim smPlcARActualP As String
Dim smPlcBalanceMsgP As String
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcClose_Click()
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilUpdateRhf As Integer
    Dim slNowDate As String
    Dim slStr As String

    mGenTextFile

    'gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slEndCurrentPeriod
    'If gDateValue(smCalDate) <= gDateValue(slEndCurrentPeriod) Then
    '    ilRet = MsgBox("Date must be after end of current period " & slEndCurrentPeriod, vbOkOnly + vbExclamation, "Error")
    '    Exit Sub
    'End If
    'Update Spf
    ilUpdateRhf = True
    Do
        ilCRet = btrGetFirst(hmSpf, tgSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
        If ilUpdateRhf Then
            tmRhf.iCode = 0
            tmRhf.sSourceType = "R"
            tmRhf.sRRP = tgSpf.sRRP
            gPDNToStr tgSpf.sRB, 2, slStr
            gStrToPDN slStr, 2, 6, tmRhf.sRB
            tmRhf.iRPRP(0) = tgSpf.iRPRP(0)
            tmRhf.iRPRP(1) = tgSpf.iRPRP(1)
            tmRhf.iRCRP(0) = tgSpf.iRCRP(0)
            tmRhf.iRCRP(1) = tgSpf.iRCRP(1)
            tmRhf.iRNRP(0) = tgSpf.iRNRP(0)
            tmRhf.iRNRP(1) = tgSpf.iRNRP(1)
            tmRhf.iUrfCode = tgUrf(0).iCode
            slNowDate = Format$(gNow(), "m/d/yy")
            gPackDate slNowDate, tmRhf.iDateEntered(0), tmRhf.iDateEntered(1)
            tmRhf.sUnused = ""
            ilRet = btrInsert(hmRhf, tmRhf, imRhfRecLen, INDEXKEY0)
            ilUpdateRhf = False
        End If
        tgSpf.iRPRP(0) = tgSpf.iRCRP(0)
        tgSpf.iRPRP(1) = tgSpf.iRCRP(1)
        tgSpf.iRCRP(0) = tgSpf.iRNRP(0)
        tgSpf.iRCRP(1) = tgSpf.iRNRP(1)
        gPackDate smNewEndPeriod, tgSpf.iRNRP(0), tgSpf.iRNRP(1)
        gStrToPDN sgThisMonthsClosing, 2, 6, tgSpf.sRB
        ilRet = btrUpdate(hmSpf, tgSpf, imSpfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    mTerminate
End Sub
Private Sub cmcClose_GotFocus()
    gCtrlGotFocus cmcClose
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
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = btrClose(hmSpf)
    btrDestroy hmSpf
    ilRet = btrClose(hmRhf)
    btrDestroy hmRhf

    Set ARReconc = Nothing   'Remove data segment

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
    Dim slPrevEndPeriodDate As String
    Dim slCurrEndPeriodDate As String
    Dim slNowDate As String
    Dim slLastInvDate As String
    Dim llRCRPDate As Long
    Dim llStdRCRPDate As Long
    Dim llStdLastInvDate As Long

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    imChgMode = False
    imBSMode = False
    mInitBox
    hmSpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        Exit Sub
    End If
    imSpfRecLen = Len(tgSpf)
    hmRhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRhf, "", sgDBPath & "Rhf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        Exit Sub
    End If
    imRhfRecLen = Len(tmRhf)
    mReconcTotals
    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), smEndPrevPeriod
    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), smEndCurrPeriod
    gUnpackDate tgSpf.iRNRP(0), tgSpf.iRNRP(1), smEndNextPeriod
    slDate = Format$(gDateValue(smEndNextPeriod) + 1, "m/d/yy")
    If tgSpf.sRRP = "S" Then    'Standard month
        smNewEndPeriod = gObtainEndStd(slDate)
    ElseIf tgSpf.sRRP = "C" Then    'Calendar
        smNewEndPeriod = gObtainEndCal(slDate)
    Else    'Corporate
        ilRet = gObtainCorpCal()            'retrieve corp calendars in memory for COF
        smNewEndPeriod = gObtainEndCorp(slDate, True)
    End If
    ARReconc.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone ARReconc

    'Moved from Collect to allow reconc to be pressed.  Used as aid to determine how to balance.
    '5/6/08: Convert the Currect period to standard broadcast month
    gUnpackDateLong tgSpf.iRCRP(0), tgSpf.iRCRP(1), llRCRPDate
    llStdRCRPDate = llRCRPDate - 15
    slDate = Format$(llStdRCRPDate, "m/d/yy")
    slDate = gObtainEndStd(slDate)
    llStdRCRPDate = gDateValue(slDate)
    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llStdLastInvDate
    If llStdRCRPDate > llStdLastInvDate Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Month must be Invoiced Prior to next Reconcile", vbOKOnly + vbExclamation, "Reconcile")
        cmcClose.Enabled = False
        Exit Sub
    End If

    'disallow closing period if trying to reconcile month before final invoicing
    '12/12/07- include corporate calendar
    'If (tgSpf.sRRP = "C") Or (tgSpf.sRRP = "S") Then
        'gSpfRead
        gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slPrevEndPeriodDate
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slLastInvDate
        If (tgSpf.sRRP <> "C") And (tgSpf.sRRP <> "S") Then
            slPrevEndPeriodDate = Format(gDateValue(slPrevEndPeriodDate) - 7, "m/d/yy")
        End If
        'Added year test for corporate calendar
        If (Month(slPrevEndPeriodDate) >= Month(slLastInvDate)) Or (Year(slPrevEndPeriodDate) > Year(slLastInvDate)) Then
            If Year(slPrevEndPeriodDate) >= Year(slLastInvDate) Then        '1-23-04
                ilRet = MsgBox("Month must be Invoiced Prior to next Reconcile", vbOKOnly + vbExclamation, "Reconcile")
                cmcClose.Enabled = False
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slCurrEndPeriodDate
        slNowDate = Format$(gNow(), "m/d/yy")
        If gDateValue(slNowDate) < gDateValue(slCurrEndPeriodDate) Then
            ilRet = MsgBox("End Date of Current Period is in the Future, Closing of Period Disallowed", vbOKOnly + vbExclamation, "Reconcile")
            cmcClose.Enabled = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    'End If
    'gCenterModalForm ARReconc
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
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
End Sub
Private Sub mReconcTotals()
    Dim slLastMonthClosing As String
    Dim slEndPrevPeriod As String
    Dim slEndCurrentPeriod As String
    Dim llEndPrevPeriod1 As Long
    Dim llEndCurrentPeriod As Long
    Dim slEndNextPeriod As String
    Dim slPredictedDollars As String
    Dim slStr As String
    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slEndPrevPeriod
    lacARPrevious.Caption = "Account Receivable Balance, Previous Closing " & slEndPrevPeriod
    gPDNToStr tgSpf.sRB, 2, slLastMonthClosing
    gFormatStr slLastMonthClosing, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
    'plcARPrevious.Caption = slStr
    smPlcARPreviousP = slStr
    plcARPrevious_Paint
    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slEndCurrentPeriod
    llEndPrevPeriod1 = gDateValue(slEndPrevPeriod) + 1
    lacARCurrent.Caption = "Current Period Transactions " & Format$(llEndPrevPeriod1, "m/d/yy") & "-" & slEndCurrentPeriod
    llEndCurrentPeriod = gDateValue(slEndCurrentPeriod)
    lacARPredicted.Caption = "Predicted Current Receivable Balance Through " & slEndCurrentPeriod
    lacARActual.Caption = "Actual Current Receivable Balance Through " & slEndCurrentPeriod

    gUnpackDate tgSpf.iRNRP(0), tgSpf.iRNRP(1), slEndNextPeriod
    gFormatStr sgCurrentDollars, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
    'plcARCurrent.Caption = slStr
    smPlcARCurrentP = slStr
    plcARCurrent_Paint
    slPredictedDollars = gAddStr(slLastMonthClosing, sgCurrentDollars)
    gFormatStr slPredictedDollars, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
    'plcARPredicted.Caption = slStr
    smPlcARPredictedP = slStr
    plcARPredicted_Paint
    gFormatStr sgThisMonthsClosing, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
    'plcARActual.Caption = slStr
    smPlcARActualP = slStr
    plcARActual_Paint
    If gCompNumberStr(sgThisMonthsClosing, slPredictedDollars) = 0 Then
        imBalanced = True
        smPlcBalanceMsgP = "* IN BALANCE *"
        plcBalanceMsg_Paint
        cmcClose.Enabled = True
        cmcCancel.Enabled = True
    Else
        imBalanced = False
        slStr = gSubStr(sgThisMonthsClosing, slPredictedDollars)
        gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN + FMTNEGATBACK, 2, slStr
        'plcBalanceMsg = "* OUT OF BALANCE BY " & slStr
        smPlcBalanceMsgP = "* OUT OF BALANCE BY " & slStr
        cmcClose.Enabled = False
        cmcCancel.Enabled = True
    End If
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
    igManUnload = YES
    Unload ARReconc
    igManUnload = NO
End Sub
Private Sub pbcClosing_Paint()
    pbcClosing.CurrentX = 1110 + 45
    pbcClosing.CurrentY = 225 - 15
    pbcClosing.Print smEndPrevPeriod
    pbcClosing.CurrentX = 1110 + 45
    pbcClosing.CurrentY = 420 - 15
    pbcClosing.Print smEndCurrPeriod
    pbcClosing.CurrentX = 1110 + 45
    pbcClosing.CurrentY = 615 - 15
    pbcClosing.Print smEndNextPeriod
    pbcClosing.CurrentX = 2145 + 45
    pbcClosing.CurrentY = 225 - 15
    pbcClosing.Print smEndCurrPeriod
    pbcClosing.CurrentX = 2145 + 45
    pbcClosing.CurrentY = 420 - 15
    pbcClosing.Print smEndNextPeriod
    pbcClosing.CurrentX = 2145 + 45
    pbcClosing.CurrentY = 615 - 15
    pbcClosing.Print smNewEndPeriod
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcBalanceMsg_Paint()
    plcBalanceMsg.CurrentX = 0
    plcBalanceMsg.CurrentY = 0
    'plcBalanceMsg.Print "*  IN BALANCE *"
    plcBalanceMsg.Print smPlcBalanceMsgP
End Sub
Private Sub plcARPrevious_Paint()
    plcARPrevious.CurrentX = 0
    plcARPrevious.CurrentY = 0
    plcARPrevious.Print smPlcARPreviousP
End Sub
Private Sub plcARCurrent_Paint()
    plcARCurrent.CurrentX = 0
    plcARCurrent.CurrentY = 0
    plcARCurrent.Print smPlcARCurrentP
End Sub
Private Sub plcARPredicted_Paint()
    plcARPredicted.Cls
    plcARPredicted.CurrentX = 0
    plcARPredicted.CurrentY = 0
    plcARPredicted.Print smPlcARPredictedP
End Sub
Private Sub plcARActual_Paint()
    plcARActual.Cls
    plcARActual.CurrentX = 0
    plcARActual.CurrentY = 0
    plcARActual.Print smPlcARActualP
End Sub

Private Sub plcPrtMsg_Paint()
    plcPrtMsg.Cls
    plcPrtMsg.CurrentX = 0
    plcPrtMsg.CurrentY = 0
    plcPrtMsg.Print "Reconcile Report Must Be Printed To Completion Prior To Closing Period"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.Cls
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Reconcile"
End Sub

Public Sub mGenTextFile()
    Dim hlMsg As Integer
    Dim slToFile As String
    Dim slMsg As String
    Dim ilRet As Integer

    'On Error GoTo mGenTextFileErr:
    slToFile = sgDBPath & "Messages\Reconcile.Txt"
    'hlMsg = FreeFile
    'Open slToFile For Append As hlMsg
    ilRet = gFileOpen(slToFile, "Append", hlMsg)
    If ilRet <> 0 Then
        MsgBox "Unable to Open " & slToFile & " Error #" & str(Err.Number)
        Close #hlMsg
        Exit Sub
    End If
    slMsg = "Reconcile run on " & Format(Now, "m/d/yy") & " at " & Format(Now, "h:mm:ssAM/PM")
    Print #hlMsg, slMsg

    Print #hlMsg, ""

    slMsg = "     " & lacARPrevious.Caption
    Do While Len(slMsg) < 60
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smPlcARPreviousP
    Print #hlMsg, slMsg

    slMsg = "     " & lacARCurrent.Caption
    Do While Len(slMsg) < 60
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smPlcARCurrentP
    Print #hlMsg, slMsg

    slMsg = "     " & lacARPredicted.Caption
    Do While Len(slMsg) < 60
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smPlcARPredictedP
    Print #hlMsg, slMsg

    slMsg = "     " & lacARActual.Caption
    Do While Len(slMsg) < 60
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smPlcARActualP
    Print #hlMsg, slMsg

    Print #hlMsg, ""

    slMsg = "     "
    slMsg = slMsg & "Ending Dates of Closing Period"
    Print hlMsg, slMsg

    slMsg = "     "
    Do While Len(slMsg) < 25
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & "Old"
    Do While Len(slMsg) < 35
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & "New"
    Print #hlMsg, slMsg

    slMsg = "     " & "Prior Period"
    Do While Len(slMsg) < 25
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smEndPrevPeriod
    Do While Len(slMsg) < 35
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smEndCurrPeriod
    Print #hlMsg, slMsg

    slMsg = "     " & "Current Period"
    Do While Len(slMsg) < 25
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smEndCurrPeriod
    Do While Len(slMsg) < 35
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smEndNextPeriod
    Print #hlMsg, slMsg

    slMsg = "     " & "Future Period"
    Do While Len(slMsg) < 25
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smEndNextPeriod
    Do While Len(slMsg) < 35
        slMsg = slMsg & " "
    Loop
    slMsg = slMsg & smNewEndPeriod
    Print #hlMsg, slMsg

    Print #hlMsg, ""

    slMsg = "     " & smPlcBalanceMsgP
    Print #hlMsg, slMsg

    Close #hlMsg
    Exit Sub
'mGenTextFileErr:
'    MsgBox "Unable to Open " & slToFile & " Error #" & str(Err.Number)
'    Close #hlMsg
'    Exit Sub
End Sub
