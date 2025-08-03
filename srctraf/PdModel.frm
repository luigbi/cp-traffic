VERSION 5.00
Begin VB.Form PdModel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   4740
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
   ScaleHeight     =   3540
   ScaleWidth      =   4740
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2475
      TabIndex        =   4
      Top             =   3195
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Top             =   3195
      Width           =   945
   End
   Begin VB.PictureBox plcModel 
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
      Height          =   2670
      Left            =   150
      ScaleHeight     =   2610
      ScaleWidth      =   4260
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   270
      Width           =   4320
      Begin VB.ListBox lbcSelection 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   4200
      End
   End
   Begin VB.Label plcScreen 
      Caption         =   "Model Pledge from:"
      Height          =   225
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   1860
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3165
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "PdModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PdModel.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PdModel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text

Dim tmFpf As FPF        'FPF record image
Dim tmFpfSrchKey1 As FPFKEY1    'FPF key record image
Dim tmFpfSrchKey2 As FPFKEY2    'FPF key record image
Dim hmFpf As Integer    'Sale Commission file handle
Dim imFpfRecLen As Integer        'FPF record length

'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Private Sub cmcCancel_Click()
    igPdCodeFpf = 0
    igPdReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()

    igPdCodeFpf = 0
    igPdReturn = 1
    If lbcSelection.ListIndex >= 0 Then
        igPdCodeFpf = lbcSelection.ItemData(lbcSelection.ListIndex)
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
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
    If (igWinStatus(COLLECTIONSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        lbcSelection.Enabled = False
    Else
        lbcSelection.Enabled = True
    End If
'    gShowBranner
    Me.KeyPreview = True
    PdModel.Refresh
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
        cmcCancel_Click
    End If
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
    imTerminate = False
    imFirstActivate = True


    Screen.MousePointer = vbHourglass
    PdModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone PdModel
    'PdModel.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    lbcSelection.ListIndex = -1
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm PdModel
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    imTerminate = True
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
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    imPopReqd = False
    'Populate with each unique effective date

    hmFpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmFpf, "", sgDBPath & "FPF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mPopulateErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: FPF.Btr)", PdModel
    On Error GoTo 0
    imFpfRecLen = Len(tmFpf)

    tmFpfSrchKey2.iFnfCode = igPledgeFnfCode
    tmFpfSrchKey2.iVefCode = igPledgeVefCode
    tmFpfSrchKey2.iEffStartDate(0) = 0
    tmFpfSrchKey2.iEffStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmFpf, tmFpf, imFpfRecLen, tmFpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmFpf.iFnfCode = igPledgeFnfCode) And (tmFpf.iVefCode = igPledgeVefCode)
        gUnpackDate tmFpf.iEffStartDate(0), tmFpf.iEffStartDate(1), slStartDate
        gUnpackDate tmFpf.iEffEndDate(0), tmFpf.iEffEndDate(1), slEndDate
        If Trim$(slEndDate) <> "" Then
            If gDateValue(slEndDate) = gDateValue("12/31/2068") Then
                slEndDate = "TFN"
            End If
        End If
        'set to zero to have date in descending order
        lbcSelection.AddItem slStartDate & "-" & slEndDate, 0
        lbcSelection.ItemData(lbcSelection.NewIndex) = Trim$(Str$(tmFpf.iCode))
        ilRet = btrGetNext(hmFpf, tmFpf, imFpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    btrExtClear hmFpf   'Clear any previous extend operation
    ilRet = btrClose(hmFpf)
    Exit Sub
mPopulateErr:
    btrExtClear hmFpf   'Clear any previous extend operation
    ilRet = btrClose(hmFpf)
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload PdModel
    Set PdModel = Nothing   'Remove data segment
    igManUnload = NO
End Sub



Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
