VERSION 5.00
Begin VB.Form RCModel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4410
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   3330
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
   ScaleHeight     =   4410
   ScaleWidth      =   3330
   Begin VB.TextBox edcDate 
      Height          =   315
      Left            =   2145
      TabIndex        =   4
      Top             =   3390
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1860
      TabIndex        =   6
      Top             =   3945
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
      Left            =   3105
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   3780
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1920
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   630
      TabIndex        =   5
      Top             =   3945
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
      ScaleWidth      =   2955
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   3015
      Begin VB.ListBox lbcRateCard 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   2895
      End
   End
   Begin VB.Label lacDate 
      Caption         =   "Select Rate Card Items Referenced in Contracts On or After"
      Height          =   600
      Left            =   90
      TabIndex        =   3
      Top             =   3030
      Width           =   2025
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   3915
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RCModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rcmodel.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCModel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim tmRateCardCode() As SORTCODE
Dim smRateCardCodeTag As String
Dim smNowDate As String
Dim smStartDate As String

Private Sub cmcCancel_Click()
    igRcfModel = 0
    igRCReturn = 0
    sgRCModelDate = ""
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDate As String

    slDate = edcDate.Text
    If slDate <> "" Then
        If Not gValidDate(slDate) Then
            Beep
            edcDate.SetFocus
            Exit Sub
        End If
    End If
    sgRCModelDate = slDate
    igRcfModel = 0
    If lbcRateCard.ListIndex >= 1 Then
        slNameCode = tmRateCardCode(lbcRateCard.ListIndex - 1).sKey    'lbcRateCardCode.List(lbcRateCard.ListIndex - 1)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        igRcfModel = Val(slCode)
    End If
    igRCReturn = 1
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
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        lbcRateCard.Enabled = False
    Else
        lbcRateCard.Enabled = True
    End If
'    gShowBranner
    Me.KeyPreview = True
    RCModel.Refresh
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

Private Sub Form_Unload(Cancel As Integer)
    Set RCModel = Nothing   'Remove data segment
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
    RCModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm RCModel
    'RCModel.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    smNowDate = Format$(gNow(), "m/d/yy")
    smStartDate = gObtainYearStartDate(0, smNowDate)
    edcDate.Text = smStartDate
    lbcRateCard.ListIndex = 1
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm RCModel
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
    Dim ilRet As Integer 'btrieve status

    imPopReqd = False
    'ilRet = gPopRateCardBox(RCModel, 0, lbcRateCard, lbcRateCardCode, -1)
    ilRet = gPopRateCardBox(RCModel, 0, lbcRateCard, tmRateCardCode(), smRateCardCodeTag, -1)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopRateCardBox)", RCModel
        On Error GoTo 0
        lbcRateCard.AddItem "[None]", 0  'Force as first item on list
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
    Erase tmRateCardCode
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload RCModel
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
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Model From Rate Card"
End Sub
