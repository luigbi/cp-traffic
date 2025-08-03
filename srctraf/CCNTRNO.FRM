VERSION 5.00
Begin VB.Form CCntrNo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1725
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   3885
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
   ScaleHeight     =   1725
   ScaleWidth      =   3885
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2145
      TabIndex        =   6
      Top             =   1290
      Width           =   1050
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
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1515
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   1245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1245
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   555
      TabIndex        =   5
      Top             =   1290
      Width           =   1050
   End
   Begin VB.PictureBox plcInv 
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
      Height          =   780
      Left            =   105
      ScaleHeight     =   720
      ScaleWidth      =   3585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   3645
      Begin VB.TextBox edcCntrNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   3
         Top             =   75
         Width           =   1545
      End
      Begin VB.Label lacTempMsg 
         Appearance      =   0  'Flat
         Caption         =   "(Temporary Number until Exported)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   435
         Width           =   3450
      End
      Begin VB.Label lacCntrNo 
         Appearance      =   0  'Flat
         Caption         =   "Contract #"
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
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1065
      End
   End
End
Attribute VB_Name = "CCntrNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CCNTRNO.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CCntrNo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract number screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer

Private Sub cmcCancel_Click()
    igTerminateReturn = 0   '0=Canceled
    sgPassValue = ""
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    slStr = edcCntrNo.Text
    If (Trim$(slStr) = "") Or (Val(slStr) = 0) Then
       Beep
       edcCntrNo.SetFocus
       Exit Sub
    Else
        'Test if previously used
    End If
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilRet As Integer
    Dim slStr As String

    gCtrlGotFocus cmcDone
    slStr = edcCntrNo.Text
    Contract.edcCntrNo.Text = slStr
    ilRet = mGetNextCntrNo(hmPassChf, lmChfRecPos, lmCntrChfCode)
    edcCntrNo.Text = Contract.edcCntrNo.Text
End Sub
Private Sub edcCntrNo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcCntrNo_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcCntrNo.Text
    slStr = Left$(slStr, edcCntrNo.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCntrNo.SelStart - edcCntrNo.SelLength)
    If gCompNumberStr(slStr, "999999999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
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
End Sub
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
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    CCntrNo.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterModalForm CCntrNo
    If tgUrf(0).iRemoteUserID > 0 Then
        ilRet = mGetNextCntrNo(hmPassChf, lmChfRecPos, lmCntrChfCode)
        edcCntrNo.Text = Contract.edcCntrNo.Text
        edcCntrNo.Enabled = False
        lacTempMsg.Caption = "(Temporary Number until Exported)"
    Else
        edcCntrNo.Text = Trim$(Contract.edcCntrNo.Text)
        lacTempMsg.Caption = "(Tab to Assign # or Enter #)"
    End If
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Unload CCntrNo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CCntrNo = Nothing   'Remove data segment
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Contract #"
End Sub
