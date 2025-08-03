VERSION 5.00
Begin VB.Form Password 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2400
   ClientLeft      =   1620
   ClientTop       =   2670
   ClientWidth     =   5010
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
   ScaleHeight     =   2400
   ScaleWidth      =   5010
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
      Left            =   2730
      TabIndex        =   9
      Top             =   1980
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
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
      Left            =   1230
      TabIndex        =   8
      Top             =   1980
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
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
      TabStop         =   0   'False
      Top             =   15
      Width           =   915
   End
   Begin VB.PictureBox plcPassword 
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   150
      ScaleHeight     =   1320
      ScaleWidth      =   4605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   465
      Width           =   4665
      Begin VB.TextBox edcVerifyPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1725
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   855
         Width           =   2775
      End
      Begin VB.TextBox edcCurrentPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1725
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   105
         Width           =   2775
      End
      Begin VB.TextBox edcNewPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1725
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lacVerifyPassword 
         Appearance      =   0  'Flat
         Caption         =   "Verify Password"
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
         TabIndex        =   6
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label lacNewPassword 
         Appearance      =   0  'Flat
         Caption         =   "New Password"
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
         TabIndex        =   4
         Top             =   525
         Width           =   1410
      End
      Begin VB.Label lacCurrentPassword 
         Appearance      =   0  'Flat
         Caption         =   "Current Password"
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
         TabIndex        =   2
         Top             =   150
         Width           =   1740
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   135
      Top             =   1950
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Password.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Password.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Password message screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smCurrentPassword As String
Private Sub cmcCancel_Click()
    UserOpt!edcHiddenPassword.Text = ""
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim slStr1 As String
    slStr = Trim$(edcCurrentPassword.Text)
    If StrComp(slStr, smCurrentPassword, 1) = 0 Then
        slStr = Trim$(edcNewPassword.Text)
        If Len(slStr) > 0 Then
            slStr1 = Trim$(edcVerifyPassword.Text)
            If StrComp(slStr, slStr1, 1) = 0 Then
                UserOpt!edcHiddenPassword.Text = slStr
            Else
                MsgBox "Verify Password incorrect, New one rejected", vbOKOnly, "Error"
                UserOpt!edcHiddenPassword.Text = ""
            End If
        Else
            MsgBox "New Password Rejected", vbOKOnly, "Error"
            UserOpt!edcHiddenPassword.Text = ""
        End If
    Else
        MsgBox "Old Password incorrect, New one rejected", vbOKOnly, "Error"
        UserOpt!edcHiddenPassword.Text = ""
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub edcCurrentPassword_KeyPress(KeyAscii As Integer)
    If (Len(edcCurrentPassword.Text) = 0) And (KeyAscii = KEYASTERISK) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcNewPassword_KeyPress(KeyAscii As Integer)
    If (Len(edcNewPassword.Text) = 0) And (KeyAscii = KEYASTERISK) Then
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
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Password = Nothing   'Remove data segment
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
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    smCurrentPassword = UserOpt!edcHiddenPassword.Text
    UserOpt!edcHiddenPassword.Text = ""
    imTerminate = False
    imChgMode = False
    imBSMode = False
    Password.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone Password
    'gCenterModalForm Password
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
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
    Unload Password
    igManUnload = NO
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Password"
End Sub
