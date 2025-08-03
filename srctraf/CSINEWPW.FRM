VERSION 5.00
Begin VB.Form CSINewPW 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4425
   ClientLeft      =   1680
   ClientTop       =   2550
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
   ScaleHeight     =   4425
   ScaleWidth      =   5010
   Begin VB.Frame frcStrongPW 
      Caption         =   "Rules for Strong Password"
      Height          =   1815
      Left            =   150
      TabIndex        =   8
      Top             =   1980
      Width           =   4680
      Begin VB.Label lacStrongPW 
         Caption         =   "1.  At least one Alpha Character, either Upper or Lower Case"
         Height          =   420
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Top             =   255
         Width           =   4335
      End
      Begin VB.Label lacStrongPW 
         Caption         =   "2.  At least 1 Number between 0 and 9."
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   10
         Top             =   705
         Width           =   4035
      End
      Begin VB.Label lacStrongPW 
         Caption         =   "3.  At least 1 Special Symbol like @, #, $, etc."
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   11
         Top             =   960
         Width           =   4110
      End
      Begin VB.Label lacStrongPW 
         Caption         =   "4.  Between 8 and 20 characters"
         Height          =   225
         Index           =   3
         Left            =   45
         TabIndex        =   12
         Top             =   1215
         Width           =   4065
      End
      Begin VB.Label lacStrongPW 
         Caption         =   "Note:  Password is case-sensitive at sign-on"
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   14
         Top             =   1485
         Width           =   4065
      End
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
      Left            =   2745
      TabIndex        =   16
      Top             =   4005
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
      Left            =   1245
      TabIndex        =   15
      Top             =   4005
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   4830
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4185
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
      Begin VB.TextBox edcVerifyPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1725
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   870
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
         Top             =   495
         Width           =   2775
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
         Top             =   915
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
         Top             =   540
         Width           =   1410
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3915
      Width           =   360
   End
End
Attribute VB_Name = "CSINewPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CSINEWPW.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CSINewPW.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Password message screen code
Option Explicit
Option Compare Text
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smCurrentPassword As String
Dim smOldPasswords(0 To 2) As String

Private Sub cmcCancel_Click()
    igExitTraffic = True
    sgPWResult = ""
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim slStr1 As String
    Dim ilMatchOld As Integer
    'hide current password question when password is ""
    Dim blTestCurrentPassword As Boolean

    ilMatchOld = False
    blTestCurrentPassword = edcCurrentPassword.Visible
    slStr1 = Trim$(smCurrentPassword)
    If blTestCurrentPassword Then
        slStr = Trim$(edcCurrentPassword.Text)
        If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
            If StrComp(slStr, slStr1, vbBinaryCompare) <> 0 Then
                MsgBox "Current Passwords Don't Match", vbOKOnly, "Error"
                edcCurrentPassword.SetFocus
                Exit Sub
            End If
        Else
            If StrComp(slStr, slStr1, vbTextCompare) <> 0 Then
                MsgBox "Current Passwords Don't Match", vbOKOnly, "Error"
                edcCurrentPassword.SetFocus
                Exit Sub
            End If
        End If
    End If
    slStr = Trim$(edcNewPassword.Text)
    If Len(slStr) <= 0 Then
        MsgBox "New Password Not Defined", vbOKOnly, "Error"
        edcNewPassword.SetFocus
        Exit Sub
    End If
    slStr = Trim$(edcNewPassword.Text)
    If StrComp(slStr, smCurrentPassword, vbTextCompare) = 0 Then
        ilMatchOld = True
    Else
        If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
            If StrComp(slStr, smOldPasswords(0), vbTextCompare) = 0 Then
                ilMatchOld = True
            End If
            If StrComp(slStr, smOldPasswords(1), vbTextCompare) = 0 Then
                ilMatchOld = True
            End If
            If StrComp(slStr, smOldPasswords(2), vbTextCompare) = 0 Then
                ilMatchOld = True
            End If
        End If
    End If
    If ilMatchOld Then
        MsgBox "New Password cannot match Old Passwords", vbOKOnly, "Error"
        edcNewPassword.SetFocus
        Exit Sub
    End If
    If Not gStrongPassword(slStr) Then
        MsgBox "New Password Violates Password Rules", vbOKOnly, "Error"
        edcNewPassword.SetFocus
        Exit Sub
    End If
    slStr = Trim$(edcNewPassword.Text)
    slStr1 = Trim$(edcVerifyPassword.Text)
    If (Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) = STRONGPASSWORD Then
        If StrComp(slStr, slStr1, vbBinaryCompare) = 0 Then
            sgPWResult = slStr
        Else
            MsgBox "Verify Password does not match New Password", vbOKOnly, "Error"
            edcVerifyPassword.SetFocus
            Exit Sub
        End If
    Else
        If StrComp(slStr, slStr1, vbTextCompare) = 0 Then
            sgPWResult = slStr
        Else
            MsgBox "Verify Password does not match New Password", vbOKOnly, "Error"
            edcVerifyPassword.SetFocus
            Exit Sub
        End If
    End If
    igExitTraffic = False
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub

Private Sub edcCurrentPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNewPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNewPassword_KeyPress(KeyAscii As Integer)
    If (Len(edcNewPassword.Text) = 0) And (KeyAscii = KEYASTERISK) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcVerifyPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcVerifyPassword_KeyPress(KeyAscii As Integer)
    If (Len(edcVerifyPassword.Text) = 0) And (KeyAscii = KEYASTERISK) Then
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
        igExitTraffic = False
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CSINewPW = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    ''Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    ''Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    ''Traffic!cdcSetup.Action = 6
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
    smCurrentPassword = Trim$(tgPWUrf.sPassword)
    sgPWResult = ""
    smOldPasswords(0) = Trim$(tgPWUrf.sOldPassword(0))
    smOldPasswords(1) = Trim$(tgPWUrf.sOldPassword(1))
    smOldPasswords(2) = Trim$(tgPWUrf.sOldPassword(2))
    imTerminate = False
    imChgMode = False
    imBSMode = False
    If ((Asc(tgSpf.sUsingFeatures2) And STRONGPASSWORD) <> STRONGPASSWORD) Then
        imcHelp.Top = plcPassword.Top + plcPassword.Height + 200
        cmcDone.Top = plcPassword.Top + plcPassword.Height + 300
        cmcCancel.Top = plcPassword.Top + plcPassword.Height + 300
        frcStrongPW.Visible = False
    End If
    CSINewPW.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterModalForm CSINewPW
    'gCenterModalForm CSINewPW
    gCenterStdAlone CSINewPW
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
    Unload CSINewPW
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
