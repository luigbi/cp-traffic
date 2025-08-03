VERSION 5.00
Begin VB.Form EngrPass 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2145
   ClientLeft      =   2970
   ClientTop       =   1440
   ClientWidth     =   4410
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   4410
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   285
      Left            =   1695
      TabIndex        =   4
      Top             =   1755
      Width           =   945
   End
   Begin VB.PictureBox plcInvNo 
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
      Height          =   1440
      Left            =   165
      ScaleHeight     =   1380
      ScaleWidth      =   4020
      TabIndex        =   2
      Top             =   225
      Width           =   4080
      Begin VB.TextBox edcPassword 
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
         Height          =   285
         Left            =   1485
         TabIndex        =   0
         Top             =   1035
         Width           =   1560
      End
      Begin VB.Label lacPassword 
         Appearance      =   0  'Flat
         Caption         =   "New Password"
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
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label lacPrev 
         Appearance      =   0  'Flat
         Caption         =   "Previous Password xxxxx, leave blank for View Mode Only"
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
         Height          =   825
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3750
      End
   End
   Begin VB.Label plcScreen 
      Caption         =   "Password"
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2070
   End
End
Attribute VB_Name = "EngrPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrPass.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the password screen code
Option Explicit
Option Compare Text
Dim imPasswordValue As Integer
Private Sub cmcDone_Click()
    Dim ilValue As Integer
    Dim ilRet As Integer
    Dim llValue As Long
    
    If (Val(edcPassword.text) = Val(sgSpecialPassword)) And (Len(edcPassword.text) = 4) Then
        igPasswordOk = True
    Else
        'ilValue = Int(10000 * Rnd(-imPasswordValue) + 1)
        'If Val(edcPassword.text) = ilValue Then
        '    igPasswordOk = True
        'Else
        '    igPasswordOk = False
        'End If
        llValue = 10 * CLng(imPasswordValue)
        ilValue = Int(10000 * Rnd(-llValue) + 1)
        If Val(edcPassword.text) = ilValue Then
            igPasswordOk = True
        Else
            igPasswordOk = False
        End If
    End If
    mTerminate
End Sub
Private Sub edcPassword_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcPassword.text
    slStr = Left$(slStr, edcPassword.SelStart) & Chr$(KeyAscii) & Right$(slStr, Len(slStr) - edcPassword.SelStart - edcPassword.SelLength)
    If Val(slStr) > 10000 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Load()
    mInit
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
    If Trim$(sgSpecialPassword) <> "" Then
        edcPassword.PasswordChar = "*"
    End If
    EngrPass.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    EngrPass.Move (Screen.Width - EngrPass.Width) \ 2, (Screen.Height - EngrPass.Height) \ 2 + 175 '+ Screen.Height \ 10
    Randomize
    imPasswordValue = Int(10000 * Rnd + 1)
    lacPrev.Caption = "Current Keycode is " & Trim$(sgPasswordAddition) & Trim$(Str$(imPasswordValue)) & ", Leave New Password Blank for View Mode Only.  Obtain New Password from Counterpoint Software"
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
    Dim ilRet As Integer
    Unload EngrPass
    Set EngrPass = Nothing   'Remove data segment
End Sub
