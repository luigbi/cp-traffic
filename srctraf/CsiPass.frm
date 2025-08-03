VERSION 5.00
Begin VB.Form CSIPass 
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
      TabIndex        =   3
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
      TabIndex        =   0
      TabStop         =   0   'False
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
         IMEMode         =   3  'DISABLE
         Left            =   1485
         PasswordChar    =   "*"
         TabIndex        =   2
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
         TabIndex        =   4
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
         TabIndex        =   1
         Top             =   120
         Width           =   3750
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   -15
      Width           =   2055
   End
End
Attribute VB_Name = "CSIPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CSIPass.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the password screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imPasswordValue As Integer
Dim smSpecial As String

Private Sub cmcDone_Click()
    Dim ilValue As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim llValue As Long
    
    llValue = 10 * CLng(imPasswordValue)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(edcPassword.Text)
    If (Val(slStr) = ilValue) Or ((StrComp(slStr, smSpecial, vbTextCompare) = 0) And (Len(slStr) = 9)) Then
        igPasswordOk = True
    Else
        igPasswordOk = False
    End If
    mTerminate
End Sub
Private Sub edcPassword_KeyPress(KeyAscii As Integer)
    Dim slStr As String
'    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
'        Beep
'        KeyAscii = 0
'        Exit Sub
'    End If
'    slStr = edcPassword.Text
'    slStr = Left$(slStr, edcPassword.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPassword.SelStart - edcPassword.SelLength)
'    If Val(slStr) > 10000 Then
'        Beep
'        KeyAscii = 0
'        Exit Sub
'    End If
End Sub

Private Sub Form_Activate()
    'Me.KeyPreview = True
    'Me.Refresh
    DoEvents
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   
    'If (KeyCode = KEYF1) Or (KeyCode = KEYF2) Or (KeyCode = KEYF3) Then
    '    gFunctionKeyBranch KeyCode
    'End If
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
    CSIPass.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    CSIPass.Move (Screen.Width - CSIPass.Width) \ 2, (Screen.Height - CSIPass.Height) \ 2 + 175 '+ Screen.Height \ 10
    'CSIPass.Show
        'imPasswordValue = tmSpf.iPassword
        'If (imPasswordValue = 8224) Or (imPasswordValue = 0) Then
            Randomize
            imPasswordValue = Int(10000 * Rnd + 1)
            mSpecialPW
        'End If
        lacPrev.Caption = "Current Keycode is" & Str$(imPasswordValue) & ".  Leave New Keycode blank for View Mode Only.  Obtain New Keycode from Counterpoint Software."
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
    Unload CSIPass
    Set CSIPass = Nothing   'Remove data segment
End Sub

Private Sub mSpecialPW()
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String
    
    slDate = Format$(Now(), "m/d/yy")
    slMonth = Month(slDate)
    slYear = Year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(Str$(ilValue))
    Do While Len(slStr) < 4
        slStr = "0" & slStr
    Loop
    smSpecial = "Login" & slStr

End Sub
