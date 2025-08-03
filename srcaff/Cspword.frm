VERSION 5.00
Begin VB.Form CSPWord 
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
         Caption         =   "New Keycode"
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
         Caption         =   "Previous Keycode xxxxx, leave blank for View Mode Only"
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
Attribute VB_Name = "CSPWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CSPWord.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the password screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
'Removed saving of password in spf
'Dim hmSpf As Integer
'Dim imSpfRecLen As Integer
'Dim tmSpf As SPF
Dim imPasswordValue As Integer
Private Sub cmcDone_Click()
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilRet As Integer
    If (Val(edcPassword.Text) = Val(sgSpecialPassword)) And (Len(edcPassword.Text) = 4) Then
        igPasswordOk = True
        igChangesAllowed = -1
    Else
        ' Dan M 4/22/09 affiliate now like traffic, uses keycode. password no longer used.
        llValue = 10 * CLng(imPasswordValue)
        ilValue = Int(10000 * Rnd(-llValue) + 1)
        'ilValue = Int(10000 * Rnd(-imPasswordValue) + 1)
        If Val(edcPassword.Text) = ilValue Then
            igPasswordOk = True
             igChangesAllowed = -1
            'Do
            '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            '    tmSpf.iPassword = Val(edcPassword.Text)
            '    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
            'Loop While ilRet = BTRV_ERR_CONFLICT
              ' Dan M 4/17/09 Limit # of changes to site options
        Else
            ilRet = mTestForChangeNumber(ilValue)
            If ilRet > 0 Then
                igPasswordOk = True
                igChangesAllowed = ilRet
            Else
              igPasswordOk = False
              igChangesAllowed = 0
            End If
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
    slStr = edcPassword.Text
    slStr = Left$(slStr, edcPassword.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPassword.SelStart - edcPassword.SelLength)
    If Val(slStr) > 10000 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    Me.KeyPreview = True
    Me.Refresh
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
    If Trim$(sgSpecialPassword) <> "" Then
        edcPassword.PasswordChar = "*"
    End If
    CSPWord.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    CSPWord.Move (Screen.Width - CSPWord.Width) \ 2, (Screen.Height - CSPWord.Height) \ 2 + 175 '+ Screen.Height \ 10
    ''CSPWord.Show
    'hmSpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    'ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'BTRV_LOCK_NONE)
    'If ilRet = BTRV_ERR_NONE Then
    '    imSpfRecLen = Len(tmSpf)
    '    'Until multi-user version of btrieve installed- test field
    '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    '    imPasswordValue = tmSpf.iPassword
    '    If (imPasswordValue = 8224) Or (imPasswordValue = 0) Then
    ' Dan M 4/20/09 added do loop to keep value 4 numbers
    Do
            Randomize
            imPasswordValue = Int(10000 * Rnd + 1)
    Loop While imPasswordValue + 100 > 9999
    '    End If
    If sgPasswordAddition <> "Disallow View" Then
        lacPrev.Caption = "Current Keycode is " & Trim$(sgPasswordAddition) & Trim$(Str$(imPasswordValue)) & ", Leave New Keycode Blank for View Mode Only.  Obtain New Keycode from Counterpoint Software"
    Else
        lacPrev.Caption = "Current Keycode is " & Trim$(Str$(imPasswordValue)) & sgCRLF & "Obtain New Keycode from Counterpoint Software"
    End If
    'End If
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
    'ilRet = btrClose(hmSpf)
    'btrDestroy hmSpf
    Unload CSPWord
End Sub
Private Function mTestForChangeNumber(ilMainPass As Integer) As Integer
Dim ilMasterPass As Integer
Dim ilChoicePass As Integer
Dim ilComparePass As Integer
Dim c As Integer
ilMasterPass = Val(edcPassword.Text)
ilChoicePass = ilMasterPass - ilMainPass
For c = 1 To 10
    ilComparePass = Int(100 * Rnd(-c) + 1)
    If ilChoicePass = ilComparePass Then
        mTestForChangeNumber = c
        Exit Function
    End If
Next c
mTestForChangeNumber = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set CSPWord = Nothing   'Remove data segment
End Sub
