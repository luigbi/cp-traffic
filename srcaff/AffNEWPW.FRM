VERSION 5.00
Begin VB.Form AffNewPW 
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frcStrongPW 
      Caption         =   "Rules for Strong Password"
      Height          =   1815
      Left            =   150
      TabIndex        =   8
      Top             =   1980
      Visible         =   0   'False
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
      Default         =   -1  'True
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
      Height          =   1485
      Left            =   150
      ScaleHeight     =   1425
      ScaleWidth      =   4605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   4665
      Begin VB.TextBox edcCurrentPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1695
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox edcVerifyPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1695
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1035
         Width           =   2775
      End
      Begin VB.TextBox edcNewPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1695
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   600
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
         Left            =   30
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1650
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
         Left            =   60
         TabIndex        =   6
         Top             =   1080
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
         Left            =   60
         TabIndex        =   4
         Top             =   645
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
Attribute VB_Name = "AffNewPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AffNewPw.frm
'
'
'
' Description:
'   This file contains the Password message screen code.
Option Explicit
Option Compare Text

Dim rst_Ust As ADODB.Recordset

Const KEYASTERISK = 42   'Asterisk key pressed

Private Sub cmcCancel_Click()
    igExitAff = True
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim slStr1 As String
    Dim slSqlUpdate As String
    Dim blPerformUpdate As Boolean
         '6/23/09:  Added SQL call to aviod assuming that rst as set from calling program (frmOptions and frmLogin).
        '          Also, changed from using rst to using rst_ust.
      'Dan M 7/24/09 set rst_ust
        SQLQuery = "SELECT * FROM ust where ucase(ustname) = '" & UCase(sgPassUserName) & "'"
        Set rst_Ust = gSQLSelectCall(SQLQuery)
        If rst_Ust.BOF And rst_Ust.EOF Then
            gMsgBox "Error finding username. Can't change password."
            cmcCancel_Click
            Exit Sub
        End If
   'current password not visible at login, where password must be blank
    If bgShowCurrentPassword Then
'        rst_Ust.Filter = "ustname = '" & frmOptions.txtSignName.text & "'"
'        If rst_Ust.BOF And rst_Ust.EOF Then
'            gMsgBox "Error finding username. Can't change password."
'            cmcCancel_Click
'            Exit Sub
'        End If
        slStr1 = Trim(rst_Ust!ustpassword)
        'rst_Ust.Filter = adFilterNone
        slStr = Trim$(edcCurrentPassword.Text)
        '5608  if 'strong password' added, old password should not be forced to be case sensitive
'        If bgStrongPassword Then
'            If StrComp(slStr, slStr1, vbBinaryCompare) <> 0 Then
'                MsgBox "Current Passwords Don't Match", vbOKOnly, "Error"
'                edcCurrentPassword.SetFocus
'                Exit Sub
'            End If
'        Else
'            If StrComp(slStr, slStr1, vbTextCompare) <> 0 Then
'                MsgBox "Current Passwords Don't Match", vbOKOnly, "Error"
'                edcCurrentPassword.SetFocus
'                Exit Sub
'            End If
'        End If
        If bgStrongPassword And gStrongPassword(slStr) Then
            If StrComp(slStr, slStr1, vbBinaryCompare) <> 0 Then
                MsgBox "Current Password is not correct", vbOKOnly, "Incorrect Password"
                edcCurrentPassword.SetFocus
                Exit Sub
            End If
        Else
            If StrComp(slStr, slStr1, vbTextCompare) <> 0 Then
                MsgBox "Current Password is not correct", vbOKOnly, "Incorrect Password"
                edcCurrentPassword.SetFocus
                Exit Sub
            End If
        End If
    End If
    slStr = Trim$(edcNewPassword.Text)
    slStr1 = Trim$(edcVerifyPassword.Text)
    If Len(slStr) <= 0 Then
        MsgBox "New Password Not Defined", vbOKOnly, "Error"
        edcNewPassword.SetFocus
        Exit Sub
    End If
    If bgStrongPassword Then
        If StrComp(slStr, slStr1, vbBinaryCompare) = 0 Then
            If Not gStrongPassword(slStr) Then
                MsgBox "New Password Violates Password Rules", vbOKOnly, "Error"
                edcNewPassword.Text = ""
                edcVerifyPassword.Text = ""
                edcNewPassword.SetFocus
                Exit Sub
            Else
                blPerformUpdate = True
            End If
        Else
            MsgBox "Verify Password does not match New Password", vbOKOnly, "Error"
            edcVerifyPassword.SetFocus
            Exit Sub
        End If
    Else
        If StrComp(slStr, slStr1, vbTextCompare) = 0 Then
            blPerformUpdate = True

        Else
            MsgBox "Verify Password does not match New Password", vbOKOnly, "Error"
            edcVerifyPassword.SetFocus
            Exit Sub
        End If
    End If
    'update only for login when password blank; if user in user options, he can choose to save later.
    If blPerformUpdate And Not bgShowCurrentPassword Then
        slSqlUpdate = "update ust set ustPassword = '" & slStr & "' where ustName = '" & rst_Ust!ustname & " ' "
        On Error GoTo ERRORBOX
        cnn.Execute (slSqlUpdate)
    End If
    sgPasswordPasser = slStr
    igExitAff = False
    mTerminate
    Exit Sub
ERRORBOX:
    MsgBox "Password could not be saved: Couldn't write to database.", vbOKOnly + vbInformation, "Error"
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

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

'Private Sub Form_Deactivate()
'    Me.KeyPreview = False
'End Sub




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
    If bgShowCurrentPassword Then
        edcCurrentPassword.Visible = True
        lacCurrentPassword.Visible = True
    Else
        edcCurrentPassword.Visible = False
        With lacCurrentPassword
            .Caption = "Your password is blank and needs to be reset."
            .Width = .Width + edcCurrentPassword.Width
            .Visible = True
        End With
   End If
    Screen.MousePointer = vbHourglass
    If bgStrongPassword Then
        frcStrongPW.Visible = True
    Else
        imcHelp.Top = plcPassword.Top + plcPassword.Height + 200
        cmcDone.Top = plcPassword.Top + plcPassword.Height + 300
        cmcCancel.Top = plcPassword.Top + plcPassword.Height + 300
    End If

    Me.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    Screen.MousePointer = vbDefault
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
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rst_Ust = Nothing
Set AffNewPW = Nothing
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Password"
End Sub
