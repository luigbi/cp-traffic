VERSION 5.00
Begin VB.Form frmSetupPassword 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please get Keycode from Counterpoint"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      Caption         =   "&Quit"
      Height          =   285
      Left            =   2610
      TabIndex        =   5
      Top             =   1680
      Width           =   945
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Enter"
      Default         =   -1  'True
      Height          =   285
      Left            =   1245
      TabIndex        =   4
      Top             =   1665
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
      Left            =   225
      ScaleHeight     =   1380
      ScaleWidth      =   4020
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
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
         TabIndex        =   1
         Top             =   1035
         Width           =   1560
      End
      Begin VB.Label lacWrong 
         Caption         =   "That password was not correct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   165
         TabIndex        =   6
         Top             =   435
         Width           =   3165
      End
      Begin VB.Label lacPrev 
         Appearance      =   0  'Flat
         Caption         =   "Previous Password xxxxx."
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
         TabIndex        =   2
         Top             =   1050
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmSetupPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim imPasswordValue As Integer
Dim smSpecial As String
Dim imNumberTries As Integer
Const MAXTRIES = 3

Private Sub cmcDone_Click()
    Dim ilValue As Integer
    Dim llValue As Long
    Dim slStr As String
    Dim ilRet As Integer
    
    lacWrong.Visible = False
    llValue = 10 * CLng(imPasswordValue)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(edcPassword.Text)
    If (Val(slStr) = ilValue) Or (StrComp(slStr, smSpecial, vbTextCompare) = 0) Then
        bgSetupPassword = True
    Else
        bgSetupPassword = False
        imNumberTries = imNumberTries + 1
        lacWrong.Visible = True
    End If
    If bgSetupPassword Then
         mTerminate
    ElseIf imNumberTries >= MAXTRIES Then
        mTerminate
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
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
    lacWrong.Visible = False
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2 + 175  '+ Screen.Height \ 10
    Randomize
    imPasswordValue = Int(10000 * Rnd + 1)
    mSpecialPW
    lacPrev.Caption = "Current Keycode is " & Trim$(Str$(imPasswordValue)) & "."
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
    Unload Me
    Set frmSetupPassword = Nothing   'Remove data segment
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
    smSpecial = slStr

End Sub
