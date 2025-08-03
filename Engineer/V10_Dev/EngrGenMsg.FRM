VERSION 5.00
Begin VB.Form EngrGenMsg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2505
   ClientLeft      =   4485
   ClientTop       =   1470
   ClientWidth     =   5310
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
   ScaleHeight     =   2505
   ScaleWidth      =   5310
   Begin VB.TextBox edcEditValue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Left            =   495
      TabIndex        =   7
      Top             =   1605
      Visible         =   0   'False
      Width           =   4395
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   3
      Left            =   3990
      TabIndex        =   4
      Top             =   2040
      Width           =   1170
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   2040
      Width           =   1170
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   2
      Top             =   2040
      Width           =   1170
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
      Left            =   -15
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2310
      Width           =   120
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label plcEditValue 
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   1290
      Visible         =   0   'False
      Width           =   5025
   End
   Begin VB.Label plcMsg 
      Height          =   885
      Left            =   135
      TabIndex        =   5
      Top             =   120
      Width           =   5025
   End
End
Attribute VB_Name = "EngrGenMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrGenMsg.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
Dim imFirstTime As Integer
Dim imFirstActivate As Integer

Private Sub cmcButton_Click(Index As Integer)
    igAnsCMC = Index
    If igEditBox = 1 Then
        sgEditValue = edcEditValue.Text
    End If
    mTerminate
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Exit Sub
    End If
    imFirstActivate = False
    If imFirstTime Then
        imFirstTime = True
        cmcButton(igDefCMC).SetFocus
    End If
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
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
    Dim ilNoButtons As Integer
    Dim ilLoop As Integer
    
    imFirstActivate = True
    imFirstTime = True
    plcMsg.Move (EngrGenMsg.Width - plcMsg.Width) \ 2, 120
    plcMsg.Caption = sgGenMsg
    For ilLoop = 0 To 3 Step 1
        If sgCMCTitle(ilLoop) <> "" Then
            cmcButton(ilLoop).Caption = sgCMCTitle(ilLoop)
            ilNoButtons = ilNoButtons + 1
        End If
    Next ilLoop
    If ilNoButtons = 2 Then
        If igEditBox = 1 Then
            cmcButton(0).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width - 60, edcEditValue.Top + edcEditValue.Height + 225
        Else
            cmcButton(0).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width - 60, plcMsg.Top + plcMsg.Height + 225
        End If
        cmcButton(1).Move plcMsg.Left + plcMsg.Width \ 2 + 60, cmcButton(0).Top
        cmcButton(0).Visible = True
        cmcButton(1).Visible = True
        cmcButton(2).Visible = False
        cmcButton(3).Visible = False
    ElseIf ilNoButtons = 3 Then
        If igEditBox = 1 Then
            cmcButton(1).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width \ 2, edcEditValue.Top + edcEditValue.Height + 225
        Else
            cmcButton(1).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width \ 2, plcMsg.Top + plcMsg.Height + 225
        End If
        cmcButton(0).Move cmcButton(1).Left - cmcButton(1).Width - 120, cmcButton(1).Top
        cmcButton(2).Move cmcButton(1).Left + cmcButton(1).Width + 120, cmcButton(1).Top
        cmcButton(0).Visible = True
        cmcButton(1).Visible = True
        cmcButton(2).Visible = True
        cmcButton(3).Visible = False
    ElseIf ilNoButtons = 4 Then
        If igEditBox = 1 Then
            cmcButton(0).Move plcMsg.Left, edcEditValue.Top + edcEditValue.Height + 225
        Else
            cmcButton(0).Move plcMsg.Left, plcMsg.Top + plcMsg.Height + 225
        End If
        cmcButton(1).Move cmcButton(0).Left + cmcButton(0).Width + 120, cmcButton(0).Top
        cmcButton(2).Move cmcButton(1).Left + cmcButton(1).Width + 120, cmcButton(0).Top
        cmcButton(3).Move cmcButton(2).Left + cmcButton(2).Width + 120, cmcButton(0).Top
        cmcButton(0).Visible = True
        cmcButton(1).Visible = True
        cmcButton(2).Visible = True
        cmcButton(3).Visible = True
    Else
        If igEditBox = 1 Then
            cmcButton(0).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width \ 2, edcEditValue.Top + edcEditValue.Height + 225
        Else
            cmcButton(0).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width \ 2, plcMsg.Top + plcMsg.Height + 225
        End If
        cmcButton(0).Visible = True
        cmcButton(1).Visible = False
        cmcButton(2).Visible = False
        cmcButton(3).Visible = False
    End If
    If igEditBox = 1 Then
        plcEditValue.Caption = sgMsgEditValue
        plcEditValue.Visible = True
        edcEditValue.Text = sgEditValue
        edcEditValue.Visible = True
    End If
    Screen.MousePointer = vbHourglass
    EngrGenMsg.Height = cmcButton(0).Top + 5 * cmcButton(0).Height / 3
    'gCenterModalForm GenMsg
    gCenterForm EngrGenMsg
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
    Unload EngrGenMsg
    Set EngrGenMsg = Nothing   'Remove data segment
End Sub
