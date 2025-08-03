VERSION 5.00
Begin VB.Form frmGenMsg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   4485
   ClientTop       =   1470
   ClientWidth     =   7140
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
   ScaleHeight     =   3090
   ScaleWidth      =   7140
   Begin VB.PictureBox pbcArial 
      Height          =   225
      Left            =   5640
      ScaleHeight     =   165
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6615
      Top             =   2505
   End
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
      Left            =   120
      TabIndex        =   6
      Top             =   1995
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   3
      Left            =   3990
      TabIndex        =   4
      Top             =   2670
      Width           =   1170
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   2670
      Width           =   1170
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   2
      Top             =   2670
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
      Default         =   -1  'True
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2670
      Width           =   1170
   End
   Begin VB.Frame frcRadio 
      Caption         =   "Frame1"
      Height          =   825
      Left            =   120
      TabIndex        =   7
      Top             =   1725
      Visible         =   0   'False
      Width           =   6450
      Begin VB.OptionButton rbcRadio 
         Caption         =   "Radio3"
         Height          =   225
         Index           =   3
         Left            =   1170
         TabIndex        =   11
         Top             =   510
         Width           =   1155
      End
      Begin VB.OptionButton rbcRadio 
         Caption         =   "Radio2"
         Height          =   225
         Index           =   2
         Left            =   75
         TabIndex        =   10
         Top             =   510
         Width           =   1155
      End
      Begin VB.OptionButton rbcRadio 
         Caption         =   "Radio1"
         Height          =   225
         Index           =   1
         Left            =   1170
         TabIndex        =   9
         Top             =   225
         Width           =   1155
      End
      Begin VB.OptionButton rbcRadio 
         Caption         =   "Radio0"
         Height          =   225
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.Label plcMsg 
      Height          =   1710
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6885
   End
End
Attribute VB_Name = "frmGenMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: frmGenMsg.Frm
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
    If igEditBox = 2 Then
        sgEditValue = ""
        If rbcRadio(0).Value Then
            sgEditValue = "0"
        ElseIf rbcRadio(1).Value Then
            sgEditValue = "1"
        ElseIf rbcRadio(2).Value Then
            sgEditValue = "2"
        ElseIf rbcRadio(3).Value Then
            sgEditValue = "3"
        End If
        If sgEditValue = "" Then
            MsgBox "Selection missing", vbExclamation + vbOKOnly, "Warning"
            Exit Sub
        End If
    End If
    mTerminate
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
    If imFirstTime Then
        imFirstTime = True
        'Moved to tmcStart
        'cmcButton(igDefCMC).SetFocus
    End If
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
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
    Dim llWidth As Long
    Dim ilEditValue As Integer

    imFirstActivate = True
    imFirstTime = True
    plcMsg.Move (frmGenMsg.Width - plcMsg.Width) \ 2, 120
    plcMsg.Caption = sgGenMsg
    llWidth = 0
    For ilLoop = 0 To 3 Step 1
        If sgCMCTitle(ilLoop) <> "" Then
            cmcButton(ilLoop).Caption = sgCMCTitle(ilLoop)
            If frmGenMsg.TextWidth(sgCMCTitle(ilLoop)) > cmcButton(ilLoop).Width Then
                cmcButton(ilLoop).Width = frmGenMsg.TextWidth(sgCMCTitle(ilLoop)) + 300
            End If
            llWidth = llWidth + cmcButton(ilLoop).Width + 60
            ilNoButtons = ilNoButtons + 1
        End If
    Next ilLoop
    llWidth = llWidth - 60
    If igEditBox <> 2 Then
        cmcButton(0).Move plcMsg.Left + (plcMsg.Width - llWidth) \ 2, plcMsg.Top + plcMsg.Height + 2 * edcEditValue.Height
    Else
        cmcButton(0).Move plcMsg.Left + (plcMsg.Width - llWidth) \ 2, plcMsg.Top + plcMsg.Height + frcRadio.Height + 225
    End If
    If ilNoButtons = 2 Then
        'cmcButton(0).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width - 60, plcMsg.Top + plcMsg.Height + 225
        'cmcButton(1).Move plcMsg.Left + plcMsg.Width \ 2 + 60, cmcButton(0).Top
        cmcButton(1).Move cmcButton(0).Left + cmcButton(0).Width + 60, cmcButton(0).Top
        cmcButton(0).Visible = True
        cmcButton(1).Visible = True
        cmcButton(2).Visible = False
        cmcButton(3).Visible = False
    ElseIf ilNoButtons = 3 Then
        'cmcButton(1).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width \ 2, plcMsg.Top + plcMsg.Height + 225
        'cmcButton(0).Move cmcButton(1).Left - cmcButton(1).Width - 120, cmcButton(1).Top
        'cmcButton(2).Move cmcButton(1).Left + cmcButton(1).Width + 120, cmcButton(1).Top
        cmcButton(1).Move cmcButton(0).Left + cmcButton(0).Width + 60, cmcButton(0).Top
        cmcButton(2).Move cmcButton(1).Left + cmcButton(1).Width + 60, cmcButton(0).Top
        cmcButton(0).Visible = True
        cmcButton(1).Visible = True
        cmcButton(2).Visible = True
        cmcButton(3).Visible = False
    ElseIf ilNoButtons = 4 Then
        'cmcButton(0).Move plcMsg.Left, plcMsg.Top + plcMsg.Height + 225
        'cmcButton(1).Move cmcButton(0).Left + cmcButton(0).Width + 120, cmcButton(0).Top
        'cmcButton(2).Move cmcButton(1).Left + cmcButton(1).Width + 120, cmcButton(0).Top
        'cmcButton(3).Move cmcButton(2).Left + cmcButton(2).Width + 120, cmcButton(0).Top
        cmcButton(1).Move cmcButton(0).Left + cmcButton(0).Width + 60, cmcButton(0).Top
        cmcButton(2).Move cmcButton(1).Left + cmcButton(1).Width + 60, cmcButton(0).Top
        cmcButton(3).Move cmcButton(2).Left + cmcButton(2).Width + 60, cmcButton(0).Top
        cmcButton(0).Visible = True
        cmcButton(1).Visible = True
        cmcButton(2).Visible = True
        cmcButton(3).Visible = True
    Else
        'cmcButton(0).Move plcMsg.Left + plcMsg.Width \ 2 - cmcButton(0).Width \ 2, plcMsg.Top + plcMsg.Height + 225
        cmcButton(0).Visible = True
        cmcButton(1).Visible = False
        cmcButton(2).Visible = False
        cmcButton(3).Visible = False
    End If
    If igEditBox = 1 Then
        edcEditValue.Text = sgEditValue
        edcEditValue.Visible = True
    End If
    If igEditBox = 2 Then
'        If sgEditValue <> "" Then
'            frcRadio.Caption = sgEditValue
'        Else
            frcRadio.Caption = ""
'        End If
        rbcRadio(0).Caption = sgRadioTitle(0)
        rbcRadio(1).Caption = sgRadioTitle(1)
        rbcRadio(2).Caption = sgRadioTitle(2)
        rbcRadio(3).Caption = sgRadioTitle(3)
        rbcRadio(0).Width = pbcArial.TextWidth(sgRadioTitle(0)) + 360
        rbcRadio(1).Width = pbcArial.TextWidth(sgRadioTitle(1)) + 360
        rbcRadio(1).Left = rbcRadio(0).Left + rbcRadio(0).Width
        If sgRadioTitle(2) <> "" Then
            rbcRadio(2).Width = pbcArial.TextWidth(sgRadioTitle(2)) + 360
            If sgRadioTitle(3) <> "" Then
                rbcRadio(3).Width = pbcArial.TextWidth(sgRadioTitle(3)) + 360
                rbcRadio(3).Left = rbcRadio(2).Left + rbcRadio(2).Width
            Else
                rbcRadio(3).Visible = False
            End If
        Else
            rbcRadio(2).Visible = False
            rbcRadio(3).Visible = False
        End If
        If sgEditValue <> "" Then
            ilEditValue = Val(sgEditValue)
            If (ilEditValue >= 0) And (ilEditValue <= 3) Then
                rbcRadio(ilEditValue).Value = True
            End If
        End If
        frcRadio.Visible = True
    Else
        frcRadio.Visible = False
    End If
    Screen.MousePointer = vbHourglass
    frmGenMsg.Height = cmcButton(0).Top + 5 * cmcButton(0).Height / 3
    'gCenterModalForm frmGenMsg
    gCenterForm frmGenMsg
    tmcStart.Enabled = True
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
    Unload frmGenMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGenMsg = Nothing   'Remove data segment
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If (igDefCMC >= 0) And (igDefCMC <= 3) Then
        cmcButton(igDefCMC).SetFocus
    End If
End Sub
