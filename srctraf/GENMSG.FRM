VERSION 5.00
Begin VB.Form GenMsg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2880
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
   ScaleHeight     =   2880
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frcRadio 
      Height          =   1575
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   6495
      Begin VB.OptionButton rdcOption 
         Caption         =   "Option4"
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton rdcOption 
         Caption         =   "Option3"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton rdcOption 
         Caption         =   "Option1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton rdcOption 
         Caption         =   "Option2"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox edcMsg 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   6705
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6495
      Top             =   1980
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
      Left            =   495
      TabIndex        =   5
      Top             =   1920
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
      Top             =   2460
      Width           =   1170
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   2460
      Width           =   1170
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   2
      Top             =   2460
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
      Top             =   2460
      Width           =   1170
   End
   Begin VB.Label plcMsg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   6705
   End
End
Attribute VB_Name = "GenMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GENMSG.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: GenMsg.Frm
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
    '9459 radio buttons? Don't set igAnsCMC
    If Not bgUseRadioButtons Then
        igAnsCMC = Index
    'radio buttons, and chose 'cancel'
    ElseIf Index = 1 Then
        igAnsCMC = -1
    End If
    If igEditBox = 1 Then
        sgEditValue = edcEditValue.Text
    End If
    mTerminate
End Sub
Private Sub rdcOption_Click(Index As Integer)
    '9459
    igAnsCMC = Index
End Sub
Private Sub edcEditValue_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    If igEditBox = 1 And lgEditBoxMaxValue > 0 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcEditValue.Text
        slStr = Left$(slStr, edcEditValue.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcEditValue.SelStart - edcEditValue.SelLength)
        If gCompNumberStr(slStr, Str(lgEditBoxMaxValue)) > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
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
    Dim ilNoButtons As Integer
    Dim illoop As Integer
    Dim llWidth As Long
    
    imFirstActivate = True
    imFirstTime = True
    plcMsg.Move (GenMsg.Width - plcMsg.Width) \ 2, 120
    plcMsg.Caption = sgGenMsg
    edcMsg.Move plcMsg.Left, plcMsg.Top, plcMsg.Width, plcMsg.Height
    edcMsg.Text = sgGenMsg
    If Len(sgGenMsg) <= 500 Then
        plcMsg.Visible = True
        edcMsg.Visible = False
    Else
        plcMsg.Visible = False
        edcMsg.Visible = True
    End If
    '9459 radio buttons
    If bgUseRadioButtons Then
     '   frcRadio.Move plcMsg.Left, plcMsg.Top + plcMsg.Height + 10
        frcRadio.Visible = True
        For illoop = 0 To 3 Step 1
            If sgCMCTitle(illoop) <> "" Then
                rdcOption(illoop).Caption = sgCMCTitle(illoop)
                If GenMsg.TextWidth(sgCMCTitle(illoop)) + 300 > rdcOption(illoop).Width Then
                    rdcOption(illoop).Width = GenMsg.TextWidth(sgCMCTitle(illoop)) + 300
                End If
                llWidth = llWidth + rdcOption(illoop).Width + 60
                ilNoButtons = ilNoButtons + 1
            End If
        Next illoop
        Select Case ilNoButtons
            Case 2
                rdcOption(3).Visible = False
                rdcOption(2).Visible = False
            Case 3
                rdcOption(3).Visible = False
            Case 4
        End Select
        ilNoButtons = 2
        llWidth = 2500 '3975
        cmcButton(0).Caption = "OK"
        cmcButton(1).Caption = "Cancel"
        igEditBox = 0
    Else
        '9459
        frcRadio.Visible = False
        edcEditValue.Move plcMsg.Left, plcMsg.Top + plcMsg.Height + 120
        llWidth = 0
        For illoop = 0 To 3 Step 1
            If sgCMCTitle(illoop) <> "" Then
                cmcButton(illoop).Caption = sgCMCTitle(illoop)
                If GenMsg.TextWidth(sgCMCTitle(illoop)) + 300 > cmcButton(illoop).Width Then
                    cmcButton(illoop).Width = GenMsg.TextWidth(sgCMCTitle(illoop)) + 300
                End If
                llWidth = llWidth + cmcButton(illoop).Width + 60
                ilNoButtons = ilNoButtons + 1
            End If
        Next illoop
    End If
    llWidth = llWidth - 60
    cmcButton(0).Move plcMsg.Left + (plcMsg.Width - llWidth) \ 2, edcEditValue.Top + edcEditValue.Height + 225
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
        If igEditBoxMaxCharacters >= 0 Then
            edcEditValue.MaxLength = igEditBoxMaxCharacters
        End If
    End If
    Screen.MousePointer = vbHourglass
    GenMsg.Height = cmcButton(0).Top + 5 * cmcButton(0).Height / 3
    'gCenterModalForm GenMsg
    gCenterModalForm GenMsg
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
    Unload GenMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    igEditBoxMaxCharacters = 0
    lgEditBoxMaxValue = 0

    Set GenMsg = Nothing   'Remove data segment
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub



Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If (igDefCMC >= 0) And (igDefCMC <= 3) Then
        On Error Resume Next
        '9459 radio buttons
        If bgUseRadioButtons Then
            rdcOption(igDefCMC).SetFocus
        Else
            cmcButton(igDefCMC).SetFocus
        End If
    End If
End Sub
