VERSION 5.00
Begin VB.Form GenSch 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   4485
   ClientTop       =   1470
   ClientWidth     =   3030
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
   ScaleHeight     =   1095
   ScaleWidth      =   3030
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   225
      Top             =   525
   End
   Begin VB.PictureBox pbcSchMess 
      Height          =   1110
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   2970
      TabIndex        =   1
      Top             =   0
      Width           =   3030
      Begin VB.Label lacSchMess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2955
         WordWrap        =   -1  'True
      End
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2310
      Width           =   120
   End
End
Attribute VB_Name = "GenSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GenSch.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: GenSch.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text

Dim imDelayReason As Integer
Dim imFirstActivate As Integer

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
    imFirstActivate = True
    lacSchMess.Caption = sgGenMsg
    Screen.MousePointer = vbHourglass
    gCenterModalForm GenSch
    tmcDelay.Enabled = True
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
    Unload GenSch
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set GenSch = Nothing   'Remove data segment
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub tmcDelay_Timer()
    Dim ilRet As Integer

    tmcDelay.Enabled = False
    If imDelayReason = 0 Then
        imDelayReason = 1
        If igDefCMC = 1 Then
            ilRet = gFeedSchSpots(False)
        ElseIf igDefCMC = 2 Then
            ilRet = gCheckAvailAttributes(igAnsCMC)
        End If
        lacSchMess.Caption = sgLF & sgCR & "Completed..."
        tmcDelay.Enabled = True
        Exit Sub
    End If
    pbcSchMess.Visible = False
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
