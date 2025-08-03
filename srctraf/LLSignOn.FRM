VERSION 5.00
Begin VB.Form LLSignOn 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2175
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   6420
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
   ScaleHeight     =   2175
   ScaleWidth      =   6420
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   1785
      Width           =   1050
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
      Height          =   105
      Left            =   45
      ScaleHeight     =   105
      ScaleWidth      =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1965
      Width           =   15
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Sign-On"
      Height          =   285
      Left            =   1845
      TabIndex        =   7
      Top             =   1785
      Width           =   1050
   End
   Begin VB.PictureBox plcInv 
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
      Height          =   1020
      Left            =   105
      ScaleHeight     =   960
      ScaleWidth      =   6120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   6180
      Begin VB.TextBox edcStudioID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   6
         Top             =   510
         Width           =   1155
      End
      Begin VB.TextBox edcSignOn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   4
         Top             =   75
         Width           =   4680
      End
      Begin VB.Label lacStudioID 
         Appearance      =   0  'Flat
         Caption         =   "Studio ID:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   555
         Width           =   975
      End
      Begin VB.Label lacSignOn 
         Appearance      =   0  'Flat
         Caption         =   "Name:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label lacReplace 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   345
      Width           =   6270
   End
   Begin VB.Label lacScreen 
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6285
   End
End
Attribute VB_Name = "LLSignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of LLSignOn.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LLSignOn.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract number screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer

Private Sub cmcCancel_Click()
    igTerminateReturn = 0   '0=Canceled
    sgLLName = ""
    sgStudioID = ""
    mTerminate
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    sgLLName = edcSignOn.Text
    sgStudioID = edcStudioID.Text
    If (Trim$(sgLLName) = "") Then
       Beep
       edcSignOn.SetFocus
       Exit Sub
    'ElseIf (Trim$(sgStudioID) = "") Then
    '   Beep
    '   edcStudioID.SetFocus
    '   Exit Sub
    End If
    igTerminateReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slStr                                                   *
'******************************************************************************************


    gCtrlGotFocus cmcDone
End Sub

Private Sub edcStudioID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStudioID_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Sub edcSignOn_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSignOn_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
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

    'If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
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
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    LLSignOn.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'LLSignOn.Left = LiveLog.Left + LiveLog.Width / 2 - LLSignOn.Width / 2
    'LLSignOn.Top = LiveLog.Top + LiveLog.Height / 2 - LLSignOn.Height / 2
    gCenterStdAlone LLSignOn
    lacScreen.Caption = sgSignOnMsg
    If sgReplaceMsg = "" Then
        lacReplace.Visible = False
        cmcDone.Caption = "Sign-On"
    Else
        lacReplace.Visible = True
        lacReplace.Caption = sgReplaceMsg
        cmcDone.Caption = "Replace"
    End If
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
    Unload LLSignOn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LLSignOn = Nothing   'Remove data segment
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub
