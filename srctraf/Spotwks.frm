VERSION 5.00
Begin VB.Form SpotWks 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2685
   ClientLeft      =   2115
   ClientTop       =   2160
   ClientWidth     =   4155
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
   ScaleHeight     =   2685
   ScaleWidth      =   4155
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
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
      Left            =   2175
      TabIndex        =   3
      Top             =   2295
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
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
      Left            =   900
      TabIndex        =   2
      Top             =   2295
      Width           =   1050
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
      Height          =   465
      Left            =   30
      ScaleHeight     =   465
      ScaleWidth      =   3990
      TabIndex        =   0
      Top             =   30
      Width           =   3990
   End
   Begin VB.PictureBox plcLine 
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   180
      ScaleHeight     =   1530
      ScaleWidth      =   3720
      TabIndex        =   1
      Top             =   540
      Width           =   3780
      Begin VB.TextBox edcNoWks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2610
         MaxLength       =   2
         TabIndex        =   8
         Top             =   825
         Width           =   825
      End
      Begin VB.Label lacDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1995
         TabIndex        =   7
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label lacWks 
         Appearance      =   0  'Flat
         Caption         =   "Number of Weeks to Extend"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   855
         Width           =   2475
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date of Extend"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   1770
      End
   End
End
Attribute VB_Name = "SpotWks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Spotwks.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpotWks.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Price Grid Calculate input screen code
Option Explicit
Option Compare Text
'Media Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imUpdateAllowed As Integer

Dim smScreenCaption As String

Private Sub cmcCancel_Click()
    igSpotWksReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    slStr = edcNoWks.Text
    If slStr = "" Then
        igSpotWksReturn = 0
    Else
        igSpotWksReturn = Val(slStr)
    End If
    mTerminate
End Sub
Private Sub edcNoWks_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNoWks.Text
    slStr = Left$(slStr, edcNoWks.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNoWks.SelStart - edcNoWks.SelLength)
    slComp = "99"
    If gCompNumberStr(slStr, slComp) > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(SPOTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        plcLine.Visible = False
        plcLine.Visible = True
    End If
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
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    SpotWks.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterModalForm SpotWks
    smScreenCaption = "Extend Weeks for " & sgSpotWksVehName
    lacDate.Caption = sgSpotWksStartDate
    edcNoWks.Text = "1"
    gCenterModalForm SpotWks
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
    igManUnload = YES
    Unload SpotWks
    igManUnload = NO
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SpotWks = Nothing   'Remove data segment
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub
