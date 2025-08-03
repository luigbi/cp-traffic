VERSION 5.00
Begin VB.Form frmGetPath 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3825
   ClientLeft      =   4485
   ClientTop       =   1470
   ClientWidth     =   7830
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
   ScaleHeight     =   3825
   ScaleWidth      =   7830
   Begin VB.DirListBox lbcFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   1320
      TabIndex        =   4
      Top             =   765
      Width           =   6000
   End
   Begin VB.DriveListBox cbcDrive 
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
      Height          =   315
      Left            =   1335
      TabIndex        =   3
      Top             =   270
      Width           =   3045
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   375
      Top             =   2940
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4065
      TabIndex        =   2
      Top             =   3390
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
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   285
      Left            =   2475
      TabIndex        =   0
      Top             =   3390
      Width           =   1170
   End
   Begin VB.Label lacFolder 
      Caption         =   "Folder:"
      Height          =   285
      Left            =   195
      TabIndex        =   6
      Top             =   795
      Width           =   1005
   End
   Begin VB.Label lacDrive 
      Caption         =   "Drive:"
      Height          =   210
      Left            =   195
      TabIndex        =   5
      Top             =   285
      Width           =   990
   End
End
Attribute VB_Name = "frmGetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: frmGetPath.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Private smInitPath As String



Private Sub cbcDrive_Change()
    lbcFolder.Path = cbcDrive.Drive
End Sub

Private Sub cmcCancel_Click()
    igGetPath = 1
    mTerminate
End Sub

Private Sub cmcDone_Click()
    Dim slStr As String
    Dim ilPos As Integer
    
    slStr = lbcFolder.Path
    If right$(slStr, 1) <> "\" Then
        sgGetPath = lbcFolder.Path & "\"
    Else
        sgGetPath = lbcFolder.Path
    End If
    If igPathType = 1 Then
        ilPos = InStr(1, Trim$(UCase$(sgGetPath)), Trim$(UCase$(smInitPath)), vbTextCompare)
        If ilPos <> 1 Then
            MsgBox "Selected folder must be a subfolder of " & smInitPath
            Exit Sub
        End If
    End If
    igGetPath = 0
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
    Dim ilPos As Integer
    Dim slStr As String
    Dim slDrive As String
    Dim slPath As String

    smInitPath = sgGetPath
    ilPos = InStr(sgGetPath, ":")
    If ilPos > 0 Then
        slDrive = Left(sgGetPath, ilPos)
        slPath = Mid$(sgGetPath, ilPos + 1)
        If right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        End If
        On Error Resume Next
        cbcDrive.Drive = slDrive
        lbcFolder.Path = slPath
        On Error GoTo 0
    End If
    imFirstActivate = True
    gCenterForm frmGetPath
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
    Unload frmGetPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGetPath = Nothing   'Remove data segment
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
End Sub
