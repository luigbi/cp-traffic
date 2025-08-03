VERSION 5.00
Begin VB.Form frmReportQueueDescription 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   4485
   ClientTop       =   1470
   ClientWidth     =   8055
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
   ScaleHeight     =   2310
   ScaleWidth      =   8055
   Begin VB.TextBox edcDescription 
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
      Left            =   1410
      MaxLength       =   100
      TabIndex        =   4
      Top             =   675
      Width           =   6285
   End
   Begin VB.CommandButton cmcButton 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   2
      Top             =   1800
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
      Caption         =   "&Add to Queue"
      Default         =   -1  'True
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label lacDescription 
      Caption         =   "Description"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   675
      Width           =   1335
   End
   Begin VB.Label lacNumberWaiting 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1245
      Width           =   7605
   End
   Begin VB.Label lacRptName 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6885
   End
End
Attribute VB_Name = "frmReportQueueDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AffReportQueueDescription.Frm
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
    If Index = 1 Then
        igRQReturnStatus = 0
    Else
        igRQReturnStatus = 1
        sgRQDescription = edcDescription.Text
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
    lacRptName.Caption = sgRQReportName
    
    Screen.MousePointer = vbHourglass
    gCenterForm frmReportQueueDescription
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
    Unload frmReportQueueDescription
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReportQueueDescription = Nothing   'Remove data segment
End Sub
