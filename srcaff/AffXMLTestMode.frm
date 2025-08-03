VERSION 5.00
Begin VB.Form frmXMLTestMode 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2505
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
   ScaleHeight     =   2505
   ScaleWidth      =   7140
   Begin VB.TextBox edcEndTag 
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
      Left            =   1380
      TabIndex        =   6
      Top             =   1365
      Width           =   5310
   End
   Begin VB.TextBox edcData 
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
      Left            =   1380
      TabIndex        =   4
      Top             =   900
      Width           =   5310
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6495
      Top             =   1980
   End
   Begin VB.TextBox edcStartTag 
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
      Left            =   1380
      TabIndex        =   2
      Top             =   435
      Width           =   5310
   End
   Begin VB.CommandButton cmcSend 
      Appearance      =   0  'Flat
      Caption         =   "&Send- Don't Ask"
      Height          =   285
      Index           =   1
      Left            =   3585
      TabIndex        =   8
      Top             =   2040
      Width           =   1905
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2310
      Width           =   120
   End
   Begin VB.CommandButton cmcSend 
      Appearance      =   0  'Flat
      Caption         =   "&Send-Ask"
      Default         =   -1  'True
      Height          =   285
      Index           =   0
      Left            =   1215
      TabIndex        =   7
      Top             =   2040
      Width           =   1905
   End
   Begin VB.Label lacForm 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   5235
   End
   Begin VB.Label Label2 
      Caption         =   "End Tag"
      Height          =   300
      Left            =   225
      TabIndex        =   5
      Top             =   1350
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Data"
      Height          =   300
      Left            =   225
      TabIndex        =   3
      Top             =   885
      Width           =   1125
   End
   Begin VB.Label lacStartTag 
      Caption         =   "Start Tag"
      Height          =   300
      Left            =   225
      TabIndex        =   1
      Top             =   420
      Width           =   1125
   End
End
Attribute VB_Name = "frmXMLTestMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: frmXMLTestMode.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
Dim imFirstTime As Integer
Dim imFirstActivate As Integer

Private Sub cmcSend_Click(Index As Integer)
    igAnsCMC = Index
    sgEditValue = edcStartTag.Text & "|" & edcData.Text & "|" & edcEndTag.Text
    mTerminate
End Sub

Private Sub edcData_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEndTag_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStartTag_GotFocus()
    gCtrlGotFocus ActiveControl
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
    ReDim slFields(0 To 3) As String
    Dim ilRet As Integer

    imFirstActivate = True
    imFirstTime = True
    Screen.MousePointer = vbHourglass
    ilRet = gParseItem(sgEditValue, 1, "|", slFields(0))
    ilRet = gParseItem(sgEditValue, 2, "|", slFields(1))
    ilRet = gParseItem(sgEditValue, 3, "|", slFields(2))
    ilRet = gParseItem(sgEditValue, 4, "|", slFields(3))
    lacForm.Caption = slFields(0)
    If slFields(1) = "" Then
        edcStartTag.Enabled = False
    End If
    edcStartTag.Text = slFields(1)
    If slFields(2) = "" Then
        edcData.Enabled = False
    End If
    edcData.Text = slFields(2)
    If slFields(3) = "" Then
        edcEndTag.Enabled = False
    End If
    edcEndTag.Text = slFields(3)
    frmXMLTestMode.Height = cmcSend(0).Top + 5 * cmcSend(0).Height / 3
    gCenterForm frmXMLTestMode
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
    Unload frmXMLTestMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmXMLTestMode = Nothing   'Remove data segment
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
End Sub
