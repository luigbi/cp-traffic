VERSION 5.00
Begin VB.Form SpotLine 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   900
   ClientTop       =   2280
   ClientWidth     =   7110
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
   ScaleHeight     =   3705
   ScaleWidth      =   7110
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
      Left            =   3675
      TabIndex        =   3
      Top             =   3270
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3270
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
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   30
      Width           =   3240
   End
   Begin VB.PictureBox plcLine 
      ForeColor       =   &H00000000&
      Height          =   2700
      Left            =   180
      ScaleHeight     =   2640
      ScaleWidth      =   6705
      TabIndex        =   1
      Top             =   285
      Width           =   6765
      Begin VB.PictureBox pbcLbcLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   2475
         Left            =   75
         ScaleHeight     =   2475
         ScaleWidth      =   6570
         TabIndex        =   6
         Top             =   75
         Width           =   6570
      End
      Begin VB.ListBox lbcLine 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2550
         Left            =   30
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   45
         Width           =   6645
      End
   End
End
Attribute VB_Name = "SpotLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Spotline.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpotLine.Frm
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
'Dim imListField(1 To 5) As Integer
Dim imListField(0 To 5) As Integer
Dim imLBCtrls As Integer



Private Sub cmcCancel_Click()
    igSpotLineReturn = -1
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim slStr As String
    Dim ilRet As Integer
    Dim slLine As String
    slStr = lbcLine.List(lbcLine.ListIndex)
    ilRet = gParseItem(slStr, 1, "|", slLine)
    ilRet = gParseItem(slLine, 2, "#", slLine)
    igSpotLineReturn = Val(Trim$(slLine))
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

Private Sub Form_Unload(Cancel As Integer)
    Set SpotLine = Nothing   'Remove data segment
End Sub

Private Sub lbcLine_Click()
    If lbcLine.ListIndex >= 0 Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
    End If
    pbcLbcLine_Paint
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
    SpotLine.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterModalForm SpotLine
    gCenterModalForm SpotLine
    imLBCtrls = 1
    imListField(1) = 15
    imListField(2) = 12 * igAlignCharWidth
    imListField(3) = 26 * igAlignCharWidth
    imListField(4) = 46 * igAlignCharWidth
    imListField(5) = 110 * igAlignCharWidth
    pbcLbcLine_Paint
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
    Unload SpotLine
    igManUnload = NO
End Sub

Private Sub lbcLine_Scroll()
    pbcLbcLine_Paint
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcLine_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLineEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 3) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilLineEnd = lbcLine.TopIndex + lbcLine.Height \ fgListHtArial825
    If ilLineEnd > lbcLine.ListCount Then
        ilLineEnd = lbcLine.ListCount
    End If
    If lbcLine.ListCount <= lbcLine.Height \ fgListHtArial825 Then
        llWidth = lbcLine.Width - 30
    Else
        llWidth = lbcLine.Width - igScrollBarWidth - 30
    End If
    pbcLbcLine.Width = llWidth
    pbcLbcLine.Cls
    llFgColor = pbcLbcLine.ForeColor
    For ilLoop = lbcLine.TopIndex To ilLineEnd - 1 Step 1
        pbcLbcLine.ForeColor = llFgColor
        If lbcLine.MultiSelect = 0 Then
            If lbcLine.ListIndex = ilLoop Then
                gPaintArea pbcLbcLine, CSng(0), CSng((ilLoop - lbcLine.TopIndex) * fgListHtArial825), CSng(pbcLbcLine.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLine.ForeColor = vbWhite
            End If
        Else
            If lbcLine.Selected(ilLoop) Then
                gPaintArea pbcLbcLine, CSng(0), CSng((ilLoop - lbcLine.TopIndex) * fgListHtArial825), CSng(pbcLbcLine.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLine.ForeColor = vbWhite
            End If
        End If
        slStr = lbcLine.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcLine.CurrentX = imListField(ilField)
            pbcLbcLine.CurrentY = (ilLoop - lbcLine.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcLine, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcLine.Print slStr
        Next ilField
        pbcLbcLine.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Extra Bonus Line Selection"
End Sub
