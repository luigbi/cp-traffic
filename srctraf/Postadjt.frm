VERSION 5.00
Begin VB.Form PostAdjt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   420
   ClientTop       =   1395
   ClientWidth     =   6885
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
   ScaleHeight     =   5490
   ScaleWidth      =   6885
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   180
      Left            =   45
      Picture         =   "Postadjt.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   105
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
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5280
      Width           =   90
   End
   Begin VB.PictureBox pbcIBSTab 
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
      Height          =   120
      Left            =   15
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   780
      Width           =   105
   End
   Begin VB.PictureBox pbcIBTab 
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
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   3810
      Width           =   60
   End
   Begin VB.VScrollBar vbcItemBill 
      Height          =   3480
      LargeChange     =   15
      Left            =   6285
      TabIndex        =   9
      Top             =   1140
      Width           =   240
   End
   Begin VB.PictureBox pbcAdj 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3510
      Left            =   420
      Picture         =   "Postadjt.frx":030A
      ScaleHeight     =   3510
      ScaleWidth      =   5880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1095
      Width           =   5880
      Begin VB.Label lacAdjFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   240
         Left            =   0
         TabIndex        =   14
         Top             =   450
         Visible         =   0   'False
         Width           =   5865
      End
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   795
      TabIndex        =   10
      Top             =   5115
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2175
      TabIndex        =   11
      Top             =   5115
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4830
      TabIndex        =   13
      Top             =   5115
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3510
      TabIndex        =   12
      Top             =   5115
      Width           =   1050
   End
   Begin VB.PictureBox plcSelect 
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
      Height          =   750
      Left            =   375
      ScaleHeight     =   690
      ScaleWidth      =   6150
      TabIndex        =   1
      Top             =   240
      Width           =   6210
      Begin VB.ComboBox cbcSelect 
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
         Height          =   300
         Left            =   60
         TabIndex        =   3
         Top             =   405
         Width           =   6090
      End
      Begin VB.ComboBox cbcAdvt 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2130
         TabIndex        =   2
         Top             =   60
         Width           =   4020
      End
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   -15
      Width           =   2295
   End
   Begin VB.PictureBox plcAdj 
      ForeColor       =   &H00000000&
      Height          =   3660
      Left            =   375
      ScaleHeight     =   3600
      ScaleWidth      =   6165
      TabIndex        =   5
      Top             =   1035
      Width           =   6225
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5025
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacTotals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5820
      TabIndex        =   15
      Top             =   4755
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6150
      Picture         =   "Postadjt.frx":B6B4
      Top             =   4905
      Width           =   480
   End
End
Attribute VB_Name = "PostAdjt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Postadjt.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostAdjt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Post Adjustments screen code
Option Explicit
Option Compare Text
Dim imUpdateAllowed As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
End Sub

Private Sub cmcUpdate_Click()
    If Not imUpdateAllowed Then
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
    'gShowBranner
    If (igWinStatus(INVOICESJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        If tgUrf(0).iSlfCode > 0 Then
            imUpdateAllowed = False
        Else
            imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed
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
        'If (cbcSelect.Enabled) And (imBoxNo > 0) Then
        '    cbcSelect.Enabled = False
        '    ilReSet = True
        'Else
        '    ilReSet = False
        'End If
        gFunctionKeyBranch KeyCode
        'If imCEBoxNo > 0 Then
        '    mCEEnableBox imCEBoxNo
        'ElseIf imDMBoxNo > 0 Then
        '    mDmEnableBox imDMBoxNo
        'Else
        '    mEnableBox imBoxNo
        'End If
        'If ilReSet Then
        '    cbcSelect.Enabled = True
        'End If
    End If

End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = cbcAdvt.ListIndex
    If ilIndex >= 0 Then
        slName = cbcAdvt.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(PostAdjt, cbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(PostAdjt, cbcAdvt, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", PostAdjt
        On Error GoTo 0
'        cbcAdvt.AddItem "[New]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcAdvt
            If gLastFound(cbcAdvt) >= 0 Then
                cbcAdvt.ListIndex = gLastFound(cbcAdvt)
            Else
                cbcAdvt.ListIndex = -1
            End If
        Else
            cbcAdvt.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    PostAdjt.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterModalForm PostAdjt
    gCenterStdAlone PostAdjt
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    igManUnload = YES
    Unload PostAdjt
    Set PostAdjt = Nothing
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcAdj_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Invoice Post Adjustments"
End Sub
