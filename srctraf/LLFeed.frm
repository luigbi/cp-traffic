VERSION 5.00
Begin VB.Form LLFeed 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   1170
   ClientTop       =   2070
   ClientWidth     =   6930
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
   ScaleHeight     =   3630
   ScaleWidth      =   6930
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1410
      Left            =   3900
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1695
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
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
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   15
         Picture         =   "LLFeed.frx":0000
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "LLFeed.frx":0CBE
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   16
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
      Left            =   3660
      TabIndex        =   13
      Top             =   3225
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
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
      Left            =   2040
      TabIndex        =   12
      Top             =   3225
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
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   1980
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1980
   End
   Begin VB.PictureBox plcTimes 
      ForeColor       =   &H00000000&
      Height          =   2040
      Left            =   195
      ScaleHeight     =   1980
      ScaleWidth      =   6375
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   6435
      Begin VB.OptionButton rbcChgTo 
         Caption         =   "National"
         Height          =   255
         Index           =   2
         Left            =   4980
         TabIndex        =   8
         Top             =   705
         Width           =   1065
      End
      Begin VB.OptionButton rbcChgTo 
         Caption         =   "Visiting"
         Height          =   255
         Index           =   1
         Left            =   3195
         TabIndex        =   7
         Top             =   705
         Width           =   1680
      End
      Begin VB.OptionButton rbcChgTo 
         Caption         =   "Home"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   705
         Width           =   1485
      End
      Begin VB.CommandButton cmcTime 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4950
         Picture         =   "LLFeed.frx":0FC8
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1125
         Width           =   195
      End
      Begin VB.TextBox edcTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1125
         Width           =   1245
      End
      Begin VB.TextBox edcTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   3
         Top             =   315
         Width           =   1245
      End
      Begin VB.CommandButton cmcTime 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4950
         Picture         =   "LLFeed.frx":10C2
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   195
      End
      Begin VB.Label lacChgTo 
         Caption         =   "Change To Feed"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   735
         Width           =   1500
      End
      Begin VB.Label lacStartTime 
         Appearance      =   0  'Flat
         Caption         =   "Change To Feed Source Start Time"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1125
         Width           =   3015
      End
      Begin VB.Label lacEndTime 
         Appearance      =   0  'Flat
         Caption         =   "Home Feed Source End Time"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   3390
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3120
      Width           =   360
   End
End
Attribute VB_Name = "LLFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of LLFeed.FRM on Fri 3/12/10 @ 11:00 AM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  lmNowDate                                                                             *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LLFeed.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text
'Dim tmRec As LPOPREC
'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim imTimeIndex As Integer

Private Sub cmcCancel_Click()
    igLLFeedReturn = False
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcTme.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                        slNowDate                     slNowTime                 *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim slTime As String
    Dim ilRes As Integer
    slTime = Trim$(edcTime(0).Text)
    If slTime = "" Then
        ilRes = MsgBox("End time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcTime(0).SetFocus
        Exit Sub
    Else
        If Not gValidTime(slTime) Then
            ilRes = MsgBox("End time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcTime(0).SetFocus
            Exit Sub
        End If
    End If
    slTime = Trim$(edcTime(1).Text)
    If slTime = "" Then
        ilRes = MsgBox("Start time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcTime(1).SetFocus
        Exit Sub
    Else
        If Not gValidTime(slTime) Then
            ilRes = MsgBox("Start time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcTime(1).SetFocus
            Exit Sub
        End If
    End If
    If (rbcChgTo(0).Value <> True) And (rbcChgTo(1).Value <> True) And (rbcChgTo(2).Value <> True) Then
        ilRes = MsgBox("Change To Feed must be specified", vbOKOnly + vbExclamation, "Incomplete")
        Exit Sub
    End If
    If (rbcChgTo(0).Value = True) Then
        sgLLCurrentFeed = "H"
    End If
    If (rbcChgTo(1).Value = True) Then
        sgLLCurrentFeed = "V"
    End If
    If (rbcChgTo(2).Value = True) Then
        sgLLCurrentFeed = "N"
    End If
    igLLFeedReturn = True
    sgLLFeedEndTime = edcTime(0).Text
    sgLLFeedStartTime = edcTime(1).Text
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcTme.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    plcTme.Visible = False
End Sub
Private Sub cmcTime_Click(Index As Integer)
    plcTme.Visible = Not plcTme.Visible
    edcTime(Index).SelStart = 0
    edcTime(Index).SelLength = Len(edcTime(Index).Text)
    edcTime(Index).SetFocus
End Sub
Private Sub cmcTime_GotFocus(Index As Integer)
    If imTimeIndex <> Index Then
        plcTme.Visible = False
    End If
    imTimeIndex = Index
    plcTme.Move plcTimes.Left + edcTime(Index).Left, plcTimes.Top + edcTime(Index).Top + edcTime(Index).Height
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcTime_Change(Index As Integer)
    Dim slStartTime As String
    Dim slEndTime As String

    If Index = 0 Then
        slEndTime = Trim$(edcTime(0).Text)
        If slEndTime <> "" Then
            If gValidTime(slEndTime) Then
                slStartTime = Trim$(edcTime(1).Text)
                If slStartTime <> "" Then
                    If gValidTime(slStartTime) Then
                        If gTimeToLong(slEndTime, False) > gTimeToLong(slStartTime, False) Then
                            edcTime(1).Text = slEndTime
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub edcTime_GotFocus(Index As Integer)
    If Not imBypassFocus Then
        If imTimeIndex <> Index Then
            plcTme.Visible = False
        End If
        imTimeIndex = Index
        plcTme.Move plcTimes.Left + edcTime(Index).Left, plcTimes.Top + edcTime(Index).Top + edcTime(Index).Height
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcTime_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTime_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTime(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        ilFound = False
        For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
            If KeyAscii = igLegalTime(ilLoop) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcTime_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcTme.Visible = Not plcTme.Visible
        End If
        edcTime(Index).SelStart = 0
        edcTime(Index).SelLength = Len(edcTime(Index).Text)
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

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LLFeed = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'*                                                                                        *
'* Local Labels (Removed)                                                                 *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slEventTitle1 As String
    Dim slEventTitle2 As String

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    imBypassFocus = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    mInitBox
    LLFeed.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone LLFeed
    slNowDate = Format$(gNow(), "m/d/yy")
    slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
    edcTime(0).Text = slNowTime
    edcTime(1).Text = slNowTime
    gGetEventTitles igLLCurrentVefCode, slEventTitle1, slEventTitle2
    rbcChgTo(0).Caption = slEventTitle2
    rbcChgTo(1).Caption = slEventTitle1
    If sgLLCurrentFeed = "H" Then
        lacEndTime.Caption = slEventTitle2 & " Feed Source End Time"
        'lacEndTime.Caption = "Home Feed Source End Time"
        ''lacStartTime.Caption = "Change To Feed Source Start Time"
        rbcChgTo(0).Enabled = False
    ElseIf sgLLCurrentFeed = "N" Then
        lacEndTime.Caption = "National Feed Source End Time"
        'lacStartTime.Caption = "Visiting Feed Source Start Time"
        rbcChgTo(2).Enabled = False
    Else
        lacEndTime.Caption = slEventTitle1 & " Feed Source End Time"
        'lacEndTime.Caption = "Visiting Feed Source End Time"
        ''lacStartTime.Caption = "Home Feed Source Start Time"
        rbcChgTo(1).Enabled = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
         
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mTerminate
'   Where:
'
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload LLFeed
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_GotFocus()
    plcTme.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcTime(imTimeIndex).SetFocus
                    'SendKeys slKey
                    gSendKeys edcTime(imTimeIndex), slKey
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub plcTimes_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Feed Source Change"
End Sub
