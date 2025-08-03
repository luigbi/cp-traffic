VERSION 5.00
Begin VB.Form PreFeed 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4770
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
   ScaleHeight     =   4770
   ScaleWidth      =   6885
   Begin VB.PictureBox pbcBusTab 
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
      Left            =   6765
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   21
      Top             =   3600
      Width           =   60
   End
   Begin VB.PictureBox pbcBusSTab 
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
      Left            =   6705
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   16
      Top             =   750
      Width           =   105
   End
   Begin VB.TextBox edcBus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   5010
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcBusArrow 
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
      Left            =   4605
      Picture         =   "PreFeed.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   270
      Top             =   4095
   End
   Begin VB.ListBox lbcZone 
      Appearance      =   0  'Flat
      Height          =   450
      ItemData        =   "PreFeed.frx":030A
      Left            =   2295
      List            =   "PreFeed.frx":031D
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.ListBox lbcDay 
      Appearance      =   0  'Flat
      Height          =   660
      ItemData        =   "PreFeed.frx":0336
      Left            =   1185
      List            =   "PreFeed.frx":034F
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2850
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2085
      Picture         =   "PreFeed.frx":036F
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   2730
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "PreFeed.frx":0469
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   24
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
            Picture         =   "PreFeed.frx":1127
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.VScrollBar vbcBus 
      Height          =   2535
      LargeChange     =   12
      Left            =   6345
      TabIndex        =   23
      Top             =   1005
      Width           =   240
   End
   Begin VB.PictureBox pbcBus 
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
      Height          =   2580
      Left            =   4905
      Picture         =   "PreFeed.frx":1431
      ScaleHeight     =   2580
      ScaleWidth      =   1455
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   975
      Width           =   1455
      Begin VB.Label lacBusFrame 
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
         Left            =   30
         TabIndex        =   22
         Top             =   525
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin VB.PictureBox plcBus 
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   4860
      ScaleHeight     =   2595
      ScaleWidth      =   1710
      TabIndex        =   17
      Top             =   930
      Width           =   1770
   End
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
      Picture         =   "PreFeed.frx":D77F
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1275
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5280
      Width           =   90
   End
   Begin VB.PictureBox pbcPFSTab 
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
      Left            =   60
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   465
      Width           =   105
   End
   Begin VB.PictureBox pbcPFTab 
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
      TabIndex        =   9
      Top             =   3810
      Width           =   60
   End
   Begin VB.VScrollBar vbcPreFeed 
      Height          =   3060
      LargeChange     =   14
      Left            =   4140
      TabIndex        =   10
      Top             =   720
      Width           =   240
   End
   Begin VB.PictureBox pbcPreFeed 
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
      Height          =   3120
      Left            =   345
      Picture         =   "PreFeed.frx":DA89
      ScaleHeight     =   3120
      ScaleWidth      =   3780
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   645
      Width           =   3780
      Begin VB.Label lacPreFeedFrame 
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
         Width           =   3780
      End
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1485
      TabIndex        =   11
      Top             =   4320
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2865
      TabIndex        =   12
      Top             =   4320
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   13
      Top             =   4320
      Width           =   1050
   End
   Begin VB.PictureBox plcPreFeed 
      ForeColor       =   &H00000000&
      Height          =   3195
      Left            =   300
      ScaleHeight     =   3135
      ScaleWidth      =   4065
      TabIndex        =   1
      Top             =   630
      Width           =   4125
   End
   Begin VB.Label lacTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   330
      TabIndex        =   26
      Top             =   3870
      Width           =   4140
   End
   Begin VB.Label lacScreen 
      Height          =   525
      Left            =   120
      TabIndex        =   25
      Top             =   45
      Width           =   4590
      WordWrap        =   -1  'True
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6195
      Picture         =   "PreFeed.frx":3410B
      Top             =   4110
      Width           =   480
   End
End
Attribute VB_Name = "PreFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PreFeed.frm on Fri 3/12/10 @ 11:00 A
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PreFeed.Frm
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
Dim hmPff As Integer
Dim tmPFF As PFF        'GSF record image
Dim imPffRecLen As Integer        'GSF record length
Dim tmPffSrchKey0 As LONGKEY0
Dim tmPffSrchKey1 As PFFKEY1
Dim tmPFFInfo() As PFFINFO
Dim hmPbf As Integer
Dim tmPbf As PBF        'GSF record image
Dim imPbfRecLen As Integer        'GSF record length
Dim tmPbfSrchKey0 As LONGKEY0
Dim tmPbfSrchKey1 As LONGKEY0
Dim tmPBFInfo() As PBFINFO
Dim tmPFCtrls(0 To 6)  As FIELDAREA     'Index zero ignoed
Dim imLBPFCtrls As Integer
Dim imPFBoxNo As Integer   'Current event name Box
Dim imPFRowNo As Integer  'Current event row
Dim tmBusCtrls(0 To 2)  As FIELDAREA    'Index zero ignored
Dim imLBBusCtrls As Integer
Dim imBusBoxNo As Integer   'Current event name Box
Dim imBusRowNo As Integer  'Current event row
Dim smPFSave() As String
Dim smPFShow() As String
Dim lmPFSave() As Long
Dim smAllBusSave() As String    'Index: 1=From Bus; 2=To Bus
Dim lmAllBusSave() As Long  'Index: 1=PbfCode; 2=Current Tie ID
Dim smBusSave() As String   'Index: 1=From Bus; 2=To Bus
Dim smBusShow() As String
Dim lmBusSave() As Long     'Index: 1=PbfCode; 2=Current Tie ID
Dim lmDelPFFCode() As Long
Dim lmDelPBFCode() As Long
Dim imSettingValue As Integer
Dim imPffChg As Integer
Dim imPbfChg As Integer
Dim imBypassSetting As Integer
Dim imChgMode As Integer
Dim imLbcArrowSetting As Integer
Dim imComboBoxIndex As Integer
Dim imBypassFocus As Integer
Dim imBSMode As Integer
Dim imTabDirection As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragShift As Integer
Dim smFocusOnPForBus As String * 1  'P=PreFeed; B=Bus
Dim lmTiePffToPbfID As Long
Dim lmCurTieID As Long

Const LBONE = 1

Const FROMSTARTTIMEINDEX = 1
Const FROMENDTIMEINDEX = 2
Const FROMDAYINDEX = 3
Const FROMZONEINDEX = 4
Const TOSTARTTIMEINDEX = 5
Const TODAYINDEX = 6

Const FROMBUSINDEX = 1
Const TOBUSINDEX = 2






Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    mClearBus
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = -1
    imPFRowNo = -1
End Sub

Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imPFBoxNo > 0 Then
            mPFEnableBox imPFBoxNo
        End If
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    mClearBus
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = -1
    imPFRowNo = -1
End Sub

Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDropDown_Click()
    Select Case imPFBoxNo
        Case FROMSTARTTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case FROMENDTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case FROMDAYINDEX
            lbcDay.Visible = Not lbcDay.Visible
        Case FROMZONEINDEX
            lbcZone.Visible = Not lbcZone.Visible
        Case TOSTARTTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case TODAYINDEX
            lbcDay.Visible = Not lbcDay.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mPFEnableBox imPFBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    mClearCtrlFields
    mPreFeedPop
    mBusPop
    mMoveRecToCtrl
    imPffChg = False
    imPbfChg = False
    imPFBoxNo = -1
    imPFRowNo = -1
    mSetCommands
    pbcPFSTab.SetFocus
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmcUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcBus_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imPFBoxNo
        Case FROMSTARTTIMEINDEX
        Case FROMENDTIMEINDEX
        Case FROMDAYINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcDay, imBSMode, imComboBoxIndex
        Case FROMZONEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcZone, imBSMode, imComboBoxIndex
        Case TOSTARTTIMEINDEX
        Case TODAYINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcDay, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKeyAscii As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    
    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imPFBoxNo
        Case FROMSTARTTIMEINDEX
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
        Case FROMENDTIMEINDEX
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
        Case FROMDAYINDEX
        Case FROMZONEINDEX
        Case TOSTARTTIMEINDEX
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
        Case TODAYINDEX
    End Select

End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imPFBoxNo
            Case FROMSTARTTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case FROMENDTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case FROMDAYINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcDay, imLbcArrowSetting
                End If
            Case FROMZONEINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcZone, imLbcArrowSetting
                End If
            Case TOSTARTTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case TODAYINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    gProcessArrowKey Shift, KeyCode, lbcDay, imLbcArrowSetting
                End If
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imPFBoxNo
            Case FROMSTARTTIMEINDEX
            Case FROMENDTIMEINDEX
            Case FROMDAYINDEX
            Case FROMZONEINDEX
            Case TOSTARTTIMEINDEX
            Case TODAYINDEX
        End Select
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
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
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
        'If (cbcSelect.Enabled) And (imPFBoxNo > 0) Then
        '    cbcSelect.Enabled = False
        '    ilReSet = True
        'Else
        '    ilReSet = False
        'End If
        gFunctionKeyBranch KeyCode
        'If imCEBoxNo > 0 Then
        '    mCEEnableBox imCEBoxNo
        'ElseIf imDMBoxNo > 0 Then
        '    mDmPFEnableBox imDMBoxNo
        'Else
        '    mPFEnableBox imPFBoxNo
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
'*******************************************************
'*                                                     *
'*      Procedure Name:mPreFeedPop                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
'*******************************************************
Private Sub mPreFeedPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilUpper As Integer
    Dim llTime As Long
    Dim slStr As String
    
    ReDim tmPFFInfo(0 To 0) As PFFINFO   'VB list box clear (list box used to retain code number so record can be found)
    ilUpper = 0
    If igPreFeedType = 0 Then
        tmPffSrchKey1.sType = "D"
    Else
        tmPffSrchKey1.sType = "E"
    End If
    tmPffSrchKey1.iVefCode = igPreFeedVefCode
    tmPffSrchKey1.sAirDay = sgPreFeedDay
    gPackDate sgPreFeedDate, tmPffSrchKey1.iStartDate(0), tmPffSrchKey1.iStartDate(1)
    ilRet = btrGetEqual(hmPff, tmPFF, imPffRecLen, tmPffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If tmPFF.sType <> tmPffSrchKey1.sType Then
            Exit Do
        End If
        If tmPFF.iVefCode <> igPreFeedVefCode Then
            Exit Do
        End If
        If tmPFF.sAirDay <> sgPreFeedDay Then
            Exit Do
        End If
        gUnpackDateLong tmPFF.iStartDate(0), tmPFF.iStartDate(1), llDate
        If llDate <> gDateValue(sgPreFeedDate) Then
            Exit Do
        End If
        gUnpackTimeLong tmPFF.iFromStartTime(0), tmPFF.iFromStartTime(1), False, llTime
        slStr = Trim$(str$(llTime))
        Do While Len(slStr) < 7
            slStr = "0" & slStr
        Loop
        tmPFFInfo(ilUpper).sKey = slStr
        tmPFFInfo(ilUpper).tPff = tmPFF
        ReDim Preserve tmPFFInfo(0 To ilUpper + 1) As PFFINFO
        ilUpper = ilUpper + 1
        ilRet = btrGetNext(hmPff, tmPFF, imPffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If UBound(tmPFFInfo) > 1 Then
        ArraySortTyp fnAV(tmPFFInfo(), 0), UBound(tmPFFInfo), 0, LenB(tmPFFInfo(0)), 0, LenB(tmPFFInfo(0).sKey), 0
    End If
    Exit Sub
mPreFeedPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBusPop                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
'*******************************************************
Private Sub mBusPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slCode As String
    
    
    ReDim tmPBFInfo(0 To 0) As PBFINFO   'VB list box clear (list box used to retain code number so record can be found)
    ilUpper = 0
    If igPreFeedType = 0 Then
        Exit Sub
    End If
    For ilLoop = LBound(tmPFFInfo) To UBound(tmPFFInfo) - 1 Step 1
        tmPbfSrchKey1.lCode = tmPFFInfo(ilLoop).tPff.lCode
        ilRet = btrGetEqual(hmPbf, tmPbf, imPbfRecLen, tmPbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            If tmPbf.lPffCode <> tmPFFInfo(ilLoop).tPff.lCode Then
                Exit Do
            End If
            slCode = Trim$(str$(tmPbf.lPffCode))
            Do While Len(slCode) < 10
                slCode = "0" & slCode
            Loop
            tmPBFInfo(ilUpper).sKey = slCode & tmPbf.sFromBus
            tmPBFInfo(ilUpper).tPbf = tmPbf
            ReDim Preserve tmPBFInfo(0 To ilUpper + 1) As PBFINFO
            ilUpper = ilUpper + 1
            ilRet = btrGetNext(hmPbf, tmPbf, imPbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
    Next ilLoop
    Exit Sub
mBusPopErr:
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
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    imLBPFCtrls = 1
    imLBBusCtrls = 1
    lacScreen.Caption = sgPreFeedScreenCaption
    imFirstActivate = True
    imSettingValue = False
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassFocus = False
    imBypassSetting = False
    imTabDirection = 0  'Left to right movement
    imPFRowNo = -1
    imPFBoxNo = -1
    imBusRowNo = -1
    imBusBoxNo = -1
    lmCurTieID = -1
    smFocusOnPForBus = ""
    ReDim smPFShow(0 To 6, 0 To 1) As String 'Values shown in program area
    ReDim smPFSave(0 To 6, 0 To 1) As String 'Values saved (program name) in program area
    ReDim lmPFSave(0 To 2, 0 To 1) As Long
    ReDim smBusShow(0 To 2, 0 To 1) As String 'Values shown in program area
    ReDim smBusSave(0 To 2, 0 To 1) As String 'Values saved (program name) in program area
    ReDim lmBusSave(0 To 2, 0 To 1) As Long
    ReDim smAllBusSave(0 To 2, 0 To 1) As String
    ReDim lmAllBusSave(0 To 2, 0 To 1) As Long
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcBusArrow.Picture = IconTraf!imcArrow.Picture
    pbcBusArrow.Width = 90
    pbcBusArrow.Height = 165
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    If igPreFeedType = 0 Then
        lacTime.Caption = "All times are Affiliate local times"
    Else
        lacTime.Caption = "All times are Feed times"
    End If
    hmPff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPff, "", sgDBPath & "PFF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PFF.Btr)", PreFeed
    On Error GoTo 0
    imPffRecLen = Len(tmPFF)
    
    hmPbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPbf, "", sgDBPath & "PBF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: PBF.Btr)", PreFeed
    On Error GoTo 0
    imPbfRecLen = Len(tmPbf)
    
    mInitBox
    
    mClearCtrlFields

    mPreFeedPop
    If imTerminate Then
        Exit Sub
    End If
    mBusPop
    mMoveRecToCtrl
    PreFeed.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterModalForm PreFeed
    gCenterStdAlone PreFeed
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub
mInitErr:
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
    Dim ilRet
    Screen.MousePointer = vbDefault
    
    igManUnload = YES
    Unload PreFeed
    igManUnload = NO
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmPFFInfo
    Erase tmPBFInfo
    Erase smPFSave
    Erase smPFShow
    Erase lmPFSave
    Erase smBusSave
    Erase smBusShow
    Erase lmBusSave
    Erase smAllBusSave
    Erase lmAllBusSave
    Erase lmDelPFFCode
    Erase lmDelPBFCode
    
    ilRet = btrClose(hmPff)
    btrDestroy hmPff
    ilRet = btrClose(hmPbf)
    btrDestroy hmPbf
    
    Set PreFeed = Nothing

End Sub

Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim ilRowNo As Integer
    Dim llPffCode As Long
    Dim llPbfCode As Long
    
    If smFocusOnPForBus <> "B" Then
        If (imPFRowNo < vbcPreFeed.Value) Or (imPFRowNo > vbcPreFeed.Value + vbcPreFeed.LargeChange) Then
            Exit Sub
        End If
        ilRowNo = imPFRowNo
        mBusSetShow imBusBoxNo
        imBusBoxNo = -1
        imBusRowNo = -1
        mClearBus
        mPFSetShow imPFBoxNo, False
        imPFBoxNo = -1
        imPFRowNo = -1
        pbcArrow.Visible = False
        lacPreFeedFrame.Visible = False
        pbcBusArrow.Visible = False
        lacBusFrame.Visible = False
        gCtrlGotFocus ActiveControl
        ilUpper = UBound(smPFSave, 2)
        llPffCode = lmPFSave(1, ilRowNo)
        If ilRowNo = ilUpper Then
            mPFInitNew ilRowNo
        Else
            If llPffCode > 0 Then
                lmDelPFFCode(UBound(lmDelPFFCode)) = llPffCode
                ReDim Preserve lmDelPFFCode(0 To UBound(lmDelPFFCode) + 1) As Long
            End If
            For ilLoop = ilRowNo To ilUpper - 1 Step 1
                For ilIndex = 1 To UBound(smPFSave, 1) Step 1
                    smPFSave(ilIndex, ilLoop) = smPFSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(smPFShow, 1) Step 1
                    smPFShow(ilIndex, ilLoop) = smPFShow(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(lmPFSave, 1) Step 1
                    lmPFSave(ilIndex, ilLoop) = lmPFSave(ilIndex, ilLoop + 1)
                Next ilIndex
            Next ilLoop
            ilUpper = UBound(smPFSave, 2)
            ReDim Preserve smPFShow(0 To 6, 0 To ilUpper - 1) As String 'Values shown in program area
            ReDim Preserve smPFSave(0 To 6, 0 To ilUpper - 1) As String    'Values saved (program name) in program area
            ReDim Preserve lmPFSave(0 To 2, 0 To ilUpper - 1) As Long    'Values saved (program name) in program area
            imPffChg = True
        End If
        mSetCommands
        lacPreFeedFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        imSettingValue = True
        vbcPreFeed.Min = LBONE  'LBound(smPFShow, 2)
        imSettingValue = True
        If UBound(smPFShow, 2) - 1 <= vbcPreFeed.LargeChange + 1 Then ' + 1 Then
            vbcPreFeed.Max = LBONE  'LBound(smPFShow, 2)
        Else
            vbcPreFeed.Max = UBound(smPFShow, 2) - vbcPreFeed.LargeChange
        End If
        imSettingValue = True
        vbcPreFeed.Value = vbcPreFeed.Min
        imSettingValue = True
        pbcPreFeed.Cls
        pbcPreFeed_Paint
    Else
        If (imBusRowNo < vbcBus.Value) Or (imBusRowNo > vbcBus.Value + vbcBus.LargeChange) Then
            Exit Sub
        End If
        ilRowNo = imBusRowNo
        mBusSetShow imBusBoxNo
        imBusBoxNo = -1
        imBusRowNo = -1
        pbcArrow.Visible = False
        lacBusFrame.Visible = False
        gCtrlGotFocus ActiveControl
        ilUpper = UBound(smBusSave, 2)
        llPbfCode = lmBusSave(1, ilRowNo)
        If ilRowNo = ilUpper Then
            mBusInitNew ilRowNo
        Else
            If llPbfCode > 0 Then
                lmDelPBFCode(UBound(lmDelPBFCode)) = llPbfCode
                ReDim Preserve lmDelPBFCode(0 To UBound(lmDelPBFCode) + 1) As Long
            End If
            For ilLoop = ilRowNo To ilUpper - 1 Step 1
                For ilIndex = 1 To UBound(smBusSave, 1) Step 1
                    smBusSave(ilIndex, ilLoop) = smBusSave(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(smBusShow, 1) Step 1
                    smBusShow(ilIndex, ilLoop) = smBusShow(ilIndex, ilLoop + 1)
                Next ilIndex
                For ilIndex = 1 To UBound(lmBusSave, 1) Step 1
                    lmBusSave(ilIndex, ilLoop) = lmBusSave(ilIndex, ilLoop + 1)
                Next ilIndex
            Next ilLoop
            ilUpper = UBound(smBusSave, 2)
            ReDim Preserve smBusShow(0 To 2, 0 To ilUpper - 1) As String 'Values shown in program area
            ReDim Preserve smBusSave(0 To 2, 0 To ilUpper - 1) As String    'Values saved (program name) in program area
            ReDim Preserve lmBusSave(0 To 2, 0 To ilUpper - 1) As Long    'Values saved (program name) in program area
            imPbfChg = True
        End If
        mSetCommands
        lacBusFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        imSettingValue = True
        vbcBus.Min = LBONE  'LBound(smBusShow, 2)
        imSettingValue = True
        If UBound(smBusShow, 2) - 1 <= vbcBus.LargeChange + 1 Then ' + 1 Then
            vbcBus.Max = LBONE  'LBound(smBusShow, 2)
        Else
            vbcBus.Max = UBound(smBusShow, 2) - vbcBus.LargeChange
        End If
        imSettingValue = True
        vbcBus.Value = vbcBus.Min
        imSettingValue = True
        pbcBus.Cls
        pbcBus_Paint
    End If
End Sub

Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash_Click
End Sub

Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If smFocusOnPForBus <> "B" Then
        If State = vbEnter Then    'Enter drag over
            lacPreFeedFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
            imcTrash.Picture = IconTraf!imcTrashOpened.Picture
        ElseIf State = vbLeave Then
            lacPreFeedFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
            imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        End If
    Else
        If State = vbEnter Then    'Enter drag over
            lacBusFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
            imcTrash.Picture = IconTraf!imcTrashOpened.Picture
        ElseIf State = vbLeave Then
            lacBusFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
            imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        End If
    End If
End Sub

Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lacScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lbcDay_Click()
    gProcessLbcClick lbcDay, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcZone_Click()
    gProcessLbcClick lbcZone, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub pbcBus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (smFocusOnPForBus <> "P") And (smFocusOnPForBus <> "B") Then
        Exit Sub
    End If
    smFocusOnPForBus = "B"
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub

Private Sub pbcBus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilBusRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    ilBusRow = vbcBus.LargeChange + 1
    If UBound(smBusSave, 2) > ilBusRow Then
        ilMaxRow = ilBusRow
    Else
        ilMaxRow = UBound(smBusSave, 2) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBBusCtrls To UBound(tmBusCtrls) Step 1
            If (X >= tmBusCtrls(ilBox).fBoxX) And (X <= (tmBusCtrls(ilBox).fBoxX + tmBusCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmBusCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmBusCtrls(ilBox).fBoxY + tmBusCtrls(ilBox).fBoxH)) Then
                    mPFSetShow imPFBoxNo, True
                    ilRowNo = ilRow + vbcBus.Value - 1
                    If ilRowNo > UBound(smBusSave, 2) Then
                        Beep
                        mBusSetFocus imBusBoxNo
                        Exit Sub
                    End If
                    If (ilBox > FROMBUSINDEX) And (Trim$(smBusSave(1, ilRowNo)) = "") Then
                        Beep
                        mBusSetFocus imBusBoxNo
                        Exit Sub
                    End If
                    mBusSetShow imBusBoxNo
                    imBusRowNo = ilRow + vbcBus.Value - 1
                    If (imBusRowNo = UBound(smBusSave, 2)) And (Trim$(smBusSave(1, imBusRowNo)) = "") Then
                        mBusInitNew imBusRowNo
                    End If
                    imBusBoxNo = ilBox
                    mBusEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mBusSetFocus imBusBoxNo
End Sub



Private Sub pbcBus_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long

    ilStartRow = vbcBus.Value '+ 1  'Top location
    ilEndRow = vbcBus.Value + vbcBus.LargeChange ' + 1
    If ilEndRow > UBound(smBusSave, 2) Then
        If Trim$(smBusShow(1, UBound(smBusShow, 2))) <> "" Then
            ilEndRow = UBound(smBusSave, 2) 'include blank row as it might have data
        Else
            ilEndRow = UBound(smBusSave, 2) - 1
        End If
    End If
    llColor = pbcBus.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smBusSave, 2) Then
            pbcBus.ForeColor = DARKPURPLE
        Else
            pbcBus.ForeColor = llColor
        End If
        For ilBox = imLBBusCtrls To UBound(tmBusCtrls) Step 1
            pbcBus.CurrentX = tmBusCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcBus.CurrentY = tmBusCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 15 '+ fgBoxInsetY
            slStr = Trim$(smBusShow(ilBox, ilRow))
            pbcBus.Print slStr
        Next ilBox
    Next ilRow
    pbcBus.ForeColor = llColor

End Sub

Private Sub pbcBusSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcBusSTab.hwnd Then
        Exit Sub
    End If
    Select Case imBusBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smBusSave, 2) = 1) Then
                imBusRowNo = 1
                mBusInitNew imBusRowNo
            Else
                If UBound(smBusSave, 2) <= vbcBus.LargeChange Then 'was <=
                    vbcBus.Max = LBONE  'LBound(smBusSave, 2)
                Else
                    vbcBus.Max = UBound(smBusSave, 2) - vbcBus.LargeChange '- 1
                End If
                imBusRowNo = 1
                If imBusRowNo >= UBound(smBusSave, 2) Then
                    mBusInitNew imBusRowNo
                End If
                imSettingValue = True
                vbcBus.Value = vbcBus.Min
                imSettingValue = False
            End If
            ilBox = FROMBUSINDEX
            imBusBoxNo = ilBox
            mBusEnableBox ilBox
            Exit Sub
        Case FROMBUSINDEX, 0
            mBusSetShow imBusBoxNo
            If (imBusBoxNo < 1) And (imBusRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = TOBUSINDEX
            If imBusRowNo <= 1 Then
                imBusBoxNo = -1
                imBusRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imBusRowNo = imBusRowNo - 1
            If imBusRowNo < vbcBus.Value Then
                imSettingValue = True
                vbcBus.Value = vbcBus.Value - 1
                imSettingValue = False
            End If
            imBusBoxNo = ilBox
            mBusEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imBusBoxNo - 1
    End Select
    mBusSetShow imBusBoxNo
    imBusBoxNo = ilBox
    mBusEnableBox ilBox
End Sub

Private Sub pbcBusTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If GetFocus() <> pbcBusTab.hwnd Then
        Exit Sub
    End If
    Select Case imBusBoxNo
        Case -1 'Tab from control prior to form area
            imBusRowNo = UBound(smBusSave, 2) - 1
            imSettingValue = True
            If imBusRowNo <= vbcBus.LargeChange + 1 Then
                vbcBus.Value = 1
            Else
                vbcBus.Value = imBusRowNo - vbcBus.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = TOBUSINDEX
        Case TOBUSINDEX
            mBusSetShow imBusBoxNo
            If mTestBusSaveFields(imBusRowNo) = NO Then
                mBusEnableBox imBusBoxNo
                Exit Sub
            End If
            'If imBusRowNo >= UBound(smBusSave, 2) Then
            '    imPbfChg = True
            '    ReDim Preserve smBusShow(1 To 2, 1 To imBusRowNo + 1) As String 'Values shown in program area
            '    ReDim Preserve smBusSave(1 To 2, 1 To imBusRowNo + 1) As String 'Values saved (program name) in program area
            '    ReDim Preserve lmBusSave(1 To 2, 1 To imBusRowNo + 1) As Long 'Values saved (program name) in program area
            '    For ilLoop = LBound(smBusShow, 1) To UBound(smBusShow, 1) Step 1
            '        smBusShow(ilLoop, imBusRowNo + 1) = ""
            '    Next ilLoop
            '    For ilLoop = LBound(smBusSave, 1) To UBound(smBusSave, 1) Step 1
            '        smBusSave(ilLoop, imBusRowNo + 1) = ""
            '    Next ilLoop
            '    For ilLoop = LBound(lmBusSave, 1) To UBound(lmBusSave, 1) Step 1
            '        lmBusSave(ilLoop, imBusRowNo + 1) = 0
            '    Next ilLoop
            'End If
            mAddBusRow
            If imBusRowNo >= UBound(smBusSave, 2) - 1 Then
                imBusRowNo = imBusRowNo + 1
                'mBusInitNew imBusRowNo
                'If UBound(smBusSave, 2) <= vbcBus.LargeChange Then 'was <=
                '    vbcBus.Max = LBound(smBusSave, 2) '- 1
                'Else
                '    vbcBus.Max = UBound(smBusSave, 2) - vbcBus.LargeChange '- 1
                'End If
            Else
                imBusRowNo = imBusRowNo + 1
            End If
            If imBusRowNo > vbcBus.Value + vbcBus.LargeChange Then
                imSettingValue = True
                vbcBus.Value = vbcBus.Value + 1
                imSettingValue = False
            End If
            If imBusRowNo >= UBound(smBusSave, 2) Then
                imBusBoxNo = 0
                mSetCommands
                'lacBusFrame.Move 0, tmBusCtrls(PROPNOINDEX).fBoxY + (imBusRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacBusFrame.Visible = True
                pbcBusArrow.Move pbcBusArrow.Left, plcBus.Top + tmBusCtrls(FROMBUSINDEX).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15) + 45
                pbcBusArrow.Visible = True
                pbcBusArrow.SetFocus
                Exit Sub
            Else
                ilBox = FROMBUSINDEX
            End If
            imBusBoxNo = ilBox
            mBusEnableBox ilBox
            Exit Sub
        Case 0
            ilBox = FROMBUSINDEX
        Case Else
            ilBox = imBusBoxNo + 1
    End Select
    mBusSetShow imBusBoxNo
    imBusBoxNo = ilBox
    mBusEnableBox ilBox
End Sub

Private Sub pbcClickFocus_GotFocus()
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    mClearBus
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = -1
    imPFRowNo = -1
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcPFSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcPFSTab.hwnd Then
        Exit Sub
    End If
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    imTabDirection = -1 'Set- Right to left
    Select Case imPFBoxNo
        Case -1 'Tab from control prior to form area
            If (UBound(smPFSave, 2) = 1) Then
                imTabDirection = 0  'Set-Left to right
                imPFRowNo = 1
                mPFInitNew imPFRowNo
            Else
                If UBound(smPFSave, 2) <= vbcPreFeed.LargeChange Then 'was <=
                    vbcPreFeed.Max = LBONE  'LBound(smPFSave, 2)
                Else
                    vbcPreFeed.Max = UBound(smPFSave, 2) - vbcPreFeed.LargeChange '- 1
                End If
                imPFRowNo = 1
                If imPFRowNo >= UBound(smPFSave, 2) Then
                    mPFInitNew imPFRowNo
                End If
                imSettingValue = True
                vbcPreFeed.Value = vbcPreFeed.Min
                imSettingValue = False
            End If
            ilBox = FROMSTARTTIMEINDEX
            imPFBoxNo = ilBox
            mPFEnableBox ilBox
            Exit Sub
        Case FROMSTARTTIMEINDEX, 0
            mPFSetShow imPFBoxNo, False
            If (imPFBoxNo < 1) And (imPFRowNo < 1) Then 'Modelled from Proposal
                Exit Sub
            End If
            ilBox = TODAYINDEX
            If imPFRowNo <= 1 Then
                mClearBus
                imPFBoxNo = -1
                imPFRowNo = -1
                cmcDone.SetFocus
                Exit Sub
            End If
            imPFRowNo = imPFRowNo - 1
            If imPFRowNo < vbcPreFeed.Value Then
                imSettingValue = True
                vbcPreFeed.Value = vbcPreFeed.Value - 1
                imSettingValue = False
            End If
            imPFBoxNo = ilBox
            mPFEnableBox ilBox
            Exit Sub
        Case FROMSTARTTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imPFBoxNo - 1
        Case FROMENDTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imPFBoxNo - 1
        Case TOSTARTTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imPFBoxNo - 1
        Case Else
            ilBox = imPFBoxNo - 1
    End Select
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = ilBox
    mPFEnableBox ilBox
End Sub

Private Sub pbcPFSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcPFTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If GetFocus() <> pbcPFTab.hwnd Then
        Exit Sub
    End If
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    imTabDirection = 0 'Set- Left to right
    Select Case imPFBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imPFRowNo = UBound(smPFSave, 2) - 1
            imSettingValue = True
            If imPFRowNo <= vbcPreFeed.LargeChange + 1 Then
                vbcPreFeed.Value = 1
            Else
                vbcPreFeed.Value = imPFRowNo - vbcPreFeed.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = TODAYINDEX
        Case TODAYINDEX
            mPFSetShow imPFBoxNo, False
            If mTestSaveFields(imPFRowNo) = NO Then
                mPFEnableBox imPFBoxNo
                Exit Sub
            End If
            mClearBus
            'If imPFRowNo >= UBound(smPFSave, 2) Then
            '    imPffChg = True
            '    ReDim Preserve smPFShow(1 To 6, 1 To imPFRowNo + 1) As String 'Values shown in program area
            '    ReDim Preserve smPFSave(1 To 6, 1 To imPFRowNo + 1) As String 'Values saved (program name) in program area
            '    ReDim Preserve lmPFSave(1 To 2, 1 To imPFRowNo + 1) As Long 'Values saved (program name) in program area
            '    For ilLoop = LBound(smPFShow, 1) To UBound(smPFShow, 1) Step 1
            '        smPFShow(ilLoop, imPFRowNo + 1) = ""
            '    Next ilLoop
            '    For ilLoop = LBound(smPFSave, 1) To UBound(smPFSave, 1) Step 1
            '        smPFSave(ilLoop, imPFRowNo + 1) = ""
            '    Next ilLoop
            '    For ilLoop = LBound(lmPFSave, 1) To UBound(lmPFSave, 1) Step 1
            '        lmPFSave(ilLoop, imPFRowNo + 1) = 0
            '    Next ilLoop
            'End If
            mAddPFRow
            If imPFRowNo >= UBound(smPFSave, 2) - 1 Then
                imPFRowNo = imPFRowNo + 1
                'mPFInitNew imPFRowNo
                'If UBound(smPFSave, 2) <= vbcPreFeed.LargeChange Then 'was <=
                '    vbcPreFeed.Max = LBound(smPFSave, 2) '- 1
                'Else
                '    vbcPreFeed.Max = UBound(smPFSave, 2) - vbcPreFeed.LargeChange '- 1
                'End If
            Else
                imPFRowNo = imPFRowNo + 1
            End If
            If imPFRowNo > vbcPreFeed.Value + vbcPreFeed.LargeChange Then
                imSettingValue = True
                vbcPreFeed.Value = vbcPreFeed.Value + 1
                imSettingValue = False
            End If
            If imPFRowNo >= UBound(smPFSave, 2) Then
                imPFBoxNo = 0
                mSetCommands
                'lacPreFeedFrame.Move 0, tmPFCtrls(PROPNOINDEX).fBoxY + (imPFRowNo - vbcProj.Value) * (fgBoxGridH + 15) - 30
                'lacPreFeedFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcPreFeed.Top + tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = FROMSTARTTIMEINDEX
            End If
            imPFBoxNo = ilBox
            mPFEnableBox ilBox
            Exit Sub
        Case 0
            ilBox = FROMSTARTTIMEINDEX
        Case FROMSTARTTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imPFBoxNo + 1
        Case FROMENDTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imPFBoxNo + 1
        Case TOSTARTTIMEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidTime(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imPFBoxNo + 1
        Case Else
            ilBox = imPFBoxNo + 1
    End Select
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = ilBox
    mPFEnableBox ilBox
End Sub

Private Sub pbcPFTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcPreFeed_GotFocus()
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
End Sub

Private Sub pbcPreFeed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    smFocusOnPForBus = "P"
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragShift = Shift
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub

Private Sub pbcPreFeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilPreFeedRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    ilPreFeedRow = vbcPreFeed.LargeChange + 1
    If UBound(smPFSave, 2) > ilPreFeedRow Then
        ilMaxRow = ilPreFeedRow
    Else
        ilMaxRow = UBound(smPFSave, 2) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBPFCtrls To UBound(tmPFCtrls) Step 1
            If (X >= tmPFCtrls(ilBox).fBoxX) And (X <= (tmPFCtrls(ilBox).fBoxX + tmPFCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmPFCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmPFCtrls(ilBox).fBoxY + tmPFCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPreFeed.Value - 1
                    If ilRowNo > UBound(smPFSave, 2) Then
                        Beep
                        mPFSetFocus imPFBoxNo
                        Exit Sub
                    End If
                    If (ilBox > FROMSTARTTIMEINDEX) And (Trim$(smPFSave(1, ilRowNo)) = "") Then
                        Beep
                        mPFSetFocus imPFBoxNo
                        Exit Sub
                    End If
                    mPFSetShow imPFBoxNo, False
                    mClearBus
                    imPFRowNo = ilRow + vbcPreFeed.Value - 1
                    If (imPFRowNo = UBound(smPFSave, 2)) And (Trim$(smPFSave(1, imPFRowNo)) = "") Then
                        mPFInitNew imPFRowNo
                    End If
                    imPFBoxNo = ilBox
                    mPFEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mPFSetFocus imPFBoxNo
End Sub

Private Sub pbcPreFeed_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long

    ilStartRow = vbcPreFeed.Value '+ 1  'Top location
    ilEndRow = vbcPreFeed.Value + vbcPreFeed.LargeChange ' + 1
    If ilEndRow > UBound(smPFSave, 2) Then
        If Trim$(smPFShow(1, UBound(smPFShow, 2))) <> "" Then
            ilEndRow = UBound(smPFSave, 2) 'include blank row as it might have data
        Else
            ilEndRow = UBound(smPFSave, 2) - 1
        End If
    End If
    llColor = pbcPreFeed.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smPFSave, 2) Then
            pbcPreFeed.ForeColor = DARKPURPLE
        Else
            pbcPreFeed.ForeColor = llColor
        End If
        For ilBox = imLBPFCtrls To UBound(tmPFCtrls) Step 1
            pbcPreFeed.CurrentX = tmPFCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPreFeed.CurrentY = tmPFCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 15 '+ fgBoxInsetY
            slStr = Trim$(smPFShow(ilBox, ilRow))
            pbcPreFeed.Print slStr
        Next ilBox
    Next ilRow
    pbcPreFeed.ForeColor = llColor
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
                    Select Case imPFBoxNo
                        Case FROMSTARTTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                         Case FROMENDTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                        Case TOSTARTTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                   End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub plcPreFeed_Click()
    pbcClickFocus.SetFocus
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    flTextHeight = pbcPreFeed.TextHeight("1") - 35

    plcPreFeed.Move 300, 600, pbcPreFeed.Width + fgPanelAdj + vbcPreFeed.Width, pbcPreFeed.Height + fgPanelAdj
    pbcPreFeed.Move plcPreFeed.Left + fgBevelX, plcPreFeed.Top + fgBevelY
    vbcPreFeed.Move pbcPreFeed.Left + pbcPreFeed.Width - 15, pbcPreFeed.Top

    'From Start Time
    gSetCtrl tmPFCtrls(FROMSTARTTIMEINDEX), 30, 375, 780, fgBoxGridH
    'From End Time
    gSetCtrl tmPFCtrls(FROMENDTIMEINDEX), 825, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY, 780, fgBoxGridH
    'From Day
    gSetCtrl tmPFCtrls(FROMDAYINDEX), 1620, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY, 435, fgBoxGridH
    'From Zone
    gSetCtrl tmPFCtrls(FROMZONEINDEX), 2070, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY, 435, fgBoxGridH
    'To Start Time
    gSetCtrl tmPFCtrls(TOSTARTTIMEINDEX), 2520, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY, 780, fgBoxGridH
    'To Day
    gSetCtrl tmPFCtrls(TODAYINDEX), 3315, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY, 435, fgBoxGridH
    
    plcBus.Move plcPreFeed.Left + plcPreFeed.Width + 300, plcPreFeed.Top + plcPreFeed.Height / 2 - plcBus.Height / 2, pbcBus.Width + fgPanelAdj + vbcBus.Width, pbcBus.Height + fgPanelAdj
    pbcBus.Move plcBus.Left + fgBevelX, plcBus.Top + fgBevelY
    vbcBus.Move pbcBus.Left + pbcBus.Width - 15, pbcBus.Top

    'From Bus
    gSetCtrl tmBusCtrls(FROMBUSINDEX), 30, 225, 690, fgBoxGridH
    'To Bus
    gSetCtrl tmBusCtrls(TOBUSINDEX), 735, tmBusCtrls(FROMBUSINDEX).fBoxY, 690, fgBoxGridH

    If igPreFeedType = 0 Then
        PreFeed.Width = 4790
        cmcDone.Left = cmcDone.Left - 1335
        cmcCancel.Left = cmcCancel.Left - 1335
        cmcUpdate.Left = cmcUpdate.Left - 1335
        imcTrash.Left = 4290
    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer

    imPffChg = False
    imPbfChg = False
    pbcPreFeed.Cls
    pbcBus.Cls
    For ilLoop = LBound(tmPFCtrls) To UBound(tmPFCtrls) Step 1
        tmPFCtrls(ilLoop).sShow = ""
    Next ilLoop
    ReDim smPFShow(0 To 6, 0 To 1) As String 'Values shown in program area
    ReDim smPFSave(0 To 6, 0 To 1) As String 'Values saved (program name) in program area
    ReDim lmPFSave(0 To 2, 0 To 1) As Long
    For ilLoop = LBound(smPFShow, 1) To UBound(smPFShow, 1) Step 1
        smPFShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smPFSave, 1) To UBound(smPFSave, 1) Step 1
        smPFSave(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(lmPFSave, 1) To UBound(lmPFSave, 1) Step 1
        lmPFSave(ilLoop, 1) = 0
    Next ilLoop
    vbcPreFeed.Min = LBONE  'LBound(smPFShow, 2)
    imSettingValue = True
    vbcPreFeed.Max = LBONE  'LBound(smPFShow, 2)
    imSettingValue = False
    
    ReDim smBusShow(0 To 2, 0 To 1) As String 'Values shown in program area
    ReDim smBusSave(0 To 2, 0 To 1) As String 'Values saved (program name) in program area
    ReDim lmBusSave(0 To 2, 0 To 1) As Long
    For ilLoop = LBound(smBusShow, 1) To UBound(smBusShow, 1) Step 1
        smBusShow(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(smBusSave, 1) To UBound(smBusSave, 1) Step 1
        smBusSave(ilLoop, 1) = ""
    Next ilLoop
    For ilLoop = LBound(lmBusSave, 1) To UBound(lmBusSave, 1) Step 1
        lmBusSave(ilLoop, 1) = 0
    Next ilLoop
    vbcBus.Min = LBONE  'LBound(smBusShow, 2)
    imSettingValue = True
    vbcBus.Max = LBONE  'LBound(smBusShow, 2)
    imSettingValue = False
    ReDim lmDelPFFCode(0 To 0) As Long
    ReDim lmDelPBFCode(0 To 0) As Long
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilRowNo As Integer
    Dim ilDay As Integer
    Dim slZone As String
    Dim slStr As String
    Dim ilPbf As Integer
    Dim ilIndex As Integer

    ilUpper = UBound(tmPFFInfo) + 1
    ReDim smPFShow(0 To 6, 0 To ilUpper) As String 'Values shown in program area
    ReDim smPFSave(0 To 6, 0 To ilUpper) As String 'Values saved (program name) in program area
    ReDim lmPFSave(0 To 2, 0 To ilUpper) As Long
    For ilIndex = 1 To ilUpper Step 1
        For ilLoop = LBound(smPFShow, 1) To UBound(smPFShow, 1) Step 1
            smPFShow(ilLoop, ilIndex) = ""
        Next ilLoop
        For ilLoop = LBound(smPFSave, 1) To UBound(smPFSave, 1) Step 1
            smPFSave(ilLoop, ilIndex) = ""
        Next ilLoop
        For ilLoop = LBound(lmPFSave, 1) To UBound(lmPFSave, 1) Step 1
            lmPFSave(ilLoop, ilIndex) = 0
        Next ilLoop
    Next ilIndex
    lmTiePffToPbfID = 1
    ReDim smAllBusSave(0 To 2, 0 To 1) As String
    ReDim lmAllBusSave(0 To 2, 0 To 1) As Long
    For ilRowNo = LBound(tmPFFInfo) To UBound(tmPFFInfo) - 1 Step 1
        'From Start Time
        gUnpackTime tmPFFInfo(ilRowNo).tPff.iFromStartTime(0), tmPFFInfo(ilRowNo).tPff.iFromStartTime(1), "A", "1", smPFSave(1, ilRowNo + 1)
        slStr = smPFSave(1, ilRowNo + 1)
        gSetShow pbcPreFeed, slStr, tmPFCtrls(FROMSTARTTIMEINDEX)
        smPFShow(FROMSTARTTIMEINDEX, ilRowNo + 1) = tmPFCtrls(FROMSTARTTIMEINDEX).sShow
        'From End Time
        gUnpackTime tmPFFInfo(ilRowNo).tPff.iFromEndTime(0), tmPFFInfo(ilRowNo).tPff.iFromEndTime(1), "A", "1", smPFSave(2, ilRowNo + 1)
        slStr = smPFSave(2, ilRowNo + 1)
        gSetShow pbcPreFeed, slStr, tmPFCtrls(FROMENDTIMEINDEX)
        smPFShow(FROMENDTIMEINDEX, ilRowNo + 1) = tmPFCtrls(FROMENDTIMEINDEX).sShow
        'From Day
        ilDay = tmPFFInfo(ilRowNo).tPff.iFromDay
        smPFSave(3, ilRowNo + 1) = Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su")
        slStr = smPFSave(3, ilRowNo + 1)
        gSetShow pbcPreFeed, slStr, tmPFCtrls(FROMDAYINDEX)
        smPFShow(FROMDAYINDEX, ilRowNo + 1) = tmPFCtrls(FROMDAYINDEX).sShow
        'Zone
        slZone = tmPFFInfo(ilRowNo).tPff.sFromZone
        smPFSave(4, ilRowNo + 1) = Switch(slZone = "E", "ET", slZone = "C", "CT", slZone = "M", "MT", slZone = "P", "PT", slZone = "A", "All")
        slStr = smPFSave(4, ilRowNo + 1)
        gSetShow pbcPreFeed, slStr, tmPFCtrls(FROMZONEINDEX)
        smPFShow(FROMZONEINDEX, ilRowNo + 1) = tmPFCtrls(FROMZONEINDEX).sShow
        'To Start Time
        gUnpackTime tmPFFInfo(ilRowNo).tPff.iToStartTime(0), tmPFFInfo(ilRowNo).tPff.iToStartTime(1), "A", "1", smPFSave(5, ilRowNo + 1)
        slStr = smPFSave(5, ilRowNo + 1)
        gSetShow pbcPreFeed, slStr, tmPFCtrls(TOSTARTTIMEINDEX)
        smPFShow(TOSTARTTIMEINDEX, ilRowNo + 1) = tmPFCtrls(TOSTARTTIMEINDEX).sShow
        'From Day
        ilDay = tmPFFInfo(ilRowNo).tPff.iToDay
        smPFSave(6, ilRowNo + 1) = Switch(ilDay = 0, "Mo", ilDay = 1, "Tu", ilDay = 2, "We", ilDay = 3, "Th", ilDay = 4, "Fr", ilDay = 5, "Sa", ilDay = 6, "Su")
        slStr = smPFSave(6, ilRowNo + 1)
        gSetShow pbcPreFeed, slStr, tmPFCtrls(TODAYINDEX)
        smPFShow(TODAYINDEX, ilRowNo + 1) = tmPFCtrls(TODAYINDEX).sShow
        lmPFSave(1, ilRowNo + 1) = tmPFFInfo(ilRowNo).tPff.lCode
        lmPFSave(2, ilRowNo + 1) = lmTiePffToPbfID
        For ilPbf = LBONE To UBound(tmPBFInfo) - 1 Step 1
            If tmPBFInfo(ilPbf).tPbf.lPffCode = tmPFFInfo(ilRowNo).tPff.lCode Then
                smAllBusSave(1, UBound(smAllBusSave, 2)) = tmPBFInfo(ilPbf).tPbf.sFromBus
                smAllBusSave(2, UBound(smAllBusSave, 2)) = tmPBFInfo(ilPbf).tPbf.sToBus
                lmAllBusSave(1, UBound(lmAllBusSave, 2)) = tmPBFInfo(ilPbf).tPbf.lCode
                lmAllBusSave(2, UBound(lmAllBusSave, 2)) = lmTiePffToPbfID
                ReDim Preserve smAllBusSave(0 To 2, 0 To UBound(smAllBusSave, 2) + 1) As String
                ReDim Preserve lmAllBusSave(0 To 2, 0 To UBound(lmAllBusSave, 2) + 1) As Long
            End If
        Next ilPbf
        lmTiePffToPbfID = lmTiePffToPbfID + 1
    Next ilRowNo
    lmPFSave(2, UBound(smPFSave, 2)) = lmTiePffToPbfID
    lmTiePffToPbfID = lmTiePffToPbfID + 1
    imSettingValue = True
    vbcPreFeed.Min = LBONE  'LBound(smPFShow, 2)
    imSettingValue = True
    If UBound(smPFShow, 2) - 1 <= vbcPreFeed.LargeChange + 1 Then ' + 1 Then
        vbcPreFeed.Max = LBONE  'LBound(smPFShow, 2)
    Else
        vbcPreFeed.Max = UBound(smPFShow, 2) - vbcPreFeed.LargeChange
    End If
    imSettingValue = True
    vbcPreFeed.Value = vbcPreFeed.Min
    imSettingValue = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    Dim slZone As String
    Dim ilPff As Integer

    ReDim tmPFFInfo(0 To UBound(smPFSave, 2) - 1) As PFFINFO
    For ilRowNo = LBONE To UBound(smPFSave, 2) - 1 Step 1
        ilPff = ilRowNo - 1
        tmPFFInfo(ilPff).tPff.lCode = lmPFSave(1, ilRowNo)
        If igPreFeedType = 0 Then
            tmPFFInfo(ilPff).tPff.sType = "D"
        Else
            tmPFFInfo(ilPff).tPff.sType = "E"
        End If
        tmPFFInfo(ilPff).tPff.iVefCode = igPreFeedVefCode
        tmPFFInfo(ilPff).tPff.sAirDay = sgPreFeedDay
        gPackDate sgPreFeedDate, tmPFFInfo(ilPff).tPff.iStartDate(0), tmPFFInfo(ilPff).tPff.iStartDate(1)
        'From Start Times
        gPackTime smPFSave(1, ilRowNo), tmPFFInfo(ilPff).tPff.iFromStartTime(0), tmPFFInfo(ilPff).tPff.iFromStartTime(1)
        'From End Time
        gPackTime smPFSave(2, ilRowNo), tmPFFInfo(ilPff).tPff.iFromEndTime(0), tmPFFInfo(ilPff).tPff.iFromEndTime(1)
        'From Day
        slDay = smPFSave(3, ilRowNo)
        tmPFFInfo(ilPff).tPff.iFromDay = Switch(slDay = "Mo", 0, slDay = "Tu", 1, slDay = "We", 2, slDay = "Th", 3, slDay = "Fr", 4, slDay = "Sa", 5, slDay = "Su", 6)
        'Zone
        slZone = smPFSave(4, ilRowNo)
        tmPFFInfo(ilPff).tPff.sFromZone = Switch(slZone = "ET", "E", slZone = "CT", "C", slZone = "MT", "M", slZone = "PT", "P", slZone = "All", "A")
        
        'To Start Times
        gPackTime smPFSave(5, ilRowNo), tmPFFInfo(ilPff).tPff.iToStartTime(0), tmPFFInfo(ilPff).tPff.iToStartTime(1)
        'From Day
        slDay = smPFSave(6, ilRowNo)
        tmPFFInfo(ilPff).tPff.iToDay = Switch(slDay = "Mo", 0, slDay = "Tu", 1, slDay = "We", 2, slDay = "Th", 3, slDay = "Fr", 4, slDay = "Sa", 5, slDay = "Su", 6)
        tmPFFInfo(ilPff).tPff.sUnused = ""
        tmPFFInfo(ilPff).lTiePffToPbfID = lmPFSave(2, ilRowNo)
    Next ilRowNo
    Exit Sub
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mPFInitNew(ilRowNo As Integer)
    Dim ilLoop As Integer
    Dim llTiePffToPdfID As Long

    llTiePffToPdfID = lmPFSave(2, ilRowNo)
    For ilLoop = LBound(smPFSave, 1) To UBound(smPFSave, 1) Step 1
        smPFSave(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smPFShow, 1) To UBound(smPFShow, 1) Step 1
        smPFShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(lmPFSave, 1) To UBound(lmPFSave, 1) Step 1
        lmPFSave(ilLoop, ilRowNo) = 0
    Next ilLoop
    If llTiePffToPdfID <= 0 Then
        lmPFSave(2, ilRowNo) = lmTiePffToPbfID
        lmTiePffToPbfID = lmTiePffToPbfID + 1
    Else
        lmPFSave(2, ilRowNo) = llTiePffToPdfID
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mBusInitNew(ilRowNo As Integer)
    Dim ilLoop As Integer

    For ilLoop = LBound(smBusSave, 1) To UBound(smBusSave, 1) Step 1
        smBusSave(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smBusShow, 1) To UBound(smBusShow, 1) Step 1
        smBusShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(lmBusSave, 1) To UBound(lmBusSave, 1) Step 1
        lmBusSave(ilLoop, ilRowNo) = 0
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPFSetShow                      *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPFSetShow(ilBoxNo As Integer, ilRetainFocus As Integer)
'
'   mPFSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If Not ilRetainFocus Then
        pbcArrow.Visible = False
        lacPreFeedFrame.Visible = False
    End If
    If (ilBoxNo < imLBPFCtrls) Or (ilBoxNo > UBound(tmPFCtrls)) Then
        Exit Sub
    End If

    If Not ilRetainFocus Then
        smFocusOnPForBus = ""
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case FROMSTARTTIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPreFeed, slStr, tmPFCtrls(ilBoxNo)
                smPFShow(ilBoxNo, imPFRowNo) = tmPFCtrls(ilBoxNo).sShow
                If Trim$(smPFSave(1, imPFRowNo)) <> slStr Then
                    imPffChg = True
                    smPFSave(1, imPFRowNo) = slStr
                End If
            End If
        Case FROMENDTIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPreFeed, slStr, tmPFCtrls(ilBoxNo)
                smPFShow(ilBoxNo, imPFRowNo) = tmPFCtrls(ilBoxNo).sShow
                If Trim$(smPFSave(2, imPFRowNo)) <> slStr Then
                    imPffChg = True
                    smPFSave(2, imPFRowNo) = slStr
                End If
            End If
        Case FROMDAYINDEX
            lbcDay.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcPreFeed, slStr, tmPFCtrls(ilBoxNo)
            smPFShow(ilBoxNo, imPFRowNo) = tmPFCtrls(ilBoxNo).sShow
            If Trim$(smPFSave(3, imPFRowNo)) <> slStr Then
                imPffChg = True
                smPFSave(3, imPFRowNo) = slStr
            End If
        Case FROMZONEINDEX
            lbcZone.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcPreFeed, slStr, tmPFCtrls(ilBoxNo)
            smPFShow(ilBoxNo, imPFRowNo) = tmPFCtrls(ilBoxNo).sShow
            If Trim$(smPFSave(4, imPFRowNo)) <> slStr Then
                imPffChg = True
                smPFSave(4, imPFRowNo) = slStr
            End If
        Case TOSTARTTIMEINDEX
            plcTme.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If gValidTime(slStr) Then
                gSetShow pbcPreFeed, slStr, tmPFCtrls(ilBoxNo)
                smPFShow(ilBoxNo, imPFRowNo) = tmPFCtrls(ilBoxNo).sShow
                If Trim$(smPFSave(5, imPFRowNo)) <> slStr Then
                    imPffChg = True
                    smPFSave(5, imPFRowNo) = slStr
                End If
            End If
        Case TODAYINDEX
            lbcDay.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcPreFeed, slStr, tmPFCtrls(ilBoxNo)
            smPFShow(ilBoxNo, imPFRowNo) = tmPFCtrls(ilBoxNo).sShow
            If Trim$(smPFSave(6, imPFRowNo)) <> slStr Then
                imPffChg = True
                smPFSave(6, imPFRowNo) = slStr
            End If
            mAddPFRow
    End Select
    mSetCommands
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mBusSetShow                     *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mBusSetShow(ilBoxNo As Integer)
'
'   mPFSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcBusArrow.Visible = False
    lacBusFrame.Visible = False
    If (ilBoxNo < imLBBusCtrls) Or (ilBoxNo > UBound(tmBusCtrls)) Then
        Exit Sub
    End If

    smFocusOnPForBus = "P"
    Select Case ilBoxNo 'Branch on box type (control)
        Case FROMBUSINDEX
            edcBus.Visible = False
            slStr = edcBus.Text
            gSetShow pbcBus, slStr, tmBusCtrls(ilBoxNo)
            smBusShow(ilBoxNo, imBusRowNo) = tmBusCtrls(ilBoxNo).sShow
            If StrComp(slStr, smBusSave(1, imBusRowNo), vbTextCompare) <> 0 Then
                imPbfChg = True
                smBusSave(1, imBusRowNo) = slStr
            End If
        Case TOBUSINDEX
            edcBus.Visible = False
            slStr = edcBus.Text
            gSetShow pbcBus, slStr, tmBusCtrls(ilBoxNo)
            smBusShow(ilBoxNo, imBusRowNo) = tmBusCtrls(ilBoxNo).sShow
            If StrComp(slStr, smBusSave(2, imBusRowNo), vbTextCompare) <> 0 Then
                imPbfChg = True
                smBusSave(2, imBusRowNo) = slStr
            End If
            mAddBusRow
    End Select
    lmBusSave(2, imBusRowNo) = lmCurTieID
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilAltered As Integer
    If (imBypassSetting) Or (Not imUpdateAllowed) Then
        Exit Sub
    End If
    ilAltered = imPffChg
    If imPbfChg Then
        ilAltered = True
    End If
    If (Not ilAltered) And (UBound(lmDelPFFCode) > LBound(lmDelPFFCode)) Then
        ilAltered = True
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields() = YES) And (ilAltered) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields() As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRowNo As Integer
    For ilRowNo = LBONE To UBound(smPFSave, 2) - 1 Step 1
        If Trim$(smPFSave(1, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smPFSave(2, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smPFSave(3, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smPFSave(4, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smPFSave(5, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
        If Trim$(smPFSave(6, ilRowNo)) = "" Then
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestBusFields                  *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestBusFields() As Integer
'
'   iRet = mTestBusFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRowNo As Integer
    For ilRowNo = LBONE To UBound(smBusSave, 2) - 1 Step 1
        If Trim$(smBusSave(1, ilRowNo)) = "" Then
            mTestBusFields = NO
            Exit Function
        End If
        If Trim$(smBusSave(2, ilRowNo)) = "" Then
            mTestBusFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestBusFields = YES
End Function
Private Sub tmcDrag_Timer()
    Dim ilPreFeedRow As Integer
    Dim ilBusRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            If smFocusOnPForBus <> "B" Then
                mBusSetShow imBusBoxNo
                imBusBoxNo = -1
                imBusRowNo = -1
                ilPreFeedRow = vbcPreFeed.LargeChange + 1
                If UBound(smPFSave, 2) > ilPreFeedRow Then
                    ilMaxRow = ilPreFeedRow
                Else
                    ilMaxRow = UBound(smPFSave, 2)
                End If
                For ilRow = 1 To ilMaxRow Step 1
                    If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY + tmPFCtrls(FROMSTARTTIMEINDEX).fBoxH)) Then
                        mClearBus
                        mPFSetShow imPFBoxNo, False
                        imPFBoxNo = -1
                        imPFRowNo = -1
                        imPFRowNo = ilRow + vbcPreFeed.Value - 1
                        lacPreFeedFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                        lacPreFeedFrame.Move 0, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacPreFeedFrame.Visible = True
                        pbcArrow.Move pbcArrow.Left, plcPreFeed.Top + tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15) + 45
                        pbcArrow.Visible = True
                        imcTrash.Enabled = True
                        lacPreFeedFrame.Drag vbBeginDrag
                        lacPreFeedFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                        lmCurTieID = lmPFSave(2, imPFRowNo)
                        mMovAllBusToCtrl lmCurTieID
                        pbcBus.Cls
                        pbcBus_Paint
                        Exit Sub
                    End If
                Next ilRow
            Else
                ilBusRow = vbcBus.LargeChange + 1
                If UBound(smBusSave, 2) > ilBusRow Then
                    ilMaxRow = ilBusRow
                Else
                    ilMaxRow = UBound(smBusSave, 2)
                End If
                For ilRow = 1 To ilMaxRow Step 1
                    If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmBusCtrls(FROMBUSINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmBusCtrls(FROMBUSINDEX).fBoxY + tmBusCtrls(FROMBUSINDEX).fBoxH)) Then
                        mBusSetShow imBusBoxNo
                        imBusBoxNo = -1
                        imBusRowNo = -1
                        imBusRowNo = ilRow + vbcBus.Value - 1
                        lacBusFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                        lacBusFrame.Move 0, tmBusCtrls(FROMBUSINDEX).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacBusFrame.Visible = True
                        pbcBusArrow.Move pbcBusArrow.Left, plcBus.Top + tmBusCtrls(FROMBUSINDEX).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15) + 45
                        pbcBusArrow.Visible = True
                        imcTrash.Enabled = True
                        lacBusFrame.Drag vbBeginDrag
                        lacBusFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                        Exit Sub
                    End If
                Next ilRow
            End If
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select

End Sub

Private Sub vbcBus_Change()
    If imSettingValue Then
        pbcBus.Cls
        pbcBus_Paint
        imSettingValue = False
    Else
        mBusSetShow imBusBoxNo
        imBusBoxNo = -1
        imBusRowNo = -1
        pbcBus.Cls
        pbcBus_Paint
    End If

End Sub

Private Sub vbcBus_GotFocus()
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    mPFSetShow imPFBoxNo, True
    gCtrlGotFocus ActiveControl
End Sub

Private Sub vbcPreFeed_Change()
    If imSettingValue Then
        pbcPreFeed.Cls
        pbcPreFeed_Paint
        imSettingValue = False
    Else
        mBusSetShow imBusBoxNo
        imBusBoxNo = -1
        imBusRowNo = -1
        mClearBus
        mPFSetShow imPFBoxNo, False
        imPFBoxNo = -1
        imPFRowNo = -1
        pbcPreFeed.Cls
        pbcPreFeed_Paint
    End If
End Sub

Private Sub vbcPreFeed_GotFocus()
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    mClearBus
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = -1
    imPFRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPFEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPFEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If ilBoxNo < imLBPFCtrls Or ilBoxNo > UBound(tmPFCtrls) Then
        Exit Sub
    End If
    If (imPFRowNo < vbcPreFeed.Value) Or (imPFRowNo >= vbcPreFeed.Value + vbcPreFeed.LargeChange + 1) Then
        mClearBus
        mPFSetShow ilBoxNo, False
        Exit Sub
    End If
    lacPreFeedFrame.Move 0, tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15) - 30
    lacPreFeedFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcPreFeed.Top + tmPFCtrls(FROMSTARTTIMEINDEX).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    smFocusOnPForBus = "P"
    Select Case ilBoxNo 'Branch on box type (control)
        Case FROMSTARTTIMEINDEX
            edcDropDown.Width = tmPFCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPreFeed, edcDropDown, tmPFCtrls(ilBoxNo).fBoxX, tmPFCtrls(ilBoxNo).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smPFSave(1, imPFRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imPFRowNo > UBound(smPFSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case FROMENDTIMEINDEX
            edcDropDown.Width = tmPFCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPreFeed, edcDropDown, tmPFCtrls(ilBoxNo).fBoxX, tmPFCtrls(ilBoxNo).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smPFSave(2, imPFRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imPFRowNo > UBound(smPFSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case FROMDAYINDEX
            lbcDay.Height = gListBoxHeight(lbcDay.ListCount, 7)
            edcDropDown.Width = tmPFCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 2
            gMoveTableCtrl pbcPreFeed, edcDropDown, tmPFCtrls(ilBoxNo).fBoxX, tmPFCtrls(ilBoxNo).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcDay.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imPFRowNo - vbcPreFeed.Value <= vbcPreFeed.LargeChange \ 2 Then
                lbcDay.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcDay.Move edcDropDown.Left, edcDropDown.Top - lbcDay.Height
            End If
            imChgMode = True
            slStr = Trim$(smPFSave(3, imPFRowNo))
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcDay
                If gLastFound(lbcDay) >= 0 Then
                    lbcDay.ListIndex = gLastFound(lbcDay)
                Else
                    lbcDay.ListIndex = 0
                End If
            Else
                lbcDay.ListIndex = 0
            End If
            If lbcDay.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcDay.List(lbcDay.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case FROMZONEINDEX
            lbcZone.Height = gListBoxHeight(lbcZone.ListCount, 7)
            edcDropDown.Width = tmPFCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 3
            gMoveTableCtrl pbcPreFeed, edcDropDown, tmPFCtrls(ilBoxNo).fBoxX, tmPFCtrls(ilBoxNo).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcZone.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imPFRowNo - vbcPreFeed.Value <= vbcPreFeed.LargeChange \ 2 Then
                lbcZone.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcZone.Move edcDropDown.Left, edcDropDown.Top - lbcZone.Height
            End If
            imChgMode = True
            slStr = Trim$(smPFSave(4, imPFRowNo))
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcZone
                If gLastFound(lbcZone) >= 0 Then
                    lbcZone.ListIndex = gLastFound(lbcZone)
                Else
                    lbcZone.ListIndex = 0
                End If
            Else
                lbcZone.ListIndex = 0
            End If
            If lbcZone.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcZone.List(lbcZone.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TOSTARTTIMEINDEX
            edcDropDown.Width = tmPFCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcPreFeed, edcDropDown, tmPFCtrls(ilBoxNo).fBoxX, tmPFCtrls(ilBoxNo).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            slStr = Trim$(smPFSave(5, imPFRowNo))
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If imPFRowNo > UBound(smPFSave, 2) - 1 Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case TODAYINDEX
            lbcDay.Height = gListBoxHeight(lbcDay.ListCount, 7)
            edcDropDown.Width = tmPFCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 2
            gMoveTableCtrl pbcPreFeed, edcDropDown, tmPFCtrls(ilBoxNo).fBoxX, tmPFCtrls(ilBoxNo).fBoxY + (imPFRowNo - vbcPreFeed.Value) * (fgBoxGridH + 15)
            edcDropDown.Left = edcDropDown.Left - cmcDropDown.Width
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcDay.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If imPFRowNo - vbcPreFeed.Value <= vbcPreFeed.LargeChange \ 2 Then
                lbcDay.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcDay.Move edcDropDown.Left, edcDropDown.Top - lbcDay.Height
            End If
            imChgMode = True
            slStr = Trim$(smPFSave(6, imPFRowNo))
            If slStr <> "" Then
                gFindMatch slStr, 0, lbcDay
                If gLastFound(lbcDay) >= 0 Then
                    lbcDay.ListIndex = gLastFound(lbcDay)
                Else
                    lbcDay.ListIndex = 0
                End If
            Else
                lbcDay.ListIndex = 0
            End If
            If lbcDay.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcDay.List(lbcDay.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    If (lmCurTieID <> -1) And (lmCurTieID <> lmPFSave(2, imPFRowNo)) And (igPreFeedType = 1) Then
        mMovCtrlToAllBus
        lmCurTieID = -1
    End If
    If (lmCurTieID = -1) And (igPreFeedType = 1) Then
        lmCurTieID = lmPFSave(2, imPFRowNo)
        mMovAllBusToCtrl lmCurTieID
        pbcBus.Cls
        pbcBus_Paint
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPFEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mBusEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    
    
    If imPFBoxNo < imLBPFCtrls Or imPFBoxNo > UBound(tmPFCtrls) Then
        Exit Sub
    End If
    If (imPFRowNo < vbcPreFeed.Value) Or (imPFRowNo >= vbcPreFeed.Value + vbcPreFeed.LargeChange + 1) Then
        mClearBus
        mPFSetShow ilBoxNo, False
        Exit Sub
    End If
    If (smFocusOnPForBus <> "P") And (smFocusOnPForBus <> "B") Then
        Exit Sub
    End If
    If ilBoxNo < imLBBusCtrls Or ilBoxNo > UBound(tmBusCtrls) Then
        Exit Sub
    End If
    If (imBusRowNo < vbcBus.Value) Or (imBusRowNo >= vbcBus.Value + vbcBus.LargeChange + 1) Then
        mBusSetShow ilBoxNo
        Exit Sub
    End If
    lacBusFrame.Move 0, tmBusCtrls(FROMBUSINDEX).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15) - 30
    lacBusFrame.Visible = True
    pbcBusArrow.Move pbcBusArrow.Left, plcBus.Top + tmBusCtrls(FROMBUSINDEX).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15) + 45
    pbcBusArrow.Visible = True
    smFocusOnPForBus = "B"
    Select Case ilBoxNo 'Branch on box type (control)
        Case FROMBUSINDEX
            edcBus.Width = tmBusCtrls(ilBoxNo).fBoxW
            edcBus.MaxLength = 5
            gMoveTableCtrl pbcBus, edcBus, tmBusCtrls(ilBoxNo).fBoxX, tmBusCtrls(ilBoxNo).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smBusSave(1, imBusRowNo))
            edcBus.Text = slStr
            edcBus.SelStart = 0
            edcBus.SelLength = Len(edcBus.Text)
            edcBus.Visible = True  'Set visibility
            edcBus.SetFocus
        Case TOBUSINDEX
            edcBus.Width = tmBusCtrls(ilBoxNo).fBoxW
            edcBus.MaxLength = 5
            gMoveTableCtrl pbcBus, edcBus, tmBusCtrls(ilBoxNo).fBoxX, tmBusCtrls(ilBoxNo).fBoxY + (imBusRowNo - vbcBus.Value) * (fgBoxGridH + 15)
            slStr = Trim$(smBusSave(2, imBusRowNo))
            edcBus.Text = slStr
            edcBus.SelStart = 0
            edcBus.SelLength = Len(edcBus.Text)
            edcBus.Visible = True  'Set visibility
            edcBus.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim slDay As String
    Dim ilFromDay As Integer
    Dim ilToDay As Integer
    
    If Trim$(smPFSave(1, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("From Start Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMSTARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Not gValidTime(smPFSave(1, ilRowNo)) Then
        Beep
        ilRes = MsgBox("From Start Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMSTARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smPFSave(2, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("From End Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMENDTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Not gValidTime(smPFSave(2, ilRowNo)) Then
        Beep
        ilRes = MsgBox("From End Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMENDTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If gTimeToLong(smPFSave(1, ilRowNo), False) > gTimeToLong(smPFSave(2, ilRowNo), True) Then
        Beep
        ilRes = MsgBox("From Start Time must be prior to From End Time", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMSTARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smPFSave(3, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("From Day must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMDAYINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smPFSave(4, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("From Zone must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = FROMZONEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smPFSave(5, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("To Start Time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = TOSTARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Not gValidTime(smPFSave(5, ilRowNo)) Then
        Beep
        ilRes = MsgBox("To Start Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = TOSTARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smPFSave(6, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("To Day must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = TODAYINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    slDay = smPFSave(3, ilRowNo)
    ilFromDay = Switch(slDay = "Mo", 0, slDay = "Tu", 1, slDay = "We", 2, slDay = "Th", 3, slDay = "Fr", 4, slDay = "Sa", 5, slDay = "Su", 6)
    slDay = smPFSave(6, ilRowNo)
    ilToDay = Switch(slDay = "Mo", 0, slDay = "Tu", 1, slDay = "We", 2, slDay = "Th", 3, slDay = "Fr", 4, slDay = "Sa", 5, slDay = "Su", 6)
    If ilToDay > ilFromDay Then
        Beep
        ilRes = MsgBox("To Day must be prior or equal to From Day", vbOKOnly + vbExclamation, "Incomplete")
        imPFBoxNo = TODAYINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestBusSaveFields              *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestBusSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestBusSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If Trim$(smBusSave(1, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("From Bus must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBusBoxNo = FROMSTARTTIMEINDEX
        mTestBusSaveFields = NO
        Exit Function
    End If
    If Trim$(smBusSave(2, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("To Bus must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBusBoxNo = FROMENDTIMEINDEX
        mTestBusSaveFields = NO
        Exit Function
    End If
    mTestBusSaveFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestBusSaveFields              *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestAllBusSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestBusSaveFields(ilRowNo)
'   Where:
'       ilRowNo (I)- row number to be checked
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If Trim$(smAllBusSave(1, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("From Bus must be specified", vbOKOnly + vbExclamation, "Incomplete")
        'imBusBoxNo = FROMSTARTTIMEINDEX
        mTestAllBusSaveFields = NO
        Exit Function
    End If
    If Trim$(smAllBusSave(2, ilRowNo)) = "" Then
        Beep
        ilRes = MsgBox("To Bus must be specified", vbOKOnly + vbExclamation, "Incomplete")
        'imBusBoxNo = FROMENDTIMEINDEX
        mTestAllBusSaveFields = NO
        Exit Function
    End If
    mTestAllBusSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPFSetFocus                     *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mPFSetFocus(ilBoxNo As Integer)
'
'   mPFSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBPFCtrls) Or (ilBoxNo > UBound(tmPFCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case FROMSTARTTIMEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case FROMENDTIMEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case FROMDAYINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case FROMZONEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TOSTARTTIMEINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TODAYINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPFSetFocus                     *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mBusSetFocus(ilBoxNo As Integer)
'
'   mPFSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If smFocusOnPForBus <> "B" Then
        Exit Sub
    End If
    If (ilBoxNo < imLBBusCtrls) Or (ilBoxNo > UBound(tmBusCtrls)) Then
        Exit Sub
    End If

    On Error Resume Next
    Select Case ilBoxNo 'Branch on box type (control)
        Case FROMBUSINDEX
            edcBus.SetFocus
        Case TOBUSINDEX
            edcBus.SetFocus
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMovAllBusToCtrl                *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMovAllBusToCtrl(llTieID As Long)
    Dim ilRow As Integer
    Dim ilUpper As Integer
    Dim slStr As String
    
    ReDim smBusSave(0 To 2, 0 To 1) As String
    ReDim smBusShow(0 To 2, 0 To 1) As String
    ReDim lmBusSave(0 To 2, 0 To 1) As Long
    ilUpper = 1
    For ilRow = LBONE To UBound(lmAllBusSave, 2) - 1 Step 1
        If lmAllBusSave(2, ilRow) = llTieID Then
            smBusSave(1, ilUpper) = smAllBusSave(1, ilRow)
            slStr = smBusSave(1, ilUpper)
            gSetShow pbcBus, slStr, tmBusCtrls(FROMBUSINDEX)
            smBusShow(FROMBUSINDEX, ilUpper) = tmBusCtrls(FROMBUSINDEX).sShow
            smBusSave(2, ilUpper) = smAllBusSave(2, ilRow)
            slStr = smBusSave(2, ilUpper)
            gSetShow pbcBus, slStr, tmBusCtrls(TOBUSINDEX)
            smBusShow(TOBUSINDEX, ilUpper) = tmBusCtrls(TOBUSINDEX).sShow
            lmBusSave(1, ilUpper) = lmAllBusSave(1, ilRow)
            lmBusSave(2, ilUpper) = lmAllBusSave(2, ilRow)
            lmAllBusSave(2, ilRow) = -1
            ReDim Preserve smBusSave(0 To 2, 0 To ilUpper + 1) As String
            ReDim Preserve smBusShow(0 To 2, 0 To ilUpper + 1) As String
            ReDim Preserve lmBusSave(0 To 2, 0 To ilUpper + 1) As Long
            ilUpper = ilUpper + 1
        End If
    Next ilRow
    imSettingValue = True
    vbcBus.Min = LBONE  'LBound(smBusShow, 2)
    imSettingValue = True
    If UBound(smBusShow, 2) - 1 <= vbcBus.LargeChange + 1 Then ' + 1 Then
        vbcBus.Max = LBONE  'LBound(smBusShow, 2)
    Else
        vbcBus.Max = UBound(smBusShow, 2) - vbcBus.LargeChange
    End If
    imSettingValue = True
    vbcBus.Value = vbcBus.Min
    imSettingValue = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMovCtrlToAllBus()
    Dim ilRow As Integer
    Dim ilFound As Integer
    Dim ilAll As Integer
    Dim ilUpper As Integer
    
    For ilRow = LBONE To UBound(smBusSave, 2) - 1 Step 1
        ilFound = False
        For ilAll = LBONE To UBound(lmAllBusSave, 2) - 1 Step 1
            If lmAllBusSave(2, ilAll) = -1 Then
                ilFound = True
                smAllBusSave(1, ilAll) = smBusSave(1, ilRow)
                smAllBusSave(2, ilAll) = smBusSave(2, ilRow)
                lmAllBusSave(1, ilAll) = lmBusSave(1, ilRow)
                lmAllBusSave(2, ilAll) = lmBusSave(2, ilRow)
                Exit For
            End If
        Next ilAll
        If Not ilFound Then
            ilUpper = UBound(smAllBusSave, 2)
            smAllBusSave(1, ilUpper) = smBusSave(1, ilRow)
            smAllBusSave(2, ilUpper) = smBusSave(2, ilRow)
            lmAllBusSave(1, ilUpper) = lmBusSave(1, ilRow)
            lmAllBusSave(2, ilUpper) = lmBusSave(2, ilRow)
            ReDim Preserve smAllBusSave(0 To 2, 0 To ilUpper + 1) As String
            ReDim Preserve lmAllBusSave(0 To 2, 0 To ilUpper + 1) As Long
        End If
    Next ilRow
End Sub


Private Sub mClearBus()
    If (lmCurTieID <> -1) And (igPreFeedType = 1) Then
        mMovCtrlToAllBus
        lmCurTieID = -1
        ReDim smBusSave(0 To 2, 0 To 1) As String
        ReDim smBusShow(0 To 2, 0 To 1) As String
        ReDim lmBusSave(0 To 2, 0 To 1) As Long
        imSettingValue = True
        vbcBus.Min = LBONE  'LBound(smBusShow, 2)
        imSettingValue = True
        vbcBus.Max = LBONE  'LBound(smBusShow, 2)
        imSettingValue = True
        vbcBus.Value = vbcBus.Min
        imSettingValue = True
        pbcBus.Cls
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:9/24/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilLoop As Integer
    Dim ilNew As Integer
    If (imPffChg Or imPbfChg) And (UBound(smPFSave, 2) > LBONE) Or (UBound(lmDelPFFCode) > LBound(lmDelPFFCode) Or UBound(smBusSave, 2) > LBONE) Or (UBound(lmDelPBFCode) > LBound(lmDelPBFCode)) Then
        If ilAsk Then
            ilNew = True
            If UBound(tmPFFInfo) > LBound(tmPFFInfo) Then
                ilNew = False
            End If
            If Not ilNew Then
                slMess = "Save Changes"
            Else
                slMess = "Add Changes"
            End If
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                Exit Function
            End If
            If ilRes = vbYes Then
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        Else
            ilRes = mSaveRec()
            mSaveRecChg = ilRes
            Exit Function
        End If
    End If
    mSaveRecChg = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/29/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilPFRowNo As Integer
    Dim ilBusRowNo As Integer
    Dim tlPbf As PBF
    
    mBusSetShow imBusBoxNo
    imBusBoxNo = -1
    imBusRowNo = -1
    mClearBus
    mPFSetShow imPFBoxNo, False
    imPFBoxNo = -1
    imPFRowNo = -1
    For ilPFRowNo = 1 To UBound(smPFSave, 2) - 1 Step 1
        If mTestSaveFields(ilPFRowNo) = NO Then
            mSaveRec = False
            imPFRowNo = ilPFRowNo
            Exit Function
        End If
        For ilBusRowNo = 1 To UBound(smAllBusSave, 2) - 1 Step 1
            If lmAllBusSave(2, ilBusRowNo) = lmPFSave(2, ilPFRowNo) Then
                If mTestAllBusSaveFields(ilBusRowNo) = NO Then
                    imBusRowNo = -1
                    imBusBoxNo = -1
                    mSaveRec = False
                    imPFRowNo = ilPFRowNo
                    Exit Function
                End If
            End If
        Next ilBusRowNo
    Next ilPFRowNo
    mMoveCtrlToRec
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    ilRet = btrBeginTrans(hmPff, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 1", vbOKOnly + vbExclamation, "Invoice")
        Exit Function
    End If
    'Remove Bus records
    If (igPreFeedType = 1) Then
        For ilBusRowNo = LBound(lmDelPBFCode) To UBound(lmDelPBFCode) - 1 Step 1
            Do
                tmPbfSrchKey0.lCode = lmDelPBFCode(ilBusRowNo)
                ilRet = btrGetEqual(hmPbf, tmPbf, imPbfRecLen, tmPbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmPff)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 2", vbOKOnly + vbExclamation, "Pre-Feed Bus")
                    Exit Function
                End If
                ilRet = btrDelete(hmPbf)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                ilCRet = btrAbortTrans(hmPff)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 3", vbOKOnly + vbExclamation, "Pre-Feed Bus")
                Exit Function
            End If
        Next ilBusRowNo
    End If
    'Remove Prefeed records
    For ilPFRowNo = LBound(lmDelPFFCode) To UBound(lmDelPFFCode) - 1 Step 1
        Do
            tmPbfSrchKey1.lCode = lmDelPFFCode(ilPFRowNo)
            ilRet = btrGetEqual(hmPbf, tmPbf, imPbfRecLen, tmPbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            ilRet = btrDelete(hmPbf)
        Loop While ilRet = BTRV_ERR_CONFLICT
        Do
            tmPffSrchKey0.lCode = lmDelPFFCode(ilPFRowNo)
            ilRet = btrGetEqual(hmPff, tmPFF, imPffRecLen, tmPffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                ilCRet = btrAbortTrans(hmPff)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 4", vbOKOnly + vbExclamation, "Pre-Feed Bus")
                Exit Function
            End If
            ilRet = btrDelete(hmPff)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            ilCRet = btrAbortTrans(hmPff)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 5", vbOKOnly + vbExclamation, "Pre-Feed Bus")
            Exit Function
        End If
    Next ilPFRowNo
    'Add PreFeed and corresponding Bus
    For ilPFRowNo = LBound(tmPFFInfo) To UBound(tmPFFInfo) - 1 Step 1
        Do
            If tmPFFInfo(ilPFRowNo).tPff.lCode > 0 Then
                tmPffSrchKey0.lCode = tmPFFInfo(ilPFRowNo).tPff.lCode
                ilRet = btrGetEqual(hmPff, tmPFF, imPffRecLen, tmPffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    ilCRet = btrAbortTrans(hmPff)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 6", vbOKOnly + vbExclamation, "Pre-Feed Bus")
                    Exit Function
                End If
                ilRet = btrUpdate(hmPff, tmPFFInfo(ilPFRowNo).tPff, imPffRecLen)
            Else
                ilRet = btrInsert(hmPff, tmPFFInfo(ilPFRowNo).tPff, imPffRecLen, INDEXKEY0)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            ilCRet = btrAbortTrans(hmPff)
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 7", vbOKOnly + vbExclamation, "Pre-Feed Bus")
            Exit Function
        End If
        If (igPreFeedType = 1) Then
            For ilBusRowNo = LBONE To UBound(smAllBusSave, 2) - 1 Step 1
                If lmAllBusSave(2, ilBusRowNo) = tmPFFInfo(ilPFRowNo).lTiePffToPbfID Then
                    tmPbf.lCode = lmAllBusSave(1, ilBusRowNo)
                    tmPbf.lPffCode = tmPFFInfo(ilPFRowNo).tPff.lCode
                    tmPbf.sFromBus = smAllBusSave(1, ilBusRowNo)
                    tmPbf.sToBus = smAllBusSave(2, ilBusRowNo)
                    tmPbf.sUnused = ""
                    Do
                        If tmPbf.lCode > 0 Then
                            tmPbfSrchKey0.lCode = tmPbf.lCode
                            ilRet = btrGetEqual(hmPbf, tlPbf, imPbfRecLen, tmPbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet <> BTRV_ERR_NONE Then
                                If ilRet >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                ilCRet = btrAbortTrans(hmPff)
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 8", vbOKOnly + vbExclamation, "Pre-Feed Bus")
                                Exit Function
                            End If
                            ilRet = btrUpdate(hmPbf, tmPbf, imPbfRecLen)
                        Else
                            ilRet = btrInsert(hmPbf, tmPbf, imPbfRecLen, INDEXKEY0)
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    If ilRet <> BTRV_ERR_NONE Then
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        ilCRet = btrAbortTrans(hmPff)
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Update Not Completed, Try Later.  Error #" & str$(ilRet) & " at 9", vbOKOnly + vbExclamation, "Pre-Feed Bus")
                        Exit Function
                    End If
                End If
            Next ilBusRowNo
        End If
    Next ilPFRowNo
    
    ilRet = btrEndTrans(hmPff)
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function

End Function

Private Sub mAddPFRow()
    Dim ilLoop As Integer
    If imPFRowNo >= UBound(smPFSave, 2) Then
        imPffChg = True
        ReDim Preserve smPFShow(0 To 6, 0 To imPFRowNo + 1) As String 'Values shown in program area
        ReDim Preserve smPFSave(0 To 6, 0 To imPFRowNo + 1) As String 'Values saved (program name) in program area
        ReDim Preserve lmPFSave(0 To 2, 0 To imPFRowNo + 1) As Long 'Values saved (program name) in program area
        For ilLoop = LBound(smPFShow, 1) To UBound(smPFShow, 1) Step 1
            smPFShow(ilLoop, imPFRowNo + 1) = ""
        Next ilLoop
        For ilLoop = LBound(smPFSave, 1) To UBound(smPFSave, 1) Step 1
            smPFSave(ilLoop, imPFRowNo + 1) = ""
        Next ilLoop
        For ilLoop = LBound(lmPFSave, 1) To UBound(lmPFSave, 1) Step 1
            lmPFSave(ilLoop, imPFRowNo + 1) = 0
        Next ilLoop
        mPFInitNew imPFRowNo + 1
        If UBound(smPFSave, 2) <= vbcPreFeed.LargeChange Then 'was <=
            vbcPreFeed.Max = LBONE  'LBound(smPFSave, 2) '- 1
        Else
            vbcPreFeed.Max = UBound(smPFSave, 2) - vbcPreFeed.LargeChange '- 1
        End If
    End If

End Sub

Private Sub mAddBusRow()
    Dim ilLoop As Integer
    
    If imBusRowNo >= UBound(smBusSave, 2) Then
        imPbfChg = True
        ReDim Preserve smBusShow(0 To 2, 0 To imBusRowNo + 1) As String 'Values shown in program area
        ReDim Preserve smBusSave(0 To 2, 0 To imBusRowNo + 1) As String 'Values saved (program name) in program area
        ReDim Preserve lmBusSave(0 To 2, 0 To imBusRowNo + 1) As Long 'Values saved (program name) in program area
        For ilLoop = LBound(smBusShow, 1) To UBound(smBusShow, 1) Step 1
            smBusShow(ilLoop, imBusRowNo + 1) = ""
        Next ilLoop
        For ilLoop = LBound(smBusSave, 1) To UBound(smBusSave, 1) Step 1
            smBusSave(ilLoop, imBusRowNo + 1) = ""
        Next ilLoop
        For ilLoop = LBound(lmBusSave, 1) To UBound(lmBusSave, 1) Step 1
            lmBusSave(ilLoop, imBusRowNo + 1) = 0
        Next ilLoop
        mBusInitNew imBusRowNo + 1
        If UBound(smBusSave, 2) <= vbcBus.LargeChange Then 'was <=
            vbcBus.Max = LBONE  'LBound(smBusSave, 2) '- 1
        Else
            vbcBus.Max = UBound(smBusSave, 2) - vbcBus.LargeChange '- 1
        End If
    End If

End Sub
