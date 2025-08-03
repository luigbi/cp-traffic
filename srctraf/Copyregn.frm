VERSION 5.00
Begin VB.Form CopyRegn 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2280
   ClientLeft      =   1020
   ClientTop       =   2595
   ClientWidth     =   5625
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   5625
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
      Height          =   315
      Left            =   165
      TabIndex        =   1
      Top             =   360
      Width           =   5190
   End
   Begin VB.TextBox edcRegion 
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
      HelpContextID   =   8
      Left            =   1380
      MaxLength       =   80
      TabIndex        =   7
      Top             =   1260
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox edcName 
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
      HelpContextID   =   8
      Left            =   240
      MaxLength       =   80
      TabIndex        =   5
      Top             =   915
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   225
      ScaleHeight     =   210
      ScaleWidth      =   1110
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3450
      TabIndex        =   11
      Top             =   1800
      Width           =   945
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   4890
      Top             =   1710
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2145
      TabIndex        =   10
      Top             =   1800
      Width           =   945
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5000
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
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
      ScaleWidth      =   90
      TabIndex        =   8
      Top             =   2400
      Width           =   90
   End
   Begin VB.PictureBox pbcSTab 
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
      Left            =   30
      ScaleHeight     =   105
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   285
      Width           =   60
   End
   Begin VB.PictureBox pbcRegion 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   195
      Picture         =   "Copyregn.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   5115
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   5115
   End
   Begin VB.PictureBox plcRegion 
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
      Height          =   825
      Left            =   150
      ScaleHeight     =   765
      ScaleWidth      =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   5220
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   855
      TabIndex        =   9
      Top             =   1800
      Width           =   945
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      ScaleHeight     =   240
      ScaleWidth      =   5295
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   5295
   End
End
Attribute VB_Name = "CopyRegn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyregn.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopyRegn.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Region Area Copy input screen code
Option Explicit
Option Compare Text
''Program library dates Field Areas
'Dim tmCtrls(1 To 5)  As FIELDAREA
'Dim imBoxNo As Integer   'Current event name Box
'Dim imState As Integer
'Dim imFirstActivate As Integer
''Region Area
'Dim tmRaf As RAF            'RAF record image
'Dim tmRafSrchKey As LONGKEY0  'RAF key record image
'Dim tmRafSrchKey2 As LONGKEY0  'RAF key record image
'Dim hmRaf As Integer        'RAF Handle
'Dim imRafRecLen As Integer      'RAF record length
'Dim imUpdateAllowed As Integer
'Dim imTerminate As Integer  'True = terminating task, False= OK
'Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
'Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
'Dim imBSMode As Integer     'Backspace flag
'Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
'Dim imSelectedIndex As Integer
'Dim imPopReqd As Integer
'Dim smRegionCodeTag As String
'Dim tmRegionCode() As SORTCODE
'
'Const NAMEINDEX = 1
'Const STATEINDEX = 2
'Const REGIONINDEX = 3
'Const ENTEREDDATEINDEX = 4
'Const DORMANTDATEINDEX = 5
'Private Sub cbcSelect_Change()
'    Dim ilLoop As Integer   'For loop control parameter
'    Dim ilRet As Integer    'Return status
'    Dim slStr As String     'Text entered
'    Dim ilIndex As Integer  'Current index selected from combo box
'    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
'        Exit Sub
'    End If
'    imChgMode = True    'Set change mode to avoid infinite loop
'    imBypassSetting = True
'    Screen.MousePointer = vbHourglass  'Wait
'    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
'    If ilRet = 0 Then
'        ilIndex = cbcSelect.ListIndex
'        If Not mReadRec(ilIndex, SETFORREADONLY) Then
'            GoTo cbcSelectErr
'        End If
'    Else
'        If ilRet = 1 Then
'            If cbcSelect.ListCount > 0 Then
'                cbcSelect.ListIndex = 0
'            Else
'                cbcSelect.ListIndex = -1
'            End If
'        End If
'        ilRet = 1   'Clear fields as no match name found
'    End If
'    pbcRegion.Cls
'    If ilRet = 0 Then
'        imSelectedIndex = cbcSelect.ListIndex
'        mMoveRecToCtrl
'    Else
'        imSelectedIndex = 0
'        mClearCtrlFields
'        If slStr <> "[New]" Then
'            edcName.Text = slStr
'        End If
'    End If
'    For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
'        If (ilLoop = ENTEREDDATEINDEX) Then
'            gUnpackDate tmRaf.iDateEntrd(0), tmRaf.iDateEntrd(1), slStr
'            gSetShow pbcRegion, slStr, tmCtrls(ilLoop)
'        ElseIf (ilLoop = DORMANTDATEINDEX) Then
'            gUnpackDate tmRaf.iDateDormant(0), tmRaf.iDateDormant(1), slStr
'            gSetShow pbcRegion, slStr, tmCtrls(ilLoop)
'        Else
'            mSetShow ilLoop  'Set show strings
'        End If
'    Next ilLoop
'    pbcRegion_Paint
'    Screen.MousePointer = vbDefault
'    imChgMode = False
'    mSetCommands
'    imBypassSetting = False
'    Exit Sub
'cbcSelectErr:
'    On Error GoTo 0
'    Screen.MousePointer = vbDefault
'    imTerminate = True
'    Exit Sub
'End Sub
'Private Sub cbcSelect_Click()
'    cbcSelect_Change    'Process change as change event is not generated by VB
'End Sub
'Private Sub cbcSelect_GotFocus()
'    Dim slSvText As String   'Save so list box can be reset
'    If imTerminate Then
'        Exit Sub
'    End If
'
'    mSetShow imBoxNo
'    imBoxNo = -1
'    slSvText = cbcSelect.Text
''    ilSvIndex = cbcSelect.ListIndex
'    mPopulate igAdfCode
'    If imTerminate Then
'        Exit Sub
'    End If
'    If cbcSelect.ListCount <= 1 Then
'        cbcSelect.ListIndex = 0
'        mClearCtrlFields
'        If pbcSTab.Enabled Then
'            pbcSTab.SetFocus
'        Else
'            cmcCancel.SetFocus
'        End If
'        Exit Sub
'    End If
'    gCtrlGotFocus ActiveControl
'    If (slSvText = "") Or (slSvText = "[New]") Then
'        cbcSelect.ListIndex = 0
'        cbcSelect_Change    'Call change so picture area repainted
'    Else
'        gFindMatch slSvText, 1, cbcSelect
'        If gLastFound(cbcSelect) > 0 Then
''            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
'            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
'                cbcSelect.ListIndex = gLastFound(cbcSelect)
'                cbcSelect_Change    'Call change so picture area repainted
'                imPopReqd = False
'            End If
'        Else
'            cbcSelect.ListIndex = 0
'            mClearCtrlFields
'            cbcSelect_Change    'Call change so picture area repainted
'        End If
'    End If
'End Sub
'Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
'    'Delete key causes the charact to the right of the cursor to be deleted
'    imBSMode = False
'End Sub
'Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
'    'Backspace character cause selected test to be deleted or
'    'the first character to the left of the cursor if no text selected
'    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
'        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
'            imBSMode = True 'Force deletion of character prior to selected text
'        End If
'    End If
'End Sub
'Private Sub cmcCancel_Click()
'    mTerminate
'End Sub
'Private Sub cmcDone_Click()
'    If imUpdateAllowed Then
'        If mSaveRecChg(True) = False Then
'            If imTerminate Then
'                cmcCancel_Click
'                Exit Sub
'            End If
'            mEnableBox imBoxNo
'            Exit Sub
'        End If
'    End If
'    mTerminate
'End Sub
'Private Sub cmcDone_GotFocus()
'    Dim ilLoop As Integer
'    mSetShow imBoxNo
'    imBoxNo = -1
'    If Not cmcSave.Enabled Then
'        'Cycle to first unanswered mandatory
'        For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) - 2 Step 1
'            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
'                Beep
'                imBoxNo = ilLoop
'                mEnableBox imBoxNo
'                Exit Sub
'            End If
'        Next ilLoop
'    End If
'    gCtrlGotFocus cmcDone
'End Sub
'Private Sub cmcSave_Click()
'    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
'    Dim llCode As Long
'    Dim ilLoop As Integer
'    Dim slNameCode As String
'    Dim slCode As String
'    Dim ilRet As Integer
'    If Not imUpdateAllowed Then
'        Exit Sub
'    End If
'    slName = Trim$(edcName.Text)   'Save name
'    If mSaveRecChg(False) = False Then
'        If imTerminate Then
'            cmcCancel_Click
'            Exit Sub
'        End If
'        mEnableBox imBoxNo
'        Exit Sub
'    End If
'    imBoxNo = -1
'    llCode = tmRaf.lCode
'    cbcSelect.Clear
'    smRegionCodeTag = ""
'    mPopulate igAdfCode
'    For ilLoop = 0 To UBound(tmRegionCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
'        slNameCode = tmRegionCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
'        ilRet = gParseItem(slNameCode, 2, "\", slCode)
'        If Val(slCode) = llCode Then
'            If cbcSelect.ListIndex = ilLoop + 1 Then
'                cbcSelect_Change
'            Else
'                cbcSelect.ListIndex = ilLoop + 1
'            End If
'            Exit For
'        End If
'    Next ilLoop
'    DoEvents
'    mSetCommands
'    If cbcSelect.Enabled Then
'        cbcSelect.SetFocus
'    Else
'        pbcClickFocus.SetFocus
'    End If
'End Sub
'Private Sub cmcSave_GotFocus()
'    gCtrlGotFocus cmcSave
'    mSetShow imBoxNo
'    imBoxNo = -1
'End Sub
'Private Sub edcName_Change()
'    mSetChg NAMEINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function
'End Sub
'Private Sub edcName_GotFocus()
'    gCtrlGotFocus edcName
'End Sub
'Private Sub edcName_KeyPress(KeyAscii As Integer)
'    Dim ilKey As Integer
'    ilKey = KeyAscii
'    If Not gCheckKeyAscii(ilKey) Then
'        KeyAscii = 0
'        Exit Sub
'    End If
'End Sub
'Private Sub edcRegion_Change()
'    mSetChg REGIONINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function
'End Sub
'Private Sub edcRegion_GotFocus()
'    gCtrlGotFocus edcRegion
'End Sub
'Private Sub edcRegion_KeyPress(KeyAscii As Integer)
'    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
'        Beep
'        KeyAscii = 0
'        Exit Sub
'    End If
'End Sub
'Private Sub Form_Activate()
'    If Not imFirstActivate Then
'        DoEvents    'Process events so pending keys are not sent to this
'                    'form when keypreview turn on
'        gShowBranner imUpdateAllowed
'        Me.KeyPreview = True
'        Exit Sub
'    End If
'    imFirstActivate = False
'    If (tgUrf(0).sRegionCopy = "V") And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
'        imUpdateAllowed = False
'        pbcRegion.Enabled = False
'        pbcSTab.Enabled = False
'        pbcTab.Enabled = False
'    Else
'        pbcRegion.Enabled = True
'        pbcSTab.Enabled = True
'        pbcTab.Enabled = True
'        imUpdateAllowed = True
'    End If
'    gShowBranner imUpdateAllowed
'    Me.KeyPreview = True
'    CopyRegn.Refresh
'End Sub
'Private Sub Form_Click()
'    pbcClickFocus.SetFocus
'End Sub
'
'Private Sub Form_Deactivate()
'    Me.KeyPreview = False
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim ilReSet As Integer
'
'    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
'        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
'            cbcSelect.Enabled = False
'            ilReSet = True
'        Else
'            ilReSet = False
'        End If
'        gFunctionKeyBranch KeyCode
'        If imBoxNo > 0 Then
'            mEnableBox imBoxNo
'        End If
'        If ilReSet Then
'            cbcSelect.Enabled = True
'        End If
'    End If
'
'End Sub
'
Private Sub Form_Load()
'    mInit
'    If imTerminate Then
'        cmcCancel_Click
'    End If
End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mClearCtrlFields                *
''*                                                     *
''*             Created:5/13/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Clear each control on the      *
''*                      screen                         *
''*                                                     *
''*******************************************************
'Private Sub mClearCtrlFields()
''
''   mClearCtrlFields
''   Where:
''
'    Dim ilLoop As Integer
'    Dim slStr As String
'
'    edcName.Text = ""
'    edcRegion.Text = ""
'    imState = -1
'    mMoveCtrlToRec False
'    slStr = ""
'    gPackDate slStr, tmRaf.iDateEntrd(0), tmRaf.iDateEntrd(1)
'    slStr = ""
'    gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
'    For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
'        tmCtrls(ilLoop).iChg = False
'    Next ilLoop
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mEnableBox                      *
''*                                                     *
''*             Created:5/13/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Enable specified control       *
''*                                                     *
''*******************************************************
'Private Sub mEnableBox(ilBoxNo As Integer)
''
''   mInitParameters ilBoxNo
''   Where:
''       ilBoxNo (I)- Number of the Control to be enabled
''
'    If ilBoxNo < LBound(tmCtrls) Or ilBoxNo > UBound(tmCtrls) Then
'        Exit Sub
'    End If
'
'
'    Select Case ilBoxNo
'        Case NAMEINDEX 'Name
'            edcName.Width = tmCtrls(ilBoxNo).fBoxW
'            edcName.MaxLength = 80
'            gMoveFormCtrl pbcRegion, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
'            edcName.Visible = True  'Set visibility
'            edcName.SetFocus
'        Case STATEINDEX 'Selling or Airing
'            If imState < 0 Then
'                imState = 0
'                tmCtrls(ilBoxNo).iChg = True
'                mSetCommands
'            End If
'            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
'            gMoveFormCtrl pbcRegion, pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
'            pbcState_Paint
'            pbcState.Visible = True
'            pbcState.SetFocus
'        Case REGIONINDEX 'Region
'            edcRegion.Width = tmCtrls(ilBoxNo).fBoxW
'            edcRegion.MaxLength = 9
'            gMoveFormCtrl pbcRegion, edcRegion, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
'            edcRegion.Visible = True  'Set visibility
'            edcRegion.SetFocus
'    End Select
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mInit                           *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Initialize modular             *
''*                                                     *
''*******************************************************
'Private Sub mInit()
''
''   mInit
''   Where:
''
'    Dim ilRet As Integer
'    imTerminate = False
'
'    Screen.MousePointer = vbHourglass
'    imFirstActivate = True
'    CopyRegn.Height = cmcDone.Top + 5 * cmcDone.Height / 3
'    mInitBox
'    gCenterStdAlone CopyRegn
'    imChgMode = False
'    imBSMode = False
'    imBypassSetting = False
'    imBoxNo = -1
'    imTabDirection = 0  'Left to right movement
'    'CopyRegn.Show
'    Screen.MousePointer = vbHourglass
'    If Not gRecLengthOk("Raf.btr", Len(tmRaf)) Then
'        imTerminate = True
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    hmRaf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", CopyRegn
'    On Error GoTo 0
'    imRafRecLen = Len(tmRaf)
'    mPopulate igAdfCode
'    If imTerminate Then
'        Exit Sub
'    End If
''    mInitDDE
'    'imcHelp.Picture = Traffic!imcHelp.Picture
'    'plcScreen.Caption = "Region Area: " & sgAdvtName
''    gCenterModalForm CopyRegn
'    Screen.MousePointer = vbDefault
'    Exit Sub
'mInitErr:
'    On Error GoTo 0
'    imTerminate = True
'    Exit Sub
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mInitBox                        *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Set mouse and control locations*
''*                                                     *
''*******************************************************
'Private Sub mInitBox()
''
''   mInitBox
''   Where:
''
'    Dim flTextHeight As Single  'Standard text height
'    flTextHeight = pbcRegion.TextHeight("1") - 35
'    'Position panel and picture areas with panel
'    plcRegion.Move 150, 720, pbcRegion.Width + fgPanelAdj
'    pbcRegion.Move plcRegion.Left + fgBevelX, plcRegion.Top + fgBevelY
'    'Name
'    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 5070, fgBoxStH
'    'State
'    gSetCtrl tmCtrls(STATEINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1680, fgBoxStH
'    'Region
'    gSetCtrl tmCtrls(REGIONINDEX), 1170, tmCtrls(STATEINDEX).fBoxY, 1260, fgBoxStH
'    'Entered Date
'    gSetCtrl tmCtrls(ENTEREDDATEINDEX), 2445, tmCtrls(STATEINDEX).fBoxY, 1320, fgBoxStH
'    'Dormant Date
'    gSetCtrl tmCtrls(DORMANTDATEINDEX), 3780, tmCtrls(STATEINDEX).fBoxY, 1320, fgBoxStH
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mMoveCtrlToRec                  *
''*                                                     *
''*             Created:5/01/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Move control values to record  *
''*                                                     *
''*******************************************************
'Private Sub mMoveCtrlToRec(ilTestChg As Integer)
''
''   mMoveCtrlToRec iTest
''   Where:
''       iTest (I)- True = only move if field changed
''                  False = move regardless of change state
''
'    Dim slStr As String
'    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
'        tmRaf.sName = edcName.Text
'    End If
'    tmRaf.sAbbr = ""
'    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
'        Select Case imState
'            Case 0  'Active
'                tmRaf.sState = "A"
'            Case 1  'Dormant
'                If tmRaf.sState <> "D" Then
'                    slStr = Format$(gNow(), "m/d/yy")
'                    gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
'                End If
'                tmRaf.sState = "D"
'            Case Else
'                tmRaf.sState = ""
'        End Select
'    End If
'    If Not ilTestChg Or tmCtrls(REGIONINDEX).iChg Then
'        tmRaf.lRegionCode = Val(edcRegion.Text)
'    End If
'    tmRaf.sType = "R"
'    tmRaf.sInclExcl = "I"
'    tmRaf.sCategory = ""
'    tmRaf.sShowNoProposal = "N"
'    tmRaf.sShowOnOrder = "N"
'    tmRaf.sShowOnInvoice = "N"
'    tmRaf.sCustom = "N"
'    tmRaf.sUnused = ""
'    'tmRaf.sUnused = ""
'
'    Exit Sub
'
'    On Error GoTo 0
'    imTerminate = True
'    Exit Sub
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mMoveRecToCtrl                  *
''*                                                     *
''*             Created:5/13/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Move record values to controls *
''*                      on the screen                  *
''*                                                     *
''*******************************************************
'Private Sub mMoveRecToCtrl()
''
''   mMoveRecToCtrl
''   Where:
''
'    Dim ilLoop As Integer
'    edcName.Text = Trim$(tmRaf.sName)
'    Select Case tmRaf.sState
'        Case "A"
'            imState = 0
'        Case "D"
'            imState = 1
'        Case Else
'            imState = -1
'    End Select
'    If tmRaf.lRegionCode > 0 Then
'        edcRegion.Text = Trim$(str$(tmRaf.lRegionCode))
'    Else
'        edcRegion.Text = ""
'    End If
'    For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
'        tmCtrls(ilLoop).iChg = False
'    Next ilLoop
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mOKName                         *
''*                                                     *
''*             Created:6/1/93        By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Test that name is unique        *
''*                                                     *
''*******************************************************
'Private Function mOKName()
'    Dim slStr As String
'    Dim ilRet As Integer
'    Dim tlRaf As RAF
'    If edcName.Text <> "" Then    'Test name
'        slStr = Trim$(edcName.Text)
'        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
'        If gLastFound(cbcSelect) <> -1 Then   'Name found
'            If gLastFound(cbcSelect) <> imSelectedIndex Then
'                If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
'                    Beep
'                    MsgBox "Advertiser Region name already defined, enter a different name", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
'                    edcName.Text = Trim$(tmRaf.sName) 'Reset text
'                    mSetShow imBoxNo
'                    mSetChg imBoxNo
'                    imBoxNo = 1
'                    mEnableBox imBoxNo
'                    mOKName = False
'                    Exit Function
'                End If
'            End If
'        End If
'        If imSelectedIndex = 0 Then
'            tmRaf.lCode = 0
'        End If
'        slStr = Trim$(edcRegion.Text)
'        tmRafSrchKey2.lCode = CLng(slStr)
'        ilRet = btrGetEqual(hmRaf, tlRaf, imRafRecLen, tmRafSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
'        If (ilRet = BTRV_ERR_NONE) And (tmRaf.lCode <> tlRaf.lCode) Then
'            Beep
'            MsgBox "Region code already defined, enter a different code", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
'            edcRegion.Text = Trim$(str$(tmRaf.lRegionCode)) 'Reset text
'            mSetShow imBoxNo
'            mSetChg imBoxNo
'            imBoxNo = REGIONINDEX
'            mEnableBox imBoxNo
'            mOKName = False
'            Exit Function
'        End If
'    End If
'    mOKName = True
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mPopulate                       *
''*                                                     *
''*             Created:7/19/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Populate advertiser regions    *
''*                      list box if required           *
''*                                                     *
''*******************************************************
'Private Sub mPopulate(ilAdfCode As Integer)
''
''   mPopulate
''   Where:
''
'    Dim ilRet As Integer
'    'Repopulate if required- if sales source changed by another user while in this screen
'    imPopReqd = False
'    ilRet = gPopRegionBox(CopyRegn, ilAdfCode, "R", True, cbcSelect, tmRegionCode(), smRegionCodeTag)
'    If ilRet <> CP_MSG_NOPOPREQ Then
'        On Error GoTo mPopulateErr
'        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", CopyRegn
'        On Error GoTo 0
'        cbcSelect.AddItem "[New]", 0  'Force as first item on list
'        imPopReqd = True
'    End If
'    Exit Sub
'mPopulateErr:
'    On Error GoTo 0
'    imTerminate = True
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mReadRec                        *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Read a record                  *
''*                                                     *
''*******************************************************
'Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer) As Integer
''
''   iRet = mReadRec(ilSelectIndex)
''   Where:
''       ilSelectIndex (I) - list box index
''       iRet (O)- True if record read,
''                 False if not read
''
'    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
'    Dim slCode As String    'Code number- so record can be found
'    Dim ilRet As Integer    'Return status
'    slNameCode = tmRegionCode(ilSelectIndex - 1).sKey    'lbcCopyRegnCode.List(ilSelectIndex - 1)
'    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'    On Error GoTo mReadRecErr
'    gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", CopyRegn
'    On Error GoTo 0
'    tmRafSrchKey.lCode = CLng(slCode)
'    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
'    On Error GoTo mReadRecErr
'    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Advertiser Region Area)", CopyRegn
'    On Error GoTo 0
'    mReadRec = True
'    Exit Function
'mReadRecErr:
'    On Error GoTo 0
'    mReadRec = False
'    Exit Function
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mSaveRec                        *
''*                                                     *
''*             Created:5/14/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Update or added record          *
''*                                                     *
''*******************************************************
'Private Function mSaveRec() As Integer
''
''   iRet = mSaveRec()
''   Where:
''       iRet (O)- True if updated or added, False if not updated or added
''
'    Dim ilRet As Integer
'    Dim slMsg As String
'    Dim slStr As String
'    mSetShow imBoxNo
'    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
'        mSaveRec = False
'        Exit Function
'    End If
'    If Not mOKName() Then
'        mSaveRec = False
'        Exit Function
'    End If
'    Screen.MousePointer = vbHourglass  'Wait
'    Do  'Loop until record updated or added
'        If imSelectedIndex <> 0 Then
'            'Reread record in so lastest is obtained
'            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
'                Screen.MousePointer = vbDefault
'                ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
'                imTerminate = True
'                mSaveRec = False
'                Exit Function
'            End If
'        End If
'        mMoveCtrlToRec True
'        If imSelectedIndex = 0 Then 'New selected
'            tmRaf.lCode = 0
'            tmRaf.iadfCode = igAdfCode
'            tmRaf.sAssigned = "N"
'            slStr = Format$(gNow(), "m/d/yy")
'            gPackDate slStr, tmRaf.iDateEntrd(0), tmRaf.iDateEntrd(1)
'            slStr = ""
'            gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
'            tmRaf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
'            ilRet = btrInsert(hmRaf, tmRaf, imRafRecLen, INDEXKEY0)
'            slMsg = "mSaveRec (btrInsert: Advertiser Region Area)"
'        Else 'Old record-Update
'            tmRaf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
'            ilRet = btrUpdate(hmRaf, tmRaf, imRafRecLen)
'            slMsg = "mSaveRec (btrUpdate: Advertiser Product Name)"
'        End If
'    Loop While ilRet = BTRV_ERR_CONFLICT
'    On Error GoTo mSaveRecErr
'    gBtrvErrorMsg ilRet, slMsg, CopyRegn
'    On Error GoTo 0
'    mSaveRec = True
'    Screen.MousePointer = vbDefault
'    Exit Function
'mSaveRecErr:
'    On Error GoTo 0
'    Screen.MousePointer = vbDefault
'    imTerminate = True
'    mSaveRec = False
'    Exit Function
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mSaveRecChg                      *
''*                                                     *
''*             Created:5/14/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Determine if record altered and*
''*                      requires updating              *
''*                                                     *
''*******************************************************
'Private Function mSaveRecChg(ilAsk As Integer) As Integer
''
''   iAsk = True
''   iRet = mSaveRecChg(iAsk)
''   Where:
''       iAsk (I)- True = Ask if changed records should be updated;
''                 False= Update record if required without asking user
''       iRet (O)- True if updated or added, False if not updated or added
''
'    Dim ilRes As Integer
'    Dim slMess As String
'    Dim ilAltered As Integer
'    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
'    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
'        If ilAltered = YES Then
'            If ilAsk Then
'                If imSelectedIndex > 0 Then
'                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
'                Else
'                    slMess = "Add " & edcName.Text
'                End If
'                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
'                If ilRes = vbCancel Then
'                    mSaveRecChg = False
'                    pbcRegion_Paint
'                    Exit Function
'                End If
'                If ilRes = vbYes Then
'                    ilRes = mSaveRec()
'                    mSaveRecChg = ilRes
'                    Exit Function
'                End If
'                If ilRes = vbNo Then
'                    cbcSelect.ListIndex = 0
'                End If
'            Else
'                ilRes = mSaveRec()
'                mSaveRecChg = ilRes
'                Exit Function
'            End If
'        End If
'    End If
'    mSaveRecChg = True
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mSetChg                         *
''*                                                     *
''*             Created:5/12/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Determine if value for a       *
''*                      control is different from the  *
''*                      record                         *
''*                                                     *
''*******************************************************
'Private Sub mSetChg(ilBoxNo As Integer)
''
''   mSetChg ilBoxNo
''   Where:
''       ilBoxNo (I)- Number of the Control whose value should be checked
''
'    Dim slStr As String
'    If ilBoxNo < LBound(tmCtrls) Or ilBoxNo > UBound(tmCtrls) Then
''        mSetCommands
'        Exit Sub
'    End If
'
'    Select Case ilBoxNo 'Branch on box type (control)
'        Case NAMEINDEX 'Name
'            gSetChgFlag tmRaf.sName, edcName, tmCtrls(ilBoxNo)
'        Case STATEINDEX
'        Case REGIONINDEX
'            slStr = Trim$(str$(tmRaf.lRegionCode))
'            gSetChgFlag slStr, edcRegion, tmCtrls(ilBoxNo)
'    End Select
'    mSetCommands
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mSetCommands                    *
''*                                                     *
''*             Created:6/30/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Set command buttons (enable or *
''*                      disabled)                      *
''*                                                     *
''*******************************************************
'Private Sub mSetCommands()
''
''   mSetCommands
''   Where:
''
'    Dim ilAltered As Integer
'    If imBypassSetting Then
'        Exit Sub
'    End If
'    If Not imUpdateAllowed Then
'        cmcSave.Enabled = False
'        cmcDone.Enabled = False
'        Exit Sub
'    End If
'    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
'    'Update button set if all mandatory fields have data and any field altered
'    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) Then
'        cmcSave.Enabled = True
'    Else
'        cmcSave.Enabled = False
'    End If
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mSetFocus                       *
''*                                                     *
''*             Created:6/28/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Set focus to specified control *
''*                                                     *
''*******************************************************
'Private Sub mSetFocus(ilBoxNo As Integer)
''
''   mSetFocus ilBoxNo
''   Where:
''       ilBoxNo (I)- Number of the Control to be enabled
''
'    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > UBound(tmCtrls)) Then
'        Exit Sub
'    End If
'
'    Select Case ilBoxNo 'Branch on box type (control)
'        Case NAMEINDEX 'Name
'            edcName.SetFocus
'        Case STATEINDEX 'State
'            pbcState.SetFocus
'        Case REGIONINDEX 'Region
'            edcRegion.SetFocus
'    End Select
'    mSetCommands
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mSetShow                        *
''*                                                     *
''*             Created:6/30/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Format user input for a control*
''*                      to be displayed on the form    *
''*                                                     *
''*******************************************************
'Private Sub mSetShow(ilBoxNo As Integer)
''
''   mSetShow ilBoxNo
''   Where:
''       ilBoxNo (I)- Number of the Control whose value should be saved
''
'    Dim slStr As String
'    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > UBound(tmCtrls)) Then
'        Exit Sub
'    End If
'
'    Select Case ilBoxNo 'Branch on box type (control)
'        Case NAMEINDEX 'Name
'            edcName.Visible = False  'Set visibility
'            slStr = edcName.Text
'            gSetShow pbcRegion, slStr, tmCtrls(ilBoxNo)
'        Case STATEINDEX 'State
'            pbcState.Visible = False  'Set visibility
'            If imState = 0 Then
'                slStr = "Active"
'            ElseIf imState = 1 Then
'                slStr = "Dormant"
'            Else
'                slStr = ""
'            End If
'            gSetShow pbcRegion, slStr, tmCtrls(ilBoxNo)
'        Case REGIONINDEX 'Region
'            edcRegion.Visible = False  'Set visibility
'            slStr = edcRegion.Text
'            gSetShow pbcRegion, slStr, tmCtrls(ilBoxNo)
'    End Select
'    mSetCommands
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mTerminate                      *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: terminate form                 *
''*                                                     *
''*******************************************************
'Private Sub mTerminate()
''
''   mTerminate
''   Where:
''
'    Screen.MousePointer = vbDefault
'    btrDestroy hmRaf
'    igManUnload = YES
'    Unload CopyRegn
'    Set CopyRegn = Nothing   'Remove data segment
'    igManUnload = NO
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mTestFields                     *
''*                                                     *
''*             Created:4/21/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Test fields for mandatory and   *
''*                     blanks                          *
''*                                                     *
''*******************************************************
'Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
''
''   iState = ALLBLANK+NOMSG   'Blank
''   iTest = TESTALLCTRLS
''   iRet = mTestFields(iTest, iState)
''   Where:
''       iTest (I)- Test all controls or control number specified
''       iState (I)- Test one of the following:
''                  ALLBLANK=All fields blank
''                  ALLMANBLANK=All mandatory
''                    field blank
''                  ALLMANDEFINED=All mandatory
''                    fields have data
''                  Plus
''                  NOMSG=No error message shown
''                  SHOWMSG=show error message
''       iRet (O)- True if all mandatory fields blank, False if not all blank
''
''
'    Dim slStr As String
'    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
'        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
'            If ilState = (ALLMANDEFINED + SHOWMSG) Then
'                imBoxNo = NAMEINDEX
'            End If
'            mTestFields = NO
'            Exit Function
'        End If
'    End If
'    If (ilCtrlNo = STATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
'        If imState = 0 Then
'            slStr = "Active"
'        ElseIf imState = 1 Then
'            slStr = "Dormant"
'        Else
'            slStr = ""
'        End If
'        If gFieldDefinedStr(slStr, "", "Active Or Dormant must be specified", tmCtrls(STATEINDEX).iReq, ilState) = NO Then
'            If ilState = (ALLMANDEFINED + SHOWMSG) Then
'                imBoxNo = STATEINDEX
'            End If
'            mTestFields = NO
'            Exit Function
'        End If
'    End If
'    If (ilCtrlNo = REGIONINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
'        If gFieldDefinedCtrl(edcRegion, "", "Region Code must be specified", tmCtrls(REGIONINDEX).iReq, ilState) = NO Then
'            If ilState = (ALLMANDEFINED + SHOWMSG) Then
'                imBoxNo = REGIONINDEX
'            End If
'            mTestFields = NO
'            Exit Function
'        End If
'    End If
'    mTestFields = YES
'End Function
'Private Sub pbcClickFocus_GotFocus()
'    mSetShow imBoxNo
'    imBoxNo = -1
'End Sub
'Private Sub pbcRegion_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim ilBox As Integer
'    For ilBox = LBound(tmCtrls) To UBound(tmCtrls) - 2 Step 1
'        If (x >= tmCtrls(ilBox).fBoxX) And (x <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
'            If (y >= tmCtrls(ilBox).fBoxY) And (y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
'                mSetShow imBoxNo
'                imBoxNo = ilBox
'                mEnableBox ilBox
'                Exit Sub
'            End If
'        End If
'    Next ilBox
'    mSetFocus imBoxNo
'End Sub
'Private Sub pbcRegion_Paint()
'    Dim ilBox As Integer
'    For ilBox = LBound(tmCtrls) To UBound(tmCtrls) Step 1
'        pbcRegion.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
'        pbcRegion.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
'        pbcRegion.Print tmCtrls(ilBox).sShow
'    Next ilBox
'End Sub
'Private Sub pbcSTab_GotFocus()
'    Dim ilBox As Integer
'    If GetFocus() <> pbcSTab.hwnd Then
'        Exit Sub
'    End If
'    imTabDirection = -1 'Set- Right to left
'    Select Case imBoxNo
'        Case -1 'Tab from control prior to form area
'            imTabDirection = 0  'Set-Left to right
'            ilBox = NAMEINDEX
'            imBoxNo = ilBox
'            mEnableBox ilBox
'            Exit Sub
'        Case NAMEINDEX 'Name (first control within header)
'            mSetShow imBoxNo
'            imBoxNo = -1
'            cmcCancel.SetFocus
'        Case Else
'            ilBox = imBoxNo - 1
'    End Select
'    mSetShow imBoxNo
'    imBoxNo = ilBox
'    mEnableBox ilBox
'End Sub
'Private Sub pbcState_GotFocus()
'    gCtrlGotFocus ActiveControl
'End Sub
'Private Sub pbcState_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
'        If imState <> 0 Then
'            tmCtrls(imBoxNo).iChg = True
'        End If
'        imState = 0
'        pbcState_Paint
'    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
'        If imState <> 1 Then
'            tmCtrls(imBoxNo).iChg = True
'        End If
'        imState = 1
'        pbcState_Paint
'    End If
'    If KeyAscii = Asc(" ") Then
'        If imState = 0 Then  'Active
'            imState = 1
'            tmCtrls(imBoxNo).iChg = True
'            pbcState_Paint
'        ElseIf imState = 1 Then  'Dormant
'            tmCtrls(imBoxNo).iChg = True
'            imState = 0  'Active
'            pbcState_Paint
'        End If
'    End If
'    mSetCommands
'End Sub
'Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If imState = 0 Then  'Active
'        tmCtrls(imBoxNo).iChg = True
'        imState = 1  'Dormant
'    ElseIf imState = 1 Then  'Dormant
'        tmCtrls(imBoxNo).iChg = True
'        imState = 0  'Active
'    End If
'    pbcState_Paint
'    mSetCommands
'End Sub
'Private Sub pbcState_Paint()
'    pbcState.Cls
'    pbcState.CurrentX = fgBoxInsetX
'    pbcState.CurrentY = 0 'fgBoxInsetY
'    Select Case imState
'        Case 0  'Active
'            pbcState.Print "Active"
'        Case 1  'Dormant
'            pbcState.Print "Dormant"
'        Case Else
'            pbcState.Print "       "
'    End Select
'End Sub
'Private Sub pbcTab_GotFocus()
'    Dim ilBox As Integer
'
'    If GetFocus() <> pbcTab.hwnd Then
'        Exit Sub
'    End If
'    imTabDirection = 0 'Set- Left to right
'    Select Case imBoxNo
'        Case -1 'Tab from control prior to form area
'            imTabDirection = -1  'Set-Right to left
'            ilBox = REGIONINDEX
'        Case REGIONINDEX 'Last control within header
'            mSetShow imBoxNo
'            imBoxNo = -1
'            cmcDone.SetFocus
'            Exit Sub
'        Case Else
'            ilBox = imBoxNo + 1
'    End Select
'    mSetShow imBoxNo
'    imBoxNo = ilBox
'    mEnableBox ilBox
'End Sub
'Private Sub plcScreen_Click()
'    pbcClickFocus.SetFocus
'End Sub
'Private Sub tmcClick_Timer()
'    If cbcSelect.ListIndex <> imSelectedIndex Then
'        cbcSelect_Change
'        'cbcSelect.SetFocus
'        Exit Sub
'    End If
'End Sub
'Private Sub plcScreen_Paint()
'    plcScreen.CurrentX = 0
'    plcScreen.CurrentY = 0
'    plcScreen.Print "Region Area: " & sgAdvtName
'End Sub
