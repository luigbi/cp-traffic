VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form InvType 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3675
   ClientLeft      =   1905
   ClientTop       =   2685
   ClientWidth     =   5220
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   5220
   Begin VB.TextBox edcGL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   165
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1770
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   735
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      Top             =   225
      Width           =   3420
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   45
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2730
      Width           =   105
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   330
      MaxLength       =   20
      TabIndex        =   4
      Top             =   675
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   3285
      Width           =   1050
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      Left            =   2115
      TabIndex        =   12
      Top             =   3285
      Width           =   1050
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
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
      Left            =   990
      TabIndex        =   11
      Top             =   3285
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   2940
      Width           =   1050
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
      Left            =   2115
      TabIndex        =   9
      Top             =   2940
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
      Left            =   990
      TabIndex        =   8
      Top             =   2940
      Width           =   1050
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   120
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   375
      Width           =   15
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   1695
      Width           =   45
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2865
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdType 
      Height          =   1440
      Left            =   270
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   765
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   2540
      _Version        =   393216
      Rows            =   14
      Cols            =   4
      FixedRows       =   2
      FixedCols       =   2
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label lbcNote 
      Caption         =   "G/L replaces G/L Receivable in Accounting Export and Primary Code Gross Sales in Great Plains Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   180
      TabIndex        =   18
      Top             =   2385
      Width           =   4665
   End
   Begin VB.Label plcScreen 
      Caption         =   "Inventory Types"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   1485
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3210
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "InvType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of InvType.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: InvType.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Event Type input screen code
Option Explicit
Option Compare Text
'Inventory Type Field Areas
Dim hmItf As Integer 'Inventory Type file handle
Dim tmItf As ITF        'ITF record image
Dim tmItfSrchKey As INTKEY0    'ITF key record image
Dim imItfRecLen As Integer        'ITF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imUpdateAllowed As Integer    'User can update records
Dim imBypassSetting As Integer
Dim imPopReqd As Integer
Dim smInitType As String
Dim smYN As String


Dim imItfChg As Integer


Dim tmInvType() As SORTCODE
Dim smInvTypeTag As String

Dim imCtrlVisible As Integer
Dim lmEnableRow As Long
Dim lmEnableCol As Long

Const ROW3INDEX = 3
Const NAMEINDEX = 2  'Name control/field
Const ROW6INDEX = 6
Const MULTIFEEDINDEX = 2
Const ROW9INDEX = 9
Const GLCASHINDEX = 2
Const ROW12INDEX = 12
Const GLTRADEINDEX = 2


Private Sub cbcSelect_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    imItfChg = False
    mClearCtrlFields
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX) = slStr
            imItfChg = True
        End If
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    Dim ilLoop As Integer
    If imTerminate Then
        Exit Sub
    End If
    mSetShow
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        For ilLoop = 1 To igDDEDelay Step 1
            DoEvents
        Next ilLoop
        If igInvTypeCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If smInitType = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = smInitType    'New name
            End If
            cbcSelect_Change
            If smInitType <> "" Then
                mSetCommands
                gFindMatch smInitType, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        pbcSTab.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCancel_Click()
    If igInvTypeCallSource <> CALLNONE Then
        igInvTypeCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
    mSetShow
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igInvTypeCallSource <> CALLNONE Then
        sgInvTypeName = grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX)
        If mSaveRecChg(False) = False Then
            sgInvTypeName = "[New]"
            If Not imTerminate Then
                mEnableBox
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox
            Exit Sub
        End If
    End If
    If igInvTypeCallSource <> CALLNONE Then
        If sgInvTypeName = "[New]" Then
            igInvTypeCallSource = CALLCANCELLED
        Else
            igInvTypeCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    mSetShow
    If Trim$(grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX)) = "" Then
        Exit Sub
    End If
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        If Not mTestFields(True) Then
            Beep
            mEnableBox
            Exit Sub
        End If
    End If
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcErase_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStamp                                                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slMsg As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(InvType, tmItf.iCode, "Iif.Btr", "IifItfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Inventory Item references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(InvType, tmItf.iCode, "Ihf.Btr", "IhfItfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Event Inventory references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmItf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = btrDelete(hmItf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", InvType
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        smInvTypeTag = ""
        mPopulate
    End If
    'Remove focus from control and make invisible
    mSetShow
    mClearCtrlFields
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    gCtrlGotFocus cmcErase
    mSetShow
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = EVENTTYPESLIST
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "InvType^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "InvType^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "InvType^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "InvType^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'InvType.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'InvType.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    ''Screen.MousePointer = vbDefault    'Default
'    RptNoSel.Show vbModal
    sgCommandStr = slStr
    RptList.Show vbModal
End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
    mSetShow
End Sub
Private Sub cmcUndo_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        mClearCtrlFields
        mMoveRecToCtrl
        mSetCommands
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    mClearCtrlFields
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus cmcUndo
    mSetShow
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = Trim$(grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX)) 'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox
        Exit Sub
    End If
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If imSvSelectedIndex <> 0 Then
        cbcSelect.Text = slName
    Else
        cbcSelect.ListIndex = 0
    End If
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcName_Change()
    grdType.CellForeColor = vbBlack
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus edcName
End Sub
Private Sub edcName_KeyPress(KeyAscii As Integer)
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
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        grdType.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        grdType.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    InvType.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If cbcSelect.Enabled Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        mEnableBox
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    If Not igManUnload Then
        mSetShow
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Erase tmInvType

    btrExtClear hmItf   'Clear any previous extend operation
    ilRet = btrClose(hmItf)
    btrDestroy hmItf
        
    Set InvType = Nothing   'Remove data segment

End Sub

Private Sub grdType_EnterCell()
    mSetShow
End Sub

Private Sub grdType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine row and col mouse up onto
    ilCol = grdType.MouseCol
    ilRow = grdType.MouseRow
    If ilCol < grdType.FixedCols Then
        grdType.Redraw = True
        Exit Sub
    End If
    If ilRow < grdType.FixedRows Then
        grdType.Redraw = True
        Exit Sub
    End If
    If (ilRow = ROW3INDEX - 1) Or (ilRow = ROW6INDEX - 1) Or (ilRow = ROW9INDEX - 1) Or (ilRow = ROW12INDEX - 1) Then
        ilRow = ilRow + 1
    End If
    If grdType.ColWidth(ilCol) <= 15 Then
        grdType.Redraw = True
        Exit Sub
    End If
    If grdType.RowHeight(ilRow) <= 15 Then
        grdType.Redraw = True
        Exit Sub
    End If
    DoEvents
    grdType.Col = ilCol
    grdType.Row = ilRow
    grdType.Redraw = True
    mEnableBox
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mClearCtrlFields
'   Where:
'
    edcName.Text = ""
    grdType.Row = grdType.FixedRows + 1
    grdType.Col = NAMEINDEX
    grdType.CellForeColor = vbBlack
    grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX) = ""
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (grdType.Row < grdType.FixedRows) Or (grdType.Row >= grdType.Rows) Or (grdType.Col < grdType.FixedCols) Or (grdType.Col >= grdType.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdType.Row
    lmEnableCol = grdType.Col

    Select Case grdType.Row
        Case ROW3INDEX
            Select Case grdType.Col
                Case NAMEINDEX 'Name
                    edcName.MaxLength = 50
                    edcName.Text = grdType.Text
            End Select
        Case ROW6INDEX
            Select Case grdType.Col
                Case MULTIFEEDINDEX
                    smYN = Trim$(grdType.Text)
                    If (smYN = "") Or (smYN = "Missing") Then
                        smYN = "Yes"
                    End If
                    pbcYN_Paint
            End Select
        Case ROW9INDEX
            Select Case grdType.Col
                Case GLCASHINDEX 'Name
                    edcGL.MaxLength = 20
                    edcGL.Text = grdType.Text
            End Select
        Case ROW12INDEX
            Select Case grdType.Col
                Case GLTRADEINDEX 'Name
                    edcGL.MaxLength = 20
                    edcGL.Text = grdType.Text
            End Select
    End Select
    mSetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
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
'*  ilLoop                        llNoRecs                                                *
'******************************************************************************************

'
'   mInitParameters
'   Where:
'
    Dim ilRet As Integer    'Error return status
    imFirstActivate = True
    imTerminate = False
    imItfChg = False
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    InvType.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone InvType
    'InvType.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imItfRecLen = Len(tmItf)  'Get and save ARF record length
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmItf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmItf, "", sgDBPath & "Itf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", InvType
    On Error GoTo 0
'    gCenterModalForm InvType
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list box to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
        mSetCommands
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                      *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
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
    grdType.Move 270, 765, 4500, 1440
    mGridLayout
    mGridColumnWidths
    mGridColumns
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    tmItf.sName = grdType.TextMatrix(ROW3INDEX, NAMEINDEX)
    If Trim$(grdType.TextMatrix(ROW6INDEX, MULTIFEEDINDEX)) = "No" Then
        tmItf.sMultiFeed = "N"
    Else
        tmItf.sMultiFeed = "Y"
    End If
    tmItf.sGLCash = grdType.TextMatrix(ROW9INDEX, GLCASHINDEX)
    tmItf.sGLTrade = grdType.TextMatrix(ROW12INDEX, GLTRADEINDEX)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mMoveRecToCtrl
'   Where:
'
    grdType.Row = ROW3INDEX
    grdType.Col = NAMEINDEX
    grdType.CellForeColor = vbBlack
    grdType.TextMatrix(ROW3INDEX, NAMEINDEX) = Trim$(tmItf.sName)
    grdType.Row = ROW6INDEX
    grdType.Col = MULTIFEEDINDEX
    grdType.CellForeColor = vbBlack
    If tmItf.sMultiFeed = "N" Then
        grdType.TextMatrix(ROW6INDEX, MULTIFEEDINDEX) = "No"
    Else
        grdType.TextMatrix(ROW6INDEX, MULTIFEEDINDEX) = "Yes"
    End If
    grdType.Row = ROW9INDEX
    grdType.Col = GLCASHINDEX
    grdType.CellForeColor = vbBlack
    grdType.TextMatrix(ROW9INDEX, GLCASHINDEX) = Trim$(tmItf.sGLCash)
    grdType.Row = ROW12INDEX
    grdType.Col = GLTRADEINDEX
    grdType.CellForeColor = vbBlack
    grdType.TextMatrix(ROW12INDEX, GLTRADEINDEX) = Trim$(tmItf.sGLTrade)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    slStr = Trim$(grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX))
    If slStr <> "" Then    'Test name
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If UCase$(Trim$(grdType.TextMatrix(grdType.FixedRows + 1, NAMEINDEX))) = UCase$(Trim$(cbcSelect.List(gLastFound(cbcSelect)))) Then
                    Beep
                    MsgBox "Inventory Type already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    grdType.Row = grdType.FixedRows + 1
                    grdType.Col = NAMEINDEX
                    mSetShow
                    mEnableBox
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKName = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone InvType, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igInvTypeCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igInvTypeCallSource = CALLNONE
    'End If
    smInitType = ""
    If igInvTypeCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            smInitType = slStr
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    imPopReqd = False
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    ilRet = gIMoveListBox(InvType, cbcSelect, tmInvType(), smInvTypeTag, "Itf.Btr", gFieldOffset("Itf", "ItfName"), 50, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", InvType
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer) As Integer
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tmInvType(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 3)", InvType
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmItfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmItf, tmItf, imItfRecLen, tmItfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", InvType
    On Error GoTo 0
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStamp                                                                               *
'******************************************************************************************

'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    mSetShow
    If Not mTestFields(True) Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmItf.iCode = 0  'Autoincrement
            ilRet = btrInsert(hmItf, tmItf, imItfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmItf, tmItf, imItfRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, InvType
    On Error GoTo 0
    smInvTypeTag = ""
    mPopulate  'Repopulate since list box not in sorted order
    cbcSelect.ListIndex = 0
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                      *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    ilAltered = imItfChg
    If mTestFields(True) Then
        If ilAltered Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
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
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        End If
    End If
    mSaveRecChg = True
End Function

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
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = imItfChg
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(False)) And (ilAltered) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (Not mTestFields(False)) And (imUpdateAllowed) Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdType.Row < grdType.FixedRows) Or (grdType.Row >= grdType.Rows) Or (grdType.Col < grdType.FixedCols) Or (grdType.Col >= grdType.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdType.Col - 1 Step 1
        llColPos = llColPos + grdType.ColWidth(ilCol)
    Next ilCol
    Select Case grdType.Row
        Case ROW3INDEX
            Select Case grdType.Col
                Case NAMEINDEX
                    edcName.Move grdType.Left + llColPos + 30, grdType.Top + grdType.RowPos(grdType.Row) + 15, grdType.ColWidth(grdType.Col) - 30, grdType.RowHeight(grdType.Row) - 15
                    edcName.Visible = True
                    edcName.SetFocus
            End Select
        Case ROW6INDEX
            Select Case grdType.Col
                Case MULTIFEEDINDEX
                    pbcYN.Move grdType.Left + llColPos + 30, grdType.Top + grdType.RowPos(grdType.Row) + 15, grdType.ColWidth(grdType.Col) - 30, grdType.RowHeight(grdType.Row) - 15
                    pbcYN.Visible = True
                    pbcYN.SetFocus
            End Select
        Case ROW9INDEX
            Select Case grdType.Col
                Case GLCASHINDEX
                    edcGL.Move grdType.Left + llColPos + 30, grdType.Top + grdType.RowPos(grdType.Row) + 15, grdType.ColWidth(grdType.Col) - 30, grdType.RowHeight(grdType.Row) - 15
                    edcGL.Visible = True
                    edcGL.SetFocus
            End Select
        Case ROW12INDEX
            Select Case grdType.Col
                Case GLTRADEINDEX
                    edcGL.Move grdType.Left + llColPos + 30, grdType.Top + grdType.RowPos(grdType.Row) + 15, grdType.ColWidth(grdType.Col) - 30, grdType.RowHeight(grdType.Row) - 15
                    edcGL.Visible = True
                    edcGL.SetFocus
            End Select
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Form user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'

    If (lmEnableRow >= grdType.FixedRows) And (lmEnableRow < grdType.Rows) Then
        Select Case lmEnableRow
            Case ROW3INDEX
                Select Case lmEnableCol
                    Case NAMEINDEX
                        edcName.Visible = False
                        If grdType.TextMatrix(lmEnableRow, lmEnableCol) <> Trim$(edcName.Text) Then
                            imItfChg = True
                        End If
                        grdType.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcName.Text)
                End Select
            Case ROW6INDEX
                Select Case lmEnableCol
                    Case MULTIFEEDINDEX
                        pbcYN.Visible = False
                        If grdType.TextMatrix(lmEnableRow, lmEnableCol) <> smYN Then
                            imItfChg = True
                        End If
                        grdType.TextMatrix(lmEnableRow, lmEnableCol) = smYN
                End Select
            Case ROW9INDEX
                Select Case lmEnableCol
                    Case GLCASHINDEX
                        edcGL.Visible = False
                        If grdType.TextMatrix(lmEnableRow, lmEnableCol) <> Trim$(edcGL.Text) Then
                            imItfChg = True
                        End If
                        grdType.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcGL.Text)
                End Select
            Case ROW12INDEX
                Select Case lmEnableCol
                    Case GLTRADEINDEX
                        edcGL.Visible = False
                        If grdType.TextMatrix(lmEnableRow, lmEnableCol) <> Trim$(edcGL.Text) Then
                            imItfChg = True
                        End If
                        grdType.TextMatrix(lmEnableRow, lmEnableCol) = Trim$(edcGL.Text)
                End Select
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
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
    Dim ilRet As Integer

    smInvTypeTag = ""

    sgDoneMsg = Trim$(str$(igInvTypeCallSource)) & "\" & sgInvTypeName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload InvType
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilShowMsg As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
    slStr = grdType.TextMatrix(ROW3INDEX, NAMEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If ilShowMsg Then
            grdType.TextMatrix(ROW3INDEX, NAMEINDEX) = "Missing"
            grdType.Row = ROW3INDEX
            grdType.Col = NAMEINDEX
            grdType.CellForeColor = vbMagenta
        End If
    End If
    slStr = grdType.TextMatrix(ROW6INDEX, MULTIFEEDINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        If ilShowMsg Then
            grdType.TextMatrix(ROW6INDEX, MULTIFEEDINDEX) = "Missing"
            grdType.Row = ROW6INDEX
            grdType.Col = MULTIFEEDINDEX
            grdType.CellForeColor = vbMagenta
        End If
    End If
    If ilError Then
        mTestFields = False
    Else
        mTestFields = True
    End If
End Function
Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub pbcSTab_GotFocus()
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        Select Case grdType.Row
            Case ROW3INDEX
                mSetShow
                cmcDone.SetFocus
                Exit Sub
            Case ROW6INDEX
                grdType.Row = ROW3INDEX
                grdType.Col = NAMEINDEX
            Case ROW9INDEX
                grdType.Row = ROW6INDEX
                grdType.Col = MULTIFEEDINDEX
            Case ROW12INDEX
                grdType.Row = ROW9INDEX
                grdType.Col = GLCASHINDEX
        End Select
        mSetShow
    Else
        grdType.Row = ROW3INDEX
        grdType.Col = NAMEINDEX
    End If
    mEnableBox
End Sub
Private Sub pbcTab_GotFocus()
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        Select Case grdType.Row
            Case ROW3INDEX
                grdType.Row = ROW6INDEX
                grdType.Col = MULTIFEEDINDEX
            Case ROW6INDEX
                grdType.Row = ROW9INDEX
                grdType.Col = GLCASHINDEX
            Case ROW9INDEX
                grdType.Row = ROW12INDEX
                grdType.Col = GLTRADEINDEX
            Case ROW12INDEX
                mSetShow
                cmcDone.SetFocus
                Exit Sub
        End Select
        mSetShow
    Else
        grdType.Row = ROW12INDEX
        grdType.Col = GLTRADEINDEX
    End If
    mEnableBox
End Sub

Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        smYN = "Yes"
        pbcYN_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smYN = "No"
        pbcYN_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smYN = "Yes" Then
            smYN = "No"
            pbcYN_Paint
        ElseIf smYN = "No" Then
            smYN = "Yes"
            pbcYN_Paint
        End If
    End If
End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smYN = "Yes" Then
        smYN = "No"
        pbcYN_Paint
    ElseIf smYN = "No" Then
        smYN = "Yes"
        pbcYN_Paint
    End If
End Sub

Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    pbcYN.Print smYN
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub mGridColumns()
    grdType.Row = ROW3INDEX - 1
    grdType.Col = NAMEINDEX
    grdType.CellFontBold = False
    grdType.CellFontName = "Arial"
    grdType.CellFontSize = 6.75
    grdType.CellForeColor = vbBlue
    grdType.TextMatrix(ROW3INDEX - 1, NAMEINDEX) = "Name"
    grdType.Row = ROW6INDEX - 1
    grdType.Col = MULTIFEEDINDEX
    grdType.CellFontBold = False
    grdType.CellFontName = "Arial"
    grdType.CellFontSize = 6.75
    grdType.CellForeColor = vbBlue
    grdType.TextMatrix(ROW6INDEX - 1, MULTIFEEDINDEX) = "Ignore Multiple Feeds"
    grdType.Row = ROW9INDEX - 1
    grdType.Col = GLCASHINDEX
    grdType.CellFontBold = False
    grdType.CellFontName = "Arial"
    grdType.CellFontSize = 6.75
    grdType.CellForeColor = vbBlue
    grdType.TextMatrix(ROW9INDEX - 1, GLCASHINDEX) = "G/L Cash"
    grdType.Row = ROW12INDEX - 1
    grdType.Col = GLTRADEINDEX
    grdType.CellFontBold = False
    grdType.CellFontName = "Arial"
    grdType.CellFontSize = 6.75
    grdType.CellForeColor = vbBlue
    grdType.TextMatrix(ROW12INDEX - 1, GLTRADEINDEX) = "G/L Trade"
End Sub

Private Sub mGridColumnWidths()
    grdType.ColWidth(2) = grdType.Width - grdType.ColWidth(0) - grdType.ColWidth(1) - grdType.ColWidth(3) - 75
End Sub
Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilCol = 0 To grdType.Cols - 1 Step 1
        grdType.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdType.RowHeight(0) = 15
    grdType.RowHeight(1) = 15
    grdType.RowHeight(2) = 135
    grdType.RowHeight(3) = fgBoxGridH
    grdType.RowHeight(4) = 15
    grdType.RowHeight(5) = 135
    grdType.RowHeight(6) = fgBoxGridH
    grdType.RowHeight(7) = 15
    grdType.RowHeight(8) = 135
    grdType.RowHeight(9) = fgBoxGridH
    grdType.RowHeight(10) = 15
    grdType.RowHeight(11) = 135
    grdType.RowHeight(12) = fgBoxGridH
    grdType.RowHeight(13) = 15
    grdType.ColWidth(0) = 15
    grdType.ColWidth(1) = 15
    grdType.ColWidth(3) = 15
    'Vertical Line
    For ilRow = 1 To 13
        grdType.Row = ilRow
        grdType.Col = 1
        grdType.CellBackColor = vbBlue
    Next ilRow
    For ilRow = 1 To 13
        grdType.Row = ilRow
        grdType.Col = 3
        grdType.CellBackColor = vbBlue
    Next ilRow
    'Horizontal
    For ilCol = 1 To 3
        grdType.Row = 1
        grdType.Col = ilCol
        grdType.CellBackColor = vbBlue
    Next ilCol
    For ilCol = 1 To 3
        grdType.Row = 4
        grdType.Col = ilCol
        grdType.CellBackColor = vbBlue
    Next ilCol
     For ilCol = 1 To 3
        grdType.Row = 7
        grdType.Col = ilCol
        grdType.CellBackColor = vbBlue
    Next ilCol
    For ilCol = 1 To 3
        grdType.Row = 10
        grdType.Col = ilCol
        grdType.CellBackColor = vbBlue
    Next ilCol
   For ilCol = 1 To 3
        grdType.Row = 13
        grdType.Col = ilCol
        grdType.CellBackColor = vbBlue
    Next ilCol
End Sub
