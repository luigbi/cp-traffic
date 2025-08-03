VERSION 5.00
Begin VB.Form BPlate 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   510
   ClientTop       =   3000
   ClientWidth     =   8490
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   8490
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
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   2805
   End
   Begin VB.ListBox lbcType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1065
      Visible         =   0   'False
      Width           =   2130
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
      Left            =   3030
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   435
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
      Left            =   4050
      Picture         =   "Bplate.frx":0000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcYN 
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
      Left            =   5250
      ScaleHeight     =   210
      ScaleWidth      =   825
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Height          =   285
      Left            =   7545
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Height          =   285
      Left            =   7275
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3255
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Height          =   285
      Left            =   7890
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3255
      Visible         =   0   'False
      Width           =   525
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3435
      Width           =   90
   End
   Begin VB.TextBox edcTitle 
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
      MaxLength       =   20
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   1845
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
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   405
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
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   90
      TabIndex        =   12
      Top             =   3165
      Width           =   90
   End
   Begin VB.TextBox edcComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   1905
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1125
      Width           =   8040
   End
   Begin VB.PictureBox pbcInfo 
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
      Height          =   2445
      Left            =   180
      Picture         =   "Bplate.frx":00FA
      ScaleHeight     =   2445
      ScaleWidth      =   8070
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   615
      Width           =   8070
      Begin VB.Label lacComment 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   1905
         Left            =   45
         TabIndex        =   10
         Top             =   525
         Width           =   7965
      End
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   1065
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   1065
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1635
      TabIndex        =   13
      Top             =   3225
      Width           =   945
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   6015
      TabIndex        =   17
      Top             =   3225
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Top             =   3225
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3825
      TabIndex        =   15
      Top             =   3225
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2730
      TabIndex        =   14
      Top             =   3225
      Width           =   945
   End
   Begin VB.PictureBox plcInfo 
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
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   8145
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   570
      Width           =   8205
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   3180
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Bplate.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BPlate.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Boilerplate input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 7)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imFirstActivate As Integer
'Dim imSave(1 To 5) As Integer   'Index:1=Type;
Dim imSave(0 To 4) As Integer   'Index:
                                '0=Type;1=Show on Proposal; 2=Show on Order; 3=Show on Spot; 4=Show on Invoice
                                '0=Yes; 1=No
Dim smComment As String 'Image save for change comparsion
Dim imBoxNo As Integer   'Current Media Box
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visibly
                                'False= Make list box invisible
Dim tmCmf As CMF        'CMF record image
Dim tmSrchKey As LONGKEY0    'CMF key record image
Dim imRecLen As Integer        'CMF record length
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmCmf As Integer 'Comment file handle
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records

Const TYPEINDEX = 1
Const TITLEINDEX = 2     'Title control/field
Const SHPROPINDEX = 3
Const SHORDERINDEX = 4
Const SHSPOTINDEX = 5
Const SHINVINDEX = 6
Const COMMENTINDEX = 7   'Comment control/field
Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
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
    pbcInfo.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcTitle.Text = slStr
        End If
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcInfo_Paint
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
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        If igCmmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgCmmTitle = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgCmmTitle  'New Name
            End If
            cbcSelect_Change
            If sgCmmTitle <> "" Then
                mSetCommands
                gFindMatch sgCmmTitle, 1, cbcSelect
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
        If pbcSTab.Enabled Then
            pbcSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
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
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCancel_Click()
    If igCmmCallSource <> CALLNONE Then
        If igCmmCallSource = CALLSOURCECONTRACT Then
            igCmmCallSource = CALLCANCELLED
        End If
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igCmmCallSource <> CALLNONE Then
        sgCmmTitle = edcTitle.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgCmmTitle = "[New]"
            If Not imTerminate Then
                mEnableBox imBoxNo
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
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    If igCmmCallSource <> CALLNONE Then
        If igCmmCallSource = CALLSOURCECONTRACT Then
            If sgCmmTitle = "[New]" Then
                igCmmCallSource = CALLCANCELLED
            Else
                igCmmCallSource = CALLDONE
            End If
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case TYPEINDEX
            lbcType.Visible = Not lbcType.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    If imSelectedIndex > 0 Then
        'Check that record is not referenced-Code missing
        ilRet = MsgBox("OK to remove " & tmCmf.sTitle, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Cmf.btr")
        ilRet = btrDelete(hmCmf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", BPlate
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcTitleCode.Tag <> "" Then
        '    If slStamp = lbcTitleCode.Tag Then
        '        lbcTitleCode.Tag = FileDateTime(sgDBPath & "Cmf.btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Cmf.btr")
            End If
        End If
        'lbcTitleCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcInfo.Cls
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
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = BOILERPLATESLIST
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "BPlate^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "BPlate^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "BPlate^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "BPlate^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'BPlate.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'BPlate.Enabled = True
    'slStr = sgDoneMsg
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    ''Screen.MousePointer = vbDefault    'Default
'    RptGen.Show vbModal
    sgCommandStr = slStr
    RptList.Show vbModal
End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUpdate_Click()
    Dim slTitle As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slTitle = edcTitle.Text   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    imBoxNo = -1
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If imSvSelectedIndex <> 0 Then
        cbcSelect.Text = slTitle
    Else
        cbcSelect.ListIndex = 0
    End If
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcComment_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcComment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcComment_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcComment_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim slStr As String
'    If ((Shift And vbCtrlMask) > 0) And (KeyCode = KEYINSERT) Then    'Move to clipboard
'        If edcComment.SelLength > 0 Then
'            slStr = Mid$(edcComment.Text, edcComment.SelStart, edcComment.SelLength)
'            Clipboard.SetText slStr, vbCFText
'        End If
'    End If
'    If ((Shift And vbShiftMask) > 0) And (KeyCode = KEYDELETE) Then   'Clear text and move to clipboard
'        If edcComment.SelLength > 0 Then
'            slStr = Mid$(edcComment.Text, edcComment.SelStart, edcComment.SelLength)
'            Clipboard.SetText slStr, vbCFText
'            slStr = Left$(edcComment.Text, edcComment.SelStart)
'            slStr = slStr & Mid$(edcComment.Text, edcComment.SelStart + edcComment.SelLength + 1)
'            edcComment.Text = slStr
'        End If
'    End If
'    If ((Shift And vbShiftMask) > 0) And (KeyCode = KEYINSERT) Then    'Move from clipboard
'        If Clipboard.GetFormat(vbCFText) Then
'            slStr = Left$(edcComment.Text, edcComment.SelStart)
'            slStr = slStr & Clipboard.GetText(vbCFText)
'            slStr = slStr & Mid$(edcComment.Text, edcComment.SelStart + edcComment.SelLength + 1)
'            edcComment.Text = slStr
'        End If
'    End If
End Sub
Private Sub edcDropDown_Change()
    Select Case imBoxNo
        Case TYPEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcType, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case TYPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcType, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcTitle_Change()
    mSetChg TITLEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcTitle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTitle_KeyPress(KeyAscii As Integer)
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
    If (igWinStatus(BOILERPLATESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcInfo.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcInfo.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BT"
    End If
    gShowBranner imUpdateAllowed
    'This loop is required to prevent a timing problem- if calling
    'with sg----- = "", then loss GotFocus to first control
    'without this loop
'    For ilLoop = 1 To 100 Step 1
'        DoEvents
'    Next ilLoop
'    gShowBranner
    Me.KeyPreview = True
    BPlate.Refresh
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
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
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
        mSetShow imBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox imBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Erase tgNameCode
    btrExtClear hmCmf   'Clear any previous extend operation
    ilRet = btrClose(hmCmf)
    btrDestroy hmCmf
    
    Set BPlate = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lacComment_Click()
    If pbcSTab.Enabled Then
        mSetShow imBoxNo
        imBoxNo = COMMENTINDEX
        mEnableBox imBoxNo
    End If
End Sub
Private Sub lbcType_Click()
    gProcessLbcClick lbcType, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    edcTitle.Text = ""
    'imSave(1) = -1
    imSave(0) = -1
    lbcType.ListIndex = -1
'    imSave(2) = -1
'    imSave(3) = -1
'    imSave(4) = -1
'    imSave(5) = -1
    imSave(1) = -1
    imSave(2) = -1
    imSave(3) = -1
    imSave(4) = -1

    edcComment.Text = ""
    smComment = ""
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case TYPEINDEX
            lbcType.height = gListBoxHeight(lbcType.ListCount, 7)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW ' + cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcInfo, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcType.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
            imChgMode = True
            'If imSave(1) < 0 Then
            If imSave(0) < 0 Then
                gFindMatch "Other Comment", 0, lbcType
                lbcType.ListIndex = gLastFound(lbcType)   '[Other]
            End If
            imComboBoxIndex = lbcType.ListIndex
            If lbcType.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcType.List(lbcType.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TITLEINDEX 'Name
            edcTitle.Width = tmCtrls(ilBoxNo).fBoxW
            edcTitle.MaxLength = 20
            gMoveFormCtrl pbcInfo, edcTitle, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcTitle.Visible = True  'Set visibility
            edcTitle.SetFocus
        Case SHPROPINDEX
            'If imSave(2) < 0 Then
            If imSave(1) < 0 Then
                'imSave(2) = 1  'No
                imSave(1) = 1  'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInfo, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SHORDERINDEX
            'If imSave(3) < 0 Then
            If imSave(2) < 0 Then
                'imSave(3) = 1  'No
                imSave(2) = 1  'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInfo, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SHSPOTINDEX
            'If imSave(4) < 0 Then
            If imSave(3) < 0 Then
                'imSave(4) = 1  'No
                imSave(3) = 1  'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInfo, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SHINVINDEX
            'If imSave(5) < 0 Then
            If imSave(4) < 0 Then
                'imSave(5) = 1  'No
                imSave(4) = 1  'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcInfo, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case COMMENTINDEX 'Comment Name
            edcComment.Width = tmCtrls(ilBoxNo).fBoxW
            edcComment.MaxLength = 5000
            edcComment.Move pbcInfo.Left + tmCtrls(ilBoxNo).fBoxX, pbcInfo.Top + tmCtrls(ilBoxNo).fBoxY + fgOffset
'            gMoveFormCtrl pbcInfo, edcComment, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcComment.Visible = True  'Set visibility
            edcComment.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
    Dim ilRet As Integer
    imLBCtrls = 1
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    BPlate.height = cmcCancel.Top + 5 * cmcCancel.height / 3
    gCenterStdAlone BPlate
    'BPlate.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE

    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imRecLen = Len(tmCmf)  'Get and save CmF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmCmf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCmf, "", sgDBPath & "Cmf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", BPlate
    On Error GoTo 0
'    gCenterModalForm BPlate
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
        mSetCommands
    End If
    'Init Type
    lbcType.AddItem "Cancellation Clause"
    lbcType.AddItem "Change Reason"
    lbcType.AddItem "Internal Comment"
    lbcType.AddItem "Line Comment"
    lbcType.AddItem "Merchandising"
    lbcType.AddItem "Other Comment"
    lbcType.AddItem "Promotion"
    If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER) Then
        lbcType.AddItem "Podcast Ad Server Buy"
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
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    flTextHeight = pbcInfo.TextHeight("1") - 35
    'Position panel and picture areas with panel
    'Position panel and picture areas with panel
    plcInfo.Move 120, 570, pbcInfo.Width + fgPanelAdj, pbcInfo.height + fgPanelAdj
    pbcInfo.Move plcInfo.Left + fgBevelX, plcInfo.Top + fgBevelY
    'Type
    gSetCtrl tmCtrls(TYPEINDEX), 30, 30, 1845, fgBoxStH
    'Title
    gSetCtrl tmCtrls(TITLEINDEX), 1890, tmCtrls(TYPEINDEX).fBoxY, 2415, fgBoxStH
    'Show on Proposal
    gSetCtrl tmCtrls(SHPROPINDEX), 4320, tmCtrls(TYPEINDEX).fBoxY, 825, fgBoxStH
    'Show on Order
    gSetCtrl tmCtrls(SHORDERINDEX), 5160, tmCtrls(TYPEINDEX).fBoxY, 630, fgBoxStH
    'Show on Spot
    gSetCtrl tmCtrls(SHSPOTINDEX), 5801, tmCtrls(TYPEINDEX).fBoxY, 1515, fgBoxStH
    'Show on Invoices
    gSetCtrl tmCtrls(SHINVINDEX), 7335, tmCtrls(TYPEINDEX).fBoxY, 720, fgBoxStH
    'Comment
    gSetCtrl tmCtrls(COMMENTINDEX), 30, tmCtrls(TYPEINDEX).fBoxY + fgStDeltaY, 8010, 2055
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim slStr As String
    If Not ilTestChg Or tmCtrls(TYPEINDEX).iChg Then
        'slStr = lbcType.List(imSave(1))
        slStr = lbcType.List(imSave(0))
        If StrComp(slStr, "Cancellation Clause", 1) = 0 Then
            tmCmf.sComType = "C"
        ElseIf StrComp(slStr, "Change Reason", 1) = 0 Then
            tmCmf.sComType = "R"
        ElseIf StrComp(slStr, "Internal Comment", 1) = 0 Then
            tmCmf.sComType = "I"
        ElseIf StrComp(slStr, "Line Comment", 1) = 0 Then
            tmCmf.sComType = "L"
        ElseIf StrComp(slStr, "Merchandising", 1) = 0 Then
            tmCmf.sComType = "M"
        ElseIf StrComp(slStr, "Other Comment", 1) = 0 Then
            tmCmf.sComType = "O"
        ElseIf StrComp(slStr, "Promotion", 1) = 0 Then
            tmCmf.sComType = "P"
        ElseIf StrComp(slStr, "Podcast CPM Buy", 1) = 0 Then
            tmCmf.sComType = "B"
        End If
    End If
    If Not ilTestChg Or tmCtrls(TITLEINDEX).iChg Then
        tmCmf.sTitle = edcTitle.Text
    End If
    If Not ilTestChg Or tmCtrls(SHPROPINDEX).iChg Then
        'Select Case imSave(2)
        Select Case imSave(1)
            Case 0  'Yes
                tmCmf.sShProp = "Y"
            Case 1  'No
                tmCmf.sShProp = "N"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SHORDERINDEX).iChg Then
        'Select Case imSave(3)
        Select Case imSave(2)
            Case 0  'Yes
                tmCmf.sShOrder = "Y"
            Case 1  'No
                tmCmf.sShOrder = "N"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SHSPOTINDEX).iChg Then
        'Select Case imSave(4)
        Select Case imSave(3)
            Case 0  'Yes
                tmCmf.sShSpot = "Y"
            Case 1  'No
                tmCmf.sShSpot = "N"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SHINVINDEX).iChg Then
        'Select Case imSave(5)
        Select Case imSave(4)
            Case 0  'Yes
                tmCmf.sShInv = "Y"
            Case 1  'No
                tmCmf.sShInv = "N"
        End Select
    End If
    If Not ilTestChg Or tmCtrls(COMMENTINDEX).iChg Then
        'tmCmf.iStrLen = Len(edcComment.Text)
        tmCmf.sComment = Trim$(edcComment.Text) & Chr$(0) '& Chr$(0) 'sgTB
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    edcTitle.Text = Trim$(tmCmf.sTitle)
    Select Case tmCmf.sComType
        Case "C"
            gFindMatch "Cancellation Clause", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "R"
            gFindMatch "Change Reason", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "I"
            gFindMatch "Internal Comment", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "L"
            gFindMatch "Line Comment", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "M"
            gFindMatch "Merchandising", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "O"
            gFindMatch "Other Comment", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "P"
            gFindMatch "Promotion", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
        Case "B"
            gFindMatch "Podcast CPM Buy", 0, lbcType
            'imSave(1) = gLastFound(lbcType)
            imSave(0) = gLastFound(lbcType)
    End Select
    'lbcType.ListIndex = imSave(1)
    lbcType.ListIndex = imSave(0)
    Select Case tmCmf.sShProp
        Case "Y"
            'imSave(2) = 0
            imSave(1) = 0
        Case "N"
            'imSave(2) = 1
            imSave(1) = 1
        Case Else
            'imSave(2) = -1
            imSave(1) = -1
    End Select
    Select Case tmCmf.sShOrder
        Case "Y"
            'imSave(3) = 0
            imSave(2) = 0
        Case "N"
            'imSave(3) = 1
            imSave(2) = 1
        Case Else
            'imSave(3) = -1
            imSave(2) = -1
    End Select
    Select Case tmCmf.sShSpot
        Case "Y"
            'imSave(4) = 0
            imSave(3) = 0
        Case "N"
            'imSave(4) = 1
            imSave(3) = 1
        Case Else
            'imSave(4) = -1
            imSave(3) = -1
    End Select
    Select Case tmCmf.sShInv
        Case "Y"
            'imSave(5) = 0
            imSave(4) = 0
        Case "N"
            'imSave(5) = 1
            imSave(4) = 1
        Case Else
            'imSave(5) = -1
            imSave(4) = -1
    End Select
    'If tmCmf.iStrLen > 0 Then
    '    edcComment.Text = Trim$(Left$(tmCmf.sComment, tmCmf.iStrLen))
    'Else
    '    edcComment.Text = ""
    'End If
    edcComment.Text = gStripChr0(tmCmf.sComment)
    smComment = edcComment.Text
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
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
    If edcTitle.Text <> "" Then    'Test name
        slStr = Trim$(edcTitle.Text)
        gFindMatch slStr, 0, cbcSelect   'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcTitle.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Boilerplate Title already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcTitle.Text = Trim$(tmCmf.sTitle) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
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
    'gInitStdAlone BPlate, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
    If igCmmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgCmmTitle = slStr
        Else
            sgCmmTitle = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    'ilRet = gLMoveListBox(BPlate, cbcSelect, lbcTitleCode, "Cmf.btr", gFieldOffset("Cmf", "CmfTitle"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gLMoveListBox(BPlate, cbcSelect, tgNameCode(), sgNameCodeTag, "Cmf.btr", gFieldOffset("Cmf", "CmfTitle"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gLMoveListBox)", BPlate
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
'   iRet = mReadRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slTitleCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slTitleCode = tgNameCode(ilSelectIndex - 1).sKey   'lbcTitleCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slTitleCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", BPlate
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmSrchKey.lCode = CLng(slCode)
    imRecLen = Len(tmCmf)  'Get and save CmF record length (the read will change the length)
    ilRet = btrGetEqual(hmCmf, tmCmf, imRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", BPlate
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
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim slName As String    'Name
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilCmRecLen As Integer
    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Cmf.btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        'ilCmRecLen = 29 + Len(Trim$(tmCmf.sComment)) + 2'5 = fixed record length; 2=Length value which is part of the variable record
        '7/14/2010 Dan change for v57 only
        'ilCmRecLen = Len(tmCmf) - Len(tmCmf.sComment) + Len(Trim$(tmCmf.sComment))
        ilCmRecLen = Len(tmCmf)
        If imSelectedIndex = 0 Then 'New selected
            tmCmf.lCode = 0  'Autoincrement
            ilRet = btrInsert(hmCmf, tmCmf, ilCmRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmCmf, tmCmf, ilCmRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, BPlate
    On Error GoTo 0
    'If lbcTitleCode.Tag <> "" Then
    '    If slStamp = lbcTitleCode.Tag Then
    '        lbcTitleCode.Tag = FileDateTime(sgDBPath & "Cmf.btr")
    '    End If
    'End If
    If sgNameCodeTag <> "" Then
        If slStamp = sgNameCodeTag Then
            sgNameCodeTag = gFileDateTime(sgDBPath & "Cmf.btr")
        End If
    End If
    If imSelectedIndex <> 0 Then
        'lbcTitleCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    cbcSelect.RemoveItem 0 'Remove [New]
    slName = Trim$(tmCmf.sTitle)
    cbcSelect.AddItem slName
    slName = tmCmf.sTitle + "\" + LTrim$(str$(tmCmf.lCode)) 'slName + "\" + LTrim$(Str$(tmCmf.lCode))
    'lbcTitleCode.AddItem slName
    gAddItemToSortCode slName, tgNameCode(), True
    cbcSelect.AddItem "[New]", 0
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
'*             Created:5/17/93       By:D. LeVine      *
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
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcTitle.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcInfo_Paint
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
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case TYPEINDEX
            Select Case tmCmf.sComType
                Case "C"
                    slStr = "Cancellation Clause"
                Case "R"
                    slStr = "Change Reason"
                Case "I"
                    slStr = "Internal Comment"
                Case "L"
                    slStr = "Line Comment"
                Case "M"
                    slStr = "Merchandising"
                Case "O"
                    slStr = "Other Comment"
                Case "P"
                    slStr = "Promotion"
                Case "B"
                    slStr = "Podcast CPM Buy"
            End Select
            gSetChgFlag slStr, lbcType, tmCtrls(ilBoxNo)
        Case TITLEINDEX 'Name
            gSetChgFlag tmCmf.sTitle, edcTitle, tmCtrls(ilBoxNo)
        Case COMMENTINDEX 'Comment
            gSetChgFlag smComment, edcComment, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
 '   If ilAltered Then
 '       cmcUndo.Enabled = True
 '   Else
 '       cmcUndo.Enabled = False
 '   End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
        If imUpdateAllowed Then
            cmcErase.Enabled = True
        Else
            cmcErase.Enabled = False
        End If
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
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slStr1 As String

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case TYPEINDEX
            lbcType.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            'imSave(1) = lbcType.ListIndex
            imSave(0) = lbcType.ListIndex
            If lbcType.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcType.List(lbcType.ListIndex)
            End If
            gSetShow pbcInfo, slStr, tmCtrls(ilBoxNo)
        Case TITLEINDEX 'Name
            edcTitle.Visible = False  'Set visibility
            slStr = edcTitle.Text
            gSetShow pbcInfo, slStr, tmCtrls(ilBoxNo)
        Case SHPROPINDEX
            pbcYN.Visible = False  'Set visibility
            'If imSave(2) = 0 Then
            If imSave(1) = 0 Then
                slStr = "Yes"
            'ElseIf imSave(2) = 1 Then
            ElseIf imSave(1) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcInfo, slStr, tmCtrls(ilBoxNo)
        Case SHORDERINDEX
            pbcYN.Visible = False  'Set visibility
            'If imSave(3) = 0 Then
            If imSave(2) = 0 Then
                slStr = "Yes"
            'ElseIf imSave(3) = 1 Then
            ElseIf imSave(2) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcInfo, slStr, tmCtrls(ilBoxNo)
        Case SHSPOTINDEX
            pbcYN.Visible = False  'Set visibility
            'If imSave(4) = 0 Then
            If imSave(3) = 0 Then
                slStr = "Yes"
            'ElseIf imSave(4) = 1 Then
            ElseIf imSave(3) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            'If imSave(1) >= 0 Then
            If imSave(0) >= 0 Then
                'slStr1 = lbcType.List(imSave(1))
                slStr1 = lbcType.List(imSave(0))
                '5/26/10: Add line comment
                'If StrComp(slStr, "Other Comment", 1) <> 0 Then
                'dan 7/13/10 change slStr to slStr1
                If (StrComp(slStr1, "Other Comment", 1) <> 0) And (StrComp(slStr1, "Line Comment", 1) <> 0) Then
                    slStr = ""
                End If
            Else
                slStr = ""
            End If
            gSetShow pbcInfo, slStr, tmCtrls(ilBoxNo)
        Case SHINVINDEX
            pbcYN.Visible = False  'Set visibility
            'If imSave(5) = 0 Then
            If imSave(4) = 0 Then
                slStr = "Yes"
            'ElseIf imSave(5) = 1 Then
            ElseIf imSave(4) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcInfo, slStr, tmCtrls(ilBoxNo)
        Case COMMENTINDEX 'Market Name
            edcComment.Visible = False  'Set visibility
            lacComment.Caption = edcComment.Text
    End Select
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
    Dim ilRet As Integer
    sgNameCodeTag = ""
    sgDoneMsg = Trim$(str$(igCmmCallSource)) & "\" & sgCmmTitle
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload BPlate
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
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
    If (ilCtrlNo = TYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcType, "", "Type must be specified", tmCtrls(TYPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TYPEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TITLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcTitle, "", "Title must be specified", tmCtrls(TITLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TITLEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SHPROPINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If imSave(2) = 0 Then
        If imSave(1) = 0 Then
            slStr = "Yes"
        'ElseIf imSave(2) = 1 Then
        ElseIf imSave(1) = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Show on Proposal must be specified", tmCtrls(SHPROPINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SHPROPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SHORDERINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If imSave(3) = 0 Then
        If imSave(2) = 0 Then
            slStr = "Yes"
        'ElseIf imSave(3) = 1 Then
        ElseIf imSave(2) = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Show on Order must be specified", tmCtrls(SHORDERINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SHPROPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SHSPOTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If imSave(4) = 0 Then
        If imSave(3) = 0 Then
            slStr = "Yes"
        'ElseIf imSave(4) = 1 Then
        ElseIf imSave(3) = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Show on Spot must be specified", tmCtrls(SHSPOTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SHPROPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SHINVINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If imSave(5) = 0 Then
        If imSave(4) = 0 Then
            slStr = "Yes"
        'ElseIf imSave(5) = 1 Then
        ElseIf imSave(4) = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Show on Invoice must be specified", tmCtrls(SHINVINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SHPROPINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMMENTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcComment, "", "Comment must be specified", tmCtrls(COMMENTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMMENTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim slStr As String

    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If ilBox = SHSPOTINDEX Then
                    'If imSave(1) >= 0 Then
                    If imSave(0) >= 0 Then
                        'slStr = lbcType.List(imSave(1))
                        slStr = lbcType.List(imSave(0))
                        '5/26/10: Add line comment
                        'If StrComp(slStr, "Other Comment", 1) <> 0 Then
                        If (StrComp(slStr, "Other Comment", 1) <> 0) And (StrComp(slStr, "Line Comment", 1) <> 0) Then
                            Beep
                            Exit Sub
                        End If
                    Else
                        Beep
                        Exit Sub
                    End If
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcInfo_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If ilBox = COMMENTINDEX Then
            lacComment.Caption = edcComment.Text
        Else
            pbcInfo.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcInfo.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcInfo.Print tmCtrls(ilBox).sShow
        End If
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If (imBoxNo <> TITLEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    Select Case imBoxNo
        Case -1
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                mSetChg 2   'Title and Type swapped positions
                ilBox = 1
            End If
        Case 1 'Type (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case SHINVINDEX
            'If imSave(1) >= 0 Then
            If imSave(0) >= 0 Then
                'slStr = lbcType.List(imSave(1))
                slStr = lbcType.List(imSave(0))
                '5/26/10: Add line comment
                'If StrComp(slStr, "Other Comment", 1) = 0 Then
                If (StrComp(slStr, "Other Comment", 1) = 0) Or (StrComp(slStr, "Line Comment", 1) = 0) Then
                    ilBox = imBoxNo - 1
                Else
                    ilBox = imBoxNo - 2
                    'If imSave(4) < 0 Then
                    If imSave(3) < 0 Then
                        'imSave(4) = 1
                        imSave(3) = 1
                    End If
                End If
            Else
                ilBox = imBoxNo - 2
                'If imSave(4) < 0 Then
                If imSave(3) < 0 Then
                    'imSave(4) = 1
                    imSave(3) = 1
                End If
            End If
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1
            ilBox = UBound(tmCtrls)
        Case SHORDERINDEX
            'If imSave(1) >= 0 Then
            If imSave(0) >= 0 Then
                'slStr = lbcType.List(imSave(1))
                slStr = lbcType.List(imSave(0))
                '5/26/10: Add line comment
                'If StrComp(slStr, "Other Comment", 1) = 0 Then
                If (StrComp(slStr, "Other Comment", 1) = 0) Or StrComp(slStr, "Line Comment", 1) = 0 Then
                    ilBox = imBoxNo + 1
                Else
                    'If imSave(4) < 0 Then
                    If imSave(3) < 0 Then
                        'imSave(4) = 1
                        imSave(3) = 1
                    End If
                    ilBox = imBoxNo + 2
                End If
            Else
                ilBox = imBoxNo + 2
                'If imSave(4) < 0 Then
                If imSave(3) < 0 Then
                    'imSave(4) = 1
                    imSave(3) = 1
                End If
            End If
        Case COMMENTINDEX
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igCmmCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    'ilIndex = imBoxNo - SHPROPINDEX + 2
    ilIndex = imBoxNo - SHPROPINDEX + 1
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imSave(ilIndex) <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSave(ilIndex) = 0
        pbcYN_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imSave(ilIndex) <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSave(ilIndex) = 1
        pbcYN_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSave(ilIndex) = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(ilIndex) = 1
            pbcYN_Paint
        ElseIf imSave(ilIndex) = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSave(ilIndex) = 0
            pbcYN_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    'ilIndex = imBoxNo - SHPROPINDEX + 2
    ilIndex = imBoxNo - SHPROPINDEX + 1
    If imSave(ilIndex) = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imSave(ilIndex) = 1
    ElseIf imSave(ilIndex) = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imSave(ilIndex) = 0
    End If
    pbcYN_Paint
    mSetCommands
End Sub
Private Sub pbcYN_Paint()
    Dim ilIndex As Integer
    'ilIndex = imBoxNo - SHPROPINDEX + 2
    ilIndex = imBoxNo - SHPROPINDEX + 1
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imSave(ilIndex) = 0 Then
        pbcYN.Print "Yes"
    ElseIf imSave(ilIndex) = 1 Then
        pbcYN.Print "No"
    Else
        pbcYN.Print "   "
    End If
End Sub
Private Sub plcInfo_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Boilerplate"
End Sub
