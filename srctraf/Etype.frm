VERSION 5.00
Begin VB.Form EType 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   1905
   ClientTop       =   2685
   ClientWidth     =   5310
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
   ScaleHeight     =   2940
   ScaleWidth      =   5310
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Top             =   225
      Width           =   2970
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2595
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2730
      Width           =   105
   End
   Begin VB.PictureBox pbcInUse 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3540
      ScaleHeight     =   210
      ScaleWidth      =   1095
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbcTime 
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
      Height          =   240
      Left            =   540
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1545
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lbcLen 
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
      Height          =   240
      Left            =   2340
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbcProg 
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
      Height          =   240
      Left            =   3495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1605
      Visible         =   0   'False
      Width           =   1290
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
      Left            =   1485
      Picture         =   "Etype.frx":0000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   480
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   645
      MaxLength       =   20
      TabIndex        =   4
      Top             =   885
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
      Left            =   3255
      TabIndex        =   17
      Top             =   2550
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
      Left            =   2130
      TabIndex        =   16
      Top             =   2550
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
      Left            =   1005
      TabIndex        =   15
      Top             =   2550
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
      Left            =   3255
      TabIndex        =   14
      Top             =   2205
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
      Left            =   2130
      TabIndex        =   13
      Top             =   2205
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
      Left            =   1005
      TabIndex        =   12
      Top             =   2205
      Width           =   1050
   End
   Begin VB.PictureBox pbcEvt 
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
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   600
      Picture         =   "Etype.frx":00FA
      ScaleHeight     =   1065
      ScaleWidth      =   4050
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   735
      Width           =   4050
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   765
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   135
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
      TabIndex        =   11
      Top             =   1695
      Width           =   45
   End
   Begin VB.PictureBox plcEvt 
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   555
      ScaleHeight     =   1110
      ScaleWidth      =   4110
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   690
      Width           =   4170
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1845
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4605
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label plcScreen 
      Caption         =   "Event Types"
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
      Left            =   165
      Top             =   2475
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "EType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Etype.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EType.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Event Type input screen code
Option Explicit
Option Compare Text
'Event Type Field Areas
Dim tmCtrls(0 To 5)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Event Type Box
Dim tmEtf As ETF        'ETF record image
Dim tmSrchKey As INTKEY0    'ETF key record image
Dim imRecLen As Integer        'ETF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmEtf As Integer 'Event Type file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imComboBoxIndex As Integer
Dim imTimeFirst As Integer  'First time at field- set default
Dim imLenFirst As Integer   'First time at field-set default
Dim imProgFirst As Integer  'First time at field- set default
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imInUse As Integer  '0=Yes; 1=No
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records

Const NAMEINDEX = 1  'Name control/field
Const INUSEINDEX = 2 'In use control/field
Const TIMEINDEX = 3  'Time format control/field
Const LENINDEX = 4   'Length format control/field
Const PROGINDEX = 5  'Prog/sponsor show format control/field
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
    pbcEvt.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcEvt_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
'    mSetCommands
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
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        For ilLoop = 1 To igDDEDelay Step 1
            DoEvents
        Next ilLoop
        If igETypeCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgETypeName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgETypeName    'New name
            End If
            cbcSelect_Change
            If sgETypeName <> "" Then
                mSetCommands
                gFindMatch sgETypeName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            pbcSTab.SetFocus
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
    If igETypeCallSource <> CALLNONE Then
        igETypeCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igETypeCallSource <> CALLNONE Then
        sgETypeName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgETypeName = "[New]"
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
    If igETypeCallSource <> CALLNONE Then
        If (sgETypeName = "[New]") Or (sgETypeName = "") Then
            igETypeCallSource = CALLCANCELLED
        Else
            igETypeCallSource = CALLDONE
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
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case TIMEINDEX
            lbcTime.Visible = Not lbcTime.Visible
        Case LENINDEX
            lbcLen.Visible = Not lbcLen.Visible
        Case PROGINDEX
            lbcProg.Visible = Not lbcProg.Visible
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
    Dim slMsg As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(EType, tmEtf.iCode, "Cif.Btr", "CifEtfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Copy Inventory references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(EType, tmEtf.iCode, "Crf.Btr", "CrfEtfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Copy Rotation references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gLICodeRefExist(EType, tmEtf.iCode, "Lef.Btr", "LefEtfCode")  'lefetfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Library Events references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(EType, tmEtf.iCode, "Enf.Btr", "EnfEtfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Event Name references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmEtf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Etf.btr")
        ilRet = btrDelete(hmEtf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", EType
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcNameCode.Tag <> "" Then
        '    If slStamp = lbcNameCode.Tag Then
        '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Etf.btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Etf.btr")
            End If
        End If
        'lbcNameCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcEvt.Cls
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
    mSetShow imBoxNo
    imBoxNo = -1
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
            slStr = "EType^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "EType^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EType^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "EType^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'EType.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'EType.Enabled = True
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
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
                GoTo cmcUndoErr
        End If
        pbcEvt.Cls
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcEvt_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcEvt.Cls
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
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = Trim$(edcName.Text)   'Save name
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
    mClearCtrlFields
    sgNameCodeTag = ""
    cbcSelect.Clear
    mPopulate  'Repopulate since list box not in sorted order
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
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcDropDown_Change()
    Select Case imBoxNo
        Case TIMEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcTime, imBSMode, imComboBoxIndex
        Case LENINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcLen, imBSMode, imComboBoxIndex
        Case PROGINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcProg, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case TIMEINDEX
            If lbcTime.ListCount = 1 Then
                lbcTime.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case LENINDEX
            If lbcLen.ListCount = 1 Then
                lbcLen.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case PROGINDEX
            If lbcProg.ListCount = 1 Then
                lbcProg.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
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
            Case TIMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcTime, imLbcArrowSetting
            Case LENINDEX
                gProcessArrowKey Shift, KeyCode, lbcLen, imLbcArrowSetting
            Case PROGINDEX
                gProcessArrowKey Shift, KeyCode, lbcProg, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcName_Change()
    mSetChg NAMEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
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
Private Sub edcName_LostFocus()
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
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
    If (igWinStatus(EVENTTYPESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcEvt.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcEvt.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    EType.Refresh
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

    sgNameCodeTag = ""
    Erase tgNameCode

    btrExtClear hmEtf   'Clear any previous extend operation
    ilRet = btrClose(hmEtf)
    btrDestroy hmEtf
    
    Set EType = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcLen_Click()
    gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLen_GotFocus()
    gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcProg_Click()
    gProcessLbcClick lbcProg, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcProg_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcTime_Click()
    gProcessLbcClick lbcTime, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcTime_GotFocus()
    gCtrlGotFocus ActiveControl
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
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    edcName.Text = ""
    imInUse = -1
    lbcTime.ListIndex = -1  'This set text to ""
    lbcLen.ListIndex = -1
    lbcProg.ListIndex = -1
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    imTimeFirst = True
    imLenFirst = True
    imProgFirst = True
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
        Case NAMEINDEX 'Name
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcEvt, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case INUSEINDEX
            If imInUse < 0 Then
                imInUse = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcInUse.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcEvt, pbcInUse, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcInUse_Paint
            pbcInUse.Visible = True
            pbcInUse.SetFocus
        Case TIMEINDEX   'Time format
            lbcTime.Height = gListBoxHeight(lbcTime.ListCount, 3)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW ' - cmcDropDown.Width
            edcDropDown.MaxLength = 17
            gMoveFormCtrl pbcEvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTime.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcTime.ListIndex < 0 Then
                lbcTime.ListIndex = 5
            End If
            imComboBoxIndex = lbcTime.ListIndex
            If lbcTime.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcTime.List(lbcTime.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case LENINDEX   'Length format
            lbcLen.Height = gListBoxHeight(lbcLen.ListCount, 3)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW ' - cmcDropDown.Width
            edcDropDown.MaxLength = 9
            gMoveFormCtrl pbcEvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcLen.ListIndex < 0 Then
                lbcLen.ListIndex = 1
            End If
            imComboBoxIndex = lbcLen.ListIndex
            If lbcLen.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcLen.List(lbcLen.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case PROGINDEX   'Program format
            lbcProg.Height = gListBoxHeight(lbcProg.ListCount, 3)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW ' - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcEvt, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProg.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcProg.ListIndex < 0 Then
                lbcProg.ListIndex = 1
            End If
            imComboBoxIndex = lbcProg.ListIndex
            If lbcProg.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcProg.List(lbcProg.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
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
'
'   mInitParameters
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Error return status
    Dim llNoRecs As Long    'Number of records
    imFirstActivate = True
    imTerminate = False
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    EType.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone EType
    'EType.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imRecLen = Len(tmEtf)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmEtf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmEtf, "", sgDBPath & "Etf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", EType
    On Error GoTo 0
    llNoRecs = btrRecords(hmEtf)
    If llNoRecs = 0 Then
        For ilLoop = 1 To 13 Step 1
            tmEtf.sTimeForm = ""
            tmEtf.sLenForm = ""
            tmEtf.sPgmForm = ""
            Select Case ilLoop
                Case 1
                    tmEtf.sType = "1"
                    tmEtf.sName = "Programs"
                    tmEtf.sTimeForm = "3"
                    tmEtf.sLenForm = "3"
                    tmEtf.sPgmForm = "1"
                    tmEtf.sInUse = "Y"
                Case 2
                    tmEtf.sType = "2"
                    tmEtf.sName = "Contract Avails"
                    tmEtf.sTimeForm = "6"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "Y"
                Case 3
                    tmEtf.sType = "3"
                    tmEtf.sName = "Open BB Avails"
                    tmEtf.sTimeForm = "1"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 4
                    tmEtf.sType = "4"
                    tmEtf.sName = "Floating Avails"
                    tmEtf.sTimeForm = "1"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 5
                    tmEtf.sType = "5"
                    tmEtf.sName = "Close BB Avails"
                    tmEtf.sTimeForm = "1"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 6
                    tmEtf.sType = "6"
                    tmEtf.sName = "Cmml Promo Avails"
                    tmEtf.sTimeForm = "6"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 7
                    tmEtf.sType = "7"
                    tmEtf.sName = "Feed Avails"
                    tmEtf.sTimeForm = "6"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 8
                    tmEtf.sType = "8"
                    tmEtf.sName = "PSA Avails"
                    tmEtf.sTimeForm = "6"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 9
                    tmEtf.sType = "9"
                    tmEtf.sName = "Promo Avails"
                    tmEtf.sTimeForm = "6"
                    tmEtf.sLenForm = "2"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "N"
                Case 10
                    tmEtf.sType = "A"
                    tmEtf.sName = "Page Skips"
                    tmEtf.sTimeForm = "7"
                    tmEtf.sLenForm = "3"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "Y"
                Case 11
                    tmEtf.sType = "B"
                    tmEtf.sName = "1 Line Space"
                    tmEtf.sTimeForm = "7"
                    tmEtf.sLenForm = "3"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "Y"
                Case 12
                    tmEtf.sType = "C"
                    tmEtf.sName = "2 Line Spaces"
                    tmEtf.sTimeForm = "7"
                    tmEtf.sLenForm = "3"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "Y"
                Case 13
                    tmEtf.sType = "D"
                    tmEtf.sName = "3 Line Spaces"
                    tmEtf.sTimeForm = "7"
                    tmEtf.sLenForm = "3"
                    tmEtf.sPgmForm = "2"
                    tmEtf.sInUse = "Y"
            End Select
            Do  'Loop until record added
                tmEtf.iCode = 0  'Autoincrement
                ilRet = btrInsert(hmEtf, tmEtf, imRecLen, INDEXKEY0)
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mInitErr
            gBtrvErrorMsg ilRet, "mInit (btrInsert)", EType
            On Error GoTo 0
        Next ilLoop
    End If
    If imTerminate Then
        Exit Sub
    End If
    lbcTime.AddItem "Event Name"
    lbcTime.AddItem "hh:mm:ss-hh:mm:ss"
    lbcTime.AddItem "hh:mm-hh:mm"
    lbcTime.AddItem "hh:mm:ss"
    lbcTime.AddItem "hh:mm"
    lbcTime.AddItem "mm:ss"
    lbcTime.AddItem "Blank"
    lbcLen.AddItem "hh:mm:ss"
    lbcLen.AddItem "hh mm'ss"""
    lbcLen.AddItem "Blank"
    lbcProg.AddItem "Event Name"
    lbcProg.AddItem "Blank"
'    gCenterModalForm EType
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list box to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
        mSetCommands
    End If
    imTimeFirst = True
    imLenFirst = True
    imProgFirst = True
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
    Dim flTextHeight As Single  'Standard text height
    flTextHeight = pbcEvt.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcEvt.Move 555, 675, pbcEvt.Width + fgPanelAdj, pbcEvt.Height + fgPanelAdj
    pbcEvt.Move plcEvt.Left + fgBevelX, plcEvt.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2895, fgBoxStH
    'In Use
    gSetCtrl tmCtrls(INUSEINDEX), 2940, tmCtrls(NAMEINDEX).fBoxY, 1095, fgBoxStH
    'Time
    gSetCtrl tmCtrls(TIMEINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + 2 * fgStDeltaY, 1800, fgBoxStH
    tmCtrls(TIMEINDEX).iReq = False
    'Length
    gSetCtrl tmCtrls(LENINDEX), 1845, tmCtrls(TIMEINDEX).fBoxY, 1080, fgBoxStH
    tmCtrls(LENINDEX).iReq = False
    'Program/Sponsor
    gSetCtrl tmCtrls(PROGINDEX), 2970, tmCtrls(TIMEINDEX).fBoxY, 1095, fgBoxStH
    tmCtrls(PROGINDEX).iReq = False
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
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmEtf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(TIMEINDEX).iChg Then
        If lbcTime.Text = "" Then
            tmEtf.sTimeForm = ""
        Else
            tmEtf.sTimeForm = Trim$(str$(lbcTime.ListIndex + 1))
        End If
    End If
    If Not ilTestChg Or tmCtrls(INUSEINDEX).iChg Then
        Select Case imInUse
            Case 0  'Yes
                tmEtf.sInUse = "Y"
            Case 1  'No
                tmEtf.sInUse = "N"
            Case Else
                tmEtf.sInUse = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(LENINDEX).iChg Then
        If lbcLen.Text = "" Then
            tmEtf.sLenForm = ""
        Else
            tmEtf.sLenForm = Trim$(str$(lbcLen.ListIndex + 1))
        End If
    End If
    If Not ilTestChg Or tmCtrls(PROGINDEX).iChg Then
        If lbcProg.Text = "" Then
            tmEtf.sPgmForm = ""
        Else
            tmEtf.sPgmForm = Trim$(str$(lbcProg.ListIndex + 1))
        End If
    End If
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
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    edcName.Text = Trim$(tmEtf.sName)
    Select Case tmEtf.sInUse
        Case "Y"
            imInUse = 0
        Case "N"
            imInUse = 1
        Case Else
            imInUse = -1
    End Select
    If tmEtf.sTimeForm = "" Then
        lbcTime.ListIndex = -1
    Else
        lbcTime.ListIndex = Val(tmEtf.sTimeForm) - 1
    End If
    If tmEtf.sLenForm = "" Then
        lbcLen.ListIndex = -1
    Else
        lbcLen.ListIndex = Val(tmEtf.sLenForm) - 1
    End If
    If tmEtf.sPgmForm = "" Then
        lbcProg.ListIndex = -1
    Else
        lbcProg.ListIndex = Val(tmEtf.sPgmForm) - 1
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    imTimeFirst = True
    imLenFirst = True
    imProgFirst = True
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
    If edcName.Text <> "" Then    'Test name
        slStr = Trim$(edcName.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Event Type already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmEtf.sName) 'Reset text
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
    'gInitStdAlone EType, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igETypeCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igETypeCallSource = CALLNONE
    'End If
    If igETypeCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgETypeName = slStr
        Else
            sgETypeName = ""
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
    imPopReqd = False
    'ilRet = gPopEvtNmByTypeBox(EType, False, True, cbcSelect, lbcNameCode)
    ilRet = gPopEvtNmByTypeBox(EType, False, True, cbcSelect, tgNameCode(), sgNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", EType
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

    slNameCode = tgNameCode(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 3)", EType
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmEtf, tmEtf, imRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", EType
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
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
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
        slStamp = gFileDateTime(sgDBPath & "Etf.btr")
        'If Len(lbcNameCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcNameCode.Tag, Len(lbcNameCode.Tag) - Len(slStamp))
        'End If
        If Len(sgNameCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(sgNameCodeTag, Len(sgNameCodeTag) - Len(slStamp))
        End If
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmEtf.iCode = 0  'Autoincrement
            tmEtf.iMerge = 0
            tmEtf.sType = "Y"
            ilRet = btrInsert(hmEtf, tmEtf, imRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmEtf, tmEtf, imRecLen)
            slMsg = "mSaveRec (btr(Update)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, EType
    On Error GoTo 0
'    If lbcNameCode.Tag <> "" Then
'        If slStamp = lbcNameCode.Tag Then
'            lbcNameCode.Tag = FileDateTime(sgDBPath & "Etf.btr")
'            If Len(slStamp) > Len(lbcNameCode.Tag) Then
'                lbcNameCode.Tag = lbcNameCode.Tag & Right$(slStamp, Len(slStamp) - Len(lbcNameCode.Tag))
'            End If
'        End If
'    End If
    'mPopulate  'Repopulate since list box not in sorted order
    'cbcSelect.ListIndex = 0
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
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcEvt_Paint
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
'*             Created:5/11/93       By:D. LeVine      *
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
        Case NAMEINDEX 'Name
            gSetChgFlag tmEtf.sName, edcName, tmCtrls(ilBoxNo)
        Case TIMEINDEX 'Time format
            If tmEtf.sTimeForm = "" Then
                gSetChgFlag tmEtf.sTimeForm, lbcTime, tmCtrls(ilBoxNo)
            Else
                slStr = lbcTime.List(Val(tmEtf.sTimeForm) - 1)
                gSetChgFlag slStr, lbcTime, tmCtrls(ilBoxNo)
            End If
        Case LENINDEX 'Length format
            If tmEtf.sLenForm = "" Then
                gSetChgFlag tmEtf.sLenForm, lbcLen, tmCtrls(ilBoxNo)
            Else
                slStr = lbcLen.List(Val(tmEtf.sLenForm) - 1)
                gSetChgFlag slStr, lbcLen, tmCtrls(ilBoxNo)
            End If
        Case PROGINDEX 'Selling or Airing or N/At
            If tmEtf.sPgmForm = "" Then
                gSetChgFlag tmEtf.sPgmForm, lbcProg, tmCtrls(ilBoxNo)
            Else
                slStr = lbcProg.List(Val(tmEtf.sPgmForm) - 1)
                gSetChgFlag slStr, lbcProg, tmCtrls(ilBoxNo)
            End If
    End Select
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
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) And (imUpdateAllowed) Then
        If imSelectedIndex > 9 Then
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
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case INUSEINDEX
            pbcInUse.SetFocus
        Case TIMEINDEX   'Time format
            edcDropDown.SetFocus
        Case LENINDEX   'Length format
            edcDropDown.SetFocus
        Case PROGINDEX   'Program format
            edcDropDown.SetFocus
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
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcEvt, slStr, tmCtrls(ilBoxNo)
        Case INUSEINDEX 'Sustaining
            pbcInUse.Visible = False  'Set visibility
            If imInUse = 0 Then
                slStr = "Yes"
            ElseIf imInUse = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcEvt, slStr, tmCtrls(ilBoxNo)
        Case TIMEINDEX 'Time format
            lbcTime.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTime.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcTime.List(lbcTime.ListIndex)
            End If
            gSetShow pbcEvt, slStr, tmCtrls(ilBoxNo)
        Case LENINDEX 'Length format
            lbcLen.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcLen.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcLen.List(lbcLen.ListIndex)
            End If
            gSetShow pbcEvt, slStr, tmCtrls(ilBoxNo)
        Case PROGINDEX 'Program
            lbcProg.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcProg.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcProg.List(lbcProg.ListIndex)
            End If
            gSetShow pbcEvt, slStr, tmCtrls(ilBoxNo)
    End Select
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

    sgDoneMsg = Trim$(str$(igETypeCallSource)) & "\" & sgETypeName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload EType
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
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = INUSEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imInUse = 0 Then
            slStr = "Yes"
        ElseIf imInUse = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes Or No must be specified for In Use", tmCtrls(INUSEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = INUSEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TIMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcTime, "", "Time format must be specified", tmCtrls(TIMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TIMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = LENINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcLen, "", "Length format must be specified", tmCtrls(LENINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = LENINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PROGINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcProg, "", "Program/Sponsor must be specified", tmCtrls(PROGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PROGINDEX
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
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcEvt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (ilBox > INUSEINDEX) And (Asc(tmEtf.sType) >= Asc("A")) And (Asc(tmEtf.sType) <= Asc("D")) Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If (ilBox = INUSEINDEX) And (Asc(tmEtf.sType) >= Asc("3")) And (Asc(tmEtf.sType) <= Asc("5")) And (tgSpf.sUsingBBs <> "Y") Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If ilBox = NAMEINDEX Then
                    If (tmEtf.sType <> "Y") And (imSelectedIndex <> 0) Then
                        Beep
                        mSetFocus imBoxNo
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
Private Sub pbcEvt_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (ilBox > INUSEINDEX) And (Asc(tmEtf.sType) >= Asc("A")) And (Asc(tmEtf.sType) <= Asc("D")) Then
            Exit For
        End If
        pbcEvt.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcEvt.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcEvt.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcInUse_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcInUse_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imInUse <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imInUse = 0
        pbcInUse_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imInUse <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imInUse = 1
        pbcInUse_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imInUse = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imInUse = 1
            pbcInUse_Paint
        ElseIf imInUse = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imInUse = 0
            pbcInUse_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcInUse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imInUse = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imInUse = 1
    ElseIf imInUse = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imInUse = 0
    End If
    pbcInUse_Paint
    mSetCommands
End Sub
Private Sub pbcInUse_Paint()
    pbcInUse.Cls
    pbcInUse.CurrentX = fgBoxInsetX
    pbcInUse.CurrentY = 0 'fgBoxInsetY
    If imInUse = 0 Then
        pbcInUse.Print "Yes"
    ElseIf imInUse = 1 Then
        pbcInUse.Print "No"
    Else
        pbcInUse.Print "   "
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imBoxNo
        Case -1
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = NAMEINDEX
                mSetCommands
            Else
                mSetChg 1
                If (Asc(tmEtf.sType) >= Asc("3")) And (Asc(tmEtf.sType) <= Asc("5")) And (tgSpf.sUsingBBs <> "Y") Then
                    ilBox = TIMEINDEX
                Else
                    ilBox = INUSEINDEX
                End If
            End If
        Case NAMEINDEX 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case INUSEINDEX
            If (tmEtf.sType <> "Y") And (imSelectedIndex <> 0) Then
                If cbcSelect.Enabled Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                Exit Sub
            Else
                ilBox = imBoxNo - 1
            End If
        Case TIMEINDEX
            If (Asc(tmEtf.sType) >= Asc("3")) And (Asc(tmEtf.sType) <= Asc("5")) And (tgSpf.sUsingBBs <> "Y") Then
                If (tmEtf.sType <> "Y") And (imSelectedIndex <> 0) Then
                    If cbcSelect.Enabled Then
                        mSetShow imBoxNo
                        imBoxNo = -1
                        cbcSelect.SetFocus
                        Exit Sub
                    End If
                    Exit Sub
                Else
                    ilBox = NAMEINDEX
                End If
            Else
                ilBox = INUSEINDEX
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
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imBoxNo
        Case -1
            imTabDirection = -1  'Set-Right to left
            If (Asc(tmEtf.sType) >= Asc("A")) And (Asc(tmEtf.sType) <= Asc("D")) Then
                ilBox = INUSEINDEX
            Else
                ilBox = UBound(tmCtrls)
            End If
        Case NAMEINDEX
            If (Asc(tmEtf.sType) >= Asc("3")) And (Asc(tmEtf.sType) <= Asc("5")) And (tgSpf.sUsingBBs <> "Y") Then
                ilBox = TIMEINDEX
            Else
                ilBox = INUSEINDEX
            End If
        Case INUSEINDEX
            If (Asc(tmEtf.sType) >= Asc("A")) And (Asc(tmEtf.sType) <= Asc("D")) Then
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igETypeCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Else
                ilBox = imBoxNo + 1
            End If
        Case UBound(tmCtrls) 'Selling or Airing (last control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igETypeCallSource = CALLNONE) Then
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
Private Sub plcEvt_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
