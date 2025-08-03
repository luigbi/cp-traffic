VERSION 5.00
Begin VB.Form AName 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   4185
   ClientTop       =   3630
   ClientWidth     =   4710
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
   ScaleHeight     =   4470
   ScaleWidth      =   4710
   Begin VB.TextBox edcSortCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   885
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   795
      TabIndex        =   1
      Top             =   315
      Width           =   3000
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4110
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2730
      Width           =   60
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1365
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   870
      MaxLength       =   20
      TabIndex        =   4
      Top             =   945
      Visible         =   0   'False
      Width           =   2820
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
      Left            =   2940
      TabIndex        =   14
      Top             =   4110
      Width           =   1050
   End
   Begin VB.CommandButton cmcMerge 
      Appearance      =   0  'Flat
      Caption         =   "&Merge into"
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
      Left            =   3930
      TabIndex        =   13
      Top             =   3225
      Visible         =   0   'False
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
      Left            =   1815
      TabIndex        =   12
      Top             =   4110
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
      Left            =   690
      TabIndex        =   11
      Top             =   4110
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
      Left            =   2940
      TabIndex        =   10
      Top             =   3735
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
      Left            =   1815
      TabIndex        =   9
      Top             =   3735
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
      Left            =   690
      TabIndex        =   8
      Top             =   3735
      Width           =   1050
   End
   Begin VB.PictureBox pbcANm 
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
      Height          =   2790
      Left            =   855
      Picture         =   "Aname.frx":0000
      ScaleHeight     =   2790
      ScaleWidth      =   2850
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   795
      Width           =   2850
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   360
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   765
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   1695
      Width           =   15
   End
   Begin VB.PictureBox plcANm 
      ForeColor       =   &H00000000&
      Height          =   2940
      Left            =   795
      ScaleHeight     =   2880
      ScaleWidth      =   2925
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   2985
   End
   Begin VB.Label plcScreen 
      Caption         =   "Avail Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   2325
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   3690
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "AName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Aname.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AName.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Avail name input screen code
Option Explicit
Option Compare Text
'Avail Name Field Areas
Dim imFirstActivate As Integer
Dim tmCtrls(0 To 12) As FIELDAREA
Dim imLBCtrls As Integer
Dim imMaxBoxNo As Integer
Dim imBoxNo As Integer   'Current Avail Name Box
Dim tmAnf As ANF        'ANF record image
Dim tmAnfSrchKey As INTKEY0    'ANF key record image
Dim imAnfRecLen As Integer        'ANF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmAnf As Integer 'Avail name file handle
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imSustain As Integer    '0= Yes; 1= No
Dim imSponsorship As Integer    '0=Yes; 1= No
Dim imRptDefault As Integer     '0=Yes; 1=No
Dim imTrafToAff As Integer      '0=Yes; 1=No
Dim imISCIExport As Integer      '0=Yes; 1=No
Dim imAudioExport As Integer      '0=Yes; 1=No
Dim imAutoExport As Integer      '0=Yes; 1=No
Dim imEventAvailsGroup As Integer '0/Blank="Standard", 1="Billboard", 2="Drop-in", and 3="Extra" - TTP 10434 - Event and Sports export (WWO)
Dim imSpotType As Integer   '0=Local; 1=Net; 2=Both
Dim imFilledBreak As Integer    '0=Always; 1=No; 2=If Partial; 3=If Adjacent Partial
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records
Dim imSSFFixReq As Integer
Dim imSvSelectedIndex As Integer
Dim imDelaySource As Integer '0=Done, 1=Update

Const NAMEINDEX = 1     'Name control/field
Const SUSTAININDEX = 2  'Sustain control/field
Const SPONSORINDEX = 3  'Sponsorship control/field
Const SORTCODEINDEX = 4
Const RPTDEFAULTINDEX = 5
Const TRAFTOAFFINDEX = 6
Const ISCIEXPORTINDEX = 7
Const AUDIOEXPORTINDEX = 8
Const AUTOEXPORTINDEX = 9
Const FILLEDBREAKINDEX = 10
Const EVENTAVAILSGROUP = 11 'TTP 10434 - Event and Sports export (WWO)
Const SPOTTYPEINDEX = 12

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
    pbcANm.Cls
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
    For ilLoop = imLBCtrls To imMaxBoxNo Step 1   'UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcANm_Paint
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
        If igANmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgANmName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgANmName  'New Name
            End If
            cbcSelect_Change
            If sgANmName <> "" Then
                mSetCommands
                gFindMatch sgANmName, 1, cbcSelect
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
    gCtrlGotFocus cbcSelect
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
    If igANmCallSource <> CALLNONE Then
        igANmCallSource = CALLCANCELLED
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
    imDelaySource = 0
    If igANmCallSource <> CALLNONE Then
        sgANmName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgANmName = "[New]"
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
    If igANmCallSource <> CALLNONE Then
        If sgANmName = "[New]" Then
            igANmCallSource = CALLCANCELLED
        Else
            igANmCallSource = CALLDONE
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
        For ilLoop = imLBCtrls To imMaxBoxNo Step 1   'UBound(tmCtrls) Step 1
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

Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim slMsg As String
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(AName, tmAnf.iCode, "Crf.Btr", "CrfAnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Rotation references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gLICodeRefExist(AName, tmAnf.iCode, "Lef.Btr", "LefAnfCode")  'lefanfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Library Events references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(AName, tmAnf.iCode, "Lhf.Btr", "LhfAnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Log History references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(AName, tmAnf.iCode, "Lst.mkd", "lstAnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Affiliate Log Spot references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(AName, tmAnf.iCode, "Rdf.Btr", "RdfAnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Daypart references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        If tgSpf.sSystemType = "R" Then
            ilRet = gIICodeRefExist(AName, tmAnf.iCode, "Fnf.Btr", "FnfAnfCode")
            If ilRet Then
                Screen.MousePointer = vbDefault
                slMsg = "Cannot erase - a Feed Name references name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmAnf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Anf.btr")
        ilRet = btrDelete(hmAnf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", AName
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcNameCode.Tag <> "" Then
        '    If slStamp = lbcNameCode.Tag Then
        '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Anf.btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Anf.btr")
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
    pbcANm.Cls
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

Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = AVAILNAMESLIST
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "AName^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "AName^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "AName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "AName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'AName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    RptList.Show vbModal
    slStr = sgDoneMsg
    'AName.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
End Sub

Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    If imSelectedIndex > 0 Then
        ilIndex = imSelectedIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To imMaxBoxNo Step 1   'UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcANm.Cls
        pbcANm_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcANm.Cls
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
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
'    slName = edcName.Text   'Save name
    imSvSelectedIndex = imSelectedIndex
    imDelaySource = 1
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
        cbcSelect.Text = slName
    Else
        cbcSelect.ListIndex = 0
    End If
    cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmAnf.iCode
    cbcSelect.Clear
    sgNameCodeTag = ""
    mPopulate
    If (imSvSelectedIndex <> 0) Or (igANmCallSource <> CALLNONE) Then
        For ilLoop = 0 To UBound(tgNameCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            slNameCode = tgNameCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = ilCode Then
                If cbcSelect.ListIndex = ilLoop + 1 Then
                    cbcSelect_Change
                Else
                    cbcSelect.ListIndex = ilLoop + 1
                End If
                Exit For
            End If
        Next ilLoop
    Else
        cbcSelect.ListIndex = 0
    End If
    mSetCommands
    cbcSelect.SetFocus
End Sub

Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcName_Change()
    mSetChg NAMEINDEX   'can't use imBoxNo as not set when edcName set via cbcSelect- altered flag set so field is saved
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus edcName
End Sub

Private Sub edcName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    '2/3/16: Disallow forward slash
    'If Not gCheckKeyAscii(ilKey) Then
    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcName_LostFocus()
    '9760
    edcName.Text = gRemoveIllegalPastedChar(edcName.Text)
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
End Sub

Private Sub edcSortCode_Change()
    mSetChg SORTCODEINDEX   'can't use imBoxNo as not set when edcName set via cbcSelect- altered flag set so field is saved
End Sub

Private Sub edcSortCode_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSortCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
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
'    Dim ilLoop As Integer
    If (igWinStatus(AVAILNAMESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcANm.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcANm.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BT"
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
'    DoEvents
    'This loop is required to prevent a timing problem- if calling
    'with sg----- = "", then loss GotFocus to first control
    'without this loop
'    For ilLoop = 1 To igDDEDelay Step 1
'        DoEvents
'    Next ilLoop
'    gShowBranner
    Me.KeyPreview = True
    AName.Refresh
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

    btrExtClear hmAnf   'Clear any previous extend operation
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf

    Set AName = Nothing   'Remove data segment
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
    edcName.Text = ""
    imSustain = -1
    imSponsorship = -1
    edcSortCode.Text = ""
    imRptDefault = -1
    imTrafToAff = -1
    imISCIExport = -1
    imAudioExport = -1
    imAutoExport = -1
    imSpotType = -1
    imEventAvailsGroup = -1
    imFilledBreak = -1
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
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxBoxNo) Then   'UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcANm, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case SUSTAININDEX 'Allow sustaining
            'mSendHelpMsg "Allow sustaining (non-sponsorship) spots"
            If imSustain < 0 Then
                imSustain = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SPONSORINDEX 'Suppress name on contract
            'mSendHelpMsg "Allow sponsorship spots"
            If imSponsorship < 0 Then
                imSponsorship = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SORTCODEINDEX 'Name
            'mSendHelpMsg "Enter avail name"
            edcSortCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcSortCode.MaxLength = 5
            gMoveFormCtrl pbcANm, edcSortCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcSortCode.Visible = True  'Set visibility
            edcSortCode.SetFocus
        Case RPTDEFAULTINDEX 'Suppress name on contract
            'mSendHelpMsg "Allow sponsorship spots"
            If imRptDefault < 0 Then
                imRptDefault = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case TRAFTOAFFINDEX 'Traffic to Affiliate
            'mSendHelpMsg "Allow sponsorship spots"
            If imTrafToAff < 0 Then
                imTrafToAff = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case ISCIEXPORTINDEX 'Suppress name on contract
            'mSendHelpMsg "Allow sponsorship spots"
            If imISCIExport < 0 Then
                imISCIExport = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case AUDIOEXPORTINDEX 'Suppress name on contract
            'mSendHelpMsg "Allow sponsorship spots"
            If imAudioExport < 0 Then
                imAudioExport = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case AUTOEXPORTINDEX 'Suppress name on contract
            'mSendHelpMsg "Allow sponsorship spots"
            If imAutoExport < 0 Then
                imAutoExport = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        'TTP 10434 - Event and Sports export (WWO)
        Case EVENTAVAILSGROUP
            If imEventAvailsGroup < 0 Then
                imEventAvailsGroup = 0
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SPOTTYPEINDEX 'Spot Type
            If imSpotType < 0 Then
                If igANmCallSource = CALLSOURCEFEED Then
                    imSpotType = 1
                Else
                    imSpotType = 2    'Both
                End If
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case FILLEDBREAKINDEX 'Filled Break
            If imFilledBreak < 0 Then
                imFilledBreak = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcANm, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
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
    AName.height = cmcReport.Top + 5 * cmcReport.height / 3
    gCenterStdAlone AName
    'AName.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imAnfRecLen = Len(tmAnf)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmAnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", AName
    On Error GoTo 0
'    gCenterModalForm AName
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
    End If
    'cbcSelect.SetFocus
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
    flTextHeight = pbcANm.TextHeight("1") - 35
    If tgSpf.sSystemType <> "R" Then
        'Hide Spot Type
        'pbcANm.Height = 1065
        'imMaxBoxNo = 5
        pbcANm.height = 2435 '2100    '1755
        imMaxBoxNo = 11
    Else
        'Show Spot Type
        'pbcANm.Height = 1410
        'imMaxBoxNo = 6
        pbcANm.height = 2790 '2445    '2100
        imMaxBoxNo = 12
    End If
    'Position panel and picture areas with panel
    plcANm.Move 795, 735, pbcANm.Width + fgPanelAdj, pbcANm.height + fgPanelAdj
    pbcANm.Move plcANm.Left + fgBevelX, plcANm.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
    'Sustaining
    gSetCtrl tmCtrls(SUSTAININDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
    'Sponsorship
    gSetCtrl tmCtrls(SPONSORINDEX), 1440, tmCtrls(SUSTAININDEX).fBoxY, 1395, fgBoxStH
    tmCtrls(SPONSORINDEX).iReq = False
    'Sort Code
    gSetCtrl tmCtrls(SORTCODEINDEX), 30, tmCtrls(SUSTAININDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
    tmCtrls(SORTCODEINDEX).iReq = False
    'Report Default
    gSetCtrl tmCtrls(RPTDEFAULTINDEX), 1440, tmCtrls(SORTCODEINDEX).fBoxY, 1395, fgBoxStH
    tmCtrls(RPTDEFAULTINDEX).iReq = False
    'Traffic to Affiliate
    gSetCtrl tmCtrls(TRAFTOAFFINDEX), 30, tmCtrls(SORTCODEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
    'ISCI Export
    gSetCtrl tmCtrls(ISCIEXPORTINDEX), 1440, tmCtrls(TRAFTOAFFINDEX).fBoxY, 1395, fgBoxStH
    'Audio Export
    gSetCtrl tmCtrls(AUDIOEXPORTINDEX), 30, tmCtrls(TRAFTOAFFINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
    'Auto Export
    gSetCtrl tmCtrls(AUTOEXPORTINDEX), 1440, tmCtrls(AUDIOEXPORTINDEX).fBoxY, 1395, fgBoxStH
    'Filled Break
    gSetCtrl tmCtrls(FILLEDBREAKINDEX), 30, tmCtrls(AUDIOEXPORTINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'Event Avails Group - TTP 10434 - Event and Sports export (WWO)
    gSetCtrl tmCtrls(EVENTAVAILSGROUP), 30, tmCtrls(FILLEDBREAKINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'Spot Type
    gSetCtrl tmCtrls(SPOTTYPEINDEX), 30, tmCtrls(EVENTAVAILSGROUP).fBoxY + fgStDeltaY, 2805, fgBoxStH
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
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmAnf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(SUSTAININDEX).iChg Then
        Select Case imSustain
            Case 0  'Yes
                If tmAnf.sSustain <> "Y" Then
                    imSSFFixReq = True
                End If
                tmAnf.sSustain = "Y"
            Case 1  'No
                If tmAnf.sSustain <> "N" Then
                    imSSFFixReq = True
                End If
                tmAnf.sSustain = "N"
            Case Else
                tmAnf.sSustain = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SPONSORINDEX).iChg Then
        Select Case imSponsorship
            Case 0  'Yes
                tmAnf.sSponsorship = "Y"
            Case 1  'No
                tmAnf.sSponsorship = "N"
            Case Else
                tmAnf.sSponsorship = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SORTCODEINDEX).iChg Then
        tmAnf.iSortCode = Val(edcSortCode.Text)
    End If
    If Not ilTestChg Or tmCtrls(RPTDEFAULTINDEX).iChg Then
        Select Case imRptDefault
            Case 0  'Yes
                tmAnf.sRptDefault = "Y"
            Case 1  'No
                tmAnf.sRptDefault = "N"
            Case Else
                tmAnf.sRptDefault = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(TRAFTOAFFINDEX).iChg Then
        Select Case imTrafToAff
            Case 0  'Yes
                tmAnf.sTrafToAff = "Y"
            Case 1  'No
                tmAnf.sTrafToAff = "N"
            Case Else
                tmAnf.sTrafToAff = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(ISCIEXPORTINDEX).iChg Then
        Select Case imISCIExport
            Case 0  'Yes
                tmAnf.sISCIExport = "Y"
            Case 1  'No
                tmAnf.sISCIExport = "N"
            Case Else
                tmAnf.sISCIExport = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(AUDIOEXPORTINDEX).iChg Then
        Select Case imAudioExport
            Case 0  'Yes
                tmAnf.sAudioExport = "Y"
            Case 1  'No
                tmAnf.sAudioExport = "N"
            Case Else
                tmAnf.sAudioExport = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(AUTOEXPORTINDEX).iChg Then
        Select Case imAutoExport
            Case 0  'Yes
                tmAnf.sAutomationExport = "Y"
            Case 1  'No
                tmAnf.sAutomationExport = "N"
            Case Else
                tmAnf.sAutomationExport = ""
        End Select
    End If
    'TTP 10434 - Event and Sports export (WWO)
    If Not ilTestChg Or tmCtrls(EVENTAVAILSGROUP).iChg Then
        Select Case imEventAvailsGroup
            Case 0  '0="Standard",
                tmAnf.iEventAvailsGroup = 0
            Case 1  '1="Billboard",
                tmAnf.iEventAvailsGroup = 1
            Case 2  ' 2="Drop-in",
                tmAnf.iEventAvailsGroup = 2
            Case 3  'I3="Extra"
                tmAnf.iEventAvailsGroup = 3
            Case Else
                tmAnf.iEventAvailsGroup = 0
        End Select
    End If
    If tgSpf.sSystemType = "R" Then
        If Not ilTestChg Or tmCtrls(SPOTTYPEINDEX).iChg Then
            Select Case imSpotType
                Case 0  'Yes
                    If tmAnf.sBookLocalFeed <> "L" Then
                        imSSFFixReq = True
                    End If
                    tmAnf.sBookLocalFeed = "L"
                Case 1  'No
                    If tmAnf.sBookLocalFeed <> "F" Then
                        imSSFFixReq = True
                    End If
                    tmAnf.sBookLocalFeed = "F"
                Case Else
                    If (tmAnf.sBookLocalFeed <> "B") And (Trim$(tmAnf.sBookLocalFeed) <> "") Then
                        imSSFFixReq = True
                    End If
                    tmAnf.sBookLocalFeed = "B"
            End Select
        End If
    Else
        tmAnf.sBookLocalFeed = "B"
    End If
    If igANmCallSource = CALLSOURCEFEED Then
        tmAnf.sBookLocalFeed = "F"
    End If
    If Not ilTestChg Or tmCtrls(FILLEDBREAKINDEX).iChg Then
        Select Case imFilledBreak
            Case 0  'Yes
                tmAnf.sFillRequired = "Y"
            Case 1  'No
                tmAnf.sFillRequired = "N"
            Case 2  'If Partial
                tmAnf.sFillRequired = "B"
            Case 3  'If Ajacent Partial
                tmAnf.sFillRequired = "A"
            Case Else
                tmAnf.sFillRequired = "Y"
        End Select
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
    edcName.Text = Trim$(tmAnf.sName)
    Select Case tmAnf.sSustain
        Case "Y"
            imSustain = 0
        Case "N"
            imSustain = 1
        Case Else
            imSustain = -1
    End Select
    Select Case tmAnf.sSponsorship
        Case "Y"
            imSponsorship = 0
        Case "N"
            imSponsorship = 1
        Case Else
            imSponsorship = -1
    End Select
    edcSortCode.Text = Trim$(str$(tmAnf.iSortCode))
    Select Case tmAnf.sRptDefault
        Case "Y"
            imRptDefault = 0
        Case "N"
            imRptDefault = 1
        Case Else
            imRptDefault = -1
    End Select
    Select Case tmAnf.sBookLocalFeed
        Case "L"
            imSpotType = 0
        Case "F"
            imSpotType = 1
        Case "B"
            imSpotType = 2
        Case Else
            imSpotType = -1
    End Select
    Select Case tmAnf.sTrafToAff
        Case "Y"
            imTrafToAff = 0
        Case "N"
            imTrafToAff = 1
        Case Else
            imTrafToAff = -1
    End Select
    Select Case tmAnf.sISCIExport
        Case "Y"
            imISCIExport = 0
        Case "N"
            imISCIExport = 1
        Case Else
            imISCIExport = -1
    End Select
    Select Case tmAnf.sAudioExport
        Case "Y"
            imAudioExport = 0
        Case "N"
            imAudioExport = 1
        Case Else
            imAudioExport = -1
    End Select
    Select Case tmAnf.sAutomationExport
        Case "Y"
            imAutoExport = 0
        Case "N"
            imAutoExport = 1
        Case Else
            imAutoExport = -1
    End Select
    If igANmCallSource = CALLSOURCEFEED Then
        imSpotType = 1
    End If
    Select Case tmAnf.sFillRequired
        Case "Y"
            imFilledBreak = 0
        Case "N"
            imFilledBreak = 1
        Case "B"
            imFilledBreak = 2
        Case "A"
            imFilledBreak = 3
        Case Else
            imFilledBreak = -1
    End Select
    imEventAvailsGroup = tmAnf.iEventAvailsGroup
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
    If edcName.Text <> "" Then    'Test name
        slStr = Trim$(edcName.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Avail Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmAnf.sName) 'Reset text
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
    'gInitStdAlone AName, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igANmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igANmCallSource = CALLNONE
    'End If
    If igANmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgANmName = slStr
        Else
            sgANmName = ""
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
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLp As Integer
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer

    imPopReqd = False
    If igANmCallSource = CALLSOURCEFEED Then
        ilFilter(0) = CHARFILTER
        slFilter(0) = "F"
        ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    Else
        ilFilter(0) = NOFILTER
        slFilter(0) = ""
        ilOffSet(0) = 0
    End If
    'ilRet = gIMoveListBox(AName, cbcSelect, lbcNameCode, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(AName, cbcSelect, tgNameCode(), sgNameCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To cbcSelect.ListCount - 1 Step 1
            slStr = Trim$(cbcSelect.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                cbcSelect.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tgNameCode) - 1 Step 1
                    tgNameCode(ilLp) = tgNameCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tgNameCode(LBound(tgNameCode) To UBound(tgNameCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", AName
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
'   iRet = ENmRead(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tgNameCode(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", AName
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmAnfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", AName
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
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilQAsked As Integer

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
    imSSFFixReq = False
    ilQAsked = False
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Anf.btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        'If (tgSpf.sSystemType = "R") And (imSSFFixReq) And (imSelectedIndex > 0) And (Not ilQAsked) Then
        If (imSSFFixReq) And (imSelectedIndex > 0) And (Not ilQAsked) Then
            ilQAsked = True
            sgGenMsg = "Indicate how Spots between Todays Date and Last Log Date that Violate the Avail changes should be Processed"
            sgCMCTitle(0) = "Pre-Empt"
            sgCMCTitle(1) = "Retain"
            sgCMCTitle(2) = "Cancel"
            igDefCMC = 0
            igEditBox = 0
            sgEditValue = ""
            GenMsg.Show vbModal
            If igAnsCMC = 2 Then
                mSaveRec = False
                Exit Function
            End If
        End If
        If imSelectedIndex = 0 Then 'New selected
            tmAnf.iCode = 0  'Autoincrement
            tmAnf.iMerge = 0
            tmAnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmAnf.iAutoCode = tmAnf.iCode
            ilRet = btrInsert(hmAnf, tmAnf, imAnfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmAnf, tmAnf, imAnfRecLen)
            slMsg = "mSaveRec (btr(Update)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, AName
    On Error GoTo 0
    If (imSelectedIndex = 0) Then   'And (tgSpf.sRemoteUsers = "Y") Then 'New selected
        Do
            'tmAnfSrchKey.iCode = tmAnf.iCode
            'ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Avail Name)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, AName
            'On Error GoTo 0
            tmAnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmAnf.iAutoCode = tmAnf.iCode
            ilRet = btrUpdate(hmAnf, tmAnf, imAnfRecLen)
            slMsg = "mSaveRec (btrUpdate:Avail Name)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, AName
        On Error GoTo 0
    End If
'    'If lbcNameCode.Tag <> "" Then
'    '    If slStamp = lbcNameCode.Tag Then
'    '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Anf.btr")
'    '    End If
'    'End If
'    If sgNameCodeTag <> "" Then
'        If slStamp = sgNameCodeTag Then
'            sgNameCodeTag = FileDateTime(sgDBPath & "Anf.btr")
'        End If
'    End If
'    If imSelectedIndex <> 0 Then
'        'lbcNameCode.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
'        cbcSelect.RemoveItem imSelectedIndex
'    End If
'    cbcSelect.RemoveItem 0 'Remove [New]
'    slName = RTrim$(tmAnf.sName)
'    cbcSelect.AddItem slName
'    slName = tmAnf.sName + "\" + LTrim$(Str$(tmAnf.iCode)) 'slName + "\" + LTrim$(Str$(tmAnf.iCode))
'    'lbcNameCode.AddItem slName
'    gAddItemToSortCode slName, tgNameCode(), True
'    cbcSelect.AddItem "[New]", 0
    If ilQAsked Then
        sgGenMsg = sgLF & sgCR & "Scheduling Avail Changes..."
        igDefCMC = 2
        GenSch.Show vbModal
    End If
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
'*             Created:4/22/93       By:D. LeVine      *
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
                    pbcANm_Paint
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
    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxBoxNo Then   'UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmAnf.sName, edcName, tmCtrls(ilBoxNo)
        Case SUSTAININDEX 'Sustain
        Case SPONSORINDEX 'Selling or Airing or N/A
        Case SORTCODEINDEX 'Name
            slStr = Trim$(str$(tmAnf.iSortCode))
            gSetChgFlag slStr, edcSortCode, tmCtrls(ilBoxNo)
        Case RPTDEFAULTINDEX 'Selling or Airing or N/A
        Case TRAFTOAFFINDEX 'Selling or Airing or N/A
        Case ISCIEXPORTINDEX 'Selling or Airing or N/A
        Case AUDIOEXPORTINDEX 'Selling or Airing or N/A
        Case AUTOEXPORTINDEX 'Selling or Airing or N/A
        Case EVENTAVAILSGROUP
        Case SPOTTYPEINDEX
        Case FILLEDBREAKINDEX
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
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) And (imUpdateAllowed) Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
    '9/12/16: Removed Merge button as no support code added to Merge.Frm
    'Merge set only if change mode
    'If (imSelectedIndex > 0) And (tgUrf(0).sMerge = "I") And (imUpdateAllowed) Then
    '    cmcMerge.Enabled = True
    'Else
    '    cmcMerge.Enabled = False
    'End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocusx                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxBoxNo) Then   'UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case SUSTAININDEX 'Allow sustaining
            pbcYN.SetFocus
        Case SPONSORINDEX 'Suppress name on contract
            pbcYN.SetFocus
        Case SORTCODEINDEX 'Name
            edcSortCode.SetFocus
        Case RPTDEFAULTINDEX 'Suppress name on contract
            pbcYN.SetFocus
        Case TRAFTOAFFINDEX 'Suppress name on contract
            pbcYN.SetFocus
        Case ISCIEXPORTINDEX 'Suppress name on contract
            pbcYN.SetFocus
        Case AUDIOEXPORTINDEX 'Suppress name on contract
            pbcYN.SetFocus
        Case AUTOEXPORTINDEX 'Suppress name on contract
            pbcYN.SetFocus
        Case EVENTAVAILSGROUP
            pbcYN.SetFocus
        Case SPOTTYPEINDEX 'Spot Type
            pbcYN.SetFocus
        Case FILLEDBREAKINDEX
            pbcYN.SetFocus
    End Select
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxBoxNo) Then   'UBound(tmCtrls)) Then
        Exit Sub
    End If

    '2/4/16: Add filter to handle the case where the name has illegal characters and it was pasted into the field
    If (ilBoxNo = NAMEINDEX) Then
        slStr = gReplaceIllegalCharacters(edcName.Text)
        edcName.Text = slStr
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case SUSTAININDEX 'Sustaining
            pbcYN.Visible = False  'Set visibility
            If imSustain = 0 Then
                slStr = "Yes"
            ElseIf imSustain = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case SPONSORINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imSponsorship = 0 Then
                slStr = "Yes"
            ElseIf imSponsorship = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case SORTCODEINDEX 'Name
            edcSortCode.Visible = False  'Set visibility
            slStr = edcSortCode.Text
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case RPTDEFAULTINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imRptDefault = 0 Then
                slStr = "Yes"
            ElseIf imRptDefault = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case TRAFTOAFFINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imTrafToAff = 0 Then
                slStr = "Yes"
            ElseIf imTrafToAff = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case ISCIEXPORTINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imISCIExport = 0 Then
                slStr = "Yes"
            ElseIf imISCIExport = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case AUDIOEXPORTINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imAudioExport = 0 Then
                slStr = "Yes"
            ElseIf imAudioExport = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case AUTOEXPORTINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imAutoExport = 0 Then
                slStr = "Yes"
            ElseIf imAutoExport = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        'TTP 10434 - Event and Sports export (WWO)
        Case EVENTAVAILSGROUP
            pbcYN.Visible = False  'Set visibility
            If imEventAvailsGroup = 0 Then
                slStr = "Standard"
            ElseIf imEventAvailsGroup = 1 Then
                slStr = "Billboard"
            ElseIf imEventAvailsGroup = 2 Then
                slStr = "Drop-in"
            ElseIf imEventAvailsGroup = 3 Then
                slStr = "Extra"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case SPOTTYPEINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imSpotType = 0 Then
                slStr = "Contract Spots Only"
            ElseIf imSpotType = 1 Then
                slStr = "Specific Feed Spots Only"
            ElseIf imSpotType = 2 Then
                slStr = "Contract & Non-Specific Feed Spots"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
        Case FILLEDBREAKINDEX 'Suppress
            pbcYN.Visible = False  'Set visibility
            If imFilledBreak = 0 Then
                slStr = "Always"
            ElseIf imFilledBreak = 1 Then
                slStr = "No"
            ElseIf imFilledBreak = 2 Then
                slStr = "If Partial"
            ElseIf imFilledBreak = 3 Then
                slStr = "If Adjacent Partial"
            Else
                slStr = ""
            End If
            gSetShow pbcANm, slStr, tmCtrls(ilBoxNo)
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
    sgDoneMsg = Trim$(str$(igANmCallSource)) & "\" & sgANmName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload AName
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
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SUSTAININDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSustain = 0 Then
            slStr = "Yes"
        ElseIf imSustain = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes Or No must be specified for Allow Sustaining", tmCtrls(SUSTAININDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SUSTAININDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SPONSORINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSponsorship = 0 Then
            slStr = "Yes"
        ElseIf imSponsorship = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes or No must be specified for Sponsorship", tmCtrls(SPONSORINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SPONSORINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SORTCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcSortCode, "", "Sort Code must be specified", tmCtrls(SORTCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SORTCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = RPTDEFAULTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imRptDefault = 0 Then
            slStr = "Yes"
        ElseIf imRptDefault = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes or No must be specified for Report Default", tmCtrls(RPTDEFAULTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = RPTDEFAULTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TRAFTOAFFINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imTrafToAff = 0 Then
            slStr = "Yes"
        ElseIf imTrafToAff = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes or No must be specified for Affiliate System", tmCtrls(TRAFTOAFFINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TRAFTOAFFINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ISCIEXPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imISCIExport = 0 Then
            slStr = "Yes"
        ElseIf imISCIExport = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes or No must be specified for ISCI Export", tmCtrls(ISCIEXPORTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ISCIEXPORTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = AUDIOEXPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imAudioExport = 0 Then
            slStr = "Yes"
        ElseIf imAudioExport = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes or No must be specified for Audio Delivery", tmCtrls(AUDIOEXPORTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = AUDIOEXPORTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = AUTOEXPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imAutoExport = 0 Then
            slStr = "Yes"
        ElseIf imAutoExport = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "A Yes or No must be specified for Automation Export", tmCtrls(AUTOEXPORTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = AUTOEXPORTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    'TTP 10434 - Event and Sports export (WWO)
    If (ilCtrlNo = EVENTAVAILSGROUP) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imEventAvailsGroup = 0 Then
            slStr = "Standard"
        ElseIf imEventAvailsGroup = 1 Then
            slStr = "Billboard"
        ElseIf imEventAvailsGroup = 2 Then
            slStr = "Drop-in"
        ElseIf imEventAvailsGroup = 3 Then
            slStr = "Extra"
        Else
            slStr = ""
        End If
    End If
    If tgSpf.sSystemType = "R" Then
        If (ilCtrlNo = SPOTTYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If imSpotType = 0 Then
                slStr = "Contract Spots Only"
            ElseIf imSpotType = 1 Then
                slStr = "Specific Feed Spots Only"
            ElseIf imSpotType = 2 Then
                slStr = "Contract & Non-Specific Feed Spots"
            Else
                slStr = ""
            End If
            If gFieldDefinedStr(slStr, "", "Spot Type must be specified", tmCtrls(SPOTTYPEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SPOTTYPEINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = FILLEDBREAKINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imFilledBreak = 0 Then
            slStr = "Always"
        ElseIf imFilledBreak = 1 Then
            slStr = "No"
        ElseIf imFilledBreak = 2 Then
            slStr = "If Partial"
        ElseIf imFilledBreak = 3 Then
            slStr = "If Adjacent Partial"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Filled Break must be specified", tmCtrls(FILLEDBREAKINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = FILLEDBREAKINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function

Private Sub pbcANm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To imMaxBoxNo Step 1    'UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub

Private Sub pbcANm_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To imMaxBoxNo Step 1    'UBound(tmCtrls) Step 1
        pbcANm.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcANm.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcANm.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub

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

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxBoxNo) Then    'UBound(tmCtrls)) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
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
                mSetChg 1
                ilBox = 2
            End If
        Case 1 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxBoxNo) Then 'UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1
            ilBox = imMaxBoxNo  'UBound(tmCtrls)
        Case imMaxBoxNo 'UBound(tmCtrls) 'Suppress (last control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igANmCallSource = CALLNONE) Then
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
    If imBoxNo = SUSTAININDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imSustain <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSustain = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imSustain <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSustain = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSustain = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSustain = 1
                pbcYN_Paint
            ElseIf imSustain = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSustain = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = SPONSORINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imSponsorship <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSponsorship = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imSponsorship <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSponsorship = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSponsorship = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSponsorship = 1
                pbcYN_Paint
            ElseIf imSponsorship = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSponsorship = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = RPTDEFAULTINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imRptDefault <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imRptDefault = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imRptDefault <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imRptDefault = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imRptDefault = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imRptDefault = 1
                pbcYN_Paint
            ElseIf imRptDefault = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imRptDefault = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = TRAFTOAFFINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imTrafToAff <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTrafToAff = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imTrafToAff <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTrafToAff = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imTrafToAff = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imTrafToAff = 1
                pbcYN_Paint
            ElseIf imTrafToAff = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imTrafToAff = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = ISCIEXPORTINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imISCIExport <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imISCIExport = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imISCIExport <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imISCIExport = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imISCIExport = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imISCIExport = 1
                pbcYN_Paint
            ElseIf imISCIExport = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imISCIExport = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = AUDIOEXPORTINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imAudioExport <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imAudioExport = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imAudioExport <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imAudioExport = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imAudioExport = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imAudioExport = 1
                pbcYN_Paint
            ElseIf imAudioExport = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imAudioExport = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = AUTOEXPORTINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imAutoExport <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imAutoExport = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imAutoExport <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imAutoExport = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imAutoExport = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imAutoExport = 1
                pbcYN_Paint
            ElseIf imAutoExport = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imAutoExport = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    'TTP 10434 - Event and Sports export (WWO)
    ElseIf imBoxNo = EVENTAVAILSGROUP Then
        If KeyAscii = Asc(" ") Then
            If imEventAvailsGroup = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imEventAvailsGroup = 1
                pbcYN_Paint
            ElseIf imEventAvailsGroup = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imEventAvailsGroup = 2
                pbcYN_Paint
            ElseIf imEventAvailsGroup = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imEventAvailsGroup = 3
                pbcYN_Paint
            ElseIf imEventAvailsGroup = 3 Then
                tmCtrls(imBoxNo).iChg = True
                imEventAvailsGroup = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = SPOTTYPEINDEX Then
        If igANmCallSource = CALLSOURCEFEED Then
            If imSpotType <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSpotType = 1
            pbcYN_Paint
        Else
            If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
                If imSpotType <> 0 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSpotType = 0
                pbcYN_Paint
            ElseIf KeyAscii = Asc("F") Or (KeyAscii = Asc("f")) Then
                If imSpotType <> 1 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSpotType = 1
                pbcYN_Paint
            ElseIf KeyAscii = Asc("B") Or (KeyAscii = Asc("b")) Then
                If imSpotType <> 2 Then
                    tmCtrls(imBoxNo).iChg = True
                End If
                imSpotType = 2
                pbcYN_Paint
            End If
            If KeyAscii = Asc(" ") Then
                If imSpotType = 0 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSpotType = 1
                    pbcYN_Paint
                ElseIf imSpotType = 1 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSpotType = 2
                    pbcYN_Paint
                ElseIf imSpotType = 2 Then
                    tmCtrls(imBoxNo).iChg = True
                    imSpotType = 0
                    pbcYN_Paint
                End If
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = FILLEDBREAKINDEX Then
        If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
            If imFilledBreak <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imFilledBreak = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imFilledBreak <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imFilledBreak = 1
            pbcYN_Paint
        ElseIf KeyAscii = Asc("I") Or (KeyAscii = Asc("i")) Then
            If imFilledBreak <> 2 Then
                tmCtrls(imBoxNo).iChg = True
                imFilledBreak = 2
            ElseIf imFilledBreak <> 3 Then
                tmCtrls(imBoxNo).iChg = True
                imFilledBreak = 3
            End If
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imFilledBreak = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imFilledBreak = 1
                pbcYN_Paint
            ElseIf imFilledBreak = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imFilledBreak = 2
                pbcYN_Paint
            ElseIf imFilledBreak = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imFilledBreak = 3
                pbcYN_Paint
            ElseIf imFilledBreak = 3 Then
                tmCtrls(imBoxNo).iChg = True
                imFilledBreak = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = SUSTAININDEX Then
        If imSustain = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSustain = 1
        ElseIf imSustain = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSustain = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = SPONSORINDEX Then
        If imSponsorship = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSponsorship = 1
        ElseIf imSponsorship = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSponsorship = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = RPTDEFAULTINDEX Then
        If imRptDefault = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imRptDefault = 1
        ElseIf imRptDefault = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imRptDefault = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = TRAFTOAFFINDEX Then
        If imTrafToAff = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imTrafToAff = 1
        ElseIf imTrafToAff = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imTrafToAff = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = ISCIEXPORTINDEX Then
        If imISCIExport = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imISCIExport = 1
        ElseIf imISCIExport = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imISCIExport = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = AUDIOEXPORTINDEX Then
        If imAudioExport = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imAudioExport = 1
        ElseIf imAudioExport = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imAudioExport = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = AUTOEXPORTINDEX Then
        If imAutoExport = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imAutoExport = 1
        ElseIf imAutoExport = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imAutoExport = 0
        End If
        pbcYN_Paint
        mSetCommands
    'TTP 10434 - Event and Sports export (WWO)
    ElseIf imBoxNo = EVENTAVAILSGROUP Then
        If imEventAvailsGroup = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imEventAvailsGroup = 1
        ElseIf imEventAvailsGroup = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imEventAvailsGroup = 2
        ElseIf imEventAvailsGroup = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imEventAvailsGroup = 3
        ElseIf imEventAvailsGroup = 3 Then
            tmCtrls(imBoxNo).iChg = True
            imEventAvailsGroup = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = SPOTTYPEINDEX Then
        If igANmCallSource = CALLSOURCEFEED Then
            If imSpotType <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSpotType = 1
        Else
            If imSpotType = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSpotType = 1
            ElseIf imSpotType = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSpotType = 2
            ElseIf imSpotType = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imSpotType = 0
            End If
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = FILLEDBREAKINDEX Then
        If imFilledBreak = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imFilledBreak = 1
        ElseIf imFilledBreak = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imFilledBreak = 2
        ElseIf imFilledBreak = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imFilledBreak = 3
        ElseIf imFilledBreak = 3 Then
            tmCtrls(imBoxNo).iChg = True
            imFilledBreak = 0
        End If
        pbcYN_Paint
        mSetCommands
   End If
End Sub

Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = SUSTAININDEX Then
        If imSustain = 0 Then
            pbcYN.Print "Yes"
        ElseIf imSustain = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = SPONSORINDEX Then
        If imSponsorship = 0 Then
            pbcYN.Print "Yes"
        ElseIf imSponsorship = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = RPTDEFAULTINDEX Then
        If imRptDefault = 0 Then
            pbcYN.Print "Yes"
        ElseIf imRptDefault = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = TRAFTOAFFINDEX Then
        If imTrafToAff = 0 Then
            pbcYN.Print "Yes"
        ElseIf imTrafToAff = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = ISCIEXPORTINDEX Then
        If imISCIExport = 0 Then
            pbcYN.Print "Yes"
        ElseIf imISCIExport = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = AUDIOEXPORTINDEX Then
        If imAudioExport = 0 Then
            pbcYN.Print "Yes"
        ElseIf imAudioExport = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = AUTOEXPORTINDEX Then
        If imAutoExport = 0 Then
            pbcYN.Print "Yes"
        ElseIf imAutoExport = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    'TTP 10434 - Event and Sports export (WWO)
    ElseIf imBoxNo = EVENTAVAILSGROUP Then
        If imEventAvailsGroup = 0 Then
            pbcYN.Print "Standard"
        ElseIf imEventAvailsGroup = 1 Then
            pbcYN.Print "Billboard"
        ElseIf imEventAvailsGroup = 2 Then
            pbcYN.Print "Drop-in"
        ElseIf imEventAvailsGroup = 3 Then
            pbcYN.Print "Extra"
        Else
            pbcYN.Print "   "
        End If
    'TTP 10434 - Event and Sports export (WWO)
    ElseIf imBoxNo = EVENTAVAILSGROUP Then
        If imEventAvailsGroup = 0 Then
            pbcYN.Print "Contract Spots Only"
        ElseIf imEventAvailsGroup = 1 Then
            pbcYN.Print "Specific Feed Spots Only"
        ElseIf imEventAvailsGroup = 2 Then
            pbcYN.Print "Contract & Non-Specific Feed Spots"
        ElseIf imEventAvailsGroup = 3 Then
            pbcYN.Print "Contract & Non-Specific Feed Spots"
        Else
            pbcYN.Print "   "
        End If
    
    ElseIf imBoxNo = SPOTTYPEINDEX Then
        If imSpotType = 0 Then
            pbcYN.Print "Contract Spots Only"
        ElseIf imSpotType = 1 Then
            pbcYN.Print "Specific Feed Spots Only"
        ElseIf imSpotType = 2 Then
            pbcYN.Print "Contract & Non-Specific Feed Spots"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = FILLEDBREAKINDEX Then
        If imFilledBreak = 0 Then
            pbcYN.Print "Always"
        ElseIf imFilledBreak = 1 Then
            pbcYN.Print "No"
        ElseIf imFilledBreak = 2 Then
            pbcYN.Print "If Partial"
        ElseIf imFilledBreak = 3 Then
            pbcYN.Print "If Adjacent Partial"
        Else
            pbcYN.Print "   "
        End If
    End If
End Sub

Private Sub plcANm_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub



