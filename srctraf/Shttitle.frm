VERSION 5.00
Begin VB.Form ShtTitle 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1950
   ClientLeft      =   1620
   ClientTop       =   3480
   ClientWidth     =   6570
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
   ScaleHeight     =   1950
   ScaleWidth      =   6570
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   195
      Top             =   1440
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
      Left            =   5880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   525
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
      Left            =   5910
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   795
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
      Left            =   5910
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1140
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
      ScaleWidth      =   75
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1770
      Width           =   75
   End
   Begin VB.TextBox edcProd 
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   2805
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
      Height          =   90
      Left            =   1530
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   735
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
      Left            =   1470
      ScaleHeight     =   90
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   1170
      Width           =   75
   End
   Begin VB.PictureBox pbcShtTitle 
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
      Height          =   375
      Left            =   1770
      Picture         =   "Shttitle.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   2850
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1650
      TabIndex        =   0
      Top             =   -30
      Width           =   1650
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1050
      TabIndex        =   9
      Top             =   1470
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Height          =   285
      Left            =   4335
      TabIndex        =   12
      Top             =   1470
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Top             =   1470
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2145
      TabIndex        =   10
      Top             =   1470
      Width           =   945
   End
   Begin VB.PictureBox plcShtTitle 
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
      Height          =   495
      Left            =   1725
      ScaleHeight     =   435
      ScaleWidth      =   2895
      TabIndex        =   5
      Top             =   780
      Width           =   2955
   End
   Begin VB.PictureBox plcSelect 
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   75
      ScaleHeight     =   345
      ScaleWidth      =   6180
      TabIndex        =   1
      Top             =   210
      Width           =   6240
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
         Left            =   3360
         TabIndex        =   3
         Top             =   0
         Width           =   2805
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
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Top             =   15
         Width           =   2805
      End
   End
End
Attribute VB_Name = "ShtTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Shttitle.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ShtTitle.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Advertiser Short Title input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 1)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmShtTitleCode() As SORTCODE
Dim smShtTitleCodeTag As String
Dim smInitAdvt As String
Dim smInitShtTitle As String
Dim tmSif As SIF        'sif record image
Dim tmSifSrchKey As LONGKEY0    'Sif key record image
Dim hmSif As Integer    'Advertiser Short Title file handle
Dim imSifRecLen As Integer        'ADF record length
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
Dim imBoxNo As Integer   'Current Media Box
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocusName As Integer
Dim imFirstFocusShtTitle As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imAdvtSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imComboBoxIndex As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records
Const NAMEINDEX = 1     'Name control/field
Private Sub cbcAdvt_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcAdvt.Text <> "" Then
            gManLookAhead cbcAdvt, imBSMode, imComboBoxIndex
        End If
        imAdvtSelectedIndex = cbcAdvt.ListIndex
        pbcShtTitle.Cls
        mClearCtrlFields
        cbcSelect.Clear 'Force population
        'lbcShtTitleCode.Tag = ""
        smShtTitleCodeTag = ""
        imChgMode = False
    End If
    mSetChg imBoxNo
End Sub
Private Sub cbcAdvt_Click()
    cbcAdvt_Change
End Sub
Private Sub cbcAdvt_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocusName Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusName = False
        If igAdvtProdCallSource <> CALLNONE Then  'If from advertiser or contract- set name and branch to control
            If smInitAdvt = "" Then
                cbcAdvt.ListIndex = 0
            Else
                cbcAdvt.Text = smInitAdvt    'Name
            End If
            cbcAdvt_Change
            If smInitAdvt <> "" Then
                mSetCommands
                gFindMatch smInitAdvt, 0, cbcAdvt
                If gLastFound(cbcAdvt) >= 0 Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
            End If
            Exit Sub
        End If
    End If
    slSvText = cbcAdvt.Text
'    ilSvIndex = cbcAdvt.ListIndex
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    If cbcAdvt.ListCount <= 0 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        cmcDone.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus ActiveControl
    If slSvText = "" Then
        cbcAdvt.ListIndex = 0
        cbcAdvt_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcAdvt
        If gLastFound(cbcAdvt) > 0 Then
'            If (ilSvIndex <> (gLastFound(cbcAdvt)) Or (ilSvIndex <> cbcAdvt.ListIndex) Then
            If (slSvText <> cbcAdvt.List(gLastFound(cbcAdvt))) Then
                cbcAdvt.ListIndex = gLastFound(cbcAdvt)
                cbcAdvt_Change    'Call change so picture area repainted
            End If
        Else
            cbcAdvt.ListIndex = 0
            mClearCtrlFields
            cbcAdvt_Change    'Call change so picture area repainted
        End If
    End If
End Sub
Private Sub cbcAdvt_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcAdvt_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcAdvt.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
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
            If cbcSelect.ListCount > 0 Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.ListIndex = -1
            End If
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcShtTitle.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcProd.Text = slStr
        End If
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcShtTitle_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
    mSetCommands
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
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
    tmcClick.Interval = 300 'Delay processing encase double click
    tmcClick.Enabled = True
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If

    mSetShow imBoxNo
    imBoxNo = -1
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocusShtTitle Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusShtTitle = False
        If igAdvtProdCallSource <> CALLNONE Then  'If from advt or contract- set name and branch to control
            If smInitShtTitle = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = smInitShtTitle    'New name
            End If
            cbcSelect_Change
            If smInitShtTitle <> "" Then
                'mSetCommands
                gFindMatch smInitShtTitle, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cbcSelect.ListIndex = gLastFound(cbcSelect)
                    cmcDone.SetFocus
                    mSetCommands
                    Exit Sub
                End If
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            mSetCommands
            Exit Sub
        End If
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
    gCtrlGotFocus ActiveControl
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
    If igAdvtProdCallSource <> CALLNONE Then
        If igAdvtProdCallSource = CALLSOURCEADVERTISER Then
            igAdvtProdCallSource = CALLCANCELLED
        End If
        If igAdvtProdCallSource = CALLSOURCECONTRACT Then
            igAdvtProdCallSource = CALLCANCELLED
        End If
        If igAdvtProdCallSource = CALLSOURCECOLLECT Then
            igAdvtProdCallSource = CALLCANCELLED
        End If
        If igAdvtProdCallSource = CALLSOURCECOPY Then
            igAdvtProdCallSource = CALLCANCELLED
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
    If igAdvtProdCallSource <> CALLNONE Then
        sgAdvtProdName = edcProd.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgAdvtName = "[New]"
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
    If igAdvtProdCallSource <> CALLNONE Then
        If igAdvtProdCallSource = CALLSOURCECONTRACT Then
            If sgAdvtProdName = "[New]" Then
                igAdvtProdCallSource = CALLCANCELLED
            Else
                igAdvtProdCallSource = CALLDONE
            End If
        End If
        If igAdvtProdCallSource = CALLSOURCEADVERTISER Then
            If sgAdvtProdName = "[New]" Then
                igAdvtProdCallSource = CALLCANCELLED
            Else
                igAdvtProdCallSource = CALLDONE
            End If
        End If
        If igAdvtProdCallSource = CALLSOURCECOLLECT Then
            If sgAdvtProdName = "[New]" Then
                igAdvtProdCallSource = CALLCANCELLED
            Else
                igAdvtProdCallSource = CALLDONE
            End If
        End If
        If igAdvtProdCallSource = CALLSOURCECOPY Then
            If sgAdvtProdName = "[New]" Then
                igAdvtProdCallSource = CALLCANCELLED
            Else
                igAdvtProdCallSource = CALLDONE
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
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim slMsg As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        ilRet = gLLCodeRefExist(ShtTitle, tmSif.lCode, "Chf.Btr", "ChfSifCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gLLCodeRefExist(ShtTitle, tmSif.lCode, "Crf.Btr", "CrfSifCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Rotation references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmSif.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Sif.Btr")
        ilRet = btrDelete(hmSif)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", ShtTitle
        On Error GoTo 0
        gGetSyncDateTime slSyncDate, slSyncTime
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "SIF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmSif.iRemoteID
'            tmDsf.lAutoCode = tmSif.lAutoCode
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcShtTitleCode.Tag <> "" Then
        '    If slStamp = lbcShtTitleCode.Tag Then
        '        lbcShtTitleCode.Tag = FileDateTime(sgDBPath & "Sif.Btr")
        '    End If
        'End If
        If smShtTitleCodeTag <> "" Then
            If slStamp = smShtTitleCodeTag Then
                smShtTitleCodeTag = gFileDateTime(sgDBPath & "Sif.Btr")
            End If
        End If
        'lbcShtTitleCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tmShtTitleCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcShtTitle.Cls
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
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    slName = edcProd.Text   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    smShtTitleCodeTag = ""
    mPopulate
    imBoxNo = -1
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If imSvSelectedIndex <> 0 Then
        cbcSelect.Text = slName
'        cbcSelect_Change    'Call change so picture area repainted
    Else
        cbcSelect.ListIndex = 0
'        mClearCtrlFields 'This is required as select_change will not be generated
    End If
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    On Error Resume Next
    cbcSelect.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcProd_Change()
    mSetChg NAMEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcProd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcProd_KeyPress(KeyAscii As Integer)
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
    If (igWinStatus(ADVERTISERSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcShtTitle.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcShtTitle.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
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
    
    Erase tmShtTitleCode
    btrExtClear hmSif   'Clear any previous extend operation
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    
    Set ShtTitle = Nothing   'Remove data segment
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
    ilIndex = imAdvtSelectedIndex
    If ilIndex >= 0 Then
        slName = cbcAdvt.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(ShtTitle, cbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(ShtTitle, cbcAdvt, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", ShtTitle
        On Error GoTo 0
        If ilIndex > 1 Then
            gFindMatch slName, 2, cbcAdvt
            If gLastFound(cbcAdvt) > 1 Then
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

    edcProd.Text = ""
    tmSif.iAdfCode = 0
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
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
            edcProd.Width = tmCtrls(ilBoxNo).fBoxW
            edcProd.MaxLength = 15  'tgSpf.iAProd
            gMoveFormCtrl pbcShtTitle, edcProd, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcProd.Visible = True  'Set visibility
            edcProd.SetFocus
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
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    ShtTitle.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone ShtTitle
    'ShtTitle.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE

    imPopReqd = False
    imSifRecLen = Len(tmSif)  'Get and save PRF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imFirstFocusName = True
    imFirstFocusShtTitle = True
    imSelectedIndex = -1
    imAdvtSelectedIndex = -1
    imBypassSetting = False
    ilRet = gParseItem(sgAdvtProdName, 1, "\", smInitAdvt)
    If ilRet <> CP_MSG_NONE Then
        smInitAdvt = ""
    Else
        smInitAdvt = Trim$(smInitAdvt)
    End If
    ilRet = gParseItem(sgAdvtProdName, 2, "\", smInitShtTitle)
    If ilRet <> CP_MSG_NONE Then
        smInitShtTitle = ""
    Else
        smInitShtTitle = Trim$(smInitShtTitle)
    End If
    hmSif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sif.Btr)", ShtTitle
    On Error GoTo 0
'    hmDsf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dsf.Btr)", ShtTitle
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    cbcAdvt.Clear 'Force population
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm ShtTitle
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
    flTextHeight = pbcShtTitle.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcShtTitle.Move 1725, 780, pbcShtTitle.Width + fgPanelAdj, pbcShtTitle.Height + fgPanelAdj
    pbcShtTitle.Move plcShtTitle.Left + fgBevelX, plcShtTitle.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim slNameCode As String  'Vehicle name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'code number
    If (imAdvtSelectedIndex >= 0) And (tmSif.iAdfCode = 0) Then
        slNameCode = tgAdvertiser(imAdvtSelectedIndex).sKey    'Traffic!lbcAdvertiser.List(imAdvtSelectedIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", ShtTitle
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmSif.iAdfCode = CInt(slCode)
    End If
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmSif.sName = edcProd.Text
    End If
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
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
    edcProd.Text = Trim$(tmSif.sName)
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
    If edcProd.Text <> "" Then    'Test name
        slStr = Trim$(edcProd.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcProd.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Advertiser Short Title already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcProd.Text = Trim$(tmSif.sName) 'Reset text
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
    'gInitStdAlone ShtTitle, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igAdvtProdCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igAdvtProdCallSource = CALLNONE
    'End If
    If igAdvtProdCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgAdvtProdName = slStr
            ilRet = gParseItem(slCommand, 5, "\", slStr)
            If ilRet = CP_MSG_NONE Then
                sgAdvtProdName = sgAdvtProdName & "\" & slStr
            End If
        Else
            sgAdvtProdName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser product    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    imPopReqd = False
    If imAdvtSelectedIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tgAdvertiser(imAdvtSelectedIndex).sKey    'Traffic!lbcAdvertiser.List(imAdvtSelectedIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mPopulateErr
    gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", ShtTitle
    On Error GoTo 0
    ilCode = CInt(slCode)
    'ilRet = gPopShortTitleBox(ShtTitle, ilCode, cbcSelect, lbcShtTitleCode)
    ilRet = gPopShortTitleBox(ShtTitle, ilCode, cbcSelect, tmShtTitleCode(), smShtTitleCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopShortTitleBox)", ShtTitle
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    slNameCode = tmShtTitleCode(ilSelectIndex - 1).sKey    'lbcShtTitleCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", ShtTitle
    On Error GoTo 0
    tmSifSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Advertiser Product)", ShtTitle
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
'*             Created:5/14/93       By:D. LeVine      *
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
    Dim llRecPos As Long
    Dim slSyncDate As String
    Dim slSyncTime As String
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
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Sif.Btr")
        'If Len(lbcShtTitleCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcShtTitleCode.Tag, Len(lbcShtTitleCode.Tag) - Len(slStamp))
        'End If
        If Len(smShtTitleCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(smShtTitleCodeTag, Len(smShtTitleCodeTag) - Len(slStamp))
        End If
        If imSelectedIndex <> 0 Then
            'Reread record in so lastest is obtained
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmSif.lCode = 0
            tmSif.iRemoteID = tgUrf(0).iRemoteUserID
            tmSif.lAutoCode = tmSif.lCode
            ilRet = btrInsert(hmSif, tmSif, imSifRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Advertiser Short Title)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmSif, tmSif, imSifRecLen)
            slMsg = "mSaveRec (btrUpdate: Advertiser ShortTitle)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, ShtTitle
    On Error GoTo 0
    ilRet = btrGetPosition(hmSif, llRecPos)
    If (imSelectedIndex = 0) Then   'And (tgSpf.sRemoteUsers = "Y") Then 'New selected
        Do
            tmSif.iRemoteID = tgUrf(0).iRemoteUserID
            tmSif.lAutoCode = tmSif.lCode
            gPackDate slSyncDate, tmSif.iSyncDate(0), tmSif.iSyncDate(1)
            gPackTime slSyncTime, tmSif.iSyncTime(0), tmSif.iSyncTime(1)
            ilRet = btrUpdate(hmSif, tmSif, imSifRecLen)
            slMsg = "mSaveRec (btrUpdate:Short Title Name)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, ShtTitle
        On Error GoTo 0
    End If
    'If smShtTitleCodeTag <> "" Then
    '    If slStamp = smShtTitleCodeTag Then
    '        smShtTitleCodeTag = FileDateTime(sgDBPath & "Sif.Btr")
    '        If Len(slStamp) > Len(smShtTitleCodeTag) Then
    '            smShtTitleCodeTag = smShtTitleCodeTag & Right$(slStamp, Len(slStamp) - Len(smShtTitleCodeTag))
    '        End If
    '    End If
    'End If
    'If imSelectedIndex <> 0 Then
    '    'lbcShtTitleCode.RemoveItem imSelectedIndex - 1
    '    gRemoveItemFromSortCode imSelectedIndex - 1, tmShtTitleCode()
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = RTrim$(tmSif.sName)
    'cbcSelect.AddItem slName
    'slName = tmSif.sName + "\" + LTrim$(Str$(tmSif.lCode))'slName + "\" + LTrim$(Str$(tmSif.lCode))
    'gAddItemToSortCode slName, tmShtTitleCode(), True
    'cbcSelect.AddItem "[New]", 0
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
'*             Created:5/14/93       By:D. LeVine      *
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
                    slMess = "Add " & edcProd.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcShtTitle_Paint
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
'*             Created:5/12/93       By:D. LeVine      *
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
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmSif.sName, edcProd, tmCtrls(ilBoxNo)
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
    If (imAdvtSelectedIndex < 0) Or (cbcSelect.Text = "") Then
        cbcAdvt.Enabled = True
        If cbcAdvt.ListIndex >= 0 Then
            cbcSelect.Enabled = True
        Else
            cbcSelect.Enabled = False
        End If
        pbcShtTitle.Enabled = False  'Disallow mouse
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    Else
        If (igWinStatus(ADVERTISERSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcShtTitle.Enabled = False  'Disallow mouse
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
        Else
            pbcShtTitle.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
        End If
        If Not ilAltered Then
            cbcSelect.Enabled = True
            cbcAdvt.Enabled = True
        Else
            cbcSelect.Enabled = False
            cbcAdvt.Enabled = False
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcProd.Visible = False  'Set visibility
            slStr = edcProd.Text
            gSetShow pbcShtTitle, slStr, tmCtrls(ilBoxNo)
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
    sgDoneMsg = Trim$(str$(igAdvtProdCallSource)) & "\" & sgAdvtProdName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload ShtTitle
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
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
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcProd, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
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
Private Sub pbcShtTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If cbcSelect.ListIndex <> imSelectedIndex Then
        cbcSelect_Change
        cbcSelect.SetFocus
        Exit Sub
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcShtTitle_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcShtTitle.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcShtTitle.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcShtTitle.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If cbcSelect.ListIndex <> imSelectedIndex Then
        cbcSelect_Change
        cbcSelect.SetFocus
        Exit Sub
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
    Select Case imBoxNo
        Case -1
            ilBox = 1
            mSetCommands
        Case 1 'Name
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
    If GetFocus() <> pbcTab.hwnd Then
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
            ilBox = NAMEINDEX
        Case NAMEINDEX
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igAdvtProdCallSource = CALLNONE) Then
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
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcShtTitle_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    If cbcSelect.ListIndex <> imSelectedIndex Then
        cbcSelect_Change
        'cbcSelect.SetFocus
        Exit Sub
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Short Title"
End Sub
