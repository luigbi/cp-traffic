VERSION 5.00
Begin VB.Form EMailContent 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6240
   ClientLeft      =   1725
   ClientTop       =   4845
   ClientWidth     =   11355
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
   ScaleHeight     =   6240
   ScaleWidth      =   11355
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   390
      Top             =   5685
   End
   Begin VB.TextBox edcContent 
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
      Height          =   3720
      Left            =   1560
      MaxLength       =   2000
      TabIndex        =   5
      Top             =   1665
      Width           =   9405
   End
   Begin VB.TextBox edcContentName 
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
      Height          =   255
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
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
      Left            =   7395
      TabIndex        =   1
      Top             =   195
      Width           =   3645
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   3435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   3435
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   4005
      TabIndex        =   6
      Top             =   5760
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   6195
      TabIndex        =   8
      Top             =   5760
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5100
      TabIndex        =   7
      Top             =   5760
      Width           =   945
   End
   Begin VB.Label lacContent 
      Caption         =   "Content"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1710
      Width           =   990
   End
   Begin VB.Label lacContentName 
      Caption         =   "Content Name"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1215
      Width           =   1395
   End
End
Attribute VB_Name = "EMailContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EMailContent.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EMailContent.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Advertiser Product input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim hmEmf As Integer    'E-Mail Content file handle
Dim tmEmf As EMF        'prf record image
Dim tmEmfSrchKey As LONGKEY0    'Prf key record image
Dim bmChgd As Boolean
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocusName As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim smInitCotentName As String
Dim imComboBoxIndex As Integer
Dim imEmfRecLen As Integer        'ADF record length
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records
Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String


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
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcContentName.Text = slStr
        End If
    End If
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

    slSvText = cbcSelect.Text
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocusName Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusName = False
        If smInitCotentName = "" Then
            cbcSelect.ListIndex = 0
        Else
            cbcSelect.Text = smInitCotentName    'New name
        End If
        cbcSelect_Change
        If smInitCotentName <> "" Then
            'mSetCommands
            gFindMatch smInitCotentName, 1, cbcSelect
            If gLastFound(cbcSelect) > 0 Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cmcDone.SetFocus
                mSetCommands
                Exit Sub
            End If
        End If
        cmcCancel.SetFocus
        mSetCommands
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
    igEMailContentCallSource = CALLCANCELLED
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    sgEMailContentName = edcContentName.Text 'Save name for returning
    If mSaveRecChg(False) = False Then
        sgAdvtName = "[New]"
        If Not imTerminate Then
            Exit Sub
        Else
            cmcCancel_Click
            Exit Sub
        End If
    End If
    If sgEMailContentName = "[New]" Then
        igEMailContentCallSource = CALLCANCELLED
    Else
        igEMailContentCallSource = CALLDONE
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        If (Trim$(edcContentName.Text) = "") Then
            Beep
            edcContentName.SetFocus
            Exit Sub
        End If
        If (Trim$(edcContent.Text) = "") Then
            Beep
            edcContent.SetFocus
            Exit Sub
        End If
    End If
End Sub


Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = Trim$(edcContentName.Text)   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        Exit Sub
    End If
    mPopulate
    Sleep 500
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If imSvSelectedIndex <> 0 Then
        cbcSelect.Text = slName
    Else
        cbcSelect.Text = slName
    End If
    cbcSelect_Change    'Call change so picture area repainted
    bmChgd = False
    mSetCommands
    If cbcSelect.Enabled Then
        cbcSelect.SetFocus
    End If
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcContent_Change()
    bmChgd = True
    mSetCommands
End Sub

Private Sub edcContentName_Change()
    bmChgd = True
    mSetCommands
End Sub
Private Sub edcContentName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcContentName_KeyPress(KeyAscii As Integer)
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
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    imUpdateAllowed = True
    Me.KeyPreview = True
    EMailContent.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
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
    Erase tmNameCode
    btrExtClear hmEmf   'Clear any previous extend operation
    ilRet = btrClose(hmEmf)
    btrDestroy hmEmf
    
    Set EMailContent = Nothing   'Remove data segment
    
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

    edcContentName.Text = ""
    edcContent.Text = ""
    tmEmf.lCode = 0
    mMoveCtrlToRec False
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
    bmChgd = False
    mInitBox
    gCenterStdAlone EMailContent
    'EMailContent.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE

    imPopReqd = False
    imEmfRecLen = Len(tmEmf)  'Get and save PRF record length
    imChgMode = False
    imBSMode = False
    imSelectedIndex = -1
    imFirstFocusName = True
    imBypassSetting = False
    smInitCotentName = sgEMailContentName
    hmEmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmEmf, "", sgDBPath & "Emf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Emf.Btr)", EMailContent
    On Error GoTo 0
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        'If smInitCotentName <> "" Then
        'Else
        '    cbcSelect.ListIndex = 0
        'End If
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
    tmEmf.sName = edcContentName.Text
    tmEmf.sContent = Trim$(edcContent.Text) & Chr$(0)
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
    edcContentName.Text = Trim$(tmEmf.sName)
    edcContent.Text = gStripChr0(tmEmf.sContent)
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
    If edcContentName.Text <> "" Then    'Test name
        slStr = Trim$(edcContentName.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcContentName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Content Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcContentName.Text = Trim$(tmEmf.sName) 'Reset text
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
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer

    imPopReqd = False
    ilfilter(0) = CHARFILTER
    slFilter(0) = sgEMailContentType
    ilOffSet(0) = gFieldOffset("Emf", "EmfType") '2
    'ilRet = gLMoveListBox(BPlate, cbcSelect, lbcTitleCode, "Cmf.btr", gFieldOffset("Cmf", "CmfTitle"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gLMoveListBox(EMailContent, cbcSelect, tmNameCode(), smNameCodeTag, "Emf.btr", gFieldOffset("Emf", "EmfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gLMoveListBox)", EMailContent
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
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    slNameCode = tmNameCode(ilSelectIndex - 1).sKey    'lbcEMailContentCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", EMailContent
    On Error GoTo 0
    tmEmfSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmEmf, tmEmf, imEmfRecLen, tmEmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: EMailContent)", EMailContent
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
    If (Trim$(edcContentName.Text) = "") Then
        mSaveRec = False
        Exit Function
    End If
    If (Trim$(edcContent.Text) = "") Then
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
        slStamp = gFileDateTime(sgDBPath & "Emf.btr")
        'If Len(lbcEMailContentCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcEMailContentCode.Tag, Len(lbcEMailContentCode.Tag) - Len(slStamp))
        'End If
        If Len(smNameCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(smNameCodeTag, Len(smNameCodeTag) - Len(slStamp))
        End If
        If imSelectedIndex <> 0 Then
            'Reread record in so lastest is obtained
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        gPackDate Format$(gNow(), "m/d/yy"), tmEmf.iEnteredDate(0), tmEmf.iEnteredDate(1)
        gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tmEmf.iEnteredTime(0), tmEmf.iEnteredTime(1)
        If imSelectedIndex = 0 Then 'New selected
            tmEmf.lCode = 0
            tmEmf.sType = sgEMailContentType
            tmEmf.iUrfCode = tgUrf(0).iCode
            ilRet = btrInsert(hmEmf, tmEmf, imEmfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Advertiser Product Name)"
        Else 'Old record-Update
            tmEmf.iUrfCode = tgUrf(0).iCode
            ilRet = btrUpdate(hmEmf, tmEmf, imEmfRecLen)
            slMsg = "mSaveRec (btrUpdate: Advertiser Product Name)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, EMailContent
    On Error GoTo 0
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
    
    If (Trim$(edcContentName.Text) = "") Then
        Beep
        edcContentName.SetFocus
        mSaveRecChg = False
        Exit Function
    End If
    If (Trim$(edcContent.Text) = "") Then
        Beep
        edcContent.SetFocus
        mSaveRecChg = False
        Exit Function
    End If
    If bmChgd Then
        If ilAsk Then
            If imSelectedIndex > 0 Then
                slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
            Else
                slMess = "Add " & edcContentName.Text
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
    'Update button set if all mandatory fields have data and any field altered
    If (bmChgd) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
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
    On Error Resume Next
    sgDoneMsg = Trim$(str$(igEMailContentCallSource)) & "\" & sgEMailContentName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload EMailContent
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If cbcSelect.ListIndex <> imSelectedIndex Then
        cbcSelect_Change
        Exit Sub
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If sgEMailContentType = "C" Then
        plcScreen.Print "Contract E-Mail Content"
    ElseIf sgEMailContentType = "P" Then
        plcScreen.Print "Proposal E-Mail Content"
    Else
        plcScreen.Print "Insertion E-Mail Content"
    End If
End Sub
