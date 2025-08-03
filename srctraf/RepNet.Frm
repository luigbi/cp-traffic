VERSION 5.00
Begin VB.Form RepNet 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
   ClientLeft      =   4185
   ClientTop       =   3630
   ClientWidth     =   6375
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
   ScaleHeight     =   5145
   ScaleWidth      =   6375
   Begin VB.TextBox edcFTPPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4980
      MaxLength       =   5
      TabIndex        =   16
      Top             =   2325
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pbcInsertionType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4560
      ScaleHeight     =   210
      ScaleWidth      =   855
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox edcIISRegSection 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   300
      MaxLength       =   80
      TabIndex        =   21
      Top             =   4050
      Visible         =   0   'False
      Width           =   5610
   End
   Begin VB.TextBox edcIISRootURL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   300
      MaxLength       =   120
      TabIndex        =   20
      Top             =   3690
      Visible         =   0   'False
      Width           =   5640
   End
   Begin VB.TextBox edcFTPExport 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   315
      MaxLength       =   120
      TabIndex        =   19
      Top             =   3345
      Visible         =   0   'False
      Width           =   5610
   End
   Begin VB.TextBox edcFTPImport 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   315
      MaxLength       =   120
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.TextBox edcFTPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   300
      MaxLength       =   20
      TabIndex        =   17
      Top             =   2670
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.TextBox edcFTPUserPW 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      IMEMode         =   3  'DISABLE
      Left            =   3135
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   2310
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox edcFTPUserID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   285
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2325
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.TextBox edcEMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   315
      MaxLength       =   80
      TabIndex        =   13
      Top             =   1965
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.TextBox edcPhoneNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3150
      MaxLength       =   25
      TabIndex        =   12
      Top             =   1635
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.TextBox edcContact 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   300
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1620
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.TextBox edcDBID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   6
      Top             =   915
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox edcSlspLast 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3165
      MaxLength       =   20
      TabIndex        =   10
      Top             =   1275
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.TextBox edcSlspFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   300
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1260
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Top             =   255
      Width           =   4005
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   135
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4875
      Width           =   60
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4995
      ScaleHeight     =   210
      ScaleWidth      =   915
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   285
      MaxLength       =   30
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   2745
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
      Left            =   3915
      TabIndex        =   25
      Top             =   4695
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
      Left            =   2670
      TabIndex        =   24
      Top             =   4695
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
      Left            =   1440
      TabIndex        =   23
      Top             =   4695
      Width           =   1050
   End
   Begin VB.PictureBox pbcRepNet 
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
      Height          =   3480
      Left            =   285
      Picture         =   "RepNet.frx":0000
      ScaleHeight     =   3480
      ScaleWidth      =   5670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   795
      Width           =   5670
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   510
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   22
      Top             =   4530
      Width           =   15
   End
   Begin VB.PictureBox plcRepNet 
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   225
      ScaleHeight     =   3555
      ScaleWidth      =   5730
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   5790
   End
   Begin VB.Label plcScreen 
      Caption         =   "Rep-Net Link"
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
      Width           =   1320
   End
End
Attribute VB_Name = "RepNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RepNet.Frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RepNet.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Avail name input screen code
Option Explicit
Option Compare Text
'Avail Name Field Areas
Dim imFirstActivate As Integer
Dim tmCtrls(0 To 17)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Avail Name Box
Dim hmNrf As Integer 'Avail name file handle
Dim tmNrf As NRF        'ANF record image
Dim tmNrfSrchKey As INTKEY0    'ANF key record image
Dim imNrfRecLen As Integer        'ANF record length
Dim tmNameCode() As SORTCODE
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imActive As Integer    '0= Yes; 1= No
Dim imTransmitPrice As Integer
Dim imInsertionType As Integer   '0=C (CSI Interface); 1=M (Manual); 2= N (neither)
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records
Dim imSvSelectedIndex As Integer
Dim imMaxBoxNo As Integer

Const NAMEINDEX = 1     'Name control/field
Const DBIDINDEX = 2  'Sustain control/field
Const INSERTIONTYPEINDEX = 3  'Sponsorship control/field
'Const ACTIVEINDEX = 4
Const SLSPFIRSTINDEX = 4
Const SLSPLASTINDEX = 5
Const CONTACTINDEX = 6
Const PHONENOINDEX = 7
Const EMAILINDEX = 8
Const TRANSMITPRICEINDEX = 9
Const FTPUSERIDINDEX = 10
Const FTPUSERPWINDEX = 11
Const FTPPORTINDEX = 12
Const FTPADDRESSINDEX = 13
Const FTPIMPORTINDEX = 14
Const FTPEXPORTINDEX = 15
Const IISROOTURLINDEX = 16
Const IISREGSECTIONINDEX = 17


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
    pbcRepNet.Cls
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
    pbcRepNet_Paint
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
        If igRNCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgRNName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgRNName  'New Name
            End If
            cbcSelect_Change
            If sgRNName <> "" Then
                mSetCommands
                gFindMatch sgRNName, 1, cbcSelect
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
    If igRNCallSource <> CALLNONE Then
        igRNCallSource = CALLCANCELLED
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
    If igRNCallSource <> CALLNONE Then
        sgRNName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgRNName = "[New]"
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
    If igRNCallSource <> CALLNONE Then
        If sgRNName = "[New]" Then
            igRNCallSource = CALLCANCELLED
        Else
            igRNCallSource = CALLDONE
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
    ilCode = tmNrf.iCode
    cbcSelect.Clear
    sgNameCodeTag = ""
    mPopulate
    If (imSvSelectedIndex <> 0) Or (igRNCallSource <> CALLNONE) Then
        For ilLoop = 0 To UBound(tmNameCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            slNameCode = tmNameCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
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

Private Sub edcContact_Change()
    mSetChg CONTACTINDEX   'can't use imBoxNo as not set when edcName set via cbcSelect- altered flag set so field is saved
End Sub

Private Sub edcContact_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDBID_Change()
    mSetChg DBIDINDEX   'can't use imBoxNo as not set when edcName set via cbcSelect- altered flag set so field is saved
End Sub

Private Sub edcDBID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEMail_Change()
    mSetChg EMAILINDEX
End Sub

Private Sub edcEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTPAddress_Change()
    mSetChg FTPADDRESSINDEX
End Sub

Private Sub edcFTPAddress_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTPExport_Change()
    mSetChg FTPEXPORTINDEX
End Sub

Private Sub edcFTPExport_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTPImport_Change()
    mSetChg FTPIMPORTINDEX
End Sub

Private Sub edcFTPImport_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTPPort_Change()
    mSetChg FTPPORTINDEX
End Sub

Private Sub edcFTPPort_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTPPort_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcFTPUserID_Change()
    mSetChg FTPUSERIDINDEX
End Sub

Private Sub edcFTPUserID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTPUserPW_Change()
    mSetChg FTPUSERPWINDEX
End Sub

Private Sub edcFTPUserPW_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcIISRegSection_Change()
    mSetChg IISREGSECTIONINDEX
End Sub

Private Sub edcIISRegSection_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcIISRootURL_Change()
    mSetChg IISROOTURLINDEX
End Sub

Private Sub edcIISRootURL_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcName_Change()
    mSetChg NAMEINDEX   'can't use imBoxNo as not set when edcName set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
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



Private Sub edcPhoneNo_Change()
    mSetChg PHONENOINDEX
End Sub

Private Sub edcPhoneNo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSlspFirst_Change()
    mSetChg SLSPFIRSTINDEX
End Sub

Private Sub edcSlspFirst_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSlspLast_Change()
    mSetChg SLSPLASTINDEX
End Sub

Private Sub edcSlspLast_GotFocus()
    gCtrlGotFocus ActiveControl
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
        pbcRepNet.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcRepNet.Enabled = True
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
    RepNet.Refresh
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

    sgNameCodeTag = ""
    Erase tmNameCode

    btrExtClear hmNrf   'Clear any previous extend operation
    ilRet = btrClose(hmNrf)
    btrDestroy hmNrf
    
    Set RepNet = Nothing   'Remove data segment

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
    imActive = -1
    imTransmitPrice = -1
    imInsertionType = -1
    edcContact.Text = ""
    edcDBID.Text = ""
    edcEMail.Text = ""
    edcFTPAddress.Text = ""
    edcFTPExport.Text = ""
    edcFTPImport.Text = ""
    edcFTPPort.Text = ""
    edcFTPUserID.Text = ""
    edcFTPUserPW.Text = ""
    edcIISRegSection.Text = ""
    edcIISRootURL.Text = ""
    edcPhoneNo.Text = ""
    edcSlspFirst.Text = ""
    edcSlspLast.Text = ""
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
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 30
            gMoveFormCtrl pbcRepNet, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case DBIDINDEX 'Name
            edcDBID.Width = tmCtrls(ilBoxNo).fBoxW
            edcDBID.MaxLength = 5
            gMoveFormCtrl pbcRepNet, edcDBID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDBID.Visible = True  'Set visibility
            edcDBID.SetFocus
        Case INSERTIONTYPEINDEX 'Suppress name on contract
            If imInsertionType < 0 Then
                imInsertionType = 0    'CSI
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcInsertionType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcRepNet, pbcInsertionType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcInsertionType_Paint
            pbcInsertionType.Visible = True
            pbcInsertionType.SetFocus
        Case TRANSMITPRICEINDEX 'Allow sustaining
            If imTransmitPrice < 0 Then
                imTransmitPrice = 0    'Yes
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcRepNet, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SLSPFIRSTINDEX 'Name
            edcSlspFirst.Width = tmCtrls(ilBoxNo).fBoxW
            edcSlspFirst.MaxLength = 20
            gMoveFormCtrl pbcRepNet, edcSlspFirst, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcSlspFirst.Visible = True  'Set visibility
            edcSlspFirst.SetFocus
        Case SLSPLASTINDEX 'Name
            edcSlspLast.Width = tmCtrls(ilBoxNo).fBoxW
            edcSlspLast.MaxLength = 20
            gMoveFormCtrl pbcRepNet, edcSlspLast, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcSlspLast.Visible = True  'Set visibility
            edcSlspLast.SetFocus
        Case CONTACTINDEX 'Name
            edcContact.Width = tmCtrls(ilBoxNo).fBoxW
            edcContact.MaxLength = 40
            gMoveFormCtrl pbcRepNet, edcContact, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcContact.Visible = True  'Set visibility
            edcContact.SetFocus
        Case PHONENOINDEX 'Name
            edcPhoneNo.Width = tmCtrls(ilBoxNo).fBoxW
            edcPhoneNo.MaxLength = 25
            gMoveFormCtrl pbcRepNet, edcPhoneNo, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcPhoneNo.Visible = True  'Set visibility
            edcPhoneNo.SetFocus
        Case EMAILINDEX 'Name
            edcEMail.Width = tmCtrls(ilBoxNo).fBoxW
            edcEMail.MaxLength = 80
            gMoveFormCtrl pbcRepNet, edcEMail, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcEMail.Visible = True  'Set visibility
            edcEMail.SetFocus
        Case FTPUSERIDINDEX 'Name
            edcFTPUserID.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTPUserID.MaxLength = 20
            gMoveFormCtrl pbcRepNet, edcFTPUserID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTPUserID.Visible = True  'Set visibility
            edcFTPUserID.SetFocus
        Case FTPUSERPWINDEX 'Name
            edcFTPUserPW.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTPUserPW.MaxLength = 20
            gMoveFormCtrl pbcRepNet, edcFTPUserPW, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTPUserPW.Visible = True  'Set visibility
            edcFTPUserPW.SetFocus
        Case FTPPORTINDEX 'Name
            edcFTPPort.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTPPort.MaxLength = 20
            gMoveFormCtrl pbcRepNet, edcFTPPort, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTPPort.Visible = True  'Set visibility
            edcFTPPort.SetFocus
        Case FTPADDRESSINDEX 'Name
            edcFTPAddress.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTPAddress.MaxLength = 20
            gMoveFormCtrl pbcRepNet, edcFTPAddress, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTPAddress.Visible = True  'Set visibility
            edcFTPAddress.SetFocus
        Case FTPIMPORTINDEX 'Name
            edcFTPImport.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTPImport.MaxLength = 120
            gMoveFormCtrl pbcRepNet, edcFTPImport, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTPImport.Visible = True  'Set visibility
            edcFTPImport.SetFocus
        Case FTPEXPORTINDEX 'Name
            edcFTPExport.Width = tmCtrls(ilBoxNo).fBoxW
            edcFTPExport.MaxLength = 120
            gMoveFormCtrl pbcRepNet, edcFTPExport, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcFTPExport.Visible = True  'Set visibility
            edcFTPExport.SetFocus
        Case IISROOTURLINDEX 'Name
            edcIISRootURL.Width = tmCtrls(ilBoxNo).fBoxW
            edcIISRootURL.MaxLength = 120
            gMoveFormCtrl pbcRepNet, edcIISRootURL, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIISRootURL.Visible = True  'Set visibility
            edcIISRootURL.SetFocus
        Case IISREGSECTIONINDEX 'Name
            edcIISRegSection.Width = tmCtrls(ilBoxNo).fBoxW
            edcIISRegSection.MaxLength = 80
            gMoveFormCtrl pbcRepNet, edcIISRegSection, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIISRegSection.Visible = True  'Set visibility
            edcIISRegSection.SetFocus
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
    RepNet.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone RepNet
    'RepNet.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imNrfRecLen = Len(tmNrf)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmNrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmNrf, "", sgDBPath & "Nrf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", RepNet
    On Error GoTo 0
'    gCenterModalForm RepNet
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
    Dim ilLoop As Integer
    Dim ilGap As Integer

    flTextHeight = pbcRepNet.TextHeight("1") - 35

    If sgRNCallType = "N" Then
        pbcRepNet.Height = 3480
        pbcRepNet.Width = 5670
        imMaxBoxNo = UBound(tmCtrls)
    ElseIf sgRNCallType = "R" Then
        pbcRepNet.Height = 375
        pbcRepNet.Width = 4260
        imMaxBoxNo = DBIDINDEX
    End If
    'Position panel and picture areas with panel
    plcRepNet.Move 225, 735, pbcRepNet.Width + fgPanelAdj, pbcRepNet.Height + fgPanelAdj
    pbcRepNet.Move plcRepNet.Left + fgBevelX, plcRepNet.Top + fgBevelY
    RepNet.Width = 2 * plcRepNet.Left + plcRepNet.Width
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
    'Database ID
    gSetCtrl tmCtrls(DBIDINDEX), 2850, tmCtrls(NAMEINDEX).fBoxY, 1395, fgBoxStH
    'Insertion Type
    gSetCtrl tmCtrls(INSERTIONTYPEINDEX), 4260, tmCtrls(NAMEINDEX).fBoxY, 1395, fgBoxStH
    ''Status
    'gSetCtrl tmCtrls(ACTIVEINDEX), 4785, tmCtrls(NAMEINDEX).fBoxY, 855, fgBoxStH
    'Salesperson First Name
    gSetCtrl tmCtrls(SLSPFIRSTINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'Salesperson Last name
    gSetCtrl tmCtrls(SLSPLASTINDEX), 2850, tmCtrls(SLSPFIRSTINDEX).fBoxY, 2790, fgBoxStH
    'Contact
    gSetCtrl tmCtrls(CONTACTINDEX), 30, tmCtrls(SLSPFIRSTINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    tmCtrls(CONTACTINDEX).iReq = False
    'Phone number
    gSetCtrl tmCtrls(PHONENOINDEX), 2850, tmCtrls(CONTACTINDEX).fBoxY, 2790, fgBoxStH
    tmCtrls(PHONENOINDEX).iReq = False
    'E-Mail
    gSetCtrl tmCtrls(EMAILINDEX), 30, tmCtrls(CONTACTINDEX).fBoxY + fgStDeltaY, 4650, fgBoxStH
    tmCtrls(EMAILINDEX).iReq = False
    'Transmit Price
    gSetCtrl tmCtrls(TRANSMITPRICEINDEX), 4695, tmCtrls(EMAILINDEX).fBoxY, 960, fgBoxStH
    'FTP User ID
    gSetCtrl tmCtrls(FTPUSERIDINDEX), 30, tmCtrls(EMAILINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    'FTP User Password
    gSetCtrl tmCtrls(FTPUSERPWINDEX), 2850, tmCtrls(FTPUSERIDINDEX).fBoxY, 1830, fgBoxStH
    'FTP Port
    gSetCtrl tmCtrls(FTPPORTINDEX), 4695, tmCtrls(FTPUSERIDINDEX).fBoxY, 960, fgBoxStH
    'FTP Address
    gSetCtrl tmCtrls(FTPADDRESSINDEX), 30, tmCtrls(FTPUSERIDINDEX).fBoxY + fgStDeltaY, 5605, fgBoxStH
    'FTP Import directory
    gSetCtrl tmCtrls(FTPIMPORTINDEX), 30, tmCtrls(FTPADDRESSINDEX).fBoxY + fgStDeltaY, 5605, fgBoxStH
    'FTP Export directory
    gSetCtrl tmCtrls(FTPEXPORTINDEX), 30, tmCtrls(FTPIMPORTINDEX).fBoxY + fgStDeltaY, 5605, fgBoxStH
    'IIS Root URL
    gSetCtrl tmCtrls(IISROOTURLINDEX), 30, tmCtrls(FTPEXPORTINDEX).fBoxY + fgStDeltaY, 5605, fgBoxStH
    'IIS Register Section
    gSetCtrl tmCtrls(IISREGSECTIONINDEX), 30, tmCtrls(IISROOTURLINDEX).fBoxY + fgStDeltaY, 5605, fgBoxStH
    If sgRNCallType = "R" Then
        For ilLoop = INSERTIONTYPEINDEX To IISREGSECTIONINDEX Step 1
            tmCtrls(ilLoop).iReq = False
        Next ilLoop
    End If
    cmcDone.Top = plcRepNet.Top + plcRepNet.Height + cmcDone.Height / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    cmcCancel.Left = RepNet.Width / 2 - cmcCancel.Width / 2
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilGap
    cbcSelect.Left = plcRepNet.Left + plcRepNet.Width - cbcSelect.Width
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
    tmNrf.sType = sgRNCallType
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmNrf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(DBIDINDEX).iChg Then
        tmNrf.sDBID = edcDBID.Text
    End If
    If Not ilTestChg Or tmCtrls(INSERTIONTYPEINDEX).iChg Then
        Select Case imInsertionType
            Case 0  'CSI
                tmNrf.sInsertionType = "C"
            Case 1  'Manual
                tmNrf.sInsertionType = "M"
            Case 2  'Neither
                tmNrf.sInsertionType = "N"
            Case Else
                tmNrf.sInsertionType = "C"
        End Select
    End If
    'If Not ilTestChg Or tmCtrls(ACTIVEINDEX).iChg Then
    '    Select Case imActive
    '        Case 0  'Yes
                tmNrf.sFTPIsOn = "Y"
    '        Case 1  'No
    '            tmNrf.sFTPIsOn = "N"
    '        Case Else
    '            tmNrf.sFTPIsOn = "Y"
    '    End Select
    'End If
    If Not ilTestChg Or tmCtrls(TRANSMITPRICEINDEX).iChg Then
        Select Case imTransmitPrice
            Case 0  'Yes
                tmNrf.sTransmitPrice = "Y"
            Case 1  'No
                tmNrf.sTransmitPrice = "N"
            Case Else
                tmNrf.sTransmitPrice = "Y"
        End Select
   End If
   If Not ilTestChg Or tmCtrls(SLSPFIRSTINDEX).iChg Then
        tmNrf.sSlspFirstName = edcSlspFirst.Text
    End If
    If Not ilTestChg Or tmCtrls(SLSPLASTINDEX).iChg Then
        tmNrf.sSlspLastName = edcSlspLast.Text
    End If
    If Not ilTestChg Or tmCtrls(CONTACTINDEX).iChg Then
        tmNrf.sContactName = edcContact.Text
    End If
    If Not ilTestChg Or tmCtrls(PHONENOINDEX).iChg Then
        tmNrf.sPhoneNo = edcPhoneNo.Text
    End If
    If Not ilTestChg Or tmCtrls(EMAILINDEX).iChg Then
        tmNrf.sEMail = edcEMail.Text
    End If
    If Not ilTestChg Or tmCtrls(FTPUSERIDINDEX).iChg Then
        tmNrf.sFTPUserID = edcFTPUserID.Text
    End If
    If Not ilTestChg Or tmCtrls(FTPUSERPWINDEX).iChg Then
        tmNrf.sFTPUserPW = edcFTPUserPW.Text
    End If
    If Not ilTestChg Or tmCtrls(FTPPORTINDEX).iChg Then
        tmNrf.iFTPPort = Val(edcFTPPort.Text)
    End If
    If Not ilTestChg Or tmCtrls(FTPADDRESSINDEX).iChg Then
        tmNrf.sFTPAddress = edcFTPAddress.Text
    End If
    If Not ilTestChg Or tmCtrls(FTPIMPORTINDEX).iChg Then
        tmNrf.sFTPImportDir = edcFTPImport.Text
    End If
    If Not ilTestChg Or tmCtrls(FTPEXPORTINDEX).iChg Then
        tmNrf.sFTPExportDir = edcFTPExport.Text
    End If
    If Not ilTestChg Or tmCtrls(IISROOTURLINDEX).iChg Then
        tmNrf.sIISRootURL = edcIISRootURL.Text
    End If
    If Not ilTestChg Or tmCtrls(IISREGSECTIONINDEX).iChg Then
        tmNrf.sIISRegSection = edcIISRegSection.Text
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
    edcName.Text = Trim$(tmNrf.sName)
    edcDBID.Text = Trim$(tmNrf.sDBID)
    Select Case tmNrf.sInsertionType
        Case "C"  'CSI
            imInsertionType = 0
        Case "M"  'Manual
            imInsertionType = 1
        Case "N"  'Neither
            imInsertionType = 2
        Case Else
            imInsertionType = 0
    End Select
    Select Case tmNrf.sFTPIsOn
        Case "Y"  'Yes
            imActive = 0
        Case "N"  'No
            imActive = 1
        Case Else
            imActive = 0
    End Select
    Select Case tmNrf.sTransmitPrice
        Case "Y"  'Yes
            imTransmitPrice = 0
        Case "N"  'No
            imTransmitPrice = 1
        Case Else
            imTransmitPrice = 0
    End Select
    edcSlspFirst.Text = Trim$(tmNrf.sSlspFirstName)
    edcSlspLast.Text = Trim$(tmNrf.sSlspLastName)
    edcContact.Text = Trim$(tmNrf.sContactName)
    edcPhoneNo.Text = Trim$(tmNrf.sPhoneNo)
    edcEMail.Text = Trim$(tmNrf.sEMail)
    edcFTPUserID.Text = Trim$(tmNrf.sFTPUserID)
    edcFTPUserPW.Text = Trim$(tmNrf.sFTPUserPW)
    edcFTPPort.Text = Trim$(tmNrf.iFTPPort)
    edcFTPAddress.Text = Trim$(tmNrf.sFTPAddress)
    edcFTPImport.Text = Trim$(tmNrf.sFTPImportDir)
    edcFTPExport.Text = Trim$(tmNrf.sFTPExportDir)
    edcIISRootURL.Text = Trim$(tmNrf.sIISRootURL)
    edcIISRegSection.Text = Trim$(tmNrf.sIISRegSection)
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
    Dim slID As String
    Dim slNameCodeDBID As String
    Dim slDBID As String
    Dim ilLoop As Integer
    Dim ilRet As Integer

    slID = Trim$(UCase$(edcDBID.Text))
    If edcName.Text <> "" Then    'Test name
        slStr = Trim$(edcName.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmNrf.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
        For ilLoop = 1 To cbcSelect.ListCount - 1 Step 1
            If ilLoop <> imSelectedIndex Then
                slNameCodeDBID = tmNameCode(ilLoop - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
                ilRet = gParseItem(slNameCodeDBID, 3, "\", slDBID)
                If Trim$(UCase$(slDBID)) = slID Then
                    Beep
                    MsgBox "Database ID already defined, enter a different ID", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmNrf.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 2
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        Next ilLoop
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
    'gInitStdAlone RepNet, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 4, "\", slStr)    'Get call source
    igRNCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igRNCallSource = CALLNONE
    'End If
    If igRNCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 5, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgRNName = slStr
        Else
            sgRNName = ""
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mPopulateErr                                                                          *
'******************************************************************************************

'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim llLen As Long

    imPopReqd = False
    ReDim tmNameCode(0 To 0) As SORTCODE
    cbcSelect.Clear
    ilRet = btrGetFirst(hmNrf, tmNrf, imNrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE)
        If tmNrf.sType = sgRNCallType Then
            tmNameCode(UBound(tmNameCode)).sKey = Trim$(tmNrf.sName) & "\" & Trim$(str$(tmNrf.iCode)) & "\" & Trim$(tmNrf.sDBID)
            ReDim Preserve tmNameCode(0 To UBound(tmNameCode) + 1) As SORTCODE
        End If
        ilRet = btrGetNext(hmNrf, tmNrf, imNrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If UBound(tmNameCode) - 1 > 0 Then
        ArraySortTyp fnAV(tmNameCode(), 0), UBound(tmNameCode), 0, LenB(tmNameCode(0)), 0, LenB(tmNameCode(0).sKey), 0
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tmNameCode) - 1 Step 1
        slNameCode = tmNameCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet = CP_MSG_NONE Then
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            cbcSelect.AddItem slName  'Add ID to list box
        End If
    Next ilLoop
    cbcSelect.AddItem "[New]", 0  'Force as first item on list
    Exit Sub
mPopulateErr: 'VBC NR
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

    slNameCode = tmNameCode(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", RepNet
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmNrfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmNrf, tmNrf, imNrfRecLen, tmNrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", RepNet
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
    ilQAsked = False
    Do  'Loop until record updated or added
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
        If imSelectedIndex = 0 Then 'New selected
            tmNrf.iCode = 0  'Autoincrement
            ilRet = btrInsert(hmNrf, tmNrf, imNrfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmNrf, tmNrf, imNrfRecLen)
            slMsg = "mSaveRec (btr(Update)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, RepNet
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
                    pbcRepNet_Paint
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > imMaxBoxNo Then   'UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmNrf.sName, edcName, tmCtrls(ilBoxNo)
        Case DBIDINDEX 'Name
            gSetChgFlag tmNrf.sDBID, edcDBID, tmCtrls(ilBoxNo)
        Case INSERTIONTYPEINDEX 'Suppress name on contract
        'Case ACTIVEINDEX 'Allow sustaining
        Case SLSPFIRSTINDEX 'Name
            gSetChgFlag tmNrf.sSlspFirstName, edcSlspFirst, tmCtrls(ilBoxNo)
        Case SLSPLASTINDEX 'Name
            gSetChgFlag tmNrf.sSlspLastName, edcSlspLast, tmCtrls(ilBoxNo)
        Case CONTACTINDEX 'Name
            gSetChgFlag tmNrf.sContactName, edcContact, tmCtrls(ilBoxNo)
        Case PHONENOINDEX 'Name
            gSetChgFlag tmNrf.sPhoneNo, edcPhoneNo, tmCtrls(ilBoxNo)
        Case EMAILINDEX 'Name
            gSetChgFlag tmNrf.sEMail, edcEMail, tmCtrls(ilBoxNo)
        Case FTPUSERIDINDEX 'Name
            gSetChgFlag tmNrf.sFTPUserID, edcFTPUserID, tmCtrls(ilBoxNo)
        Case FTPUSERPWINDEX 'Name
            gSetChgFlag tmNrf.sFTPUserPW, edcFTPUserPW, tmCtrls(ilBoxNo)
        Case FTPPORTINDEX 'Name
            gSetChgFlag Trim$(str$(tmNrf.iFTPPort)), edcFTPPort, tmCtrls(ilBoxNo)
        Case FTPADDRESSINDEX 'Name
            gSetChgFlag tmNrf.sFTPAddress, edcFTPAddress, tmCtrls(ilBoxNo)
        Case FTPIMPORTINDEX 'Name
            gSetChgFlag tmNrf.sFTPImportDir, edcFTPImport, tmCtrls(ilBoxNo)
        Case FTPEXPORTINDEX 'Name
            gSetChgFlag tmNrf.sFTPExportDir, edcFTPExport, tmCtrls(ilBoxNo)
        Case IISROOTURLINDEX 'Name
            gSetChgFlag tmNrf.sIISRootURL, edcIISRootURL, tmCtrls(ilBoxNo)
        Case IISREGSECTIONINDEX 'Name
            gSetChgFlag tmNrf.sIISRegSection, edcIISRegSection, tmCtrls(ilBoxNo)
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
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
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
        Case DBIDINDEX 'Name
            edcDBID.SetFocus
        Case INSERTIONTYPEINDEX 'Suppress name on contract
            pbcInsertionType.SetFocus
        'Case ACTIVEINDEX 'Allow sustaining
        '    pbcYN.SetFocus
        Case SLSPFIRSTINDEX 'Name
            edcSlspFirst.SetFocus
        Case SLSPLASTINDEX 'Name
            edcSlspLast.SetFocus
        Case CONTACTINDEX 'Name
            edcContact.SetFocus
        Case PHONENOINDEX 'Name
            edcPhoneNo.SetFocus
        Case EMAILINDEX 'Name
            edcEMail.SetFocus
        Case FTPUSERIDINDEX 'Name
            edcFTPUserID.SetFocus
        Case FTPUSERPWINDEX 'Name
            edcFTPUserPW.SetFocus
        Case FTPPORTINDEX 'Name
            edcFTPPort.SetFocus
        Case FTPADDRESSINDEX 'Name
            edcFTPAddress.SetFocus
        Case FTPIMPORTINDEX 'Name
            edcFTPImport.SetFocus
        Case FTPEXPORTINDEX 'Name
            edcFTPExport.SetFocus
        Case IISROOTURLINDEX 'Name
            edcIISRootURL.SetFocus
        Case IISREGSECTIONINDEX 'Name
            edcIISRegSection.SetFocus
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

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case DBIDINDEX 'Name
            edcDBID.Visible = False  'Set visibility
            slStr = edcDBID.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case INSERTIONTYPEINDEX 'Suppress
            pbcInsertionType.Visible = False  'Set visibility
            If imInsertionType = 0 Then
                slStr = "CSI Link"
            ElseIf imInsertionType = 1 Then
                slStr = "Manual"
            Else
                slStr = ""
            End If
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case TRANSMITPRICEINDEX 'Sustaining
            pbcYN.Visible = False  'Set visibility
            If imTransmitPrice = 0 Then
                slStr = "Yes"
            ElseIf imTransmitPrice = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case SLSPFIRSTINDEX 'Name
            edcSlspFirst.Visible = False  'Set visibility
            slStr = edcSlspFirst.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case SLSPLASTINDEX 'Name
            edcSlspLast.Visible = False  'Set visibility
            slStr = edcSlspLast.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case CONTACTINDEX 'Name
            edcContact.Visible = False  'Set visibility
            slStr = edcContact.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case PHONENOINDEX 'Name
            edcPhoneNo.Visible = False  'Set visibility
            slStr = edcPhoneNo.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case EMAILINDEX 'Name
            edcEMail.Visible = False  'Set visibility
            slStr = edcEMail.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case FTPUSERIDINDEX 'Name
            edcFTPUserID.Visible = False  'Set visibility
            slStr = edcFTPUserID.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case FTPUSERPWINDEX 'Name
            edcFTPUserPW.Visible = False  'Set visibility
            slStr = edcFTPUserPW.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case FTPPORTINDEX 'Name
            edcFTPPort.Visible = False  'Set visibility
            slStr = edcFTPPort.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case FTPADDRESSINDEX 'Name
            edcFTPAddress.Visible = False  'Set visibility
            slStr = edcFTPAddress.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case FTPIMPORTINDEX 'Name
            edcFTPImport.Visible = False  'Set visibility
            slStr = edcFTPImport.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case FTPEXPORTINDEX 'Name
            edcFTPExport.Visible = False  'Set visibility
            slStr = edcFTPExport.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case IISROOTURLINDEX 'Name
            edcIISRootURL.Visible = False  'Set visibility
            slStr = edcIISRootURL.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
        Case IISREGSECTIONINDEX 'Name
            edcIISRegSection.Visible = False  'Set visibility
            slStr = edcIISRegSection.Text
            gSetShow pbcRepNet, slStr, tmCtrls(ilBoxNo)
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

    sgDoneMsg = Trim$(str$(igRNCallSource)) & "\" & sgRNName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RepNet
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
    If (ilCtrlNo = DBIDINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDBID, "", "Database ID must be specified", tmCtrls(DBIDINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DBIDINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = INSERTIONTYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imInsertionType = 0 Then
            slStr = "CSI"
        ElseIf imInsertionType = 1 Then
            slStr = "Manual"
        ElseIf imInsertionType = 2 Then
            slStr = "Neither"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Insertion Type must be specified", tmCtrls(INSERTIONTYPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = INSERTIONTYPEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    'imInsertionType defined only for sgRNCallType = "N".
    'for sgRNCallType = "R", all fields from Insertion to end set to Optional in mInitBox
    If ((imInsertionType = 0) And (sgRNCallType = "N")) Or (sgRNCallType = "R") Then
        If (ilCtrlNo = SLSPFIRSTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcSlspFirst, "", "Salesperson First must be specified", tmCtrls(SLSPFIRSTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SLSPFIRSTINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = SLSPLASTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcSlspLast, "", "Salesperson Last must be specified", tmCtrls(SLSPLASTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SLSPLASTINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = CONTACTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcContact, "", "Contact Name must be specified", tmCtrls(CONTACTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CONTACTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PHONENOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcPhoneNo, "", "Contact Phone Number must be specified", tmCtrls(PHONENOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PHONENOINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = EMAILINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcEMail, "", "Contact E-Mail must be specified", tmCtrls(EMAILINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = EMAILINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    'For sgRNCallType = "R":  Fields set to Required = No and imInsertionType not set
    If ((imInsertionType = 0) And (sgRNCallType = "N")) Or (sgRNCallType = "R") Then
        If (ilCtrlNo = FTPUSERIDINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcFTPUserID, "", "FTP User ID must be specified", tmCtrls(FTPUSERIDINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FTPUSERIDINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = FTPUSERPWINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcFTPUserPW, "", "FTP User Password must be specified", tmCtrls(FTPUSERPWINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FTPUSERPWINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = FTPPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcFTPPort, "", "FTP Port must be specified", tmCtrls(FTPPORTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FTPPORTINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = FTPADDRESSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcFTPAddress, "", "FTP Address must be specified", tmCtrls(FTPADDRESSINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FTPADDRESSINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = FTPIMPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcFTPImport, "", "FTP Import directory must be specified", tmCtrls(FTPIMPORTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FTPIMPORTINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = FTPEXPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcFTPExport, "", "FTP Export directory must be specified", tmCtrls(FTPEXPORTINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FTPEXPORTINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = IISROOTURLINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcIISRootURL, "", "IIS Root URL must be specified", tmCtrls(IISROOTURLINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = IISROOTURLINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = IISREGSECTIONINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcIISRegSection, "", "IIS Register Section must be specified", tmCtrls(IISREGSECTIONINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = IISREGSECTIONINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    mTestFields = YES
End Function
Private Sub pbcRepNet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub pbcRepNet_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To imMaxBoxNo Step 1    'UBound(tmCtrls) Step 1
        pbcRepNet.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcRepNet.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcRepNet.Print tmCtrls(ilBox).sShow
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
    If GetFocus() <> pbcSTab.hwnd Then
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
    If GetFocus() <> pbcTab.hwnd Then
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
            If (cmcUpdate.Enabled) And (igRNCallSource = CALLNONE) Then
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

Private Sub pbcInsertionType_KeyPress(KeyAscii As Integer)
    If imBoxNo = INSERTIONTYPEINDEX Then
        If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            If imInsertionType <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imInsertionType = 0
            pbcInsertionType_Paint
        ElseIf KeyAscii = Asc("M") Or (KeyAscii = Asc("m")) Then
            If imInsertionType <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imInsertionType = 1
            pbcInsertionType_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imInsertionType <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imInsertionType = 2
            pbcInsertionType_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imInsertionType = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imInsertionType = 1
                pbcInsertionType_Paint
            ElseIf imInsertionType = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imInsertionType = 2
                pbcInsertionType_Paint
            ElseIf imInsertionType = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imInsertionType = 0
                pbcInsertionType_Paint
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcInsertionType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = INSERTIONTYPEINDEX Then
        If imInsertionType = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imInsertionType = 1
        ElseIf imInsertionType = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imInsertionType = 2
        ElseIf imInsertionType = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imInsertionType = 0
        End If
        pbcInsertionType_Paint
        mSetCommands
    End If
End Sub

Private Sub pbcInsertionType_Paint()
    pbcInsertionType.Cls
    pbcInsertionType.CurrentX = fgBoxInsetX
    pbcInsertionType.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = INSERTIONTYPEINDEX Then
        If imInsertionType = 0 Then
            pbcInsertionType.Print "CSI Link"
        ElseIf imInsertionType = 1 Then
            pbcInsertionType.Print "Manual"
        ElseIf imInsertionType = 2 Then
            pbcInsertionType.Print "Neither"
        Else
            pbcInsertionType.Print "   "
        End If
    End If
End Sub

Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If imBoxNo = TRANSMITPRICEINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imTransmitPrice <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTransmitPrice = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imTransmitPrice <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTransmitPrice = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imTransmitPrice = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imTransmitPrice = 1
                pbcYN_Paint
            ElseIf imTransmitPrice = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imTransmitPrice = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = TRANSMITPRICEINDEX Then
        If imTransmitPrice = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imTransmitPrice = 1
        ElseIf imTransmitPrice = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imTransmitPrice = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = TRANSMITPRICEINDEX Then
        If imTransmitPrice = 0 Then
            pbcYN.Print "Yes"
        ElseIf imTransmitPrice = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
End Sub
Private Sub plcRepNet_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub



