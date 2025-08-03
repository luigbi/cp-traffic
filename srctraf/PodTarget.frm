VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PodTarget 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   840
   ClientTop       =   1635
   ClientWidth     =   9360
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
   ScaleHeight     =   5475
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmcPodItem 
      Appearance      =   0  'Flat
      Caption         =   "&Ad Server Items"
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
      Left            =   5640
      TabIndex        =   14
      Top             =   5040
      Width           =   1530
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   5900
      TabIndex        =   1
      Top             =   240
      Width           =   3180
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   7400
      ScaleHeight     =   210
      ScaleWidth      =   1470
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3600
      Width           =   105
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   4920
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   480
      MaxLength       =   20
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2715
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
      Left            =   4440
      TabIndex        =   13
      Top             =   5040
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
      Left            =   3330
      TabIndex        =   12
      Top             =   5040
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
      Left            =   2205
      TabIndex        =   11
      Top             =   5040
      Width           =   1050
   End
   Begin VB.PictureBox pbcSoff 
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
      Height          =   375
      Left            =   360
      Picture         =   "PodTarget.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   8535
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   720
      Width           =   8535
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   1765
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
      TabIndex        =   5
      Top             =   1000
      Width           =   45
   End
   Begin VB.PictureBox plcSoff 
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   360
      ScaleHeight     =   360
      ScaleWidth      =   8520
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   675
      Width           =   8580
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox plcPodTarget 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   360
      ScaleHeight     =   3623.848
      ScaleMode       =   0  'User
      ScaleWidth      =   8901.845
      TabIndex        =   21
      Top             =   1320
      Width           =   8655
      Begin VB.PictureBox pbcFilterTab 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   4560
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   8
         Top             =   0
         Width           =   135
      End
      Begin VB.ListBox lbcCategory 
         Height          =   2595
         Left            =   0
         TabIndex        =   6
         Top             =   600
         Width           =   2235
      End
      Begin VB.TextBox edcFilter 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox lbcItems 
         Height          =   2205
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPodTargetSelected 
         Height          =   2775
         Left            =   4800
         TabIndex        =   10
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   10
         Cols            =   5
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
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
         _Band(0).Cols   =   5
      End
      Begin VB.Label lblItem 
         Caption         =   "Items"
         Height          =   196
         Left            =   2520
         TabIndex        =   24
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lacPodTargetSelected 
         Caption         =   "Selected"
         Height          =   196
         Left            =   5167
         TabIndex        =   23
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lacPodTargetList 
         Caption         =   "Category"
         Height          =   196
         Left            =   428
         TabIndex        =   22
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label plcScreen 
      Caption         =   "Ad Server Targeting"
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
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   1755
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "PodTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PodTarget.frm on Wed 12/31/20 @ 2:00 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 2020 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PodTarget.Frm
'
' Release: 8.1
'
' Description:
'   This file contains the Sales Office input screen code
Option Explicit
Option Compare Text
'Sales Office Field Areas

Dim tmCtrls(0 To 2)  As FIELDAREA
Dim imLBCtrls As Integer
Dim smNameCodeTag As String
Dim tgRegionArea() As RAFPodTarget
Dim imBoxNo As Integer   'Current Pod Target Box
Dim imState As Integer  '0=Active; 1=Dormant
Dim tmRaf As RAFPodTarget        'Raf Record image
Dim tmRafSrchKey As RAFPodTargetKEY0    'Raf key record image
Dim imRafRecLen As Integer        'SOF record length
Dim imUpdateAllowed As Integer    'User can update records
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imIgnoreClick As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstActivate As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imInitNoRows As Integer
Dim tgSEFPodTarget() As SEFPodTarget
Dim tgSelectedPodTarget() As SEFPodTarget
Dim ilAdfCode As Integer
Dim sAdfName As String

Dim imSelectedGridAltered As Integer
Dim tgPodCategory() As PODITEMCATEGORY
Dim tgPodItemsTarget() As PODITEMS
Dim imPodCategoryCode As Integer
Dim imFormModified As Integer

Private imLastListColSorted As Integer
Private imLastListSort As Integer
Private lmLastListClickedRow As Long

Private imLastSelectedColSorted As Integer
Private imLastSelectedSort As Integer
Private lmLastSelectedClickedRow As Long


Const NAMEINDEX = 1     'Name control/field
Const STATEINDEX = 2    'Status Control
'grid Index
Const CATEGORYINDEX = 0
Const ITEMINDEX = 1
Const THFCODEINDEX = 2
Const SORTINDEX = 3
Const SELECTEDINDEX = 4

Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    'If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        'Exit Sub
    'End If
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
    pbcSoff.Cls
    
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
    mPopulatePodItems
    'sort default by category
    mSortTargetSelectedCol 1
    mSortTargetSelectedCol 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcSoff_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    imSelectedGridAltered = False
    mSetCommands
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
    'mPopulate
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
        If igPodTargetCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgPodTargetName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgPodTargetName   'New name
            End If
            cbcSelect_Change
            If sgPodTargetName <> "" Then
                mSetCommands
                gFindMatch sgPodTargetName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                       ' cmcDone.SetFocus
                       cbcSelect.SetFocus
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
    'mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
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
    If igPodTargetCallSource <> CALLNONE Then
    
        If imFormModified = 1 Then
            sgPodTargetName = edcName.Text
            igPodTargetCallSource = CALLDONE
        Else
            igPodTargetCallSource = CALLCANCELLED
        End If
       
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
    
    If Not mOKSelectedTarget() Then
        Exit Sub
    End If
    
    If igPodTargetCallSource <> CALLNONE Then
        sgPodTargetName = edcName.Text 'Save name for returning
        If mSaveRecChg(True) = False Then
            sgPodTargetName = "[New]"
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
    If igPodTargetCallSource <> CALLNONE Then
        igPodTargetCallSource = CALLDONE
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
                
                If imBoxNo = 1 Then
                     MsgBox "Pod Target Name cannot be empty.", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                ElseIf imBoxNo = 2 Then
                     MsgBox "Pod Target Status cannot be empty.", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                End If
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcUpdate_Click()
    Dim imSvSelectedIndex As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
'    slName = Trim$(edcName.Text)   'Save name
    
    If Not mOKSelectedTarget() Then
        Exit Sub
    End If

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
    
    ilCode = tmRaf.lCode
    cbcSelect.Clear
    smNameCodeTag = edcName.Text
    mPopulate
    If imSvSelectedIndex <> 0 Then
        For ilLoop = 0 To UBound(tgRegionArea) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            If tgRegionArea(ilLoop).lCode = ilCode Then
                cbcSelect.ListIndex = ilLoop + 1
                Exit For
            End If
        Next ilLoop
    Else
          For ilLoop = 0 To UBound(tgRegionArea) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
            If Trim$(tgRegionArea(ilLoop).sName) = smNameCodeTag Then
                cbcSelect.ListIndex = ilLoop + 1
                Exit For
            End If
        Next ilLoop
    End If
    mSetCommands
    cbcSelect.SetFocus
    imFormModified = 1
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
    If (KeyAscii = KEYSLASH) Then
        Beep
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
    If (igWinStatus(PROPOSALSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSoff.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcSoff.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
    
    Me.KeyPreview = True
    PodTarget.Refresh
    cmcUpdate.Enabled = False
    pbcFilterTab.Left = lblItem.Left + lblItem.Width + 20
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
    Set PodTarget = Nothing   'Remove data segment
    
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
'*             Created:31/12/20       By:L.Bianchi     *
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
    imState = -1
    Erase tgSEFPodTarget
    Erase tgSelectedPodTarget
    tmRaf.lCode = 0
    tmRaf.sName = ""
    mClearGrid grdPodTargetSelected
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
            gMoveFormCtrl pbcSoff, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case STATEINDEX 'Selling or Airing
            If imState < 0 Then
                imState = 0
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSoff, pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcState_Paint
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:31/12/20      By:L.Bianchi      *
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
    Dim ilRet As Integer    'Return Status
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    
    mInitBox
    gCenterStdAlone PodTarget
    Screen.MousePointer = vbHourglass
    imPopReqd = False
    imFirstFocus = True
    imRafRecLen = Len(tmRaf)  'Get and save SOFF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = True
    imIgnoreClick = False
    imBypassSetting = False
    edcName.MaxLength = 80
    imFormModified = 0
    
    On Error GoTo 0
    If imTerminate Then
        Exit Sub
    End If
    If imTerminate Then
        Exit Sub
    End If
    mPopulatePodCategory
    If (Not tgPodCategory) <> -1 Then
        mLoadPodItems
    End If
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    mMoveRecToCtrl
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
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
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
    Dim height As Integer
    Dim rowHeight As Long
    Dim ilRow As Integer
    imLBCtrls = 1
    
    If sAdfName <> "" Then
        plcScreen.Caption = plcScreen.Caption & " : " & sAdfName
        plcScreen.Width = Len(plcScreen.Caption) * 156
    End If
    
    
    PodTarget.height = lgCurrVRes * 15 * 0.66
    flTextHeight = pbcSoff.TextHeight("1") - 35
    'Position panel and picture areas with panel
    
    PodTarget.Width = lgCurrHRes * 15 * 0.66
    cbcSelect.Move PodTarget.Width - (cbcSelect.Width + 300), fgPanelAdj * 4
    plcSoff.Move fgPanelAdj * 4, cbcSelect.Top + cbcSelect.height + fgPanelAdj * 2, pbcSoff.Width + fgPanelAdj, pbcSoff.height + fgPanelAdj
  

    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 7005, fgBoxStH
    'State
    gSetCtrl tmCtrls(STATEINDEX), 7060, tmCtrls(NAMEINDEX).fBoxY, 1470, fgBoxStH
    pbcSoff.Move plcSoff.Left + fgBevelX, plcSoff.Top + fgBevelY
    
    'pod target areas
    plcPodTarget.Move fgPanelAdj * 4, plcSoff.Top + plcSoff.height + fgPanelAdj * 2, PodTarget.Width - fgPanelAdj * 8, PodTarget.height - (cmcCancel.height + plcSoff.Top + plcSoff.height + fgPanelAdj * 8)

    lacPodTargetList.Move 0, 0
    height = CSng(plcPodTarget.height - lacPodTargetList.height + fgPanelAdj)
    lbcCategory.Move 0, lacPodTargetList.height + fgPanelAdj, (plcPodTarget.Width * 0.25)
    lbcCategory.height = height + fgPanelAdj
    
    lblItem.Move lbcCategory.Width + fgPanelAdj * 4, 0
    edcFilter.Move lbcCategory.Width + fgPanelAdj * 4, lblItem.height + fgPanelAdj, (plcPodTarget.Width * 0.25)
    lbcItems.Move lbcCategory.Width + fgPanelAdj * 4, edcFilter.Top + fgPanelAdj + edcFilter.height, (plcPodTarget.Width * 0.25)
    height = CSng(plcPodTarget.height - (edcFilter.Top + edcFilter.height) + fgPanelAdj)
    lbcItems.height = height
    
    lacPodTargetSelected.Move edcFilter.Left + edcFilter.Width + fgPanelAdj * 4, 0
    height = CSng(plcPodTarget.height - lacPodTargetSelected.height + fgPanelAdj)
    grdPodTargetSelected.Move edcFilter.Left + edcFilter.Width + fgPanelAdj * 4, lacPodTargetSelected.height + fgPanelAdj, (plcPodTarget.Width * 0.5) - fgPanelAdj * 2, height
    
    cmcDone.Move (PodTarget.Width / 2) - (cmcCancel.Width + cmcDone.Width + fgBevelX + fgPanelAdj), plcPodTarget.Top + plcPodTarget.height + fgPanelAdj * 2
    cmcCancel.Move cmcDone.Left + cmcDone.Width + fgPanelAdj, plcPodTarget.Top + plcPodTarget.height + (fgPanelAdj * 2)
    cmcUpdate.Move cmcCancel.Left + cmcCancel.Width + fgPanelAdj, plcPodTarget.Top + plcPodTarget.height + (fgPanelAdj * 2)
    cmcPodItem.Move cmcUpdate.Left + cmcUpdate.Width + fgPanelAdj, plcPodTarget.Top + plcPodTarget.height + (fgPanelAdj * 2)
    
    If (igWinStatus(PODITEMSLIST) = 0) Or (igGGFlag = 0) Then
        cmcPodItem.Enabled = False
    Else
        cmcPodItem.Enabled = True
    End If
    
    height = fgBoxGridH
    rowHeight = height
    mSetGridColumns
    mSetGridTitles
    gGrid_FillWithRows grdPodTargetSelected, rowHeight
    For ilRow = 0 To grdPodTargetSelected.Rows - 1 Step 1
        grdPodTargetSelected.rowHeight(ilRow) = fgBoxGridH
    Next ilRow
    imInitNoRows = (grdPodTargetSelected.height \ rowHeight) - grdPodTargetSelected.FixedRows - 1
End Sub
Private Sub mSetGridColumns()
    Dim ilCol As Integer
   
    grdPodTargetSelected.ColWidth(THFCODEINDEX) = 0
    grdPodTargetSelected.ColWidth(SORTINDEX) = 0
    grdPodTargetSelected.ColWidth(SELECTEDINDEX) = 0
    grdPodTargetSelected.ColWidth(CATEGORYINDEX) = grdPodTargetSelected.Width * 0.5
    grdPodTargetSelected.ColWidth(ITEMINDEX) = grdPodTargetSelected.Width * 0.5
    grdPodTargetSelected.ColWidth(CATEGORYINDEX) = grdPodTargetSelected.Width - GRIDSCROLLWIDTH
     For ilCol = 0 To ITEMINDEX Step 1
        If ilCol <> CATEGORYINDEX Then
            grdPodTargetSelected.ColWidth(CATEGORYINDEX) = grdPodTargetSelected.ColWidth(CATEGORYINDEX) - grdPodTargetSelected.ColWidth(ilCol)
        End If
    Next ilCol
    gGrid_AlignAllColsLeft grdPodTargetSelected
End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    grdPodTargetSelected.TextMatrix(0, CATEGORYINDEX) = "Category"
    grdPodTargetSelected.TextMatrix(0, ITEMINDEX) = "Item"
    grdPodTargetSelected.Row = grdPodTargetSelected.FixedRows - 1
    For llCol = CATEGORYINDEX To ITEMINDEX Step 1
       
        grdPodTargetSelected.Col = llCol
        grdPodTargetSelected.CellFontBold = False
        grdPodTargetSelected.CellFontName = "Arial"
        grdPodTargetSelected.CellFontSize = 6.75
        grdPodTargetSelected.CellForeColor = vbBlue
        grdPodTargetSelected.CellBackColor = LIGHTBLUE

    Next llCol
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
    Dim ilRow As Integer
    Dim ilRet As Integer    'Return call status
    Dim ilSelectedPodTargetIndex As Integer
    
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmRaf.sName = Trim$(edcName.Text)
    End If
    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
        Select Case imState
            Case 0  'Active
                tmRaf.rState = "A"
                tmRaf.sDormantDate = ""
            Case 1  'Dormant
                tmRaf.rState = "D"
                tmRaf.sDormantDate = Format(Now, sgSQLDateForm)
        End Select
    End If
    tmRaf.sChangeDate = Format(Now, sgSQLDateForm)
    For ilRow = grdPodTargetSelected.FixedRows To grdPodTargetSelected.Rows - 1
        If grdPodTargetSelected.TextMatrix(ilRow, CATEGORYINDEX) = "" Then
            Exit For
        End If
        Dim dbSEFIndex As Integer
        ReDim Preserve tgSelectedPodTarget(ilSelectedPodTargetIndex)
        tgSelectedPodTarget(ilSelectedPodTargetIndex).lCode = 0
        tgSelectedPodTarget(ilSelectedPodTargetIndex).iSequence = ilSelectedPodTargetIndex + 1
        tgSelectedPodTarget(ilSelectedPodTargetIndex).lRafCode = tmRaf.lCode
        tgSelectedPodTarget(ilSelectedPodTargetIndex).lThfCode = Val(grdPodTargetSelected.TextMatrix(ilRow, THFCODEINDEX))
        dbSEFIndex = gSearchSEF(tgSelectedPodTarget(ilSelectedPodTargetIndex).lThfCode)
        If dbSEFIndex <> -1 Then
            tgSelectedPodTarget(ilSelectedPodTargetIndex).lCode = tgSEFPodTarget(dbSEFIndex).lCode
            tgSEFPodTarget = RemoveSEFItem(dbSEFIndex)
        End If
        ilSelectedPodTargetIndex = ilSelectedPodTargetIndex + 1
    Next ilRow
    
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
'*             Created:31/12/20       By:L.Bianchi     *
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
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Sales source code number
    edcName.Text = Trim$(tmRaf.sName)
    Select Case tmRaf.rState
        Case "A"
            imState = 0
            tmRaf.sDormantDate = ""
        Case "D"
            imState = 1
        Case Else
            imState = -1
    End Select
    
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    If imSelectedIndex > 0 Then
       mClearGrid grdPodTargetSelected
       mShowSelectedCategoryItems
    End If
    
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:31/12/20        By:L.Bianchi    *
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
                    MsgBox "Region Area already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmRaf.sName) 'Reset text
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

Private Function mOKSelectedTarget()
    If grdPodTargetSelected.TextMatrix(1, CATEGORYINDEX) = "" Then
        Beep
        MsgBox "Selected Items are Empty. Please select at least one item. ", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        mOKSelectedTarget = False
        Exit Function
    End If
    mOKSelectedTarget = True
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
    'gInitStdAlone PodTarget, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igPodTargetCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igPodTargetCallSource = CALLNONE
    'End If
    ilRet = gParseItem(slCommand, 6, "\", slStr)
    ilAdfCode = CInt(slStr) 'Get Advertiser code
    ilRet = gParseItem(slCommand, 5, "\", sAdfName) 'Get Advertiser Name
    If igPodTargetCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgPodTargetName = slStr
        Else
            sgPodTargetName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
    'cbcSelect.AddItem "[New]", 0  'Force as first item on list
    gPopRegionArea
    imPopReqd = True
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
'*             Created:31/12/20       By:L.Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer)
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    tmRaf.sName = tgRegionArea(ilSelectIndex - 1).sName
    tmRaf.rState = tgRegionArea(ilSelectIndex - 1).rState
    tmRaf.lCode = tgRegionArea(ilSelectIndex - 1).lCode
    tmRaf.sChangeDate = tgRegionArea(ilSelectIndex - 1).sChangeDate
    tmRaf.sDormantDate = Trim$(tgRegionArea(ilSelectIndex - 1).sDormantDate)
    tmRaf.iAdfCode = tgRegionArea(ilSelectIndex - 1).iAdfCode
    mReadSEF
    On Error GoTo mReadRecErr
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
'*            Created:31/12/20      By:L.Bianchi       *
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
    mMoveCtrlToRec True
    'write Save code
    If imSelectedIndex = 0 Then 'New selected
        tmRaf.lCode = 0  'Autoincrement
        tmRaf.rType = "P"
        tmRaf.iUrfCode = tgUrf(0).iCode
        tmRaf.iAdfCode = ilAdfCode
        tmRaf.lCode = mSaveRaf()
    Else 'Old record-Update
        mUpdateRaf
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

Private Function mUpdateRaf() As Long
    Dim slQuery As String
    Dim llCount As Long
    slQuery = "Update RAF_Region_Area SET rafName = '" & gFixQuote(Trim(tmRaf.sName)) & "',"
    slQuery = slQuery & "rafState = '" & tmRaf.rState & "', rafDateEntrd = '" & tmRaf.sChangeDate & "'"
    If Trim$(tmRaf.sDormantDate) <> "" Then
        slQuery = slQuery & ", rafDateDormant = '" & tmRaf.sDormantDate & "'"
    Else
        slQuery = slQuery & ", rafDateDormant = NULL"
    End If
    
    slQuery = slQuery & " WHERE rafCode = " & tmRaf.lCode
    
    If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
       gHandleError "TrafficErrors.txt", "mUpdateRaf"
       mUpdateRaf = llCount
       Exit Function
    End If
End Function

Private Function mSaveRaf() As Long
 Dim slQuery As String
    slQuery = "INSERT INTO RAF_Region_Area(rafCode, rafAdfCode ,rafName, rafState, rafType,rafUrfCode, rafDateEntrd"
    If Trim$(tmRaf.sDormantDate) <> "" Then
        slQuery = slQuery & ", rafDateDormant) VALUES("
    Else
        slQuery = slQuery & ")VALUES("
    End If
    
    slQuery = slQuery & "replace" & ","
    slQuery = slQuery & tmRaf.iAdfCode & ",'"
    slQuery = slQuery & gFixQuote(Trim(tmRaf.sName)) & "','"
    slQuery = slQuery & tmRaf.rState & "','"
    slQuery = slQuery & tmRaf.rType & "',"
    slQuery = slQuery & tmRaf.iUrfCode
    slQuery = slQuery & ",'" & tmRaf.sChangeDate & "'"
    
    If Trim$(tmRaf.sDormantDate) <> "" Then
        slQuery = slQuery & ",'" & tmRaf.sDormantDate & "')"
        Else
        slQuery = slQuery & ")"
    End If
    mSaveRaf = gInsertAndReturnCode(slQuery, "RAF_Region_Area", "rafCode", "replace")
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
        If ilAltered = YES Or imSelectedGridAltered Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcSoff_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    If imSelectedGridAltered And mSaveRecChg Then
                        mSavePodTarget
                        imSelectedGridAltered = False
                    End If
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                If imSelectedGridAltered And mSaveRecChg Then
                    mSavePodTarget
                    imSelectedGridAltered = False
                End If
                Exit Function
            End If
        End If
    End If
    If imSelectedIndex > 0 Then
        If imSelectedGridAltered Then
            mMoveCtrlToRec True
            mSavePodTarget
            Erase tgSelectedPodTarget
            imSelectedGridAltered = False
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
            gSetChgFlag tmRaf.sName, edcName, tmCtrls(ilBoxNo)
        Case STATEINDEX
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
        'cmcDone.Enabled = True
    ElseIf imSelectedGridAltered And tmCtrls(1).sShow <> "" And tmCtrls(2).sShow <> "" Then
        cmcUpdate.Enabled = True
        'cmcDone.Enabled = True
    Else
        cmcUpdate.Enabled = False
        'cmcDone.Enabled = False
    End If
    
    If cbcSelect.ListIndex > 0 Then
     cmcDone.Enabled = True
    ElseIf edcName.Text <> "" And imState <> -1 Then
     cmcDone.Enabled = True
    Else
     cmcDone.Enabled = False
    End If
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
    Dim ilPos As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case STATEINDEX 'State
            pbcState.Visible = False  'Set visibility
            If imState = 0 Then
                slStr = "Active"
            ElseIf imState = 1 Then
                slStr = "Dormant"
            Else
                slStr = ""
            End If
            pbcState_Paint
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:31/12/20       By:L.Bianchi     *
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
    sgDoneMsg = Trim$(str$(igPodTargetCallSource)) & "\" & sgPodTargetName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload PodTarget
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:31/12/29       By:L.Bianchi     *
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
    If (ilCtrlNo = STATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imState = 0 Then
            slStr = "Active"
        ElseIf imState = 1 Then
            slStr = "Dormant"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Active Or Dormant must be specified", tmCtrls(STATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    
    mTestFields = YES
End Function

Private Sub lbcCategory_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    imIgnoreClick = False
    For iLoop = 0 To lbcCategory.ListCount - 1 Step 1
        If lbcCategory.Selected(iLoop) Then
            imPodCategoryCode = lbcCategory.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    
    If iCount = 1 Then
        lbcItems.Clear
        edcFilter.Text = ""
        mFillPodItems
    End If
End Sub

Private Sub lbcItems_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    Dim lIndex As Integer
    Dim lRIndex As Integer
    lRIndex = -1
    
    If imIgnoreClick = True Then
     Exit Sub
    End If
    
    If (Not tgSelectedPodTarget) = -1 Then
        iCount = -1
    Else
        iCount = UBound(tgSelectedPodTarget)
    End If
    For iLoop = 0 To lbcItems.ListCount - 1 Step 1
        If lbcItems.Selected(iLoop) Then
            iCount = iCount + 1
            ReDim Preserve tgSelectedPodTarget(iCount)
            tgSelectedPodTarget(iCount).lThfCode = lbcItems.ItemData(iLoop)
            tgSelectedPodTarget(iCount).lMnfCode = imPodCategoryCode
            lIndex = gBinarySearchPodCategory(imPodCategoryCode)
            tgSelectedPodTarget(iCount).sCategoryName = tgPodCategory(lIndex).sName
            lIndex = gBinarySearchPodItem(lbcItems.ItemData(iLoop))
            tgSelectedPodTarget(iCount).sPodItemName = tgPodItemsTarget(lIndex).ItemName
            tgSelectedPodTarget(iCount).iPodItemIndex = lIndex
            lRIndex = iLoop
        End If
    Next iLoop
    If lRIndex > -1 Then
        imSelectedGridAltered = True
        mAddTargetItemGrid
        lbcItems.RemoveItem lRIndex
        mSetCommands
    End If
    
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
Private Sub pbcSoff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim flAdj As Single
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY + flAdj) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH + flAdj) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcSoff_Paint()
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcSoff.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSoff.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSoff.Print tmCtrls(ilBox).sShow
    Next ilBox
    llColor = pbcSoff.ForeColor
    slFontName = pbcSoff.FontName
    flFontSize = pbcSoff.FontSize
    pbcSoff.ForeColor = BLUE
    pbcSoff.FontBold = False
    pbcSoff.FontSize = 7
    pbcSoff.FontName = "Arial"
    pbcSoff.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    pbcSoff.CurrentX = tmCtrls(NAMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcSoff.CurrentY = tmCtrls(NAMEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcSoff.Print "Name"
    pbcSoff.CurrentX = tmCtrls(STATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcSoff.CurrentY = tmCtrls(STATEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcSoff.Print "Status"
    pbcSoff.FontSize = flFontSize
    pbcSoff.FontName = slFontName
    pbcSoff.FontSize = flFontSize
    pbcSoff.ForeColor = llColor
    pbcSoff.FontBold = True
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-right to left
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1
                imTabDirection = 0  'Set-Left to right
               ' If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                    ilBox = 1
                    mSetCommands
               ' Else
               '    mSetChg 1
               '    ilBox = 2
               'End If
            Case NAMEINDEX 'Name (first control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcState_GotFocus()
    imBoxNo = STATEINDEX
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imState <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 0
        pbcState_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imState <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 1
        pbcState_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imState = 0 Then  'Active
            imState = 1
            tmCtrls(imBoxNo).iChg = True
            pbcState_Paint
        ElseIf imState = 1 Then  'Dormant
            tmCtrls(imBoxNo).iChg = True
            imState = 0  'Active
            pbcState_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imState <= 0 Then  'Active
        tmCtrls(imBoxNo).iChg = True
        imState = 1  'Dormant
    ElseIf imState = 1 Then  'Dormant
        tmCtrls(imBoxNo).iChg = True
        imState = 0  'Active
    End If
    pbcState_Paint
    mSetCommands
End Sub
Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    Select Case imState
        Case 0  'Active
            pbcState.Print "Active"
        Case 1  'Dormant
            pbcState.Print "Dormant"
        Case Else
            pbcState.Print "       "
    End Select
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcTab.hWnd Then
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
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1
                imTabDirection = -1  'Set-Right to left
                ilBox = UBound(tmCtrls)
            Case UBound(tmCtrls) 'last control
                mSetShow imBoxNo
                imBoxNo = -1
                lbcCategory.SetFocus
                'If (cmcUpdate.Enabled) And (igPodTargetCallSource = CALLNONE) Then
                    'cmcUpdate.SetFocus
                'Else
                    'cmcDone.SetFocus
                'End If
                Exit Sub
            Case Else
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSoff_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
End Sub

Private Sub mAddCateoryITemToGrid(grdAdd As MSHFlexGrid, llRow As Long, lIndex As Integer)
    If llRow >= grdAdd.Rows Then
        grdAdd.AddItem "", llRow
    Else
        If grdAdd.TextMatrix(llRow, CATEGORYINDEX) <> "" Then
            grdAdd.AddItem "", llRow
        End If
    End If
    grdAdd.TextMatrix(llRow, CATEGORYINDEX) = Trim$(tgSEFPodTarget(lIndex).sCategoryName)
    grdAdd.TextMatrix(llRow, ITEMINDEX) = Trim$(tgSEFPodTarget(lIndex).sPodItemName)
    grdAdd.TextMatrix(llRow, THFCODEINDEX) = Trim$(tgSEFPodTarget(lIndex).lThfCode)
    grdAdd.rowHeight(lIndex) = fgBoxGridH
    llRow = llRow + 1
End Sub

Private Sub grdPodTargetSelected_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llCurrentRow As Long
    Dim llCol As Long
    Dim ilIndex As Integer
    Dim ilCode As Long
    Dim ilPodIndex As Integer
    If Y < grdPodTargetSelected.rowHeight(0) Then
        grdPodTargetSelected.Col = grdPodTargetSelected.MouseCol
        mSortTargetSelectedCol grdPodTargetSelected.Col
        grdPodTargetSelected.Row = 0
        grdPodTargetSelected.Col = THFCODEINDEX
        Exit Sub
    End If
    llCurrentRow = grdPodTargetSelected.MouseRow
    llCol = grdPodTargetSelected.MouseCol
    If llCurrentRow < grdPodTargetSelected.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdPodTargetSelected.FixedRows Then
        If grdPodTargetSelected.TextMatrix(llCurrentRow, CATEGORYINDEX) <> "" Then
            ilIndex = llCurrentRow - grdPodTargetSelected.FixedRows
            ilCode = tgSelectedPodTarget(ilIndex).lMnfCode
            
            If imPodCategoryCode = ilCode Then
                ilPodIndex = gBinarySearchPodItem(tgSelectedPodTarget(ilIndex).lThfCode)
                tgSelectedPodTarget = RemoveItem(ilIndex)
                If ilPodIndex <> -1 Then
                  mPopulatePodItems
                End If
            Else
                tgSelectedPodTarget = RemoveItem(ilIndex)
            End If
            
            grdPodTargetSelected.RemoveItem llCurrentRow
            If grdPodTargetSelected.Rows - grdPodTargetSelected.FixedRows <= imInitNoRows Then
                 grdPodTargetSelected.AddItem "", grdPodTargetSelected.Rows
                 grdPodTargetSelected.rowHeight(grdPodTargetSelected.Rows - 1) = fgBoxGridH
            End If
            imSelectedGridAltered = True
            mSetCommands
        End If
    End If
End Sub

Private Sub mPaintRowColor(grdPaint As MSHFlexGrid, llRow As Long)
    Dim llCol As Long
    grdPaint.Row = llRow
    For llCol = CATEGORYINDEX To ITEMINDEX Step 1
        grdPaint.Col = llCol
        If grdPaint.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdPaint.CellBackColor = vbWhite
        Else
            grdPaint.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol
End Sub

Private Sub mSortTargetSelectedCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    'imSelectedGridAltered = True
    mSetCommands
     
    For llRow = grdPodTargetSelected.FixedRows To grdPodTargetSelected.Rows - 1 Step 1
        slStr = Trim$(grdPodTargetSelected.TextMatrix(llRow, CATEGORYINDEX))
        If slStr <> "" Then
             If ilCol = ITEMINDEX Then
                slSort = grdPodTargetSelected.TextMatrix(llRow, ITEMINDEX)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            Else
                slSort = UCase$(Trim$(grdPodTargetSelected.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdPodTargetSelected.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastSelectedColSorted) Or ((ilCol = imLastSelectedColSorted) And (imLastSelectedSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPodTargetSelected.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPodTargetSelected.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastSelectedColSorted Then
        imLastSelectedColSorted = SORTINDEX
    Else
        imLastSelectedColSorted = -1
        imLastSelectedSort = -1
    End If
    gGrid_SortByCol grdPodTargetSelected, CATEGORYINDEX, SORTINDEX, imLastSelectedColSorted, imLastSelectedSort
    imLastSelectedColSorted = ilCol
End Sub

Private Sub mClearGrid(grd As MSHFlexGrid)
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    'gGrid_Clear grdAirPlay, False
    'Set color within cells
    For llRow = grd.FixedRows To grd.Rows - 1 Step 1
        For llCol = CATEGORYINDEX To SELECTEDINDEX Step 1
            grd.TextMatrix(llRow, llCol) = ""
            grd.Row = llRow
            grd.Col = llCol
            grd.CellBackColor = vbWhite
        Next llCol
    Next llRow
    grd.TopRow = grd.FixedRows
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopRegionArea                  *
'*                                                     *
'*            Created:01/06/21      By:L. Bianchi      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: populate Region Area           *
'*                                                     *
'*******************************************************

Public Function gPopRegionArea() As Integer
    Dim llUpper As Long
    Dim rst As ADODB.Recordset
    Dim slStamp As String
    Dim llMax As Long
    Dim i As Integer

    slStamp = gFileDateTime(sgDBPath & "raf.Mkd")
    If Not gFileChgd("raf.mkd") Then
        gPopRegionArea = True
        Exit Function
    End If
    
    SQLQuery = "SELECT rafCode, rafAdfCode, rafState,rafName, rafDateEntrd, rafDateDormant FROM RAF_Region_Area WHERE rafType = 'P' and rafAdfCode = " & ilAdfCode & " order by rafName"
    Set rst = gSQLSelectCall(SQLQuery)
    llUpper = 0
    llMax = 100
    ReDim tgRegionArea(0 To llMax) As RAFPodTarget
    
    While Not rst.EOF
        tgRegionArea(llUpper).lCode = rst!rafCode
        tgRegionArea(llUpper).sName = Trim$(rst!rafName)
        tgRegionArea(llUpper).rState = Trim$(rst!rafState)
        tgRegionArea(llUpper).iAdfCode = rst!rafAdfCode
        If Not IsNull(rst!rafDateEntrd) Then
            tgRegionArea(llUpper).sChangeDate = Format(rst!rafDateEntrd, sgSQLDateForm)
        End If
        
        If Not IsNull(rst!rafDateDormant) Then
            tgRegionArea(llUpper).sDormantDate = Format(rst!rafDateDormant, sgSQLDateForm)
        End If
        llUpper = llUpper + 1
        If llUpper >= llMax Then
            llMax = llMax + 100
            ReDim Preserve tgRegionArea(0 To llMax) As RAFPodTarget
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tgRegionArea(0 To llUpper) As RAFPodTarget
    
    
    cbcSelect.Clear
    cbcSelect.AddItem "[New]", 0

    For i = 0 To UBound(tgRegionArea) - 1 Step 1
        cbcSelect.AddItem Trim$(tgRegionArea(i).sName)
        cbcSelect.ItemData(cbcSelect.NewIndex) = tgRegionArea(i).lCode
    Next i
    
    gFileChgdUpdate "raf.mkd", False
    gPopRegionArea = True
    rst.Close
    
End Function

Private Sub mSavePodTarget()
    Dim iIndex As Integer
    Dim ilBound As Integer, iUBound As Integer
    Dim slQuery As String
    Dim llCount As Long
    
    If (Not tgSEFPodTarget) <> -1 Then
        ilBound = LBound(tgSEFPodTarget)
        iUBound = UBound(tgSEFPodTarget)
        
        For iIndex = ilBound To iUBound
            slQuery = "DELETE FROM SEF_Split_Entity WHERE sefCode = " & tgSEFPodTarget(iIndex).lCode
            If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
                gHandleError "TrafficErrors.txt", "PodTarget-mSavePodTarget"
                Exit Sub
            End If
        Next iIndex
        slQuery = ""
    End If
    
    If (Not tgSelectedPodTarget) = -1 Then
        Exit Sub
    End If
    ilBound = LBound(tgSelectedPodTarget)
    iUBound = UBound(tgSelectedPodTarget)
    
    For iIndex = ilBound To iUBound
        If tgSelectedPodTarget(iIndex).lCode = 0 Then
            slQuery = "INSERT INTO SEF_Split_Entity(sefRafCode,sefLongCode,sefSeqNo, sefCategory, sefInclExcl) Values("
            slQuery = slQuery & tmRaf.lCode & ","
            slQuery = slQuery & tgSelectedPodTarget(iIndex).lThfCode & ","
            slQuery = slQuery & tgSelectedPodTarget(iIndex).iSequence & ","
            slQuery = slQuery & "'P',"
            slQuery = slQuery & "'I')"
        Else
            slQuery = "Update SEF_Split_Entity SET sefSeqNo = " & tgSelectedPodTarget(iIndex).iSequence & "WHERE sefCode =" & tgSelectedPodTarget(iIndex).lCode
        End If
        If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
            gHandleError "TrafficErrors.txt", "PodTarget-mSavePodTarget"
            Exit Sub
        End If
    Next iIndex
End Sub

Private Sub mReadSEF()

    Dim llUpper As Long
    Dim rst As ADODB.Recordset
    Dim slStamp As String
    Dim i As Integer
    Erase tgSEFPodTarget
    Erase tgSelectedPodTarget
    slStamp = gFileDateTime(sgDBPath & "SEF.Mkd")
    If Not gFileChgd("SEF.mkd") Then
        Exit Sub
    End If
    
    SQLQuery = "select sefCode, sefRafCode, sefLongCode, mnfName, thfName, thfCategoryMnfCode,sefSeqNo from SEF_Split_Entity"
    SQLQuery = SQLQuery & " INNER JOIN thf_Target_Header on SEF_Split_Entity.sefLongCode = thf_Target_Header.thfCode"
    SQLQuery = SQLQuery & " INNER JOIN MNF_Multi_Names on thf_Target_Header.thfCategoryMnfCode = MNF_Multi_Names.mnfCode"
    SQLQuery = SQLQuery & " WHERE SEF_Split_Entity.sefCategory = 'P' and SEF_Split_Entity.sefRafCode =  " & tmRaf.lCode & " order by sefSeqNo"
    Set rst = gSQLSelectCall(SQLQuery)
    llUpper = 0
    While Not rst.EOF
        ReDim Preserve tgSEFPodTarget(0 To llUpper) As SEFPodTarget
        tgSEFPodTarget(llUpper).iSequence = rst!sefSeqNo
        tgSEFPodTarget(llUpper).lThfCode = rst!sefLongCode
        tgSEFPodTarget(llUpper).lRafCode = rst!sefRafCode
        tgSEFPodTarget(llUpper).lCode = rst!sefCode
        tgSEFPodTarget(llUpper).sCategoryName = Trim$(rst!mnfName)
        tgSEFPodTarget(llUpper).sPodItemName = Trim$(rst!thfName)
        tgSEFPodTarget(llUpper).lMnfCode = rst!thfCategoryMnfCode
        llUpper = llUpper + 1
        rst.MoveNext
    Wend
    If (Not tgSEFPodTarget) <> -1 Then
        ReDim tgSelectedPodTarget(UBound(tgSEFPodTarget))
        For i = 0 To UBound(tgSEFPodTarget) Step 1
            tgSelectedPodTarget(i) = tgSEFPodTarget(i)
        Next i
    End If
    gFileChgdUpdate "SEF.mkd", False
    rst.Close
End Sub

Public Function gSearchSEF(lThfCode As Long) As Integer
    Dim Index As Integer
    gSearchSEF = -1
    If (Not tgSEFPodTarget) = -1 Then
        Exit Function
    End If
    For Index = LBound(tgSEFPodTarget) To UBound(tgSEFPodTarget)
        If tgSEFPodTarget(Index).lThfCode = lThfCode Then
            gSearchSEF = Index
            Exit For
        End If
    Next
End Function


Private Function RemoveItem(item As Integer) As SEFPodTarget()
  Dim iIndex As Integer
  Dim SEFDest() As SEFPodTarget
  Dim ilBound As Integer, iUBound As Integer
  'find the boundaries of the source array
  
  ilBound = LBound(tgSelectedPodTarget)
  iUBound = UBound(tgSelectedPodTarget)
  If iUBound = 0 Then
    RemoveItem = SEFDest
    Exit Function
  End If
  'set boundaries for the resulting array
  ReDim SEFDest(ilBound To iUBound - 1) As SEFPodTarget
  'copy items which remain
  For iIndex = ilBound To item - 1
    SEFDest(iIndex) = tgSelectedPodTarget(iIndex)
  Next iIndex
  'skip the removed item
  'and copy the remaining items, with destination index-1
  For iIndex = item + 1 To iUBound
    SEFDest(iIndex - 1) = tgSelectedPodTarget(iIndex)
  Next iIndex
  'return the result
  RemoveItem = SEFDest
End Function

Private Function RemoveSEFItem(item As Integer) As SEFPodTarget()
  Dim iIndex As Integer
  Dim SEFDest() As SEFPodTarget
  Dim ilBound As Integer, iUBound As Integer
  'find the boundaries of the source array
  
  ilBound = LBound(tgSEFPodTarget)
  iUBound = UBound(tgSEFPodTarget)
  If iUBound = 0 Then
    RemoveSEFItem = SEFDest
    Exit Function
  End If
  'set boundaries for the resulting array
  ReDim SEFDest(ilBound To iUBound - 1) As SEFPodTarget
  'copy items which remain
  For iIndex = ilBound To item - 1
    SEFDest(iIndex) = tgSEFPodTarget(iIndex)
  Next iIndex
  'skip the removed item
  'and copy the remaining items, with destination index-1
  For iIndex = item + 1 To iUBound
    SEFDest(iIndex - 1) = tgSEFPodTarget(iIndex)
  Next iIndex
  'return the result
  RemoveSEFItem = SEFDest
End Function

Private Sub mShowSelectedCategoryItems()
    Dim ilLoop As Integer
    Dim llRow As Long
    If (Not tgSelectedPodTarget) = -1 Then
        Exit Sub
    End If
    llRow = grdPodTargetSelected.FixedRows
    For ilLoop = 0 To UBound(tgSelectedPodTarget) Step 1
         mAddCateoryITemToGrid grdPodTargetSelected, llRow, ilLoop
    Next ilLoop
    imLastListColSorted = -1
    imLastListSort = -1
    'mSortExcludeCol CALLLETTERSINDEX
    'mGenOK
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
End Sub

Private Sub cmcPodItem_Click()
        Dim slStr As String
        igPodTargetCallSource = CALLSOURCEPODTARGET
        If igTestSystem Then
            slStr = "PodTarget^Test\" & sgUserName & "\" & Trim$(str$(igPodTargetCallSource))
        Else
            slStr = "PodTarget^Prod\" & sgUserName & "\" & Trim$(str$(igPodTargetCallSource))
        End If
        
        sgCommandStr = slStr
        On Error Resume Next
        PodItem.Show vbModal
        slStr = sgDoneMsg
        
        mPopulatePodCategory
        If imSelectedIndex > 0 Then
           mReadSEF
           mClearGrid grdPodTargetSelected
           mShowSelectedCategoryItems
        End If
End Sub

Private Sub mEmptyRows(grdCtrl As MSHFlexGrid)
   Dim llRows As Long
    Dim llCols As Long
    Dim llFillNoRow As Long
    llFillNoRow = grdCtrl.height \ fgBoxGridH - grdCtrl.FixedRows - 1

    For llRows = grdCtrl.FixedRows To grdCtrl.FixedRows + llFillNoRow Step 1
        Do While llRows >= grdCtrl.Rows
            grdCtrl.AddItem ""
            For llCols = 0 To grdCtrl.Cols - 1 Step 1
                grdCtrl.TextMatrix(llRows, llCols) = ""
            Next llCols
        Loop
    Next llRows
End Sub

Private Sub mPopulatePodCategory()
    Dim Index As Integer
    Dim iLoop As Integer
    Dim MNF As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT mnfCode, mnfName FROM MNF_Multi_Names WHERE mnfType = '5' order by LOWER(mnfName)"
    Set MNF = gSQLSelectCall(SQLQuery)
    
    While Not MNF.EOF
        ReDim Preserve tgPodCategory(Index)
        tgPodCategory(Index).iCode = MNF!mnfCode
        tgPodCategory(Index).sName = Trim$(MNF!mnfName)
        Index = Index + 1
        MNF.MoveNext
    Wend
    MNF.Close
    lbcCategory.Clear
    If (Not tgPodCategory) <> -1 Then
        For iLoop = 0 To UBound(tgPodCategory) Step 1
            lbcCategory.AddItem tgPodCategory(iLoop).sName
            lbcCategory.ItemData(lbcCategory.NewIndex) = tgPodCategory(iLoop).iCode
        Next iLoop
        If (imPodCategoryCode <> 0) Then
            Index = gBinarySearchPodCategory(imPodCategoryCode)
            lbcCategory.ListIndex = Index
        Else
            lbcCategory.ListIndex = 0
            imPodCategoryCode = lbcCategory.ItemData(0)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mFillPodItems()
    mLoadPodItems
    mPopulatePodItems
End Sub

Private Sub mPopulatePodItems()
    Dim iLoop As Integer
    Dim iCode As Integer
    lbcItems.Clear
    
    If (Not tgPodItemsTarget) <> -1 Then
    For iLoop = 0 To UBound(tgPodItemsTarget) Step 1
        If tgPodItemsTarget(iLoop).iMnfCode = imPodCategoryCode Then
            If gSearchSelectedPodItems(tgPodItemsTarget(iLoop).iCode) = -1 Then
                lbcItems.AddItem tgPodItemsTarget(iLoop).ItemName
                lbcItems.ItemData(lbcItems.NewIndex) = tgPodItemsTarget(iLoop).iCode
            End If
        End If
    Next iLoop
    End If
End Sub

Private Sub mLoadPodItems()
    Dim Index As Integer
    Dim THF As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    SQLQuery = "select thfCode, thfName, thfCategoryMnfCode from thf_Target_Header WHERE thfCategoryMnfCode = " & imPodCategoryCode & " order by LOWER(thfName)"
    Set THF = gSQLSelectCall(SQLQuery)
    
    While Not THF.EOF
        ReDim Preserve tgPodItemsTarget(Index)
        tgPodItemsTarget(Index).iCode = THF!thfCode
        tgPodItemsTarget(Index).ItemName = Trim$(THF!thfName)
        tgPodItemsTarget(Index).iMnfCode = THF!thfCategoryMnfCode
        Index = Index + 1
        THF.MoveNext
    Wend
    THF.Close
    Screen.MousePointer = vbDefault
End Sub

Private Function gBinarySearchPodItem(iCode As Long) As Long
   Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If (Not tgPodItemsTarget) = -1 Then
        gBinarySearchPodItem = -1
        Exit Function
    End If
    llMin = LBound(tgPodItemsTarget)
    llMax = UBound(tgPodItemsTarget)
    Do While llMin <= llMax
        
        If iCode = tgPodItemsTarget(llMin).iCode Then
            'found the match
            gBinarySearchPodItem = llMin
            Exit Function
        End If
        llMin = llMin + 1
    Loop
    gBinarySearchPodItem = -1
    Exit Function

End Function
Private Function gBinarySearchPodCategory(iCode As Integer) As Integer
   Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgPodCategory)
    llMax = UBound(tgPodCategory)
    Do While llMin <= llMax
        If iCode = tgPodCategory(llMin).iCode Then
            'found the match
            gBinarySearchPodCategory = llMin
            Exit Function
        End If
        llMin = llMin + 1
    Loop
    gBinarySearchPodCategory = -1
    Exit Function
End Function

Public Sub pbcFilterTab_GotFocus()
     imIgnoreClick = False
     lbcItems_Click
     edcFilter.Text = ""
     edcFilter.SetFocus
End Sub


Private Sub mAddTargetItemGrid()
    Dim llNextRow As Long
    Dim iCount As Integer
    If (Not tgSelectedPodTarget) <> -1 Then
        iCount = UBound(tgSelectedPodTarget)
    Else
        iCount = 0
    End If
    llNextRow = iCount + grdPodTargetSelected.FixedRows
    If llNextRow >= grdPodTargetSelected.Rows Then
        grdPodTargetSelected.AddItem "", llNextRow
    Else
        If grdPodTargetSelected.TextMatrix(llNextRow, CATEGORYINDEX) <> "" Then
            grdPodTargetSelected.AddItem "", llNextRow
        End If
    End If
    grdPodTargetSelected.TextMatrix(llNextRow, CATEGORYINDEX) = tgSelectedPodTarget(iCount).sCategoryName
    grdPodTargetSelected.TextMatrix(llNextRow, ITEMINDEX) = tgSelectedPodTarget(iCount).sPodItemName
    grdPodTargetSelected.TextMatrix(llNextRow, THFCODEINDEX) = tgSelectedPodTarget(iCount).lThfCode
    grdPodTargetSelected.TextMatrix(llNextRow, SELECTEDINDEX) = tgSelectedPodTarget(iCount).iPodItemIndex
    grdPodTargetSelected.rowHeight(llNextRow) = fgBoxGridH
End Sub

Private Function gSearchSelectedPodItems(iCode As Integer) As Integer
   Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    If (Not tgSelectedPodTarget) = -1 Then
        gSearchSelectedPodItems = -1
        Exit Function
    End If
    
    llMin = LBound(tgSelectedPodTarget)
    llMax = UBound(tgSelectedPodTarget)
    Do While llMin <= llMax
        
        If iCode = tgSelectedPodTarget(llMin).lThfCode Then
            gSearchSelectedPodItems = llMin
            Exit Function
        End If
        llMin = llMin + 1
    Loop
    gSearchSelectedPodItems = -1
    Exit Function
End Function

Private Sub edcFilter_GotFocus()
    edcFilter.BackColor = &HFFFF00
End Sub

Private Sub edcFilter_LostFocus()
    edcFilter.BackColor = &H80000005
    imIgnoreClick = False
    
End Sub

Private Sub edcFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    imIgnoreClick = True
End Sub

Private Sub edcFilter_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcFilter.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcFilter_Change()
    gMatchLookAhead edcFilter, lbcItems, imBSMode, imComboBoxIndex
    imIgnoreClick = False
End Sub
