VERSION 5.00
Begin VB.Form AdServerFilter 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5475
   ClientLeft      =   13245
   ClientTop       =   7605
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
   Begin VB.TextBox edcContract 
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox edcYear 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox plcPodTarget 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   240
      ScaleHeight     =   4139.239
      ScaleMode       =   0  'User
      ScaleWidth      =   8531.577
      TabIndex        =   18
      Top             =   840
      Width           =   8295
      Begin VB.CheckBox cksContract 
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.ListBox lbcContracts 
         Height          =   2400
         Left            =   5400
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
      Begin VB.PictureBox pbcFilterTab 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   3720
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lbcAdserver 
         Height          =   2595
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   2235
      End
      Begin VB.ListBox lbcAdvertiser 
         Height          =   2595
         Left            =   2520
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox cksAdvertiser 
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblContract 
         Caption         =   "All Contracts"
         Height          =   255
         Left            =   5760
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblAdvertiser 
         Caption         =   "All Advertisers"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblAdServerVendor 
         Caption         =   "Ad Server Vendors"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox cbcMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   3855
      TabIndex        =   2
      Top             =   240
      Width           =   1580
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8640
      TabIndex        =   16
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3600
      Width           =   105
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7200
      Top             =   4920
   End
   Begin VB.CommandButton cmcPost 
      Appearance      =   0  'Flat
      Caption         =   "&Post"
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
      Left            =   4320
      TabIndex        =   10
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
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Width           =   1050
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   1765
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   11
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
      TabIndex        =   12
      Top             =   1695
      Width           =   45
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblOr 
      Caption         =   "Or"
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblContractNumber 
      Caption         =   "Contract #"
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblYear 
      Caption         =   "Post : Year"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label plcScreen 
      Caption         =   "Ad Server Posting"
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
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   1995
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
Attribute VB_Name = "AdServerFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AdServerFilter.frm on Wed 12/31/20 @ 2:00 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 2020 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AdServerFilter.Frm
'
' Release: 8.1
'
' Description:
'   This file contains the Sales Office input screen code
Option Explicit
Option Compare Text
'Sales Office Field Areas


Dim imUpdateAllowed As Integer    'User can update records
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imIgnoreClick As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstActivate As Integer
Dim imFirstFocus As Integer 'True = cbcMonth has not had focus yet, used to branch to another control
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imPopReqd As Integer         'Flag indicating if cbcMonth was populated
Dim imBypassSetting As Integer      'In cbcMonth--- bypass mSetCommands (when user entering new name, don't want cbcMonth disabled)
Dim imAdvCheckAllProcess As Integer
Dim imAdvChkAll As Integer
Dim imContractChkAll As Integer
Dim imAdvChkTemp As Integer
Dim imContractChkTemp As Integer

Dim Months(12) As String
Dim imChfadfCodes As String
Dim imChfCntrNo As String
Dim imMonthSelectedIndex As Integer
Dim imMaxRowListItem As Integer

Dim tgAdServerVendor() As AVF
Dim imAdfCode As Integer

Dim tgAdvertiser() As ADF

Dim tgPodItemsTarget() As PODITEMS
Dim imPodCategoryCode As Integer
Dim imFormModified As Integer
Dim imSetting As Integer

Private Sub cbcMonth_DropDown()
    tmcClick.Enabled = False
    imSelectDelay = False
End Sub

Private Sub cbcMonth_GotFocus()
    If cbcMonth.Text = "" Then
    ' get the default vehicle from this global var
        gFindMatch month(0), 0, cbcMonth
        If gLastFound(cbcMonth) >= 0 Then
            cbcMonth.ListIndex = gLastFound(cbcMonth)
        Else
            If cbcMonth.ListCount > 0 Then
                cbcMonth.ListIndex = 0
            End If
        End If
    Else
        cbcMonth.ListIndex = imMonthSelectedIndex
    End If
    gCtrlGotFocus cbcMonth
    tmcClick.Enabled = False
End Sub
Private Sub cbcMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcMonth_KeyPress(KeyAscii As Integer)
    tmcClick.Enabled = False
    imSelectDelay = False
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcMonth.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcMonth_Change()
    Dim slStr As String
    slStr = Trim$(cbcMonth.Text)
    If slStr <> "" Then
            gManLookAhead cbcMonth, imBSMode, imMonthSelectedIndex
            If cbcMonth.ListIndex >= 0 Then
                imChfCntrNo = ""
                tmcClick.Enabled = False
                imSelectDelay = True
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
                If (imChfadfCodes <> "") Then
                    mPopulateContracts
                    If imContractChkAll Then
                        imSetting = True
                        cksContract_Click
                    End If
                End If
                imSetting = False
                mSetCommands
                Exit Sub
            End If
    End If
    
End Sub


Private Sub cbcMonth_Click()
    imMonthSelectedIndex = cbcMonth.ListIndex
    cbcMonth_Change
End Sub

Private Sub cksAdvertiser_Click()
    Dim ilLoop As Integer
    
    'If imAdvChkAll = True Then
    '    imAdvChkAll = False
    '    End Sub
    'End If
    
    
    If cksAdvertiser.Value = vbChecked Then
        If lbcAdvertiser.ListCount >= 1 Then
            For ilLoop = lbcAdvertiser.ListCount - 1 To 0 Step -1
                imAdvCheckAllProcess = True
                If ilLoop = 0 Then
                    imAdvCheckAllProcess = False
                End If
                If lbcAdvertiser.Selected(ilLoop) = True Then
                    lbcAdvertiser_Click
                Else
                    lbcAdvertiser.Selected(ilLoop) = True
                End If
                
            Next ilLoop
        End If
        imAdvChkAll = True
    Else
         imAdvChkAll = False
         If lbcAdvertiser.ListCount >= 1 And imAdvChkTemp = 0 Then
            For ilLoop = lbcAdvertiser.ListCount - 1 To 0 Step -1
                imAdvCheckAllProcess = True
                If ilLoop = 0 Then
                    imAdvCheckAllProcess = False
                End If
                If lbcAdvertiser.Selected(ilLoop) = False Then
                    lbcAdvertiser_Click
                Else
                    lbcAdvertiser.Selected(ilLoop) = False
                End If
                
            Next ilLoop
        End If
       
    End If
    
End Sub

Private Sub cksContract_Click()
    Dim ilLoop As Integer
    If cksContract.Value = vbChecked Then
        If lbcContracts.ListCount >= 1 Then
            For ilLoop = lbcContracts.ListCount - 1 To 0 Step -1
                If lbcContracts.Selected(ilLoop) = True Then
                    lbcContracts_Click
                Else
                    lbcContracts.Selected(ilLoop) = True
                End If
            Next ilLoop
           
        End If
        imContractChkAll = True
    Else
         imContractChkAll = False
         If lbcContracts.ListCount >= 1 And imContractChkTemp = 0 Then
            For ilLoop = lbcContracts.ListCount - 1 To 0 Step -1
                If lbcContracts.Selected(ilLoop) = False Then
                    lbcContracts_Click
                Else
                    lbcContracts.Selected(ilLoop) = False
                End If
            Next ilLoop
        End If
    
    End If
    
End Sub

Private Sub cmcCancel_Click()
    If igPodTargetCallSource <> CALLNONE Then
        If imFormModified = 1 Then
            sgPodTargetName = "test"
            igPodTargetCallSource = CALLDONE
        Else
            igMNmCallSource = CALLCANCELLED
        End If
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub


Private Sub cmcPost_GotFocus()
    'gCtrlGotFocus cmcPost
    'mSetShow imBoxNo
    'imBoxNo = -1
    
End Sub

Private Sub cmcPost_Click()
    Dim iRet As Integer
    Dim slMessage As String
    If edcContract.Text <> "" Then
        sgManualPostCpmFilter = imAdfCode & "\" & edcYear.Text & "\" & (imMonthSelectedIndex + 1) & "\" & Months(imMonthSelectedIndex) & "\" & imChfadfCodes & "\" & "(" & edcContract.Text & ")"
        iRet = mContractValid
        If Not iRet Then
            slMessage = "No matching records for the selected contract for the entered month and year"
            iRet = MsgBox(slMessage, vbOKOnly + vbQuestion)
            Exit Sub
        End If
    Else
        sgManualPostCpmFilter = imAdfCode & "\" & edcYear.Text & "\" & (imMonthSelectedIndex + 1) & "\" & Months(imMonthSelectedIndex) & "\" & imChfadfCodes & "\" & imChfCntrNo
    End If
    
    PostCpmManualCntrl.Show vbModal
End Sub

Private Sub edcContract_GotFocus()
    edcContract.BackColor = &HFFFF00
End Sub

Private Sub edcContract_KeyUp(KeyCode As Integer, Shift As Integer)
    mSetCommands
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub


Private Sub edcYear_GotFocus()
     edcYear.BackColor = &HFFFF00
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
    gShowBranner imUpdateAllowed
    mSetCommands
    
    Me.KeyPreview = True
    AdServerFilter.Refresh
    cmcPost.Enabled = False
    
End Sub
Private Sub Form_Click()
    'pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcMonth.Enabled) Then
            cbcMonth.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
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

    Set AdServerFilter = Nothing   'Remove data segment
    
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
    If imTerminate Then
        Exit Sub
    End If
    
    mInitBox
    gCenterStdAlone AdServerFilter
    Screen.MousePointer = vbHourglass
    imPopReqd = False
    imFirstFocus = True
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = True
    imIgnoreClick = False
    imBypassSetting = False
    imFormModified = 0
    imAdfCode = -1
    imAdvChkTemp = 0
    imContractChkTemp = 0
    imSetting = 0
    On Error GoTo 0
    If imTerminate Then
        Exit Sub
    End If
    If imTerminate Then
        Exit Sub
    End If
    
    'mPopulatePodCategory
    mPopulateAdServer
    mPopulateAdvertiser
    'If (Not tgPodCategory) <> -1 Then
    '    mLoadPodItems
    ' End If
    cbcMonth.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcMonth.ListIndex = 0 'This will generate a select_change event
        imMonthSelectedIndex = cbcMonth.ListIndex
        mSetCommands
    End If
    mInitializeFilter
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
    
    AdServerFilter.height = lgCurrVRes * 10 * 0.66
    flTextHeight = 0
    AdServerFilter.Width = lgCurrHRes * 10 * 0.66
    'Position panel and picture areas with panel
    lblYear.Move fgPanelAdj * 4, fgPanelAdj * 4 + fgBevelY
    edcYear.Move lblYear.Left + lblYear.Width, fgPanelAdj * 4
    
    lblMonth.Move edcYear.Left + edcYear.Width + fgBevelX * 4, fgPanelAdj * 4 + fgBevelY
    cbcMonth.Move lblMonth.Left + lblMonth.Width + fgBevelX * 2, fgPanelAdj * 4
    
    lblContractNumber.Move cbcMonth.Left + cbcMonth.Width + fgBevelX * 4, fgPanelAdj * 4 + fgBevelY
    edcContract.Move lblContractNumber.Left + lblContractNumber.Width + fgBevelX * 2, fgPanelAdj * 4
    
    lblOr.Move edcContract.Left + edcContract.Width + fgBevelX * 4, fgPanelAdj * 4 + fgBevelY
    
    plcPodTarget.Move fgPanelAdj * 4, cbcMonth.Top + cbcMonth.height + fgPanelAdj * 2, AdServerFilter.Width - fgPanelAdj * 8, AdServerFilter.height - (cmcCancel.height + cbcMonth.Top + cbcMonth.height + fgPanelAdj * 8)

    lblAdServerVendor.height = 255
    lblAdServerVendor.Move 0, fgPanelAdj + fgBevelX
    height = CSng(plcPodTarget.height + fgPanelAdj) - lblAdServerVendor.height
    lbcAdserver.Move 0, lblAdServerVendor.height + fgPanelAdj * 2, ((plcPodTarget.Width - fgPanelAdj * 4) * 0.3)
    lbcAdserver.height = height + fgPanelAdj
    
    
    lbcAdvertiser.Move lbcAdserver.Width + fgPanelAdj * 4, lbcAdserver.Top, ((plcPodTarget.Width - fgPanelAdj * 4) * 0.35)
    lbcAdvertiser.height = height + fgPanelAdj
    lblAdvertiser.height = 255
    cksAdvertiser.Move lbcAdvertiser.Left, fgPanelAdj
    lblAdvertiser.Move cksAdvertiser.Left + cksAdvertiser.Width, fgPanelAdj + fgBevelX
    
    
    lbcContracts.Move lbcAdvertiser.Width + lbcAdserver.Width + fgPanelAdj * 8, lbcAdserver.Top, ((plcPodTarget.Width - fgPanelAdj * 4) * 0.35)
    lbcContracts.height = height + fgPanelAdj
    lblContract.height = 255
    cksContract.Move lbcContracts.Left, fgPanelAdj
    lblContract.Move cksContract.Left + cksAdvertiser.Width, fgPanelAdj + fgBevelX
    
    
    'plcPodTarget.height = cksContract.Top + cksContract.height
    cmcCancel.Move (AdServerFilter.Width / 2) - (cmcCancel.Width + fgPanelAdj), plcPodTarget.Top + plcPodTarget.height + fgPanelAdj * 2
    cmcPost.Move cmcCancel.Left + cmcCancel.Width + fgPanelAdj, cmcCancel.Top
    imMaxRowListItem = CInt(lbcContracts.height / fgListHtSerif825)
    
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*            Created:15/2/21      By:L.Bianchi     *
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
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    gPopMonth
    imPopReqd = True
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    imSetting = 0
    Unload AdServerFilter
    igManUnload = NO
End Sub

Private Sub lbcAdserver_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    'imIgnoreClick = False
    imAdfCode = -1
    For iLoop = 0 To lbcAdserver.ListCount - 1 Step 1
        If lbcAdserver.Selected(iLoop) Then
            'imPodCategoryCode = lbcAdserver.ItemData(iLoop)
            imAdfCode = lbcAdserver.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    
    'If iCount = 1 Then
        'lbcAdvertiser.Clear
        'edcFilter.Text = ""
       ' mFillPodItems
    'End If
End Sub

Private Sub lbcAdvertiser_Click()
    Dim iLoop As Integer
    Dim iSelectedAny As Integer
    imChfadfCodes = ""
    imChfCntrNo = ""
    If imAdvCheckAllProcess Then
        Exit Sub
    End If

    If imAdvChkAll = True Then
        imAdvChkTemp = 1
        cksAdvertiser.Value = vbUnchecked
    End If
    imAdvChkTemp = 0
     
    For iLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
        If lbcAdvertiser.Selected(iLoop) Then
            imChfadfCodes = imChfadfCodes & lbcAdvertiser.ItemData(iLoop) & ","
            If Not iSelectedAny Then
                iSelectedAny = True
            End If
        End If
    Next iLoop
    If iSelectedAny Then
        edcContract.Text = ""
    End If
    If (imChfadfCodes <> "") Then
        imChfadfCodes = "(" & Left$(imChfadfCodes, Len(imChfadfCodes) - 1) & ")"
        mPopulateContracts
        If imContractChkAll Then
            imSetting = True
            cksContract_Click
        End If
    Else
        lbcContracts.Clear
        imChfCntrNo = ""
    End If
    imSetting = False
    mSetCommands
End Sub

Private Sub lbcContracts_Click()
    Dim iLoop As Integer
    imChfCntrNo = ""
    
    If imContractChkAll = True And Not imSetting Then
        imContractChkTemp = 1
        cksContract.Value = vbUnchecked
    End If
    imContractChkTemp = 0
    
    For iLoop = 0 To lbcContracts.ListCount - 1 Step 1
        If lbcContracts.Selected(iLoop) Then
            imChfCntrNo = imChfCntrNo & lbcContracts.ItemData(iLoop) & ","
        End If
    Next iLoop
    If (imChfCntrNo <> "") Then
        imChfCntrNo = "(" & Left$(imChfCntrNo, Len(imChfCntrNo) - 1) & ")"
    End If
    mSetCommands
End Sub


Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
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

Private Function gPopMonth()
    Months(0) = "January"
    Months(1) = "February"
    Months(2) = "March"
    Months(3) = "April"
    Months(4) = "May"
    Months(5) = "June"
    Months(6) = "July"
    Months(7) = "August"
    Months(8) = "September"
    Months(9) = "October"
    Months(10) = "November"
    Months(11) = "December"

    Dim i As Integer
    cbcMonth.Clear
    For i = 0 To 11 Step 1
        cbcMonth.AddItem Trim$(Months(i))
        cbcMonth.ItemData(cbcMonth.NewIndex) = i
    Next i
End Function

Private Sub mPopulateAdServer()
    Dim Index As Integer
    Dim iLoop As Integer
    Dim AVF As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT avfcode, avfname FROM avf_AdVendor order by avfname"
    Set AVF = gSQLSelectCall(SQLQuery)
    
    While Not AVF.EOF
        ReDim Preserve tgAdServerVendor(Index)
        tgAdServerVendor(Index).iCode = AVF!avfCode
        tgAdServerVendor(Index).sName = Trim$(AVF!avfName)
        Index = Index + 1
        AVF.MoveNext
    Wend
    AVF.Close
    lbcAdserver.Clear
    If (Not tgAdServerVendor) <> -1 Then
        For iLoop = 0 To UBound(tgAdServerVendor) Step 1
            lbcAdserver.AddItem tgAdServerVendor(iLoop).sName
            lbcAdserver.ItemData(lbcAdserver.NewIndex) = tgAdServerVendor(iLoop).iCode
        Next iLoop
            lbcAdserver.ListIndex = 0
            imAdfCode = lbcAdserver.ItemData(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mPopulateAdvertiser()
    Dim Index As Integer
    Dim iLoop As Integer
    Dim ADV As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT adfcode, adfname FROM ADF_Advertisers order by adfname"
    Set ADV = gSQLSelectCall(SQLQuery)
    
    While Not ADV.EOF
        ReDim Preserve tgAdvertiser(Index)
        tgAdvertiser(Index).iCode = ADV!adfCode
        tgAdvertiser(Index).sName = Trim$(ADV!adfName)
        Index = Index + 1
        ADV.MoveNext
    Wend
    ADV.Close
    lbcAdvertiser.Clear
    If (Not tgAdvertiser) <> -1 Then
        For iLoop = 0 To UBound(tgAdvertiser) Step 1
            lbcAdvertiser.AddItem tgAdvertiser(iLoop).sName
            lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = tgAdvertiser(iLoop).iCode
        Next iLoop
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub edcYear_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String
        'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcYear.Text
        slStr = Left$(slStr, edcYear.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcYear.SelStart - edcYear.SelLength)
        slComp = "9999"
        If gCompNumberStr(slStr, slComp) > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    
End Sub

Private Sub edcContract_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String
        'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcContract.Text
        slStr = Left$(slStr, edcContract.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcContract.SelStart - edcContract.SelLength)
         ''' TTP 10565 BEGIN JJB
        slComp = "99999999"
        'slComp = "9999"
        ''' TTP 10565 END
        If gCompNumberStr(slStr, slComp) > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    
End Sub

Private Sub edcContract_LostFocus()
    edcContract.BackColor = &H80000005
    Dim ilLoop As Integer
    
    If edcContract.Text = "" Then
        Exit Sub
    End If
    
    If cksAdvertiser.Value = vbChecked Then
        cksAdvertiser.Value = vbUnchecked
    Else
         If lbcAdvertiser.ListCount > 1 Then
            For ilLoop = lbcAdvertiser.ListCount - 1 To 0 Step -1
                imAdvCheckAllProcess = True
                If ilLoop = 0 Then
                    imAdvCheckAllProcess = False
                End If
                If lbcAdvertiser.Selected(ilLoop) = False Then
                    lbcAdvertiser_Click
                Else
                    lbcAdvertiser.Selected(ilLoop) = False
                End If
            Next ilLoop
        End If
    End If
End Sub

Private Sub edcYear_KeyUp(KeyCode As Integer, Shift As Integer)
    mSetCommands
End Sub

Private Sub edcYear_Lostfocus()
  edcYear.BackColor = &H80000005
  
  If (imChfadfCodes <> "") Then
        mPopulateContracts
        If imContractChkAll Then
            imSetting = True
            cksContract_Click
            imSetting = False
        End If
  End If
  mSetCommands
End Sub

Private Sub mSetCommands()
    If (edcYear.Text = "" Or Len(edcYear.Text) < 4) Or (imChfCntrNo = "" And edcContract.Text = "") Then
        cmcPost.Enabled = False
    Else
        cmcPost.Enabled = True
    End If
End Sub

Public Function ArrayIsInitialized(ByRef Arr As Variant) As Boolean
    On Error Resume Next
    ArrayIsInitialized = False
    If UBound(Arr) >= 0 Then If err.Number = 0 Then ArrayIsInitialized = True
End Function

Private Sub mPopulateContracts()
    
    Dim CHF As ADODB.Recordset
    Dim Index As Integer
    Dim iLoop As Integer
    Dim month As Integer
    Dim year As String
    Dim daysInMonth As Integer
    Dim ver As String
    
    Dim StartDate_CAL As String
    Dim EndDate_CAL As String
    
    Dim StartDate_STD As String
    Dim EndDate_STD As String
    
    Screen.MousePointer = vbHourglass
    
'   BEGIN TTP 10671 2023-03-21 JJB
'   NOTE:  The original code was only filtering for CALENDAR dates in the query and not accounting
'          for Standard broadcast calendar dates.  The below code accounts for both methods but does NOT
'          account for WEEKLY billing cycles.

    month = imMonthSelectedIndex + 1
    year = edcYear.Text
    
    If year = "" Or Len(year) < 4 Or imChfadfCodes = "" Then Exit Sub

    ' Calendar Dates
    StartDate_CAL = gObtainStartCal(year & "-" & month & "-" & "15")
    EndDate_CAL = gObtainEndCal(year & "-" & month & "-" & "15")
    ' Standard Dates
    StartDate_STD = gObtainStartStd(year & "-" & month & "-" & "15")
    EndDate_STD = gObtainEndStd(year & "-" & month & "-" & "15")
    
    SQLQuery = " select "
    SQLQuery = SQLQuery & vbCrLf & "     chfProduct, "
    SQLQuery = SQLQuery & vbCrLf & "     chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "     chfExtRevNo, "
    SQLQuery = SQLQuery & vbCrLf & "     chfCntRevNo, "
    SQLQuery = SQLQuery & vbCrLf & "     chfstatus ,"
    SQLQuery = SQLQuery & vbCrLf & "     (select Count(pcfCode) As adServerCount from pcf_Pod_CPM_Cntr WHERE pcfChfCOde = chfCode) "
    SQLQuery = SQLQuery & vbCrLf & " from "
    SQLQuery = SQLQuery & vbCrLf & "     CHF_Contract_Header "
    SQLQuery = SQLQuery & vbCrLf & " where "
    SQLQuery = SQLQuery & vbCrLf & "     chfDelete       = 'N' and "
    SQLQuery = SQLQuery & vbCrLf & "     chfSchStatus    = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "     adServerCount    > 0  and "
    SQLQuery = SQLQuery & vbCrLf & "     chfadfCode in " & imChfadfCodes
    SQLQuery = SQLQuery & vbCrLf & "     and "
    SQLQuery = SQLQuery & vbCrLf & "     ( "
    SQLQuery = SQLQuery & vbCrLf & "         ( "
    SQLQuery = SQLQuery & vbCrLf & "             chfBillCycle = 'S' and "
    SQLQuery = SQLQuery & vbCrLf & "             ("
    SQLQuery = SQLQuery & vbCrLf & "                 ('" & Format(StartDate_STD, "yyyy-m-d") & "' between chfStartDate and chfEndDate ) "
    SQLQuery = SQLQuery & vbCrLf & "                 or "
    SQLQuery = SQLQuery & vbCrLf & "                 ('" & Format(EndDate_STD, "yyyy-m-d") & "' between chfStartDate and chfEndDate ) "
    SQLQuery = SQLQuery & vbCrLf & "                 or "
    SQLQuery = SQLQuery & vbCrLf & "                 ('" & Format(StartDate_STD, "yyyy-mm-dd") & "' < chfStartDate and '" & Format(EndDate_STD, "yyyy-mm-dd") & "' > chfEndDate) "
    SQLQuery = SQLQuery & vbCrLf & "             )"
    SQLQuery = SQLQuery & vbCrLf & "         ) "
    SQLQuery = SQLQuery & vbCrLf & "         Or "
    SQLQuery = SQLQuery & vbCrLf & "         ( "
    SQLQuery = SQLQuery & vbCrLf & "             chfBillCycle = 'C' and "
    SQLQuery = SQLQuery & vbCrLf & "             ("
    SQLQuery = SQLQuery & vbCrLf & "                 ('" & Format(StartDate_CAL, "yyyy-m-d") & "' between chfStartDate and chfEndDate ) "
    SQLQuery = SQLQuery & vbCrLf & "                 or "
    SQLQuery = SQLQuery & vbCrLf & "                 ('" & Format(EndDate_CAL, "yyyy-m-d") & "' between chfStartDate and chfEndDate ) "
    SQLQuery = SQLQuery & vbCrLf & "                 or "
    SQLQuery = SQLQuery & vbCrLf & "                 ('" & Format(StartDate_CAL, "yyyy-mm-dd") & "' < chfStartDate and '" & Format(EndDate_CAL, "yyyy-mm-dd") & "' > chfEndDate) "
    SQLQuery = SQLQuery & vbCrLf & "             )"
    SQLQuery = SQLQuery & vbCrLf & "         ) "
    SQLQuery = SQLQuery & vbCrLf & "     ) "
    SQLQuery = SQLQuery & vbCrLf & " order by "
    SQLQuery = SQLQuery & vbCrLf & "     chfCntrNo desc, "
    SQLQuery = SQLQuery & vbCrLf & "     chfCntRevNo desc, "
    SQLQuery = SQLQuery & vbCrLf & "     chfExtRevNo desc "
    
'   END TTP 10671 2023-03-21 JJB

    Set CHF = gSQLSelectCall(SQLQuery)
    
    lbcContracts.Clear
    While Not CHF.EOF
         ver = ""
         If (CHF!chfstatus = "W") Or (CHF!chfstatus = "C") Or (CHF!chfstatus = "I") Or (CHF!chfstatus = "D") Then
            If (CHF!chfCntRevNo > 0) Then
                ver = " R" & Trim$(str$(CHF!chfCntRevNo)) & "-" & Trim$(str$(CHF!chfExtRevNo))
            Else
                ver = " V" & Trim$(str$(CHF!chfCntRevNo))
            End If
        Else
            ver = " R" & Trim$(str$(CHF!chfCntRevNo)) & "-" & Trim$(str$(CHF!chfExtRevNo))
        End If
        
        lbcContracts.AddItem CHF!chfCntrNo & " " & ver & " " & CHF!chfProduct
        lbcContracts.ItemData(lbcContracts.NewIndex) = CHF!chfCntrNo
        Index = Index + 1
        CHF.MoveNext
    Wend
    
    CHF.Close
    Screen.MousePointer = vbDefault
    
End Sub

Private Function mContractValid() As Integer

    Dim CHF As ADODB.Recordset
    Dim month As Integer
    Dim year As String
    Dim daysInMonth As Integer
    Dim StartDate As String
    Dim EndDate As String
    Dim sContract As String
    Dim ilHasRecord As Integer
    
    month = imMonthSelectedIndex + 1
    year = edcYear.Text
    
    daysInMonth = CStr(Day(DateSerial(year, month + 1, 1) - 1))
    
    StartDate = year & "-" & month & "-" & "1"
    EndDate = year & "-" & month & "-" & daysInMonth
    
    sContract = edcContract.Text
    
    Screen.MousePointer = vbHourglass
    
    SQLQuery = "select chfCntrNo,(select Count(pcfCode) As adServerCount from pcf_Pod_CPM_Cntr WHERE pcfChfCOde = chfCode)" & vbCrLf
    SQLQuery = SQLQuery & "from CHF_Contract_Header where chfDelete = 'N' and chfSchStatus = 'F' and adServerCount > 0" & vbCrLf
    SQLQuery = SQLQuery & "and chfCntrNo in (" & sContract & ")" & vbCrLf
    SQLQuery = SQLQuery & "and (( '" & StartDate & "' between chfStartDate and chfEndDate )" & vbCrLf
    SQLQuery = SQLQuery & "or ('" & EndDate & "' between chfStartDate and chfEndDate )" & vbCrLf
    SQLQuery = SQLQuery & "or ('" & Format(StartDate, "yyyy-mm-dd") & "' < chfStartDate and '" & Format(EndDate, "yyyy-mm-dd") & "' > chfEndDate)) " & vbCrLf
    Set CHF = gSQLSelectCall(SQLQuery)
    
    mContractValid = False
    
    While Not CHF.EOF
        mContractValid = True
        CHF.MoveNext
    Wend
    CHF.Close
    Screen.MousePointer = vbDefault

End Function

Private Sub lblAdvertiser_Click()
    If cksAdvertiser.Value = vbChecked Then
    cksAdvertiser.Value = vbUnchecked
    Else
    cksAdvertiser.Value = vbChecked
    End If
    
End Sub

Private Sub lblContract_Click()
    If cksContract.Value = vbChecked Then
        cksContract.Value = vbUnchecked
    Else
        cksContract.Value = vbChecked
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitializeFilter               *
'*                                                     *
'*             Created:05/05/21      By:L.Bianchi      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Filter Value        *
'*                                                     *
'*******************************************************
Private Sub mInitializeFilter()
    Dim SPF As ADODB.Recordset
    Dim dateSet As Integer
    Dim lastInvoicedDate As Date
    
    dateSet = False
    SQLQuery = "select spfBLastStdMnth  from SPF_Site_Options"
    
    Set SPF = gSQLSelectCall(SQLQuery)
    
    While Not SPF.EOF
        If Not IsNull(SPF("spfBLastStdMnth")) Then
            lastInvoicedDate = SPF!spfBLastStdMnth
            lastInvoicedDate = DateAdd("m", 1, lastInvoicedDate)
            dateSet = True
        End If
        SPF.MoveNext
    Wend
    SPF.Close
    If Not dateSet Then
        lastInvoicedDate = Now
    End If
    
    edcYear.Text = Format(lastInvoicedDate, "yyyy")
    cbcMonth.ListIndex = CInt(Format(lastInvoicedDate, "mm")) - 1
    
End Sub
