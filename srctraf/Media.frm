VERSION 5.00
Begin VB.Form Media 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4440
   ClientLeft      =   990
   ClientTop       =   2610
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
   ScaleHeight     =   4440
   ScaleWidth      =   5310
   Begin VB.TextBox edcIPumpNameSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   615
      MaxLength       =   60
      TabIndex        =   16
      Top             =   3075
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.TextBox edcIPumpNetworkID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3690
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox edcIPumpEventType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   14
      Top             =   2745
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox edcIPumpSuffix 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   1605
      MaxLength       =   6
      TabIndex        =   13
      Top             =   2730
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox edcIPumpPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   600
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcVehicle 
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
      Left            =   2895
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2925
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   165
      Width           =   1830
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4740
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2355
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4725
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2685
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4740
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2985
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
      ScaleWidth      =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3120
      Width           =   90
   End
   Begin VB.ListBox lbcInv 
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
      Left            =   585
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1845
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ListBox lbcTape 
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
      Left            =   2685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ListBox lbcCart 
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
      Left            =   600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   2025
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
      Left            =   1890
      Picture         =   "Media.frx":0000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   870
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1110
      ScaleHeight     =   210
      ScaleWidth      =   1395
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1785
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcAssignNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   19
      Top             =   1845
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcGroupNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcAssignNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   615
      MaxLength       =   5
      TabIndex        =   18
      Top             =   1845
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox edcPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2655
      MaxLength       =   6
      TabIndex        =   5
      Top             =   810
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   600
      MaxLength       =   6
      TabIndex        =   4
      Top             =   810
      Visible         =   0   'False
      Width           =   2025
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
      Left            =   3225
      TabIndex        =   26
      Top             =   3975
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
      Left            =   2085
      TabIndex        =   25
      Top             =   3975
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
      Left            =   960
      TabIndex        =   24
      Top             =   3975
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
      Left            =   3225
      TabIndex        =   23
      Top             =   3630
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
      Left            =   2100
      TabIndex        =   22
      Top             =   3630
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
      Left            =   975
      TabIndex        =   21
      Top             =   3630
      Width           =   1050
   End
   Begin VB.PictureBox pbcMcd 
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
      Height          =   2670
      Left            =   570
      Picture         =   "Media.frx":00FA
      ScaleHeight     =   2670
      ScaleWidth      =   4110
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   660
      Width           =   4110
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   195
      ScaleHeight     =   135
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   615
      Width           =   150
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   300
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   20
      Top             =   1575
      Width           =   105
   End
   Begin VB.PictureBox plcMcd 
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   525
      ScaleHeight     =   2715
      ScaleWidth      =   4170
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   4230
   End
   Begin VB.Label plcScreen 
      Caption         =   "Media Definitions"
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
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   1560
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   2805
      Width           =   360
   End
End
Attribute VB_Name = "Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Media.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Media.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Media input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim tmInvCode() As SORTCODE
Dim smInvCodeTag As String
'Media Field Areas
Dim tmCtrls(0 To 16)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Media Box
Dim tmMcf As MCF        'MCF record image
Dim tmMMcf() As MCF     'Retain all Media codes to check that name not used
Dim tmCif As CIF        'CIF record image
Dim tmSrchKey As INTKEY0    'MCF key record image
Dim imRecLen As Integer        'MCF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmMcf As Integer 'Media file handle
Dim hmCif As Integer 'Inventory file handle
Dim hmMef As Integer 'Media Extra file handle
Dim tmMef As MEF        'MEF record image
Dim imMefRecLen As Integer        'MEF record length
Dim tmMefSrchKey As INTKEY0    'MEF key record image
Dim tmMefSrchKey1 As MEFKEY1    'MEF key record image
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imScript As Integer '0=Yes; 1=No
Dim imReUse As Integer  '0=Yes; 1=No
Dim imSuppressExport As Integer  '0=Yes; 1=No
Dim imSortCart As Integer   '0=Last Used Date; 1=Cart Number
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imUpdateAllowed As Integer    'User can update records
Dim smVehicle As String
Dim smOrigAssign(0 To 1) As String

Dim tmVehicleCode() As SORTCODE
Dim smVehicleCodeTag As String

Dim imMaxCtrl As Integer

Const NAMEINDEX = 1     'Description control/field
Const PREFIXINDEX = 2   'Prefix control/field
Const REUSEINDEX = 3    'Reuse control/index
Const SORTCARTINDEX = 4     'Sort Cart control/index
Const CARTINDEX = 5     'Cart control/index
Const SUPPRESSEXPORTINDEX = 6     'Tape control/index
Const LOWNOINDEX = 7    'Lowest Assign No control/field
Const HIGHNOINDEX = 8   'Highest Assign No control/field
Const VEHINDEX = 9
Const SCRIPTINDEX = 10  'Script control/index
Const GROUPNOINDEX = 11 'Group # control/field
Const IPUMPPREFIXINDEX = 12
Const IPUMPSUFFIXINDEX = 13
Const IPUMPEVENTTYPEINDEX = 14
Const IPUMPNETWORKIDINDEX = 15
Const IPUMPNAMESPACEINDEX = 16

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
    pbcMcd.Cls
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
    For ilLoop = imLBCtrls To imMaxCtrl Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcMcd_Paint
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
        If igMcdCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgMcdName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgMcdName  'New Name
            End If
            cbcSelect_Change
            If sgMcdName <> "" Then
                mSetCommands
                gFindMatch sgMcdName, 1, cbcSelect
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
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or Trim$(cbcSelect.Text) = "" Or imPopReqd Then
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
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cmcCancel_Click()
    If igMcdCallSource <> CALLNONE Then
'        If igMcdCallSource = CALLSOURCEENAME Then
'            igMcdCallSource = CALLCANCELLED
'        End If
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
    If igMcdCallSource <> CALLNONE Then
        sgMcdName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgMcdName = "[New]"
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
    If igMcdCallSource <> CALLNONE Then
'        If igMcdCallSource = CALLSOURCEENAME Then
'            If sgMcdName = "[New]" Then
'                igMcdCallSource = CALLCANCELLED
'            Else
'                igMcdCallSource = CALLDONE
'            End If
'        End If
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
        For ilLoop = imLBCtrls To imMaxCtrl Step 1
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
        Case CARTINDEX
            lbcCart.Visible = Not lbcCart.Visible
        'Case SUPPRESSEXPORTINDEX
        '    lbcTape.Visible = Not lbcTape.Visible
        Case VEHINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
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
        ilRet = gIICodeRefExist(Media, tmMcf.iCode, "Cif.Btr", "CifMcfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Inventory references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Media, tmMcf.iCode, "Fnf.Btr", "FnfMcfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Name references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmMcf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Mcf.btr")
        ilRet = btrDelete(hmMcf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Media
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        ''If lbcNameCode.Tag <> "" Then
        ''    If slStamp = lbcNameCode.Tag Then
        ''        lbcNameCode.Tag = FileDateTime(sgDBPath & "Mcf.btr")
        ''    End If
        ''End If
        'If sgNameCodeTag <> "" Then
        '    If slStamp = sgNameCodeTag Then
        '        sgNameCodeTag = gFileDateTime(sgDBPath & "Mcf.btr")
        '    End If
        'End If
        ''lbcNameCode.RemoveItem imSelectedIndex - 1
        'gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        'cbcSelect.RemoveItem imSelectedIndex
        mPopulate
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcMcd.Cls
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
    igRptCallType = MEDIADEFINITIONSLIST
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Media^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "Media^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Media^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "Media^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'Media.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'Media.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    ''Screen.MousePointer = vbDefault    'Default
    sgCommandStr = slStr
    RptList.Show vbModal
End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus ActiveControl
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
        pbcMcd.Cls
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To imMaxCtrl Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcMcd_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcMcd.Cls
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
    slName = edcName.Text   'Save name
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
    '    cbcSelect.Text = slName
        gFindMatch slName, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            cbcSelect.ListIndex = gLastFound(cbcSelect)
            cbcSelect.Text = cbcSelect.List(cbcSelect.ListIndex)
        Else
            cbcSelect.ListIndex = 0
        End If
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
Private Sub edcAssignNo_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcAssignNo_GotFocus(Index As Integer)
    gCtrlGotFocus edcAssignNo(Index)
End Sub
Private Sub edcAssignNo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And (KeyAscii <> KEYNEG) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And ((KeyAscii < KEYUA) Or (KeyAscii > KEYUZ)) And ((KeyAscii < KEYLA) Or (KeyAscii > KEYLZ)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_Change()
    Select Case imBoxNo
        Case CARTINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCart, imBSMode, imComboBoxIndex
        'Case SUPPRESSEXPORTINDEX
        '    imLbcArrowSetting = True
        '    gMatchLookAhead edcDropDown, lbcTape, imBSMode, imComboBoxIndex
        Case VEHINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case CARTINDEX
            If lbcCart.ListCount = 1 Then
                lbcCart.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        'Case SUPPRESSEXPORTINDEX
        '    If lbcTape.ListCount = 1 Then
        '        lbcTape.ListIndex = 0
        '        'If imTabDirection = -1 Then  'Right To Left
        '        '    pbcSTab.SetFocus
        '        'Else
        '        '    pbcTab.SetFocus
        '        'End If
        '        'Exit Sub
        '    End If
        Case VEHINDEX
            If lbcVehicle.ListCount = 1 Then
                lbcVehicle.ListIndex = 0
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
            Case CARTINDEX
                gProcessArrowKey Shift, KeyCode, lbcCart, imLbcArrowSetting
            'Case SUPPRESSEXPORTINDEX
            '    gProcessArrowKey Shift, KeyCode, lbcTape, imLbcArrowSetting
            Case VEHINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcGroupNo_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcGroupNo_GotFocus()
    gCtrlGotFocus edcGroupNo
End Sub
Private Sub edcGroupNo_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcGroupNo.Text
    slStr = Left$(slStr, edcGroupNo.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGroupNo.SelStart - edcGroupNo.SelLength)
    If gCompNumberStr(slStr, "99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcIPumpEventType_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcIPumpEventType_GotFocus()
    gCtrlGotFocus edcGroupNo
End Sub

Private Sub edcIPumpNameSpace_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcIPumpNameSpace_GotFocus()
    gCtrlGotFocus edcGroupNo
End Sub

Private Sub edcIPumpNetworkID_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcIPumpNetworkID_GotFocus()
    gCtrlGotFocus edcGroupNo
End Sub

Private Sub edcIPumpPrefix_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcIPumpPrefix_GotFocus()
    gCtrlGotFocus edcGroupNo
End Sub

Private Sub edcIPumpSuffix_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcIPumpSuffix_GotFocus()
    gCtrlGotFocus edcGroupNo
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcName_Change()
    mSetChg NAMEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
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
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub edcName_LostFocus()
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
End Sub
Private Sub edcPrefix_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcPrefix_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcPrefix_KeyPress(KeyAscii As Integer)
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
    If (igWinStatus(MEDIADEFINITIONSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcMcd.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcMcd.Enabled = True
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
'    For ilLoop = 1 To igDDEDelay Step 1
'        DoEvents
'    Next ilLoop
'    gShowBranner
    Me.KeyPreview = True
    Media.Refresh
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
        '    mEnableBox imBoxNo
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

    Erase tmInvCode
    Erase tmMMcf

    btrExtClear hmMef   'Clear any previous extend operation
    ilRet = btrClose(hmMef)
    btrDestroy hmMef
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    
    Set Media = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcCart_Click()
    gProcessLbcClick lbcCart, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcCart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcTape_Click()
    gProcessLbcClick lbcTape, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcTape_GotFocus()
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
    edcName.Text = ""
    edcPrefix.Text = ""
    imReUse = -1
    imSortCart = -1
    imSuppressExport = -1
    lbcCart.ListIndex = -1
    lbcTape.ListIndex = -1
    For ilLoop = 0 To 1 Step 1
        edcAssignNo(ilLoop).Text = ""
        smOrigAssign(ilLoop) = ""
    Next ilLoop
    lbcVehicle.ListIndex = -1
    smVehicle = ""
    imScript = -1    'This sets Text = ""
    edcGroupNo.Text = ""
    edcIPumpPrefix.Text = ""
    edcIPumpSuffix.Text = ""
    edcIPumpEventType.Text = ""
    edcIPumpNetworkID.Text = ""
    edcIPumpNameSpace.Text = ""
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrl) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 6
            gMoveFormCtrl pbcMcd, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case PREFIXINDEX 'Market Name
            edcPrefix.Width = tmCtrls(ilBoxNo).fBoxW
            edcPrefix.MaxLength = 6
            gMoveFormCtrl pbcMcd, edcPrefix, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcPrefix.Visible = True  'Set visibility
            edcPrefix.SetFocus
        Case REUSEINDEX 'Reuse numbers
            If imReUse < 0 Then
                imReUse = 1    'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcMcd, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SORTCARTINDEX 'Sort Cart Inventory code
            If imSortCart < 0 Then
                imSortCart = 0    'Last Used Date
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcMcd, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case CARTINDEX   'Cart disposition
            lbcCart.Height = gListBoxHeight(lbcCart.ListCount, 4)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 17
            gMoveFormCtrl pbcMcd, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcCart.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcCart.ListIndex < 0 Then
                lbcCart.ListIndex = 0
            End If
            imComboBoxIndex = lbcCart.ListIndex
            If lbcCart.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcCart.List(lbcCart.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SUPPRESSEXPORTINDEX   'Tape disposition
            'lbcTape.Height = gListBoxHeight(lbcTape.ListCount, 4)
            'edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            'edcDropDown.MaxLength = 17
            'gMoveFormCtrl pbcMcd, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            'cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            'lbcTape.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            'imChgMode = True
            'If lbcTape.ListIndex < 0 Then
            '    lbcTape.ListIndex = 0
            'End If
            'imComboBoxIndex = lbcTape.ListIndex
            'If lbcTape.ListIndex < 0 Then
            '    edcDropDown.Text = ""
            'Else
            '    edcDropDown.Text = lbcTape.List(lbcTape.ListIndex)
            'End If
            'imChgMode = False
            'edcDropDown.SelStart = 0
            'edcDropDown.SelLength = Len(edcDropDown.Text)
            'edcDropDown.Visible = True
            'cmcDropDown.Visible = True
            'edcDropDown.SetFocus
            If imSuppressExport < 0 Then
                imSuppressExport = 1    'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcMcd, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case LOWNOINDEX 'AssignNo
            edcAssignNo(0).Width = tmCtrls(ilBoxNo).fBoxW
            edcAssignNo(0).MaxLength = 5
            gMoveFormCtrl pbcMcd, edcAssignNo(0), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAssignNo(0).Visible = True  'Set visibility
            edcAssignNo(0).SetFocus
        Case HIGHNOINDEX 'AssignNo
            edcAssignNo(1).Width = tmCtrls(ilBoxNo).fBoxW
            edcAssignNo(1).MaxLength = 5
            gMoveFormCtrl pbcMcd, edcAssignNo(1), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAssignNo(1).Visible = True  'Set visibility
            edcAssignNo(1).SetFocus
        Case VEHINDEX
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcMcd, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            imChgMode = True
            If lbcVehicle.ListIndex < 0 Then
                lbcVehicle.ListIndex = 0
            End If
            imComboBoxIndex = lbcVehicle.ListIndex
            If lbcVehicle.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SCRIPTINDEX '# Full Months
            If imScript < 0 Then
                imScript = 1    'No
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcMcd, pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case GROUPNOINDEX 'Dial Position
            edcGroupNo.Width = tmCtrls(ilBoxNo).fBoxW
            edcGroupNo.MaxLength = 5
            gMoveFormCtrl pbcMcd, edcGroupNo, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcGroupNo.Visible = True  'Set visibility
            edcGroupNo.SetFocus
        Case IPUMPPREFIXINDEX 'Dial Position
            edcIPumpPrefix.Width = tmCtrls(ilBoxNo).fBoxW
            edcIPumpPrefix.MaxLength = 6
            gMoveFormCtrl pbcMcd, edcIPumpPrefix, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIPumpPrefix.Visible = True  'Set visibility
            edcIPumpPrefix.SetFocus
        Case IPUMPSUFFIXINDEX 'Dial Position
            edcIPumpSuffix.Width = tmCtrls(ilBoxNo).fBoxW
            edcIPumpSuffix.MaxLength = 6
            gMoveFormCtrl pbcMcd, edcIPumpSuffix, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIPumpSuffix.Visible = True  'Set visibility
            edcIPumpSuffix.SetFocus
        Case IPUMPEVENTTYPEINDEX 'Dial Position
            edcIPumpEventType.Width = tmCtrls(ilBoxNo).fBoxW
            edcIPumpEventType.MaxLength = 2
            gMoveFormCtrl pbcMcd, edcIPumpEventType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIPumpEventType.Visible = True  'Set visibility
            edcIPumpEventType.SetFocus
        Case IPUMPNETWORKIDINDEX 'Dial Position
            edcIPumpNetworkID.Width = tmCtrls(ilBoxNo).fBoxW
            edcIPumpNetworkID.MaxLength = 2
            gMoveFormCtrl pbcMcd, edcIPumpNetworkID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIPumpNetworkID.Visible = True  'Set visibility
            edcIPumpNetworkID.SetFocus
        Case IPUMPNAMESPACEINDEX 'Dial Position
            edcIPumpNameSpace.Width = tmCtrls(ilBoxNo).fBoxW
            edcIPumpNameSpace.MaxLength = 60
            gMoveFormCtrl pbcMcd, edcIPumpNameSpace, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcIPumpNameSpace.Visible = True  'Set visibility
            edcIPumpNameSpace.SetFocus
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
    imFirstActivate = True
    imTerminate = False
    Screen.MousePointer = vbHourglass
    imLBCtrls = 0
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    Media.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone Media
    'Media.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imDoubleClickName = False
    imLbcMouseDown = False
    imRecLen = Len(tmMcf)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmMcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MCF.BTR)", Media
    On Error GoTo 0
    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CIF.BTR)", Media
    On Error GoTo 0
    hmMef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMef, "", sgDBPath & "Mef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MEF.BTR)", Media
    On Error GoTo 0
    imMefRecLen = Len(tmMef)
    lbcCart.AddItem "Not Applicable"
    lbcCart.AddItem "Save"
    lbcCart.AddItem "Purge"
    lbcCart.AddItem "Ask after Expired"
    lbcTape.AddItem "Not Applicable"
    lbcTape.AddItem "Return"
    lbcTape.AddItem "Destroy"
    lbcTape.AddItem "Ask after Expired"
'    gCenterModalForm Media
'    Traffic!plcHelp.Caption = ""
    lbcVehicle.Clear
    mVehPop
    cbcSelect.Clear  'Force list to be populated
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
'*      Procedure Name:mInitBox                        *
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
    flTextHeight = pbcMcd.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcMcd.Move 525, 600, pbcMcd.Width + fgPanelAdj, pbcMcd.Height + fgPanelAdj
    pbcMcd.Move plcMcd.Left + fgBevelX, plcMcd.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2025, fgBoxStH
    'Prefix
    gSetCtrl tmCtrls(PREFIXINDEX), 2070, tmCtrls(NAMEINDEX).fBoxY, 2025, fgBoxStH
    'Reuse
    gSetCtrl tmCtrls(REUSEINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2025, fgBoxStH
    'Temporary
    gSetCtrl tmCtrls(SORTCARTINDEX), 2070, tmCtrls(REUSEINDEX).fBoxY, 2025, fgBoxStH
    'Cart
    gSetCtrl tmCtrls(CARTINDEX), 30, tmCtrls(REUSEINDEX).fBoxY + fgStDeltaY, 2025, fgBoxStH
    'Tape
    gSetCtrl tmCtrls(SUPPRESSEXPORTINDEX), 2070, tmCtrls(CARTINDEX).fBoxY, 2025, fgBoxStH
    'AssignNo
    gSetCtrl tmCtrls(LOWNOINDEX), 30, tmCtrls(CARTINDEX).fBoxY + fgStDeltaY, 2025, fgBoxStH
    tmCtrls(LOWNOINDEX).iReq = False
    gSetCtrl tmCtrls(HIGHNOINDEX), 2070, tmCtrls(LOWNOINDEX).fBoxY, 2025, fgBoxStH
    tmCtrls(HIGHNOINDEX).iReq = False
    'Vehicle
    gSetCtrl tmCtrls(VEHINDEX), 30, tmCtrls(LOWNOINDEX).fBoxY + fgStDeltaY, 2025, fgBoxStH
    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
        tmCtrls(VEHINDEX).iReq = True
    Else
        tmCtrls(VEHINDEX).iReq = False
    End If
    'Script
    gSetCtrl tmCtrls(SCRIPTINDEX), 2070, tmCtrls(VEHINDEX).fBoxY, 1395, fgBoxStH
    'Group number
    gSetCtrl tmCtrls(GROUPNOINDEX), 3480, tmCtrls(VEHINDEX).fBoxY, 615, fgBoxStH
    tmCtrls(GROUPNOINDEX).iReq = False
    
    'iPump Prefix
    gSetCtrl tmCtrls(IPUMPPREFIXINDEX), 30, 1965, 1005, fgBoxStH
    tmCtrls(IPUMPPREFIXINDEX).iReq = False
    'iPump Suffix
    gSetCtrl tmCtrls(IPUMPSUFFIXINDEX), 1050, tmCtrls(IPUMPPREFIXINDEX).fBoxY, 1005, fgBoxStH
    tmCtrls(IPUMPSUFFIXINDEX).iReq = False
    'iPump Event Type
    gSetCtrl tmCtrls(IPUMPEVENTTYPEINDEX), 2070, tmCtrls(IPUMPPREFIXINDEX).fBoxY, 1005, fgBoxStH
    tmCtrls(IPUMPEVENTTYPEINDEX).iReq = False
    'iPump Network ID
    gSetCtrl tmCtrls(IPUMPNETWORKIDINDEX), 3090, tmCtrls(IPUMPPREFIXINDEX).fBoxY, 1005, fgBoxStH
    tmCtrls(GROUPNOINDEX).iReq = False
    'iPump Name Space
    gSetCtrl tmCtrls(IPUMPNAMESPACEINDEX), 30, tmCtrls(IPUMPPREFIXINDEX).fBoxY + fgStDeltaY, 4050, fgBoxStH
    tmCtrls(IPUMPNAMESPACEINDEX).iReq = False
    
    If (Asc(tgSpf.sUsingFeatures10) And WegenerIPump) <> WegenerIPump Then
        imMaxCtrl = GROUPNOINDEX
        Media.Height = 3405
        plcMcd.Height = 1875
        pbcMcd.Height = 1770
        cmcDone.Top = 2580
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        cmcErase.Top = cmcDone.Top + cmcDone.Height + 60
        cmcUndo.Top = cmcErase.Top
        cmcReport.Top = cmcErase.Top
        
    Else
        imMaxCtrl = IPUMPNAMESPACEINDEX
        Media.Height = 4530
        plcMcd.Height = 2775
        pbcMcd.Height = 2670
        cmcDone.Top = 3630
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        cmcErase.Top = cmcDone.Top + cmcDone.Height + 60
        cmcUndo.Top = cmcErase.Top
        cmcReport.Top = cmcErase.Top
    End If
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
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String

    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmMcf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(PREFIXINDEX).iChg Then
        tmMcf.sPrefix = edcPrefix.Text
    End If
    If Not ilTestChg Or tmCtrls(REUSEINDEX).iChg Then
        Select Case imReUse
            Case 0  'Yes
                tmMcf.sReuse = "Y"
            Case 1  'No
                tmMcf.sReuse = "N"
            Case Else
                tmMcf.sReuse = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SORTCARTINDEX).iChg Then
        Select Case imSortCart
            Case 0  'Last Purge Date
                tmMcf.sSortCart = "D"
            Case 1  'Cart #
                tmMcf.sSortCart = "C"
            Case Else
                tmMcf.sSortCart = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(CARTINDEX).iChg Then
        Select Case lbcCart.ListIndex
            Case 0  'N/A
                tmMcf.sCartDisp = "N"
            Case 1  'Save
                tmMcf.sCartDisp = "S"
            Case 2  'Purge
                tmMcf.sCartDisp = "P"
            Case 3  'Ask
                tmMcf.sCartDisp = "A"
            Case Else
                tmMcf.sCartDisp = ""
        End Select
    End If
    If Not ilTestChg Or tmCtrls(SUPPRESSEXPORTINDEX).iChg Then
        'Select Case lbcTape.ListIndex
        '    Case 0  'N/A
        '        tmMcf.sTapeDisp = "N"
        '    Case 1  'Return
        '        tmMcf.sTapeDisp = "R"
        '    Case 2  'Destroy
        '        tmMcf.sTapeDisp = "D"
        '    Case 3  'Ask
        '        tmMcf.sTapeDisp = "A"
        '    Case Else
        '        tmMcf.sTapeDisp = ""
        'End Select
        Select Case imSuppressExport
            Case 0  'Yes
                tmMcf.sSuppressOnExport = "Y"
            Case 1  'No
                tmMcf.sSuppressOnExport = "N"
            Case Else
                tmMcf.sSuppressOnExport = ""
        End Select
    End If
    If tmMcf.sReuse = "Y" Then
        If Not ilTestChg Or tmCtrls(LOWNOINDEX).iChg Then
            tmMcf.sAssignNo(0) = edcAssignNo(0).Text
        End If
        If Not ilTestChg Or tmCtrls(HIGHNOINDEX).iChg Then
            tmMcf.sAssignNo(1) = edcAssignNo(1).Text
        End If
    Else
        For ilLoop = 0 To 1 Step 1
            tmMcf.sAssignNo(ilLoop) = ""
        Next ilLoop
    End If
    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
        If Not ilTestChg Or tmCtrls(VEHINDEX).iChg Then
            tmMcf.iVefCode = 0
            If lbcVehicle.ListIndex > 0 Then
                slNameCode = tmVehicleCode(lbcVehicle.ListIndex - 1).sKey  'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmMcf.iVefCode = Val(slCode)
            End If
        End If
    Else
        tmMcf.iVefCode = 0
    End If
    If Not ilTestChg Or tmCtrls(SCRIPTINDEX).iChg Then
        Select Case imScript
            Case 0  'Yes
                tmMcf.sScript = "Y"
            Case 1  'No
                tmMcf.sScript = "N"
            Case Else
                tmMcf.sScript = ""
        End Select
    End If
    'If Not ilTestChg Or tmCtrls(GROUPNOINDEX).iChg Then
    '    tmMcf.iGroupNo = Val(edcGroupNo.Text)
    'End If
    If Not ilTestChg Or tmCtrls(IPUMPPREFIXINDEX).iChg Then
        tmMef.sPrefix = edcIPumpPrefix.Text
    End If
    If Not ilTestChg Or tmCtrls(IPUMPSUFFIXINDEX).iChg Then
        tmMef.sSuffix = edcIPumpSuffix.Text
    End If
    If Not ilTestChg Or tmCtrls(IPUMPEVENTTYPEINDEX).iChg Then
        tmMef.sEventType = edcIPumpEventType.Text
    End If
    If Not ilTestChg Or tmCtrls(IPUMPNETWORKIDINDEX).iChg Then
        tmMef.sNetworkID = edcIPumpNetworkID.Text
    End If
    If Not ilTestChg Or tmCtrls(IPUMPNAMESPACEINDEX).iChg Then
        tmMef.sNameSpace = edcIPumpNameSpace.Text
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                *
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
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    edcName.Text = Trim$(tmMcf.sName)
    edcPrefix.Text = Trim$(tmMcf.sPrefix)
    Select Case tmMcf.sReuse
        Case "Y"
            imReUse = 0
        Case "N"
            imReUse = 1
        Case Else
            imReUse = -1
    End Select
    Select Case tmMcf.sSortCart
        Case "D"
            imSortCart = 0
        Case "C"
            imSortCart = 1
        Case Else
            imSortCart = -1
    End Select
    Select Case tmMcf.sCartDisp
        Case "N"
            lbcCart.ListIndex = 0
        Case "S"
            lbcCart.ListIndex = 1
        Case "P"
            lbcCart.ListIndex = 2
        Case "A"
            lbcCart.ListIndex = 3
        Case Else
            lbcCart.ListIndex = -1
    End Select
    'Select Case tmMcf.sTapeDisp
    '    Case "N"
    '        lbcTape.ListIndex = 0
    '    Case "R"
    '        lbcTape.ListIndex = 1
    '    Case "D"
    '        lbcTape.ListIndex = 2
    '    Case "A"
    '        lbcTape.ListIndex = 3
    '    Case Else
    '        lbcTape.ListIndex = -1
    'End Select
    Select Case tmMcf.sSuppressOnExport
        Case "Y"
            imSuppressExport = 0
        Case "N"
            imSuppressExport = 1
        Case Else
            imSuppressExport = -1
    End Select
    For ilLoop = 0 To 1 Step 1
        edcAssignNo(ilLoop).Text = Trim$(tmMcf.sAssignNo(ilLoop))
        smOrigAssign(ilLoop) = Trim$(tmMcf.sAssignNo(ilLoop))
    Next ilLoop
    lbcVehicle.ListIndex = -1
    smVehicle = ""
    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
        If tmMcf.iVefCode > 0 Then
            For ilLoop = 0 To UBound(tmVehicleCode) - 1 Step 1  'lbcVehicleCode.ListCount - 1 Step 1
                slNameCode = tmVehicleCode(ilLoop).sKey  'lbcVehicleCode.List(gLastFound(lbcVehicle))
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tmMcf.iVefCode = Val(slCode) Then
                    lbcVehicle.ListIndex = ilLoop + 1
                    smVehicle = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
            Next ilLoop
        Else
            lbcVehicle.ListIndex = 0
            smVehicle = lbcVehicle.List(lbcVehicle.ListIndex)
        End If
    End If
    Select Case tmMcf.sScript
        Case "Y"
            imScript = 0
        Case "N"
            imScript = 1
        Case Else
            imScript = -1
    End Select
    'edcGroupNo.Text = Trim$(Str$(tmMcf.iGroupNo))
    edcIPumpPrefix.Text = Trim$(tmMef.sPrefix)
    edcIPumpSuffix.Text = Trim$(tmMef.sSuffix)
    edcIPumpEventType.Text = Trim$(tmMef.sEventType)
    edcIPumpNetworkID.Text = Trim$(tmMef.sNetworkID)
    edcIPumpNameSpace.Text = Trim$(tmMef.sNameSpace)
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
    Dim ilLoop As Integer
    Dim ilDuplicateName As Integer

    If edcName.Text <> "" Then    'Test name
        slStr = Trim$(edcName.Text)
        'gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        'If gLastFound(cbcSelect) <> -1 Then   'Name found
        '    If gLastFound(cbcSelect) <> imSelectedIndex Then
        '        If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
        '            Beep
        '            MsgBox "Media already defined, enter a different name", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
        '            edcName.Text = Trim$(tmMcf.sName) 'Reset text
        '            mSetShow imBoxNo
        '            mSetChg imBoxNo
        '            imBoxNo = 1
        '            mEnableBox imBoxNo
        '            mOKName = False
        '            Exit Function
        '        End If
        '    End If
        'End If
        For ilLoop = 0 To UBound(tmMMcf) - 1 Step 1
            If StrComp(Trim$(tmMMcf(ilLoop).sName), slStr, vbTextCompare) = 0 Then
                ilDuplicateName = True
                If imSelectedIndex > 0 Then
                    If cbcSelect.ItemData(imSelectedIndex) = tmMMcf(ilLoop).iCode Then
                        ilDuplicateName = False
                    End If
                End If
                If ilDuplicateName Then
                    Beep
                    MsgBox "Media already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    If cbcSelect.ListIndex > 0 Then
                        edcName.Text = Trim$(tmMcf.sName) 'Reset text
                    End If
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
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
    'gInitStdAlone Media, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igMcdCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igMcdCallSource = CALLNONE
    'End If
    If igMcdCallSource <> CALLNONE Then  'set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgMcdName = slStr
        Else
            sgMcdName = ""
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
    Dim ilAdd As Integer
    Dim ilVef As Integer

    imPopReqd = False
    'ilFilter(0) = NOFILTER
    'slFilter(0) = ""
    'ilOffset(0) = 0
    ''ilRet = gIMoveListBox(Media, cbcSelect, lbcNameCode, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    'ilRet = gIMoveListBox(Media, cbcSelect, tgNameCode(), sgNameCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    'If ilRet <> CP_MSG_NOPOPREQ Then
    '    On Error GoTo mPopulateErr
    '    gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", Media
    '    On Error GoTo 0
    '    cbcSelect.AddItem "[New]", 0  'Force as first item on list
    '    imPopReqd = False
    'End If
    cbcSelect.Clear
    ReDim tmMMcf(0 To 0) As MCF
    ilRet = btrGetFirst(hmMcf, tmMcf, imRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
    Do While ilRet = BTRV_ERR_NONE
        ilAdd = True
        If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
            If tmMcf.iVefCode > 0 Then
                ilVef = gBinarySearchVef(tmMcf.iVefCode)
                If ilVef = -1 Then
                    ilAdd = False
                End If
            End If
        End If
        If ilAdd Then
            cbcSelect.AddItem Trim$(tmMcf.sName)
            cbcSelect.ItemData(cbcSelect.NewIndex) = tmMcf.iCode
        End If
        tmMMcf(UBound(tmMMcf)) = tmMcf
        ReDim Preserve tmMMcf(0 To UBound(tmMMcf) + 1) As MCF
        ilRet = btrGetNext(hmMcf, tmMcf, imRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    cbcSelect.AddItem "[New]", 0
    cbcSelect.ItemData(cbcSelect.NewIndex) = 0
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                    slCode                                                  *
'******************************************************************************************

'
'   iRet = ENmRead(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    'slNameCode = tgNameCode(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
    'On Error GoTo mReadRecErr
    'gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", Media
    'On Error GoTo 0
    'slCode = Trim$(slCode)
    'tmSrchKey.iCode = CInt(slCode)
    tmSrchKey.iCode = cbcSelect.ItemData(ilSelectIndex)
    ilRet = btrGetEqual(hmMcf, tmMcf, imRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Media
    On Error GoTo 0
    If (Asc(tgSpf.sUsingFeatures10) And WegenerIPump) = WegenerIPump Then
        tmMefSrchKey1.iMcfCode = tmMcf.iCode
        ilRet = btrGetEqual(hmMef, tmMef, imMefRecLen, tmMefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, ilForUpdate)
        If ilRet <> BTRV_ERR_NONE Then
            tmMef.iCode = 0
            tmMef.sPrefix = ""
            tmMef.sSuffix = ""
            tmMef.sEventType = ""
            tmMef.sNetworkID = ""
            tmMef.sNameSpace = ""
        End If
    End If
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slName                                                                                *
'******************************************************************************************

'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilCifRecLen As Integer
    Dim ilPos As Integer
    Dim ilForm As Integer   '0=xxx; 1=axxx; 2=a-xx; 3=xx-xx
    Dim slLowest As String
    Dim slHighest As String
    Dim llLowest1 As Long
    Dim llHighest1 As Long
    Dim llLowest2 As Long
    Dim llHighest2 As Long
    Dim ilNoChar1 As Integer
    Dim ilNoChar2 As Integer
    Dim llLoop1 As Long
    Dim llLoop2 As Long
    Dim slStr As String
    Dim slNameCode As String
    Dim slStr1 As String
    Dim llTest As Long
    Dim ilFound As Integer
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
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
        slStamp = gFileDateTime(sgDBPath & "Mcf.btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmMcf.iCode = 0  'Autoincrement
            ilRet = btrInsert(hmMcf, tmMcf, imRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmMcf, tmMcf, imRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Media
    On Error GoTo 0
    If imReUse = 0 Then 'Generate numbers if required
        ilCifRecLen = Len(tmCif)
        slLowest = Trim$(edcAssignNo(0).Text)
        slHighest = Trim$(edcAssignNo(1).Text)
        If (imSelectedIndex = 0) Or ((imSelectedIndex <> 0) And ((StrComp(slLowest, smOrigAssign(0), vbTextCompare) <> 0) Or (StrComp(slHighest, smOrigAssign(1), vbTextCompare) <> 0))) Then
            ilPos = InStr(1, slLowest, "-")
            If ((Asc(slLowest) >= KEYLA) And (Asc(slLowest) <= KEYLZ)) Or ((Asc(slLowest) >= KEYUA) And (Asc(slLowest) <= KEYUZ)) Then
                If ilPos = 0 Then
                    ilForm = 1
                    llLowest1 = Asc(slLowest)
                    llLowest2 = Asc(slHighest)
                    llHighest1 = Val(right$(slLowest, Len(slLowest) - 1))
                    llHighest2 = Val(right$(slHighest, Len(slHighest) - 1))
                    ilNoChar1 = 1
                    ilNoChar2 = Len(slLowest) - 1
                Else
                    ilForm = 2
                    llLowest1 = Asc(slLowest)
                    llLowest2 = Asc(slHighest)
                    llHighest1 = Val(right$(slLowest, Len(slLowest) - ilPos))
                    llHighest2 = Val(right$(slHighest, Len(slHighest) - ilPos))
                    ilNoChar1 = 1
                    ilNoChar2 = Len(slLowest) - ilPos
                End If
            ElseIf ilPos <> 0 Then
                ilForm = 3
                llLowest1 = Val(Left$(slLowest, ilPos - 1))
                llLowest2 = Val(Left$(slHighest, ilPos - 1))
                llHighest1 = Val(right$(slLowest, Len(slLowest) - ilPos))
                llHighest2 = Val(right$(slHighest, Len(slHighest) - ilPos))
                ilNoChar1 = ilPos - 1
                ilNoChar2 = Len(slLowest) - ilPos
            Else
                ilForm = 0
                llLowest1 = Val(slLowest)
                llLowest2 = Val(slHighest)
                ilNoChar1 = Len(slLowest)
                ilNoChar2 = 0
                llHighest1 = 1
                llHighest2 = 1
            End If
            'lbcInvCode.Clear
            'lbcInvCode.Tag = ""
            ReDim tmInvCode(0 To 0) As SORTCODE
            smInvCodeTag = ""
            'lbcInv.Clear
            If imSelectedIndex <> 0 Then    'build inventory compare list
                'Read all records saving matching media
                ilfilter(0) = INTEGERFILTER
                slFilter(0) = Trim$(str$(tmMcf.iCode))
                ilOffSet(0) = gFieldOffset("Cif", "CifMcfCode") '4
                ''ilRet = gIMoveListBox(Media, lbcInv, lbcInvCode, "Cif.btr", gFieldOffset("Cif", "CifName"), 5, ilFilter(), slFilter(), ilOffset())
                'ilRet = gLMoveListBox(Media, lbcInv, tmInvCode(), smInvCodeTag, "Cif.btr", gFieldOffset("Cif", "CifName"), 5, ilFilter(), slFilter(), ilOffset())
                ilRet = gLPopListBox(Media, tmInvCode(), smInvCodeTag, "Cif.btr", gFieldOffset("Cif", "CifName"), 5, ilfilter(), slFilter(), ilOffSet())
                For llTest = 0 To UBound(tmInvCode) - 1 Step 1
                    DoEvents
                    slNameCode = tmInvCode(llTest).sKey    'lbcMster.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 1, "\", slStr1)
                    tmInvCode(llTest).sKey = slStr1
                Next llTest
            End If
            'Clear record
            tmCif.sCut = ""
            tmCif.sReel = ""
            tmCif.iLen = 0
            tmCif.iEtfCode = 0
            tmCif.iEnfCode = 0
            tmCif.iAdfCode = 0
            tmCif.lcpfCode = 0
            tmCif.iMnfComp(0) = 0
            tmCif.iMnfComp(1) = 0
            tmCif.iMnfAnn = 0
            tmCif.lCsfCode = 0
            tmCif.iNoTimesAir = 0
            tmCif.sCartDisp = tmMcf.sCartDisp
            tmCif.sTapeDisp = "N"   'tmMcf.sTapeDisp
            tmCif.sPurged = "P"
            tmCif.iPurgeDate(0) = 0
            tmCif.iPurgeDate(1) = 0
            slStr = Format$(gNow(), "m/d/yy")
            gPackDate slStr, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
            tmCif.iUsedDate(0) = 0
            tmCif.iUsedDate(1) = 0
            tmCif.iRotStartDate(0) = 0
            tmCif.iRotStartDate(1) = 0
            tmCif.iRotEndDate(0) = 0
            tmCif.iRotEndDate(1) = 0
            tmCif.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            tmCif.sPrint = "N"
            ilRet = btrBeginTrans(hmCif, 1000)
            For llLoop1 = llLowest1 To llLowest2 Step 1
                For llLoop2 = llHighest1 To llHighest2 Step 1
                    DoEvents
                    slStr = ""  'Form the number
                    Select Case ilForm
                        Case 0  'xxxx
                            slStr = Trim$(str$(llLoop1))
                            Do While Len(slStr) < ilNoChar1
                                slStr = "0" & slStr
                            Loop
                        Case 1  'axxx
                            slStr = Trim$(str$(llLoop2))
                            Do While Len(slStr) < ilNoChar2
                                slStr = "0" & slStr
                            Loop
                            slStr = Chr$(llLoop1) & slStr
                        Case 2  'a-xxx
                            slStr = Trim$(str$(llLoop2))
                            Do While Len(slStr) < ilNoChar2
                                slStr = "0" & slStr
                            Loop
                            slStr = Chr$(llLoop1) & "-" & slStr
                        Case 3  'xx-xx
                            slStr = Trim$(str$(llLoop2))
                            Do While Len(slStr) < ilNoChar2
                                slStr = "0" & slStr
                            Loop
                            slStr = Trim$(str$(llLoop1)) & "-" & slStr
                            Do While Len(slStr) < ilNoChar1 + ilNoChar2 + 1
                                slStr = "0" & slStr
                            Loop
                    End Select
                    'gFindMatch slStr, 0, lbcInv
                    'If gLastFound(lbcInv) = -1 Then
                    ilFound = False
                    For llTest = 0 To UBound(tmInvCode) - 1 Step 1
                        DoEvents
                        'slNameCode = tmInvCode(llTest).sKey    'lbcMster.List(ilLoop)
                        'ilRet = gParseItem(slNameCode, 1, "\", slStr1)
                        slStr1 = Trim$(tmInvCode(llTest).sKey)
                        If StrComp(slStr, slStr1, vbTextCompare) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next llTest
                    If Not ilFound Then
                        Do  'Loop until record updated or added
                            tmCif.lCode = 0  'Autoincrement
                            tmCif.iMcfCode = tmMcf.iCode
                            tmCif.sName = slStr
                            ilRet = btrInsert(hmCif, tmCif, ilCifRecLen, INDEXKEY0)
                            slMsg = "mSaveRec (btrInsert)"
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit For
                        End If
                    End If
                Next llLoop2
                If ilRet <> BTRV_ERR_NONE Then
                    Exit For
                End If
            Next llLoop1
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = MsgBox("Media Generation Not Completed, Try Later", vbOKOnly + vbExclamation, "Media")
                ilRet = btrAbortTrans(hmCif)
            Else
                ilRet = btrEndTrans(hmCif)
            End If
        End If
    End If

    If (Asc(tgSpf.sUsingFeatures10) And WegenerIPump) = WegenerIPump Then
        If (imSelectedIndex = 0) Or (tmMef.iCode = 0) Then 'New selected
            tmMef.iCode = 0  'Autoincrement
            tmMef.iMcfCode = tmMcf.iCode
            ilRet = btrInsert(hmMef, tmMef, imMefRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: MEF)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmMef, tmMef, imMefRecLen)
            slMsg = "mSaveRec (btrUpdate: MEF)"
        End If
    End If
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Media
    On Error GoTo 0

    ''If lbcNameCode.Tag <> "" Then
    ''    If slStamp = lbcNameCode.Tag Then
    ''        lbcNameCode.Tag = FileDateTime(sgDBPath & "Mcf.btr")
    ''    End If
    ''End If
    'If sgNameCodeTag <> "" Then
    '    If slStamp = sgNameCodeTag Then
    '        sgNameCodeTag = gFileDateTime(sgDBPath & "Mcf.btr")
    '    End If
    'End If
    'If imSelectedIndex <> 0 Then
    '    'lbcNameCode.RemoveItem imSelectedIndex - 1
    '    gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = RTrim$(tmMcf.sName)
    'cbcSelect.AddItem slName
    'slName = tmMcf.sName + "\" + LTrim$(Str$(tmMcf.iCode)) 'slName + "\" + LTrim$(Str$(tmMcf.iCode))
    ''lbcNameCode.AddItem slName
    'gAddItemToSortCode slName, tgNameCode(), True
    'cbcSelect.AddItem "[New]", 0
    mPopulate
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
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcMcd_Paint
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrl) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmMcf.sName, edcName, tmCtrls(ilBoxNo)
        Case PREFIXINDEX 'Market Name
            gSetChgFlag tmMcf.sPrefix, edcPrefix, tmCtrls(ilBoxNo)
        Case REUSEINDEX 'Reuse
        Case SORTCARTINDEX  'Temporary disposition
        Case CARTINDEX
            Select Case tmMcf.sCartDisp
                Case "N"
                    slStr = lbcCart.List(0)
                Case "S"
                    slStr = lbcCart.List(1)
                Case "P"
                    slStr = lbcCart.List(2)
                Case "A"
                    slStr = lbcCart.List(3)
                Case Else
                    slStr = ""
            End Select
            gSetChgFlag slStr, lbcCart, tmCtrls(ilBoxNo)
        Case SUPPRESSEXPORTINDEX
            'Select Case tmMcf.sTapeDisp
            '    Case "N"
            '        slStr = lbcTape.List(0)
            '    Case "R"
            '        slStr = lbcTape.List(1)
            '    Case "D"
            '        slStr = lbcTape.List(2)
            '    Case "A"
            '        slStr = lbcTape.List(3)
            '    Case Else
            '        slStr = ""
            'End Select
            'gSetChgFlag slStr, lbcTape, tmCtrls(ilBoxNo)
        Case LOWNOINDEX 'AssignNo
            gSetChgFlag tmMcf.sAssignNo(0), edcAssignNo(0), tmCtrls(ilBoxNo)
        Case HIGHNOINDEX 'AssignNo
            gSetChgFlag tmMcf.sAssignNo(1), edcAssignNo(1), tmCtrls(ilBoxNo)
        Case VEHINDEX
            If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
                gSetChgFlag smVehicle, lbcVehicle, tmCtrls(ilBoxNo)
            End If
        Case SCRIPTINDEX 'Script
        'Case GROUPNOINDEX 'Dial Position
        '    gSetChgFlag Trim$(Str$(tmMcf.iGroupNo)), edcGroupNo, tmCtrls(ilBoxNo)
        Case IPUMPPREFIXINDEX 'Name
            gSetChgFlag tmMef.sPrefix, edcIPumpPrefix, tmCtrls(ilBoxNo)
        Case IPUMPSUFFIXINDEX 'Name
            gSetChgFlag tmMef.sSuffix, edcIPumpSuffix, tmCtrls(ilBoxNo)
        Case IPUMPEVENTTYPEINDEX 'Name
            gSetChgFlag tmMef.sEventType, edcIPumpEventType, tmCtrls(ilBoxNo)
        Case IPUMPNETWORKIDINDEX 'Name
            gSetChgFlag tmMef.sNetworkID, edcIPumpNetworkID, tmCtrls(ilBoxNo)
        Case IPUMPNAMESPACEINDEX 'Name
            gSetChgFlag tmMef.sNameSpace, edcIPumpNameSpace, tmCtrls(ilBoxNo)
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
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus specified control    *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrl) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case PREFIXINDEX 'Market Name
            edcPrefix.SetFocus
        Case REUSEINDEX 'Reuse numbers
            pbcYN.SetFocus
        Case SORTCARTINDEX 'Temporary media code
            pbcYN.SetFocus
        Case CARTINDEX   'Cart disposition
            edcDropDown.SetFocus
        Case SUPPRESSEXPORTINDEX   'Suppress Automation Export
            pbcYN.SetFocus
        Case LOWNOINDEX 'AssignNo
            edcAssignNo(0).SetFocus
        Case HIGHNOINDEX 'AssignNo
            edcAssignNo(1).SetFocus
        Case VEHINDEX   'Tape disposition
            edcDropDown.SetFocus
        Case SCRIPTINDEX '# Full Months
            pbcYN.SetFocus
        Case GROUPNOINDEX 'Dial Position
            edcGroupNo.SetFocus
        Case IPUMPPREFIXINDEX 'Dial Position
            edcIPumpPrefix.SetFocus
        Case IPUMPSUFFIXINDEX 'Dial Position
            edcIPumpSuffix.SetFocus
        Case IPUMPEVENTTYPEINDEX 'Dial Position
            edcIPumpEventType.SetFocus
        Case IPUMPNETWORKIDINDEX 'Dial Position
            edcIPumpNetworkID.SetFocus
        Case IPUMPNAMESPACEINDEX 'Dial Position
            edcIPumpNameSpace.SetFocus
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrl) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case PREFIXINDEX 'Market Name
            edcPrefix.Visible = False  'Set visibility
            slStr = edcPrefix.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case REUSEINDEX 'Reuse
            pbcYN.Visible = False  'Set visibility
            If imReUse = 0 Then
                slStr = "Yes"
            ElseIf imReUse = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case SORTCARTINDEX 'Reuse
            pbcYN.Visible = False  'Set visibility
            If imSortCart = 0 Then
                slStr = "Last Used"
            ElseIf imSortCart = 1 Then
                slStr = "Cart #"
            Else
                slStr = ""
            End If
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case CARTINDEX 'Cart disposition
            lbcCart.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcCart.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcCart.List(lbcCart.ListIndex)
            End If
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case SUPPRESSEXPORTINDEX 'Suppress on Automation Export
            pbcYN.Visible = False  'Set visibility
            If imSuppressExport = 0 Then
                slStr = "Yes"
            ElseIf imSuppressExport = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case LOWNOINDEX 'Assign No
            edcAssignNo(0).Visible = False
            slStr = edcAssignNo(0).Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case HIGHNOINDEX 'Assign No
            edcAssignNo(1).Visible = False
            slStr = edcAssignNo(1).Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case VEHINDEX 'Vehicle
            lbcVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
                If lbcVehicle.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcVehicle.List(lbcVehicle.ListIndex)
                End If
            Else
                slStr = ""
            End If
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case SCRIPTINDEX 'Script
            pbcYN.Visible = False  'Set visibility
            If imScript = 0 Then
                slStr = "Yes"
            ElseIf imScript = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case GROUPNOINDEX 'Dial Position
            edcGroupNo.Visible = False  'Set visibility
            slStr = edcGroupNo.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case IPUMPPREFIXINDEX 'Dial Position
            edcIPumpPrefix.Visible = False  'Set visibility
            slStr = edcIPumpPrefix.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case IPUMPSUFFIXINDEX 'Dial Position
            edcIPumpSuffix.Visible = False  'Set visibility
            slStr = edcIPumpSuffix.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case IPUMPEVENTTYPEINDEX 'Dial Position
            edcIPumpEventType.Visible = False  'Set visibility
            slStr = edcIPumpEventType.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case IPUMPNETWORKIDINDEX 'Dial Position
            edcIPumpNetworkID.Visible = False  'Set visibility
            slStr = edcIPumpNetworkID.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
        Case IPUMPNAMESPACEINDEX 'Dial Position
            edcIPumpNameSpace.Visible = False  'Set visibility
            slStr = edcIPumpNameSpace.Text
            gSetShow pbcMcd, slStr, tmCtrls(ilBoxNo)
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

    sgDoneMsg = Trim$(str$(igMcdCallSource)) & "\" & sgMcdName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Media
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
    Dim ilRes As Integer    'Result of MsgBox
    Dim slStr As String
    Dim slLowest As String
    Dim slHighest As String
    Dim slOldHighest As String
    Dim ilLowest As Integer
    Dim ilHighest As Integer
    Dim ilCharCount As Integer
    Dim ilNoNeg As Integer
    Dim ilAlpha As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PREFIXINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcPrefix, "", "Prefix must be specified", tmCtrls(PREFIXINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PREFIXINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = REUSEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imReUse = 0 Then
            slStr = "Yes"
        ElseIf imReUse = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Reuse expired number must be specified", tmCtrls(REUSEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = REUSEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SORTCARTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSortCart = 0 Then
            slStr = "Last Used"
        ElseIf imSortCart = 1 Then
            slStr = "Cart #"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Sort Cart Inventory must be specified", tmCtrls(SORTCARTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SORTCARTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CARTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcCart, "", "Cart Disposition must be specified", tmCtrls(CARTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CARTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SUPPRESSEXPORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imSuppressExport = 0 Then
            slStr = "Yes"
        ElseIf imSuppressExport = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Suppress on Automation Export must be specified", tmCtrls(REUSEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SUPPRESSEXPORTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = LOWNOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imReUse = 0 Then
            tmCtrls(LOWNOINDEX).iReq = True
            tmCtrls(HIGHNOINDEX).iReq = True
        Else
            tmCtrls(LOWNOINDEX).iReq = False
            tmCtrls(HIGHNOINDEX).iReq = False
        End If
        If gFieldDefinedCtrl(edcAssignNo(0), "", "Lowest number to assign must be specified", tmCtrls(LOWNOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = LOWNOINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = HIGHNOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imReUse = 0 Then
            tmCtrls(LOWNOINDEX).iReq = True
            tmCtrls(HIGHNOINDEX).iReq = True
        Else
            tmCtrls(LOWNOINDEX).iReq = False
            tmCtrls(HIGHNOINDEX).iReq = False
        End If
        If gFieldDefinedCtrl(edcAssignNo(1), "", "Highest number to assign must be specified", tmCtrls(HIGHNOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = HIGHNOINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (imReUse = 0) And (ilCtrlNo = TESTALLCTRLS) And ((ilState And (ALLMANDEFINED + SHOWMSG)) = ALLMANDEFINED + SHOWMSG) Then 'Check that lowest and highest number are the same form
        'Allow forms xxx; axxx; a-xx; xx-xx
        slLowest = edcAssignNo(0).Text
        slHighest = edcAssignNo(1).Text
        If Len(slLowest) <> Len(slHighest) Then
            ilRes = MsgBox("Lowest # and Higest # must be in the same form", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = LOWNOINDEX
            mTestFields = NO
            Exit Function
        End If
        ilCharCount = 1
        ilNoNeg = 0
        ilAlpha = False
        Do While Len(slLowest) > 0
            ilLowest = Asc(slLowest)
            ilHighest = Asc(slHighest)
            If (ilLowest >= KEYUA) And (ilLowest <= KEYUZ) Then
                If ((ilHighest < KEYUA) Or (ilHighest > KEYUZ)) And ((ilHighest < KEYLA) Or (ilHighest > KEYLZ)) Then
                    ilRes = MsgBox("Lowest # and Higest # must be in the same form", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                If ilCharCount > 1 Then
                    ilRes = MsgBox("Character can only be at the start of the number", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                ilAlpha = True
            ElseIf (ilLowest >= KEYLA) And (ilLowest <= KEYLZ) Then
                If ((ilHighest < KEYUA) Or (ilHighest > KEYUZ)) And ((ilHighest < KEYLA) Or (ilHighest > KEYLZ)) Then
                    ilRes = MsgBox("Lowest # and Higest # must be in the same form", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                If ilCharCount > 1 Then
                    ilRes = MsgBox("Character can only be at the start of the number", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                ilAlpha = True
            ElseIf (ilLowest >= KEY0) And (ilLowest <= KEY9) Then
                If (ilHighest < KEY0) And (ilHighest > KEY9) Then
                    ilRes = MsgBox("Lowest # and Higest # must be in the same form", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
            ElseIf ilLowest = KEYNEG Then
                ilNoNeg = ilNoNeg + 1
                If ilNoNeg > 1 Then
                    ilRes = MsgBox("Only one '-' allowed in the number", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                If (ilHighest <> KEYNEG) Then
                    ilRes = MsgBox("Lowest # and Higest # must be in the same form", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                If ilCharCount = 1 Then
                    ilRes = MsgBox("'-' can't be at the start of the number", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                If (ilAlpha) And (ilCharCount <> 2) Then
                    ilRes = MsgBox("The '-' must be right after the character", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
            End If
            slLowest = Mid$(slLowest, 2)
            slHighest = Mid$(slHighest, 2)
            ilCharCount = ilCharCount + 1
        Loop
        'The form can't change
        If imSelectedIndex > 0 Then
            If mReadRec(imSelectedIndex, SETFORREADONLY) Then
                slHighest = Trim$(edcAssignNo(1).Text)
                slOldHighest = Trim$(tmMcf.sAssignNo(1))
                If Len(slOldHighest) <> Len(slHighest) Then
                    ilRes = MsgBox("The form of the # can't be changed", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                ilPos1 = InStr(slHighest, "-")
                ilPos2 = InStr(slOldHighest, "-")
                If ilPos1 <> ilPos2 Then
                    ilRes = MsgBox("The form of the # can't be changed", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = LOWNOINDEX
                    mTestFields = NO
                    Exit Function
                End If
                ilPos1 = Asc(slHighest)
                ilPos2 = Asc(slOldHighest)
                If (ilPos1 >= KEYUA) And (ilPos1 <= KEYUZ) Then
                    If (ilPos2 < KEYUA) Or (ilPos2 < KEYUZ) Then
                        ilRes = MsgBox("The form of the # can't be changed", vbOKOnly + vbExclamation, "Incomplete")
                        imBoxNo = LOWNOINDEX
                        mTestFields = NO
                        Exit Function
                    End If
                End If
                If (ilPos1 >= KEYLA) And (ilPos1 <= KEYLZ) Then
                    If (ilPos2 < KEYLA) Or (ilPos2 < KEYLZ) Then
                        ilRes = MsgBox("The form of the # can't be changed", vbOKOnly + vbExclamation, "Incomplete")
                        imBoxNo = LOWNOINDEX
                        mTestFields = NO
                        Exit Function
                    End If
                End If
                If (ilPos2 >= KEYUA) And (ilPos2 <= KEYUZ) Then
                    If (ilPos1 < KEYUA) Or (ilPos1 < KEYUZ) Then
                        ilRes = MsgBox("The form of the # can't be changed", vbOKOnly + vbExclamation, "Incomplete")
                        imBoxNo = LOWNOINDEX
                        mTestFields = NO
                        Exit Function
                    End If
                End If
                If (ilPos2 >= KEYLA) And (ilPos2 <= KEYLZ) Then
                    If (ilPos1 < KEYLA) Or (ilPos1 < KEYLZ) Then
                        ilRes = MsgBox("The form of the # can't be changed", vbOKOnly + vbExclamation, "Incomplete")
                        imBoxNo = LOWNOINDEX
                        mTestFields = NO
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    If (ilCtrlNo = SCRIPTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imScript = 0 Then
            slStr = "Yes"
        ElseIf imScript = 1 Then
            slStr = "No"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Script must be specified", tmCtrls(SCRIPTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SCRIPTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = GROUPNOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcGroupNo, "", "Group number must be specified", tmCtrls(GROUPNOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = GROUPNOINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function

Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
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
Private Sub pbcMcd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To imMaxCtrl Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (ilBox = LOWNOINDEX) And (imReUse = 1) Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If (ilBox = HIGHNOINDEX) And (imReUse = 1) Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If (ilBox = VEHINDEX) Then
                    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) <> MEDIACODEBYVEH Then
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
    mSetFocus imBoxNo
End Sub
Private Sub pbcMcd_Paint()
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
        llColor = pbcMcd.ForeColor
        slFontName = pbcMcd.FontName
        flFontSize = pbcMcd.FontSize
        pbcMcd.ForeColor = BLUE
        pbcMcd.FontBold = False
        pbcMcd.FontSize = 7
        pbcMcd.FontName = "Arial"
        pbcMcd.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcMcd.CurrentX = tmCtrls(VEHINDEX).fBoxX + 15  'fgBoxInsetX
        pbcMcd.CurrentY = tmCtrls(VEHINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcMcd.Print "Vehicle"
        pbcMcd.FontSize = flFontSize
        pbcMcd.FontName = slFontName
        pbcMcd.FontSize = flFontSize
        pbcMcd.ForeColor = llColor
        pbcMcd.FontBold = True
    End If
    For ilBox = imLBCtrls To imMaxCtrl Step 1
        pbcMcd.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcMcd.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        If ((ilBox = LOWNOINDEX) Or (ilBox = HIGHNOINDEX)) And (imReUse = 1) Then
            pbcMcd.Print "    "
        Else
            pbcMcd.Print tmCtrls(ilBox).sShow
        End If
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxCtrl) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    imTabDirection = -1  'Set-right to left
    ilBox = imBoxNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1
                imTabDirection = 0  'Set-Left to right
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
            Case HIGHNOINDEX
                If imReUse = 1 Then
                    ilFound = False
                End If
                ilBox = LOWNOINDEX
            Case SCRIPTINDEX
                If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) <> MEDIACODEBYVEH Then
                    If imReUse = 1 Then
                        ilFound = False
                    End If
                    ilBox = HIGHNOINDEX
                Else
                    ilBox = VEHINDEX
                End If
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxCtrl) Then
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
                ilBox = imMaxCtrl
            Case SUPPRESSEXPORTINDEX
                If imReUse = 1 Then
                    ilFound = False
                End If
                ilBox = LOWNOINDEX
            Case LOWNOINDEX
                If imReUse = 1 Then
                    ilFound = False
                End If
                ilBox = HIGHNOINDEX
            Case HIGHNOINDEX
                If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) <> MEDIACODEBYVEH Then
                    ilBox = SCRIPTINDEX
                Else
                    ilBox = VEHINDEX
                End If
            Case imMaxCtrl  'GROUPNOINDEX
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igMcdCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Case Else
                ilBox = ilBox + 1
        End Select
    Loop While Not ilFound
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If imBoxNo = SCRIPTINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imScript <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imScript = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imScript <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imScript = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imScript = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imScript = 1
                pbcYN_Paint
            ElseIf imScript = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imScript = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = REUSEINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imReUse <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imReUse = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imReUse <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imReUse = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imReUse = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imReUse = 1
                pbcYN_Paint
            ElseIf imReUse = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imReUse = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = SORTCARTINDEX Then
        If KeyAscii = Asc("L") Or (KeyAscii = Asc("l")) Then
            If imSortCart <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSortCart = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            If imSortCart <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSortCart = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSortCart = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSortCart = 1
                pbcYN_Paint
            ElseIf imSortCart = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSortCart = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    ElseIf imBoxNo = SUPPRESSEXPORTINDEX Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imSuppressExport <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSuppressExport = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imReUse <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSuppressExport = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSuppressExport = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSuppressExport = 1
                pbcYN_Paint
            ElseIf imSuppressExport = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSuppressExport = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = SCRIPTINDEX Then
        If imScript = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imScript = 1
        ElseIf imScript = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imScript = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = REUSEINDEX Then
        If imReUse = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imReUse = 1
        ElseIf imReUse = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imReUse = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = SORTCARTINDEX Then
        If imSortCart = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSortCart = 1
        ElseIf imSortCart = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSortCart = 0
        End If
        pbcYN_Paint
        mSetCommands
    ElseIf imBoxNo = SUPPRESSEXPORTINDEX Then
        If imSuppressExport = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSuppressExport = 1
        ElseIf imSuppressExport = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSuppressExport = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imBoxNo = SCRIPTINDEX Then
        If imScript = 0 Then
            pbcYN.Print "Yes"
        ElseIf imScript = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = REUSEINDEX Then
        If imReUse = 0 Then
            pbcYN.Print "Yes"
        ElseIf imReUse = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    ElseIf imBoxNo = SORTCARTINDEX Then
        If imSortCart = 0 Then
            pbcYN.Print "Last Used"
        ElseIf imSortCart = 1 Then
            pbcYN.Print "Cart #"
        Else
            pbcYN.Print "         "
        End If
    ElseIf imBoxNo = SUPPRESSEXPORTINDEX Then
        If imSuppressExport = 0 Then
            pbcYN.Print "Yes"
        ElseIf imSuppressExport = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
End Sub
Private Sub plcMcd_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(Media, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, tmVehicleCode(), smVehicleCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Media
        On Error GoTo 0
        lbcVehicle.AddItem "[All]", 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
