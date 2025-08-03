VERSION 5.00
Begin VB.Form NmAddr 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4215
   ClientLeft      =   1230
   ClientTop       =   2955
   ClientWidth     =   4845
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
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   10
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4215
   ScaleWidth      =   4845
   Begin VB.TextBox edcMediaType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   1605
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox pbcSepInvByVeh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   270
      ScaleHeight     =   210
      ScaleWidth      =   765
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox edcContactFax 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3570
      MaxLength       =   30
      TabIndex        =   17
      Top             =   3210
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox pbcISCI 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   675
      ScaleHeight     =   210
      ScaleWidth      =   765
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox edcContactPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2790
      MaxLength       =   30
      TabIndex        =   16
      Top             =   3300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox edcContactName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2310
      MaxLength       =   40
      TabIndex        =   15
      Top             =   2985
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox edcWebSite 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   660
      MaxLength       =   70
      TabIndex        =   13
      Top             =   2070
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.TextBox edcFTP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   570
      MaxLength       =   70
      TabIndex        =   10
      Top             =   2370
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.TextBox edcEMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   615
      MaxLength       =   70
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   495
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "cbcSelect"
      Top             =   315
      Width           =   3855
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4365
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3915
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   15
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2310
      Width           =   30
   End
   Begin VB.TextBox edcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   3
      Left            =   1050
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1590
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.TextBox edcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   990
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1470
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.TextBox edcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   870
      MaxLength       =   25
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.TextBox edcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   750
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.TextBox edcID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   600
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   3630
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
      HelpContextID   =   1
      Left            =   960
      TabIndex        =   19
      Top             =   3465
      Width           =   945
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
      HelpContextID   =   2
      Left            =   1980
      TabIndex        =   20
      Top             =   3465
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      HelpContextID   =   3
      Left            =   3000
      TabIndex        =   21
      Top             =   3465
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Enabled         =   0   'False
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
      HelpContextID   =   4
      Left            =   960
      TabIndex        =   22
      Top             =   3825
      Width           =   945
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
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
      HelpContextID   =   5
      Left            =   1980
      TabIndex        =   23
      Top             =   3825
      Width           =   945
   End
   Begin VB.CommandButton cmcMerge 
      Appearance      =   0  'Flat
      Caption         =   "&Merge into"
      Enabled         =   0   'False
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
      HelpContextID   =   6
      Left            =   4575
      TabIndex        =   24
      Top             =   2775
      Visible         =   0   'False
      Width           =   1020
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
      HelpContextID   =   7
      Left            =   4275
      TabIndex        =   25
      Top             =   3165
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   18
      Top             =   1725
      Width           =   15
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   1695
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   165
      Width           =   15
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4620
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3435
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4575
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3645
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   2445
      Index           =   2
      Left            =   810
      Picture         =   "Nmaddr.frx":0000
      ScaleHeight     =   2445
      ScaleWidth      =   4545
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.PictureBox pbcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1410
      Index           =   1
      Left            =   525
      Picture         =   "Nmaddr.frx":24882
      ScaleHeight     =   1410
      ScaleWidth      =   3675
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.PictureBox pbcNmAd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1755
      Index           =   0
      Left            =   195
      Picture         =   "Nmaddr.frx":27708
      ScaleHeight     =   1755
      ScaleWidth      =   3675
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.PictureBox plcNmAd 
      Height          =   2520
      Left            =   90
      ScaleHeight     =   2460
      ScaleWidth      =   4560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   795
      Width           =   4620
   End
   Begin VB.Label plcScreen 
      Caption         =   "Lock Box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2760
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   2745
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "NmAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Nmaddr.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Nmaddr.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Name/Address input screen code
Option Explicit
Option Compare Text
'Dim NmAd As Form
Dim smArfCallType As String 'L=Lock Box; A=EDI Service; C=Content Provider; P=Producer
Dim imPaintIndex As Integer '0=Log with Pledges; 1=Logs without Pledges (require Time Mapping); 2=C.P.
'NmAddr Box Field Areas
Dim tmCtrls(0 To 12)  As FIELDAREA   'Control fields
Dim imLBCtrls As Integer
Dim imMaxCtrls As Integer
Dim imBoxNo As Integer  'Current NmAddr Box
Dim tmArf As ARF       'ARF record image
Dim tmSrchKey As INTKEY0    'ARF key record image
Dim imRecLen As Integer        'ARF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmArf As Integer 'Name and address file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer 'True=cbcSelect has not had focus yet, used to branch to another control
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim smScreenCaption As String
Dim imUpdateAllowed As Integer    'User can update records
Dim imExportISCITo As Integer   '0=E-Mail; 1=FTP; 2=Manual (Neither)
Dim imSepInvByVeh As String     '0=Yes; 1=No

Const IDINDEX = 1       'ID control/index
Const NAMEADDRINDEX = 2 'Name/address control/field
Const EMAILINDEX = 6
Const FTPINDEX = 7
Const EXPORTISCIINDEX = 8
Const WEBSITEINDEX = 9
Const CONTACTNAMEINDEX = 10
Const CONTACTPHONEINDEX = 11
Const CONTACTFAXINDEX = 12

Const AIDINDEX = 1       'ID control/index
Const AMEDIATYPEINDEX = 2
Const ANAMEADDRINDEX = 3
Const ASEPINVBYVEHINDEX = 7

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
        If Not mReadRec(ilIndex, SETFORWRITE) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcNmAd(imPaintIndex).Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcID.Text = slStr
        End If
    End If
    For ilLoop = LBound(tmCtrls) To imMaxCtrls Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcNmAd_Paint imPaintIndex
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
Private Sub cbcSelect_DblClick()
    'Currently you can't get a double click event on a drop down
    cbcSelect_Click
    imBoxNo = -1
    pbcSTab.SetFocus
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
        If igArfCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgArfName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgArfName    'New name
            End If
            cbcSelect_Change
            If sgArfName <> "" Then
                mSetCommands
                If smArfCallType <> "A" Then
                    gFindMatch sgArfName, 1, cbcSelect
                    If gLastFound(cbcSelect) > 0 Then
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                Else
                    gFindMatch sgArfName, 0, cbcSelect
                    If gLastFound(cbcSelect) > 0 Then
                        cmcDone.SetFocus
                        Exit Sub
                    End If
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
    If smArfCallType <> "A" Then
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
    Else
        If cbcSelect.ListCount < 1 Then
            cmcCancel.SetFocus
            Exit Sub
        End If
        gCtrlGotFocus cbcSelect
        gFindMatch slSvText, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
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
    If igArfCallSource <> CALLNONE Then
        igArfCallSource = CALLCANCELLED
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
    If igArfCallSource <> CALLNONE Then
        sgArfName = edcID.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgArfName = "[New]"
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
    If igArfCallSource <> CALLNONE Then
        If sgArfName = "[New]" Then
            igArfCallSource = CALLCANCELLED
        Else
            igArfCallSource = CALLDONE
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
        For ilLoop = LBound(tmCtrls) To imMaxCtrls Step 1
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
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Adf.Btr", "AdfArfLkCode")  'adfArfLkCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Advertiser references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Adf.Btr", "AdfArfContrCode")  'adfArfContrCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Advertiser references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Adf.Btr", "AdfArfInvCode")  'adfarfInvCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Advertiser references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Agf.Btr", "AgfArfLkCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Agency references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Agf.Btr", "AgfArfCntrCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Agency references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Agf.Btr", "AgfArfInvCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Agency references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Eit.Mkd", "EitCommProvArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Affiliate ISCI Sent (eit) references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Fnf.Btr", "FnfProdArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Name references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Fnf.Btr", "FnfNetArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Name references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Vpf.Btr", "VpfFTPArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Option address references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Vpf.Btr", "VpfProducerArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Option address references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Vpf.Btr", "VpfProgProvArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Option address references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Vpf.Btr", "VpfCommProvArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Option address references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Vpf.Btr", "VpfAutoExptArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Option address references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(NmAddr, tmArf.iCode, "Vpf.Btr", "VpfAutoImptArfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Option address references this ID"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmArf.sID, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Arf.btr")
        ilRet = btrDelete(hmArf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", NmAddr
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcIDCode.Tag <> "" Then
        '    If slStamp = lbcIDCode.Tag Then
        '        lbcIDCode.Tag = FileDateTime(sgDBPath & "Arf.btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Arf.btr")
            End If
        End If
        'lbcIDCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcNmAd(imPaintIndex).Cls
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

Private Sub cmcMerge_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
End Sub

Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
'10071
'Private Sub cmcReport_Click()
'    Dim slStr As String
'    'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
'    '    Exit Sub
'    'End If
'    Select Case smArfCallType
'        Case "L"
'            igRptCallType = LOCKBOXESLIST
'        Case "A"
'            igRptCallType = EDISERVICESLIST
'        Case "C"
'            igRptCallType = 0
'        Case "P"
'            igRptCallType = 0
'        Case "N"
'            igRptCallType = 0
'        Case "K"
'            igRptCallType = 0
'    End Select
'    ''Screen.MousePointer = vbHourGlass  'Wait
'    'igChildDone = False
'    'edcLinkSrceDoneMsg.Text = ""
'    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
'        If igTestSystem Then
'            slStr = "NmAddr^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
'        Else
'            slStr = "NmAddr^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
'        End If
'    'Else
'    '    If igTestSystem Then
'    '        slStr = "NmAddr^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
'    '    Else
'    '        slStr = "NmAddr^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
'    '    End If
'    'End If
'    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
'    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
'    'NmAddr.Enabled = False
'    'Do While Not igChildDone
'    '    DoEvents
'    'Loop
'    'slStr = sgDoneMsg
'    'NmAddr.Enabled = True
'    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
'    'For ilLoop = 0 To 10
'    '    DoEvents
'    'Next ilLoop
'    ''Screen.MousePointer = vbDefault    'Default
'    sgCommandStr = slStr
'    RptList.Show vbModal
'End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If ((ilIndex > 0) And (smArfCallType <> "A")) Or ((ilIndex >= 0) And (smArfCallType = "A")) Then
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        pbcNmAd(imPaintIndex).Cls
        mMoveRecToCtrl
        For ilLoop = LBound(tmCtrls) To imMaxCtrls Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcNmAd_Paint imPaintIndex
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcNmAd(imPaintIndex).Cls
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
    Dim slID As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slID = Trim$(edcID.Text)   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imBoxNo = -1
'    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
'        cbcSelect.Text = slID
'    Else
'        cbcSelect.ListIndex = 0
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
'    mSetCommands
    ilCode = tmArf.iCode
    cbcSelect.Clear
    sgNameCodeTag = ""
    mPopulate
    For ilLoop = 0 To UBound(tgNameCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tgNameCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Val(slCode) = ilCode Then
            If smArfCallType <> "A" Then
                If cbcSelect.ListIndex = ilLoop + 1 Then
                    cbcSelect_Change
                Else
                    cbcSelect.ListIndex = ilLoop + 1
                End If
            Else
                If cbcSelect.ListIndex = ilLoop Then
                    cbcSelect_Change
                Else
                    cbcSelect.ListIndex = ilLoop
                End If
            End If
            Exit For
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    mSetCommands
    If cbcSelect.Enabled Then
        cbcSelect.SetFocus
    Else
        cmcDone.SetFocus
    End If
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub edcContactFax_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcContactFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcContactName_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcContactName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcContactPhone_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcContactPhone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEMail_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFTP_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcFTP_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcID_Change()
    mSetChg IDINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcID_GotFocus()
    gCtrlGotFocus edcID
End Sub
Private Sub edcID_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcID_LostFocus()
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKID()
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcMediaType_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcMediaType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNmAd_Change(iIndex As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcNmAd_GotFocus(iIndex As Integer)
    gCtrlGotFocus edcNmAd(iIndex)
End Sub
Private Sub edcNmAd_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    Dim slResult As String
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If (smArfCallType <> "L") And (smArfCallType <> "A") And (Index = 0) Then
        slStr = Chr(KeyAscii)
        slResult = gFileNameFilter(slStr)
        If slResult <> slStr Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcWebSite_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcWebSite_GotFocus()
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
    If sgArfCallType = "L" Then
        pbcNmAd(1).Visible = True
        If (igWinStatus(LOCKBOXESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcNmAd(1).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcNmAd(1).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    ElseIf sgArfCallType = "A" Then
        pbcNmAd(0).Visible = True
        If (igWinStatus(EDISERVICESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcNmAd(0).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcNmAd(0).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    ElseIf sgArfCallType = "C" Then
        pbcNmAd(2).Visible = True
        If (igWinStatus(VEHICLESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcNmAd(2).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcNmAd(2).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    ElseIf sgArfCallType = "P" Then
        pbcNmAd(2).Visible = True
        If (igWinStatus(VEHICLESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcNmAd(2).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcNmAd(2).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    ElseIf sgArfCallType = "N" Then
        pbcNmAd(2).Visible = True
        If (igWinStatus(FEEDNAMELIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcNmAd(2).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcNmAd(2).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    ElseIf sgArfCallType = "K" Then
        pbcNmAd(2).Visible = True
        If (igWinStatus(FEEDNAMELIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcNmAd(2).Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcNmAd(2).Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    NmAddr.Refresh
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
    Erase tgNameCode

    btrExtClear hmArf   'Clear any previous extend operation
    ilRet = btrClose(hmArf)
    btrDestroy hmArf

    Set NmAddr = Nothing   'Remove data segment

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
'*             Created:4/20/93       By:D. LeVine      *
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

    edcID.Text = ""
    edcMediaType.Text = ""
    For ilLoop = 0 To 3 Step 1
        edcNmAd(ilLoop).Text = ""
    Next ilLoop
    edcEMail.Text = ""
    edcFTP.Text = ""
    edcWebSite.Text = ""
    imExportISCITo = -1
    imSepInvByVeh = -1
    edcContactName.Text = ""
    edcContactPhone.Text = ""
    mMoveCtrlToRec False
    For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > imMaxCtrls) Then
        Exit Sub
    End If

    If (smArfCallType = "A") Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case AIDINDEX 'ID
                edcID.Width = tmCtrls(ilBoxNo).fBoxW
                edcID.MaxLength = 10
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcID.Visible = True  'Set visibility
                edcID.SetFocus
            Case AMEDIATYPEINDEX
                edcMediaType.Width = tmCtrls(ilBoxNo).fBoxW
                edcMediaType.MaxLength = 1
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcMediaType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcMediaType.Visible = True
                edcMediaType.SetFocus
            Case ANAMEADDRINDEX 'Name & Address
                edcNmAd(0).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(0).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(0), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(0).Visible = True  'Set visibility
                edcNmAd(0).SetFocus
            Case ANAMEADDRINDEX + 1 'Name & Address
                edcNmAd(1).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(1).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(1), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(1).Visible = True  'Set visibility
                edcNmAd(1).SetFocus
            Case ANAMEADDRINDEX + 2 'Name & Address
                edcNmAd(2).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(2).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(2), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(2).Visible = True  'Set visibility
                edcNmAd(2).SetFocus
            Case ANAMEADDRINDEX + 3 'Name & Address
                edcNmAd(3).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(3).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(3), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(3).Visible = True  'Set visibility
                edcNmAd(3).SetFocus
            Case ASEPINVBYVEHINDEX
                If imSepInvByVeh < 0 Then
                    imSepInvByVeh = 1    'No
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcSepInvByVeh.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcNmAd(imPaintIndex), pbcSepInvByVeh, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcSepInvByVeh_Paint
                pbcSepInvByVeh.Visible = True
                pbcSepInvByVeh.SetFocus
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case IDINDEX 'ID
                edcID.Width = tmCtrls(ilBoxNo).fBoxW
                If (smArfCallType = "L") Then  'Lock Box
                    edcID.MaxLength = 10
                Else
                    edcID.MaxLength = 40
                End If
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcID, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcID.Visible = True  'Set visibility
                edcID.SetFocus
            Case NAMEADDRINDEX  'Name & Address
                edcNmAd(0).Width = tmCtrls(ilBoxNo).fBoxW
                If (smArfCallType = "L") Then  'Lock Box
                    edcNmAd(0).MaxLength = 25
                Else
                    edcNmAd(0).MaxLength = 10
                    If edcNmAd(0).Text = "" Then
                        edcNmAd(0).Text = gFileNameFilter(Left(edcID, 10))
                    End If
                End If
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(0), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(0).Visible = True  'Set visibility
                edcNmAd(0).SetFocus
            Case NAMEADDRINDEX + 1 'Name & Address
                edcNmAd(1).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(1).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(1), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(1).Visible = True  'Set visibility
                edcNmAd(1).SetFocus
            Case NAMEADDRINDEX + 2 'Name & Address
                edcNmAd(2).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(2).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(2), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(2).Visible = True  'Set visibility
                edcNmAd(2).SetFocus
            Case NAMEADDRINDEX + 3 'Name & Address
                edcNmAd(3).Width = tmCtrls(ilBoxNo).fBoxW
                edcNmAd(3).MaxLength = 25
                gMoveFormCtrl pbcNmAd(imPaintIndex), edcNmAd(3), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcNmAd(3).Visible = True  'Set visibility
                edcNmAd(3).SetFocus
            Case EMAILINDEX
                edcEMail.Width = tmCtrls(ilBoxNo).fBoxW
                edcEMail.MaxLength = 70
                gMoveFormCtrl pbcNmAd(2), edcEMail, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcEMail.Visible = True  'Set visibility
                edcEMail.SetFocus
            Case FTPINDEX
                edcFTP.Width = tmCtrls(ilBoxNo).fBoxW
                edcFTP.MaxLength = 70
                gMoveFormCtrl pbcNmAd(2), edcFTP, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcFTP.Visible = True  'Set visibility
                edcFTP.SetFocus
            Case EXPORTISCIINDEX
                If imExportISCITo < 0 Then
                    imExportISCITo = 2    'Manual
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcISCI.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcNmAd(2), pbcISCI, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcISCI_Paint
                pbcISCI.Visible = True
                pbcISCI.SetFocus
            Case WEBSITEINDEX
                edcWebSite.Width = tmCtrls(ilBoxNo).fBoxW
                edcWebSite.MaxLength = 70
                gMoveFormCtrl pbcNmAd(2), edcWebSite, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcWebSite.Visible = True  'Set visibility
                edcWebSite.SetFocus
            Case CONTACTNAMEINDEX
                edcContactName.Width = tmCtrls(ilBoxNo).fBoxW
                edcContactName.MaxLength = 40
                gMoveFormCtrl pbcNmAd(2), edcContactName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcContactName.Visible = True  'Set visibility
                edcContactName.SetFocus
            Case CONTACTPHONEINDEX
                edcContactPhone.Width = tmCtrls(ilBoxNo).fBoxW
                edcContactPhone.MaxLength = 30
                gMoveFormCtrl pbcNmAd(2), edcContactPhone, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcContactPhone.Visible = True  'Set visibility
                edcContactPhone.SetFocus
            Case CONTACTFAXINDEX
                edcContactFax.Width = tmCtrls(ilBoxNo).fBoxW
                edcContactFax.MaxLength = 20
                gMoveFormCtrl pbcNmAd(2), edcContactFax, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                edcContactFax.Visible = True  'Set visibility
                edcContactFax.SetFocus
        End Select
    End If
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    imLBCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    smArfCallType = sgArfCallType
    mInitBox
    NmAddr.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone NmAddr
    'NmAddr.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    If smArfCallType = "L" Then
        imPaintIndex = 1
'        Set NmAd = NmAd_L
        smScreenCaption = "Lock Boxes"
    ElseIf smArfCallType = "A" Then
        imPaintIndex = 0
'        Set NmAd = NmAd_A
        smScreenCaption = "EDI Services"
    ElseIf smArfCallType = "C" Then
        imPaintIndex = 2
        smScreenCaption = "Content Provider"
    ElseIf smArfCallType = "P" Then
        imPaintIndex = 2
        smScreenCaption = "Producer"
    ElseIf smArfCallType = "N" Then
        imPaintIndex = 2
        smScreenCaption = "Network/Rep"
    ElseIf smArfCallType = "K" Then
        imPaintIndex = 2
        smScreenCaption = "Feed Producer"
    End If
    If (smArfCallType = "L") Or (smArfCallType = "A") Then 'Lock Box
        edcID.MaxLength = 10
        edcNmAd(0).MaxLength = 25
    Else
        edcID.MaxLength = 40
        edcNmAd(0).MaxLength = 10
    End If
    edcNmAd(1).MaxLength = 25
    edcNmAd(2).MaxLength = 25
    edcNmAd(3).MaxLength = 25

    imRecLen = Len(tmArf)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current NmAddr Box to N/A
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmArf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmArf, "", sgDBPath & "arf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", NmAddr
    On Error GoTo 0
'    gCenterModalForm NmAddr
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
        mSetCommands
    End If
    'plcScreen.Cls
    'plcScreen_Paint
    plcScreen.Caption = smScreenCaption
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                     *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    flTextHeight = pbcNmAd(0).TextHeight("1")
    'Position panel and picture areas with panel

    If (smArfCallType = "L") Or (smArfCallType = "A") Then 'Lock Box
        If (smArfCallType = "A") And (tgSpf.sBCombine <> "N") Then
            imMaxCtrls = 6
            pbcNmAd(0).Height = 1755
            NmAddr.Height = 3210 + fgBoxStH + 15
            plcNmAd.Move 510, 795, pbcNmAd(0).Width + fgPanelAdj, pbcNmAd(0).Height + fgPanelAdj
            cmcDone.Top = 2400 + fgBoxStH + 15
            cmcUndo.Top = 2760 + fgBoxStH + 15
        Else
            imMaxCtrls = 5
            pbcNmAd(0).Height = 1410
            NmAddr.Height = 3210
            plcNmAd.Move 510, 795, pbcNmAd(0).Width + fgPanelAdj, pbcNmAd(0).Height + fgPanelAdj
            cmcDone.Top = 2400
            cmcUndo.Top = 2760
        End If
        pbcNmAd(0).Move plcNmAd.Left + fgBevelX, plcNmAd.Top + fgBevelY
        pbcNmAd(1).Move plcNmAd.Left + fgBevelX, plcNmAd.Top + fgBevelY
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        '9/12/16: Removed Merge button as no support code added to Merge.Frm
        'cmcErase.Top = cmcDone.Top
        cmcErase.Top = cmcUndo.Top
        'cmcMerge.Top = cmcUndo.Top
        cmcReport.Top = cmcUndo.Top
    'Set either Lock box or EDI service picture visible
    'Know done in active
'    If smArfCallType = "L" Then 'Lock Box
'        pbcNmAd(1).Visible = True
'        pbcNmAd(0).Visible = False
'    Else    'EDI Service
'        pbcNmAd(0).Visible = True
'        pbcNmAd(1).Visible = False
'    End If
        If (smArfCallType = "A") Then
            'ID
            gSetCtrl tmCtrls(AIDINDEX), 30, 30, 2475, fgBoxStH
            'Media Type
            gSetCtrl tmCtrls(AMEDIATYPEINDEX), 2520, 30, 1125, fgBoxStH
            tmCtrls(AMEDIATYPEINDEX).iReq = False
            'Name & Address
            gSetCtrl tmCtrls(ANAMEADDRINDEX), 30, tmCtrls(AIDINDEX).fBoxY + fgStDeltaY, 3630, fgBoxStH
            gSetCtrl tmCtrls(ANAMEADDRINDEX + 1), 30, tmCtrls(ANAMEADDRINDEX).fBoxY + flTextHeight, tmCtrls(ANAMEADDRINDEX).fBoxW, flTextHeight
            gSetCtrl tmCtrls(ANAMEADDRINDEX + 2), 30, tmCtrls(ANAMEADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(ANAMEADDRINDEX).fBoxW, flTextHeight
            gSetCtrl tmCtrls(ANAMEADDRINDEX + 3), 30, tmCtrls(ANAMEADDRINDEX + 2).fBoxY + flTextHeight, tmCtrls(ANAMEADDRINDEX).fBoxW, flTextHeight
            If (tgSpf.sBCombine <> "N") Then
                imMaxCtrls = 7
                'Separate EDI Invoices by Vehicle
                gSetCtrl tmCtrls(ASEPINVBYVEHINDEX), 30, 1410, 3630, fgBoxStH
            Else
                imMaxCtrls = 6
            End If
        Else
            imMaxCtrls = 5
            'ID
            gSetCtrl tmCtrls(IDINDEX), 30, 30, 3630, fgBoxStH
            'Name & Address
            gSetCtrl tmCtrls(NAMEADDRINDEX), 30, tmCtrls(IDINDEX).fBoxY + fgStDeltaY, 3630, fgBoxStH
            gSetCtrl tmCtrls(NAMEADDRINDEX + 1), 30, tmCtrls(NAMEADDRINDEX).fBoxY + flTextHeight, tmCtrls(NAMEADDRINDEX).fBoxW, flTextHeight
            gSetCtrl tmCtrls(NAMEADDRINDEX + 2), 30, tmCtrls(NAMEADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(NAMEADDRINDEX).fBoxW, flTextHeight
            gSetCtrl tmCtrls(NAMEADDRINDEX + 3), 30, tmCtrls(NAMEADDRINDEX + 2).fBoxY + flTextHeight, tmCtrls(NAMEADDRINDEX).fBoxW, flTextHeight
            'If (smArfCallType = "A") Then
            '    If (tgSpf.sBCombine <> "N") Then
            '        'Separate EDI Invoices by Vehicle
            '        gSetCtrl tmCtrls(SEPINVBYVEHINDEX), 30, 1410, 3630, fgBoxStH
            '    End If
            'End If
        End If
    Else
        flTextHeight = flTextHeight - 35
        NmAddr.Height = 4305
        plcNmAd.Move 105, 795, pbcNmAd(2).Width + fgPanelAdj, pbcNmAd(2).Height + fgPanelAdj
        pbcNmAd(2).Move plcNmAd.Left + fgBevelX, plcNmAd.Top + fgBevelY
        cmcDone.Top = 3465
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        '9/12/16: Removed Merge button as no support code added to Merge.Frm
        'cmcErase.Top = cmcDone.Top
        cmcUndo.Top = 3825
        cmcErase.Top = cmcUndo.Top
        'cmcMerge.Top = cmcUndo.Top
        cmcReport.Top = cmcUndo.Top
        imMaxCtrls = 12
        'Name
        gSetCtrl tmCtrls(IDINDEX), 30, 30, 3435, fgBoxStH
        'Abbreviation
        gSetCtrl tmCtrls(NAMEADDRINDEX), 3480, 30, 1035, fgBoxStH
        'Address
        gSetCtrl tmCtrls(NAMEADDRINDEX + 1), 30, tmCtrls(IDINDEX).fBoxY + fgStDeltaY, 4485, fgBoxStH
        tmCtrls(NAMEADDRINDEX + 1).iReq = False
        gSetCtrl tmCtrls(NAMEADDRINDEX + 2), 30, tmCtrls(NAMEADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(NAMEADDRINDEX + 1).fBoxW, flTextHeight
        tmCtrls(NAMEADDRINDEX + 2).iReq = False
        gSetCtrl tmCtrls(NAMEADDRINDEX + 3), 30, tmCtrls(NAMEADDRINDEX + 2).fBoxY + flTextHeight, tmCtrls(NAMEADDRINDEX + 1).fBoxW, flTextHeight
        tmCtrls(NAMEADDRINDEX + 3).iReq = False
        'E-Mail
        gSetCtrl tmCtrls(EMAILINDEX), 30, 1065, 4485, fgBoxStH
        tmCtrls(EMAILINDEX).iReq = False
        'FTP
        gSetCtrl tmCtrls(FTPINDEX), 30, tmCtrls(EMAILINDEX).fBoxY + fgStDeltaY, 4485, fgBoxStH
        tmCtrls(EMAILINDEX).iReq = False
        'Export ISCI To
        gSetCtrl tmCtrls(EXPORTISCIINDEX), 30, tmCtrls(FTPINDEX).fBoxY + fgStDeltaY, 945, fgBoxStH
        tmCtrls(EXPORTISCIINDEX).iReq = False
        'Web Site
        gSetCtrl tmCtrls(WEBSITEINDEX), 990, tmCtrls(EXPORTISCIINDEX).fBoxY, 3525, fgBoxStH
        tmCtrls(WEBSITEINDEX).iReq = False
        'Contact Name
        gSetCtrl tmCtrls(CONTACTNAMEINDEX), 30, tmCtrls(EXPORTISCIINDEX).fBoxY + fgStDeltaY, 1890, fgBoxStH
        tmCtrls(CONTACTNAMEINDEX).iReq = False
        'Contact Phone
        gSetCtrl tmCtrls(CONTACTPHONEINDEX), 1935, tmCtrls(CONTACTNAMEINDEX).fBoxY, 1425, fgBoxStH
        tmCtrls(CONTACTPHONEINDEX).iReq = False
        'Contact Fax
        gSetCtrl tmCtrls(CONTACTFAXINDEX), 3375, tmCtrls(CONTACTNAMEINDEX).fBoxY, 1140, fgBoxStH
        tmCtrls(CONTACTFAXINDEX).iReq = False
    End If
    cbcSelect.Left = plcNmAd.Left
    cbcSelect.Width = plcNmAd.Width
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
    Dim ilLoop As Integer
    tmArf.sEDIMediaType = ""
    If (smArfCallType = "L") Then 'Lock Box
        If Not ilTestChg Or tmCtrls(IDINDEX).iChg Then
            tmArf.sID = edcID.Text
        End If
        For ilLoop = 0 To 3 Step 1
            If Not ilTestChg Or tmCtrls(NAMEADDRINDEX + ilLoop).iChg Then
                tmArf.sNmAd(ilLoop) = edcNmAd(ilLoop).Text
            End If
        Next ilLoop
    ElseIf (smArfCallType = "A") Then 'Agency
        If Not ilTestChg Or tmCtrls(AIDINDEX).iChg Then
            tmArf.sID = edcID.Text
        End If
        If Not ilTestChg Or tmCtrls(AMEDIATYPEINDEX).iChg Then
            tmArf.sEDIMediaType = edcMediaType.Text
        End If
        For ilLoop = 0 To 3 Step 1
            If Not ilTestChg Or tmCtrls(ANAMEADDRINDEX + ilLoop).iChg Then
                tmArf.sNmAd(ilLoop) = edcNmAd(ilLoop).Text
            End If
        Next ilLoop
        If (tgSpf.sBCombine <> "N") Then
            If Not ilTestChg Or tmCtrls(ASEPINVBYVEHINDEX).iChg Then
                If imSepInvByVeh = 0 Then
                    tmArf.sSepInvByVeh = "Y"
                Else
                    tmArf.sSepInvByVeh = "N"
                End If
            End If
        Else
            tmArf.sSepInvByVeh = "N"
        End If
    Else
        If Not ilTestChg Or tmCtrls(IDINDEX).iChg Then
            tmArf.sName = edcID.Text
        End If
        If Not ilTestChg Or tmCtrls(NAMEADDRINDEX).iChg Then
            tmArf.sID = edcNmAd(0).Text
        End If
        For ilLoop = 1 To 3 Step 1
            If Not ilTestChg Or tmCtrls(NAMEADDRINDEX + ilLoop).iChg Then
                tmArf.sNmAd(ilLoop - 1) = edcNmAd(ilLoop).Text
            End If
        Next ilLoop
        tmArf.sNmAd(3) = ""
        If Not ilTestChg Or tmCtrls(EMAILINDEX).iChg Then
            tmArf.sEMail = edcEMail.Text
        End If
        If Not ilTestChg Or tmCtrls(FTPINDEX).iChg Then
            tmArf.sFTP = edcFTP.Text
        End If
        If Not ilTestChg Or tmCtrls(EXPORTISCIINDEX).iChg Then
            If imExportISCITo = 0 Then
                tmArf.sSendISCITo = "E"
            ElseIf imExportISCITo = 1 Then
                tmArf.sSendISCITo = "F"
            Else
                tmArf.sSendISCITo = "M"
            End If
        End If
        If Not ilTestChg Or tmCtrls(WEBSITEINDEX).iChg Then
            tmArf.sWebSite = edcWebSite.Text
        End If
        If Not ilTestChg Or tmCtrls(CONTACTNAMEINDEX).iChg Then
            tmArf.sContactName = edcContactName.Text
        End If
        If Not ilTestChg Or tmCtrls(CONTACTPHONEINDEX).iChg Then
            tmArf.sContactPhone = edcContactPhone.Text
        End If
        If Not ilTestChg Or tmCtrls(CONTACTFAXINDEX).iChg Then
            tmArf.sContactFax = edcContactFax.Text
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
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
    If (smArfCallType = "L") Then 'Lock Box
        edcID.Text = Trim$(tmArf.sID)
        For ilLoop = 0 To 3 Step 1
            edcNmAd(ilLoop).Text = Trim$(tmArf.sNmAd(ilLoop))
        Next ilLoop
    ElseIf (smArfCallType = "A") Then 'Agency
        edcID.Text = Trim$(tmArf.sID)
        edcMediaType.Text = Trim(tmArf.sEDIMediaType)
        For ilLoop = 0 To 3 Step 1
            edcNmAd(ilLoop).Text = Trim$(tmArf.sNmAd(ilLoop))
        Next ilLoop
        If tmArf.sSepInvByVeh = "Y" Then
            imSepInvByVeh = 0
        Else
            imSepInvByVeh = 1
        End If
    Else
        edcID.Text = Trim$(tmArf.sName)
        edcNmAd(0).Text = Trim$(tmArf.sID)
        For ilLoop = 0 To 2 Step 1
            edcNmAd(ilLoop + 1).Text = Trim$(tmArf.sNmAd(ilLoop))
        Next ilLoop
        edcEMail.Text = Trim$(tmArf.sEMail)
        edcFTP.Text = Trim$(tmArf.sFTP)
        If tmArf.sSendISCITo = "E" Then
            imExportISCITo = 0
        ElseIf tmArf.sSendISCITo = "F" Then
            imExportISCITo = 1
        Else
            imExportISCITo = 2
        End If
        edcWebSite.Text = Trim$(tmArf.sWebSite)
        edcContactName.Text = Trim$(tmArf.sContactName)
        edcContactPhone.Text = Trim$(tmArf.sContactPhone)
    End If
    For ilLoop = LBound(tmCtrls) To imMaxCtrls Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKID                           *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that ID is unique          *
'*                                                     *
'*******************************************************
Private Function mOKID()
    Dim slStr As String
    If edcID.Text <> "" Then    'Test name
        slStr = edcID.Text
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If edcID.Text = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    If smArfCallType = "L" Then 'Lock Box
                        MsgBox "City already defined, enter a different city", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    ElseIf smArfCallType = "A" Then
                        MsgBox "Service name already specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    ElseIf smArfCallType = "C" Then
                        MsgBox "Content Provider already specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    ElseIf smArfCallType = "P" Then
                        MsgBox "Producer already specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    ElseIf smArfCallType = "N" Then
                        MsgBox "Network/Rep already specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    ElseIf smArfCallType = "K" Then
                        MsgBox "Feed Producer already specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    End If
                    If (smArfCallType = "L") Or (smArfCallType = "A") Then 'Lock Box
                        edcID.Text = Trim$(tmArf.sID) 'Reset text
                    Else
                        edcID.Text = Trim$(tmArf.sName) 'Reset text
                    End If
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKID = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKID = True
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
    'gInitStdAlone NmAddr, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", sgArfCallType)    'Get call type "L" or "A"
    ilRet = gParseItem(slCommand, 4, "\", slStr)    'Get call source
    igArfCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igArfCallSource = CALLNONE
    '    sgArfCallType = "L"
    'End If
    If igArfCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 5, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgArfName = slStr
        Else
            sgArfName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    ReDim ilOffset(0) As Integer
    imPopReqd = False
    ilfilter(0) = CHARFILTER
    slFilter(0) = smArfCallType
    ilOffset(0) = gFieldOffset("Arf", "ArfType") '2
    'ilRet = gIMoveListBox(NmAddr, cbcSelect, lbcIDCode, "Arf.Btr", gFieldOffset("arf", "ArfID"), 10, ilFilter(), slFilter(), ilOffset())
    If (smArfCallType = "L") Or (smArfCallType = "A") Then 'Lock Box
        ilRet = gIMoveListBox(NmAddr, cbcSelect, tgNameCode(), sgNameCodeTag, "Arf.Btr", gFieldOffset("arf", "ArfID"), 10, ilfilter(), slFilter(), ilOffset())
    Else
        ilRet = gIMoveListBox(NmAddr, cbcSelect, tgNameCode(), sgNameCodeTag, "Arf.Btr", gFieldOffset("arf", "ArfName"), 40, ilfilter(), slFilter(), ilOffset())
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", NmAddr
        On Error GoTo 0
        If smArfCallType <> "A" Then
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
        End If
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
Private Function mReadRec(ilSelectIndex, ilForUpdate As Integer)
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slIDCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    If smArfCallType <> "A" Then
        slIDCode = tgNameCode(ilSelectIndex - 1).sKey  'lbcIDCode.List(ilSelectIndex - 1)
    Else
        slIDCode = tgNameCode(ilSelectIndex).sKey   'lbcIDCode.List(ilSelectIndex - 1)
    End If
    ilRet = gParseItem(slIDCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", NmAddr
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmArf, tmArf, imRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", NmAddr
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
'*             Created:4/21/93       By:D. LeVine      *
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
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Arf.btr")
        'If Len(lbcIDCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcIDCode.Tag, Len(lbcIDCode.Tag) - Len(slStamp))
        'End If
        If Len(sgNameCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(sgNameCodeTag, Len(sgNameCodeTag) - Len(slStamp))
        End If
        If ((imSelectedIndex > 0) And (smArfCallType <> "A")) Or ((imSelectedIndex >= 0) And (smArfCallType = "A")) Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        If (imSelectedIndex = 0) And (smArfCallType <> "A") Then 'New selected
            tmArf.iCode = 0  'Autoincrement
            tmArf.sType = smArfCallType 'L=Lock box; A=EDI service; S=Sales Office
            tmArf.iMerge = 0   'Merge code number
            ilRet = btrInsert(hmArf, tmArf, imRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmArf, tmArf, imRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, NmAddr
    On Error GoTo 0
    'If lbcIDCode.Tag <> "" Then
    '    If slStamp = lbcIDCode.Tag Then
    '        lbcIDCode.Tag = FileDateTime(sgDBPath & "Arf.btr")
    '        If Len(slStamp) > Len(lbcIDCode.Tag) Then
    '            lbcIDCode.Tag = lbcIDCode.Tag & Right$(slStamp, Len(slStamp) - Len(lbcIDCode.Tag))
    '        End If
    '    End If
    'End If
'    If sgNameCodeTag <> "" Then
'        If slStamp = sgNameCodeTag Then
'            sgNameCodeTag = gFileDateTime(sgDBPath & "Arf.btr")
'            If Len(slStamp) > Len(sgNameCodeTag) Then
'                sgNameCodeTag = sgNameCodeTag & right$(slStamp, Len(slStamp) - Len(sgNameCodeTag))
'            End If
'        End If
'    End If
'    If imSelectedIndex <> 0 Then
'        'lbcIDCode.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
'        cbcSelect.RemoveItem imSelectedIndex
'    End If
'    cbcSelect.RemoveItem 0 'Remove [New]
'    slID = RTrim$(tmArf.sID)
'    cbcSelect.AddItem slID
'    slID = tmArf.sID + "\" + LTrim$(Str$(tmArf.iCode)) 'slID + "\" + LTrim$(Str$(tmArf.iCode))
'    'lbcIDCode.AddItem slID
'    gAddItemToSortCode slID, tgNameCode(), True
'    cbcSelect.AddItem "[New]", 0
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
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
                If ((imSelectedIndex > 0) And (smArfCallType <> "A")) Or ((imSelectedIndex >= 0) And (smArfCallType = "A")) Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcID.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcNmAd_Paint imPaintIndex
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
    Dim ilLoop As Integer   'For loop control parameter
    If ilBoxNo < LBound(tmCtrls) Or ilBoxNo > imMaxCtrls Then
'        mSetCommands
        Exit Sub
    End If

    If (smArfCallType = "A") Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case AIDINDEX 'ID
                gSetChgFlag tmArf.sID, edcID, tmCtrls(ilBoxNo)
            Case AMEDIATYPEINDEX
                gSetChgFlag tmArf.sEDIMediaType, edcMediaType, tmCtrls(ilBoxNo)
            Case ANAMEADDRINDEX
                gSetChgFlag tmArf.sNmAd(0), edcNmAd(0), tmCtrls(ANAMEADDRINDEX)
            Case ANAMEADDRINDEX + 1 To ANAMEADDRINDEX + 3 'Name & Address
                For ilLoop = 0 To 3 Step 1  'Set visibility
                    gSetChgFlag tmArf.sNmAd(ilLoop), edcNmAd(ilLoop), tmCtrls(ANAMEADDRINDEX + ilLoop)
                Next ilLoop
            Case ASEPINVBYVEHINDEX
                'Set as part of pbcSepInvByVeh
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case IDINDEX 'ID
                If (smArfCallType = "L") Then  'Lock Box
                    gSetChgFlag tmArf.sID, edcID, tmCtrls(ilBoxNo)
                Else
                    gSetChgFlag tmArf.sName, edcID, tmCtrls(ilBoxNo)
                End If
            Case NAMEADDRINDEX
                If (smArfCallType = "L") Then 'Lock Box
                    gSetChgFlag tmArf.sNmAd(0), edcNmAd(0), tmCtrls(NAMEADDRINDEX)
                Else
                    gSetChgFlag tmArf.sID, edcNmAd(0), tmCtrls(NAMEADDRINDEX)
                End If
            Case NAMEADDRINDEX + 1 To NAMEADDRINDEX + 3 'Name & Address
                For ilLoop = 0 To 3 Step 1  'Set visibility
                    gSetChgFlag tmArf.sNmAd(ilLoop), edcNmAd(ilLoop), tmCtrls(NAMEADDRINDEX + ilLoop)
                Next ilLoop
            Case EMAILINDEX
                gSetChgFlag tmArf.sEMail, edcEMail, tmCtrls(ilBoxNo)
            Case FTPINDEX
                gSetChgFlag tmArf.sFTP, edcFTP, tmCtrls(ilBoxNo)
            Case EXPORTISCIINDEX
                'In side of plcISCI
            Case WEBSITEINDEX
                gSetChgFlag tmArf.sWebSite, edcWebSite, tmCtrls(ilBoxNo)
            Case CONTACTNAMEINDEX
                gSetChgFlag tmArf.sContactName, edcContactName, tmCtrls(ilBoxNo)
            Case CONTACTPHONEINDEX
                gSetChgFlag tmArf.sContactPhone, edcContactPhone, tmCtrls(ilBoxNo)
            Case CONTACTFAXINDEX
                gSetChgFlag tmArf.sContactFax, edcContactFax, tmCtrls(ilBoxNo)
        End Select
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) Then
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
    If smArfCallType <> "A" Then
        If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
            cmcErase.Enabled = True
        Else
            cmcErase.Enabled = False
        End If
    Else
        cmcErase.Enabled = False
    End If
'Merge not coded- therefore disallow
'    'Merge set only if change mode
'    If (imSelectedIndex > 0) And (tgUrf(0).sMerge = "I") Then
'        cmcMerge.Enabled = True
'    Else
'        cmcMerge.Enabled = False
'    End If
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
'*             Created:4/20/93       By:D. LeVine      *
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
    If (ilBoxNo < LBound(tmCtrls)) Or (ilBoxNo > imMaxCtrls) Then
        Exit Sub
    End If

    If (smArfCallType = "A") Then
        Select Case ilBoxNo 'Branch on box type (control)
            Case AIDINDEX 'ID
                edcID.Visible = False  'Set visibility
                slStr = edcID.Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case AMEDIATYPEINDEX
                edcMediaType.Visible = False  'Set visibility
                slStr = edcMediaType.Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case ANAMEADDRINDEX 'Name & Address
                edcNmAd(0).Visible = False
                slStr = edcNmAd(0).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case ANAMEADDRINDEX + 1 'Name & Address
                edcNmAd(1).Visible = False
                slStr = edcNmAd(1).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                If slStr = "" Then
                    edcNmAd(2).Text = ""
                    slStr = edcNmAd(2).Text
                    gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(NAMEADDRINDEX + 2)
                    gPaintArea pbcNmAd(imPaintIndex), tmCtrls(NAMEADDRINDEX + 2).fBoxX, tmCtrls(NAMEADDRINDEX + 2).fBoxY, tmCtrls(NAMEADDRINDEX + 2).fBoxW - 15, tmCtrls(NAMEADDRINDEX + 2).fBoxH - 15, WHITE
                    edcNmAd(3).Text = ""
                    slStr = edcNmAd(3).Text
                    gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(NAMEADDRINDEX + 3)
                    gPaintArea pbcNmAd(imPaintIndex), tmCtrls(NAMEADDRINDEX + 3).fBoxX, tmCtrls(NAMEADDRINDEX + 3).fBoxY, tmCtrls(NAMEADDRINDEX + 3).fBoxW - 15, tmCtrls(NAMEADDRINDEX + 3).fBoxH - 15, WHITE
                End If
            Case ANAMEADDRINDEX + 2 'Name & Address
                edcNmAd(2).Visible = False
                slStr = edcNmAd(2).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                If slStr = "" Then
                    edcNmAd(3).Text = ""
                    slStr = edcNmAd(3).Text
                    gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(NAMEADDRINDEX + 3)
                    gPaintArea pbcNmAd(imPaintIndex), tmCtrls(NAMEADDRINDEX + 3).fBoxX, tmCtrls(NAMEADDRINDEX + 3).fBoxY, tmCtrls(NAMEADDRINDEX + 3).fBoxW - 15, tmCtrls(NAMEADDRINDEX + 3).fBoxH - 15, WHITE
                End If
            Case ANAMEADDRINDEX + 3 'Name & Address
                edcNmAd(3).Visible = False
                slStr = edcNmAd(3).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case ASEPINVBYVEHINDEX 'E-Mail
                pbcSepInvByVeh.Visible = False  'Set visibility
                If imSepInvByVeh = 0 Then
                    slStr = "Yes"
                ElseIf imSepInvByVeh = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
        End Select
    Else
        Select Case ilBoxNo 'Branch on box type (control)
            Case IDINDEX 'ID
                edcID.Visible = False  'Set visibility
                slStr = edcID.Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case NAMEADDRINDEX 'Name & Address
                edcNmAd(0).Visible = False
                slStr = edcNmAd(0).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case NAMEADDRINDEX + 1 'Name & Address
                edcNmAd(1).Visible = False
                slStr = edcNmAd(1).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                If slStr = "" Then
                    edcNmAd(2).Text = ""
                    slStr = edcNmAd(2).Text
                    gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(NAMEADDRINDEX + 2)
                    gPaintArea pbcNmAd(imPaintIndex), tmCtrls(NAMEADDRINDEX + 2).fBoxX, tmCtrls(NAMEADDRINDEX + 2).fBoxY, tmCtrls(NAMEADDRINDEX + 2).fBoxW - 15, tmCtrls(NAMEADDRINDEX + 2).fBoxH - 15, WHITE
                    edcNmAd(3).Text = ""
                    slStr = edcNmAd(3).Text
                    gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(NAMEADDRINDEX + 3)
                    gPaintArea pbcNmAd(imPaintIndex), tmCtrls(NAMEADDRINDEX + 3).fBoxX, tmCtrls(NAMEADDRINDEX + 3).fBoxY, tmCtrls(NAMEADDRINDEX + 3).fBoxW - 15, tmCtrls(NAMEADDRINDEX + 3).fBoxH - 15, WHITE
                End If
            Case NAMEADDRINDEX + 2 'Name & Address
                edcNmAd(2).Visible = False
                slStr = edcNmAd(2).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
                If slStr = "" Then
                    edcNmAd(3).Text = ""
                    slStr = edcNmAd(3).Text
                    gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(NAMEADDRINDEX + 3)
                    gPaintArea pbcNmAd(imPaintIndex), tmCtrls(NAMEADDRINDEX + 3).fBoxX, tmCtrls(NAMEADDRINDEX + 3).fBoxY, tmCtrls(NAMEADDRINDEX + 3).fBoxW - 15, tmCtrls(NAMEADDRINDEX + 3).fBoxH - 15, WHITE
                End If
            Case NAMEADDRINDEX + 3 'Name & Address
                edcNmAd(3).Visible = False
                slStr = edcNmAd(3).Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case EMAILINDEX 'E-Mail
                edcEMail.Visible = False  'Set visibility
                slStr = edcEMail.Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case FTPINDEX 'FTP
                edcFTP.Visible = False  'Set visibility
                slStr = edcFTP.Text
                gSetShow pbcNmAd(imPaintIndex), slStr, tmCtrls(ilBoxNo)
            Case EXPORTISCIINDEX
                pbcISCI.Visible = False  'Set visibility
                If imExportISCITo = 0 Then
                    slStr = "E-Mail"
                ElseIf imExportISCITo = 1 Then
                    slStr = "FTP"
                ElseIf imExportISCITo = 2 Then
                    slStr = "Manual"
                Else
                    slStr = ""
                End If
                gSetShow pbcNmAd(2), slStr, tmCtrls(ilBoxNo)
            Case WEBSITEINDEX 'Web Site
                edcWebSite.Visible = False  'Set visibility
                slStr = edcWebSite.Text
                gSetShow pbcNmAd(2), slStr, tmCtrls(ilBoxNo)
            Case CONTACTNAMEINDEX 'Web Site
                edcContactName.Visible = False  'Set visibility
                slStr = edcContactName.Text
                gSetShow pbcNmAd(2), slStr, tmCtrls(ilBoxNo)
            Case CONTACTPHONEINDEX 'Web Site
                edcContactPhone.Visible = False  'Set visibility
                slStr = edcContactPhone.Text
                gSetShow pbcNmAd(2), slStr, tmCtrls(ilBoxNo)
            Case CONTACTFAXINDEX 'Web Site
                edcContactFax.Visible = False  'Set visibility
                slStr = edcContactFax.Text
                gSetShow pbcNmAd(2), slStr, tmCtrls(ilBoxNo)
        End Select
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   FacTerminate
'   Where:
'
    sgDoneMsg = Trim$(Str$(igArfCallSource)) & "\" & sgArfName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload NmAddr
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if field defined           *                 *
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
    Dim slMess As String    'Message string
    If smArfCallType = "A" Then
        If (ilCtrlNo = AIDINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            slMess = "Service Name must be specified"
            If gFieldDefinedCtrl(edcID, "", slMess, tmCtrls(AIDINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = AIDINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = AMEDIATYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcMediaType, "", "Media Type must be specified", tmCtrls(AMEDIATYPEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = ANAMEADDRINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (ilCtrlNo = ANAMEADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcNmAd(0), "", "Name and Address must be specified", tmCtrls(ANAMEADDRINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = ANAMEADDRINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    Else
        If (ilCtrlNo = IDINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If smArfCallType = "L" Then 'Lock Box
                slMess = "Lock Box City must be specified"
            ElseIf smArfCallType = "C" Then
                slMess = "Content Provider Name must be specified"
            ElseIf smArfCallType = "P" Then
                slMess = "Producer Name must be specified"
            ElseIf smArfCallType = "N" Then
                slMess = "Network/Rep Name must be specified"
            ElseIf smArfCallType = "K" Then
                slMess = "Feed Producer Name must be specified"
            End If
            If gFieldDefinedCtrl(edcID, "", slMess, tmCtrls(IDINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = IDINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
        If (smArfCallType = "L") Then 'Lock Box
            If (ilCtrlNo = NAMEADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
                If gFieldDefinedCtrl(edcNmAd(0), "", "Name and Address must be specified", tmCtrls(NAMEADDRINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = NAMEADDRINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        Else
            If (ilCtrlNo = NAMEADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
                If gFieldDefinedCtrl(edcNmAd(0), "", "Abbreviation must be specified", tmCtrls(NAMEADDRINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = NAMEADDRINDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
    End If
    mTestFields = YES
End Function

Private Sub pbcISCI_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcISCI_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
'        If imExportISCITo <> 0 Then
'            tmCtrls(imBoxNo).iChg = True
'        End If
'        imExportISCITo = 0
'        pbcISCI_Paint
'    ElseIf KeyAscii = Asc("F") Or (KeyAscii = Asc("f")) Then
'        If imExportISCITo <> 1 Then
'            tmCtrls(imBoxNo).iChg = True
'        End If
'        imExportISCITo = 1
'        pbcISCI_Paint
'    ElseIf KeyAscii = Asc("M") Or (KeyAscii = Asc("m")) Then
'        If imExportISCITo <> 1 Then
'            tmCtrls(imBoxNo).iChg = True
'        End If
'        imExportISCITo = 2
'        pbcISCI_Paint
'    End If
'    If KeyAscii = Asc(" ") Then
'        If imExportISCITo = 0 Then
'            tmCtrls(imBoxNo).iChg = True
'            imExportISCITo = 1
'            pbcISCI_Paint
'        ElseIf imExportISCITo = 1 Then
'            tmCtrls(imBoxNo).iChg = True
'            imExportISCITo = 2
'            pbcISCI_Paint
'        ElseIf imExportISCITo = 2 Then
'            tmCtrls(imBoxNo).iChg = True
'            imExportISCITo = 0
'            pbcISCI_Paint
'        End If
'    End If
    mSetCommands
End Sub

Private Sub pbcISCI_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If imExportISCITo = 0 Then
'        tmCtrls(imBoxNo).iChg = True
'        imExportISCITo = 1
'    ElseIf imExportISCITo = 1 Then
'        tmCtrls(imBoxNo).iChg = True
'        imExportISCITo = 2
'    ElseIf imExportISCITo = 2 Then
'        tmCtrls(imBoxNo).iChg = True
'        imExportISCITo = 0
'    End If
    pbcISCI_Paint
    mSetCommands
End Sub

Private Sub pbcISCI_Paint()
    pbcISCI.Cls
    pbcISCI.CurrentX = fgBoxInsetX
    pbcISCI.CurrentY = 0 'fgBoxInsetY
    If imExportISCITo = 0 Then
        pbcISCI.Print "E-Mail"
    ElseIf imExportISCITo = 1 Then
        pbcISCI.Print "FTP"
    ElseIf imExportISCITo = 2 Then
        pbcISCI.Print "Manual"
    Else
        pbcISCI.Print "   "
    End If
End Sub

Private Sub pbcNmAd_MouseUp(iIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim flAdj As Single
    If (smArfCallType = "A") Then
        If imBoxNo = AIDINDEX Then
            If Not mOKID() Then
                Exit Sub
            End If
        End If
        For ilBox = LBound(tmCtrls) To imMaxCtrls Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
                If (ilBox = ANAMEADDRINDEX + 1) Or (ilBox = ANAMEADDRINDEX + 2) Or (ilBox = ANAMEADDRINDEX + 3) Then
                    flAdj = fgBoxInsetY
                Else
                    flAdj = 0
                End If
                If (Y >= tmCtrls(ilBox).fBoxY + flAdj) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH + flAdj) Then
                    'If (smArfCallType = "A") And (ilBox = AIDINDEX) Then
                    '    Beep
                    '    Exit Sub
                    'End If
                    mSetShow imBoxNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Else
        If imBoxNo = IDINDEX Then
            If Not mOKID() Then
                Exit Sub
            End If
        End If
        For ilBox = LBound(tmCtrls) To imMaxCtrls Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
                If (ilBox = NAMEADDRINDEX + 1) Or (ilBox = NAMEADDRINDEX + 2) Or (ilBox = NAMEADDRINDEX + 3) Then
                    flAdj = fgBoxInsetY
                Else
                    flAdj = 0
                End If
                If (Y >= tmCtrls(ilBox).fBoxY + flAdj) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH + flAdj) Then
                    mSetShow imBoxNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    End If
End Sub
Private Sub pbcNmAd_Paint(ilIndex As Integer)
    Dim ilBox As Integer
    For ilBox = LBound(tmCtrls) To imMaxCtrls Step 1
        pbcNmAd(ilIndex).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcNmAd(ilIndex).CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcNmAd(ilIndex).Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub pbcSepInvByVeh_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If imSepInvByVeh <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSepInvByVeh = 0
        pbcSepInvByVeh_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If imSepInvByVeh <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imSepInvByVeh = 1
        pbcSepInvByVeh_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imSepInvByVeh = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSepInvByVeh = 1
            pbcSepInvByVeh_Paint
        ElseIf imSepInvByVeh = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSepInvByVeh = 0
            pbcSepInvByVeh_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcSepInvByVeh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imSepInvByVeh = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imSepInvByVeh = 1
    ElseIf imSepInvByVeh = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imSepInvByVeh = 0
    End If
    pbcSepInvByVeh_Paint
    mSetCommands
End Sub

Private Sub pbcSepInvByVeh_Paint()
    pbcSepInvByVeh.Cls
    pbcSepInvByVeh.CurrentX = fgBoxInsetX
    pbcSepInvByVeh.CurrentY = 0 'fgBoxInsetY
    If imSepInvByVeh = 0 Then
        pbcSepInvByVeh.Print "Yes"
    ElseIf imSepInvByVeh = 1 Then
        pbcSepInvByVeh.Print "No"
    Else
        pbcSepInvByVeh.Print "   "
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If (smArfCallType = "A") Then
        If imBoxNo = AIDINDEX Then
            If Not mOKID() Then
                Exit Sub
            End If
        End If
        If (imBoxNo >= LBound(tmCtrls)) And (imBoxNo <= imMaxCtrls) Then
            If (imBoxNo <> AIDINDEX) Or (Not cbcSelect.Enabled) Then
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
                    mSetChg 1
                    ilBox = 2
                Else
                    mSetChg 1
                    ilBox = 2
                End If
            Case 1 'Name (last control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
            Case 2
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 2
            Case Else
                ilBox = imBoxNo - 1
        End Select
    Else
        If imBoxNo = IDINDEX Then
            If Not mOKID() Then
                Exit Sub
            End If
        End If
        If (imBoxNo >= LBound(tmCtrls)) And (imBoxNo <= imMaxCtrls) Then
            If (imBoxNo <> IDINDEX) Or (Not cbcSelect.Enabled) Then
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
            Case 1 'Name (last control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
            Case 2
                ilBox = imBoxNo - 1
            Case Else
                ilBox = imBoxNo - 1
        End Select
    End If
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If (smArfCallType = "A") Then
        If imBoxNo = AIDINDEX Then
            If Not mOKID() Then
                Exit Sub
            End If
        End If
        If (imBoxNo >= LBound(tmCtrls)) And (imBoxNo <= imMaxCtrls) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
        Select Case imBoxNo
            Case -1
                If (tgSpf.sBCombine <> "N") Then
                    ilBox = ASEPINVBYVEHINDEX
                Else
                    ilBox = ANAMEADDRINDEX + 3
                End If
            Case AIDINDEX  'ID
                ilBox = AMEDIATYPEINDEX
            Case AMEDIATYPEINDEX  'ID
                ilBox = ANAMEADDRINDEX
            Case ANAMEADDRINDEX + 1 'Name & Address (last control)
                If (edcNmAd(1).Text = "") Then
                    If (tgSpf.sBCombine <> "N") Then
                        ilBox = ASEPINVBYVEHINDEX
                    Else
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If (cmcUpdate.Enabled) And (igArfCallSource = CALLNONE) Then
                            cmcUpdate.SetFocus
                        Else
                            cmcDone.SetFocus
                        End If
                        Exit Sub
                    End If
                Else
                    ilBox = imBoxNo + 1
                End If
            Case ANAMEADDRINDEX + 2 'Name & Address (last control)
                If (edcNmAd(2).Text = "") Then
                    If (tgSpf.sBCombine <> "N") Then
                        ilBox = ASEPINVBYVEHINDEX
                    Else
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If (cmcUpdate.Enabled) And (igArfCallSource = CALLNONE) Then
                            cmcUpdate.SetFocus
                        Else
                            cmcDone.SetFocus
                        End If
                        Exit Sub
                    End If
                Else
                    ilBox = imBoxNo + 1
                End If
            Case imMaxCtrls 'Name & Address (last control)
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igArfCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Case Else
                ilBox = imBoxNo + 1
        End Select
    Else
        If imBoxNo = IDINDEX Then
            If Not mOKID() Then
                Exit Sub
            End If
        End If
        If (imBoxNo >= LBound(tmCtrls)) And (imBoxNo <= imMaxCtrls) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
        Select Case imBoxNo
            Case -1
                If (smArfCallType = "L") Then 'Lock Box
                    ilBox = NAMEADDRINDEX + 3
                Else
                    ilBox = CONTACTFAXINDEX
                End If
            Case IDINDEX  'ID
                ilBox = NAMEADDRINDEX
            Case NAMEADDRINDEX + 1 'Name & Address (last control)
                If (edcNmAd(1).Text = "") And ((smArfCallType = "L")) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    If (cmcUpdate.Enabled) And (igArfCallSource = CALLNONE) Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                Else
                    ilBox = imBoxNo + 1
                End If
            Case NAMEADDRINDEX + 2 'Name & Address (last control)
                If (edcNmAd(2).Text = "") And ((smArfCallType = "L")) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    If (cmcUpdate.Enabled) And (igArfCallSource = CALLNONE) Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                Else
                    ilBox = imBoxNo + 1
                End If
            Case imMaxCtrls 'Name & Address (last control)
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igArfCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            Case Else
                ilBox = imBoxNo + 1
        End Select
    End If
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcNmAd_Paint()
    'plcNmAd.CurrentX = 0
    'plcNmAd.CurrentY = 0
    'plcNmAd.Print "Panel3D2"
End Sub
