VERSION 5.00
Begin VB.Form RptSets 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   345
   ClientTop       =   1680
   ClientWidth     =   9210
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   9210
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   3645
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   105
      Width           =   5490
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   9120
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   0
      Width           =   105
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
      Left            =   5940
      TabIndex        =   20
      Top             =   5640
      Width           =   1050
   End
   Begin VB.CommandButton cmcSave 
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
      Left            =   4725
      TabIndex        =   19
      Top             =   5640
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
      Left            =   3510
      TabIndex        =   18
      Top             =   5640
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
      Left            =   2280
      TabIndex        =   17
      Top             =   5640
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   -30
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   1095
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox plcRptSets 
      ForeColor       =   &H00000000&
      Height          =   5085
      Left            =   90
      ScaleHeight     =   5025
      ScaleWidth      =   8985
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   495
      Width           =   9045
      Begin VB.PictureBox plcRptState 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   165
         ScaleHeight     =   225
         ScaleWidth      =   2730
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   450
         Width           =   2730
         Begin VB.OptionButton rbcRptState 
            Caption         =   "Dormant"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   1110
         End
         Begin VB.OptionButton rbcRptState 
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   495
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.ListBox lbcSelRpt 
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
         Height          =   1920
         Left            =   5190
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1170
         Width           =   3690
      End
      Begin VB.ListBox lbcUnselRpt 
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
         Height          =   1920
         Left            =   150
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   3690
      End
      Begin VB.TextBox edcRptDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   150
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3375
         Width           =   8730
      End
      Begin VB.TextBox edcDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   870
         Left            =   4650
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   75
         Width           =   4230
      End
      Begin VB.PictureBox pbcRptSample 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1050
         Index           =   0
         Left            =   150
         ScaleHeight     =   1020
         ScaleWidth      =   8445
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3930
         Width           =   8475
         Begin VB.PictureBox pbcRptSample 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   1
            Left            =   15
            ScaleHeight     =   165
            ScaleWidth      =   6060
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   -15
            Width           =   6060
         End
      End
      Begin VB.VScrollBar vbcRptSample 
         Height          =   1050
         LargeChange     =   1050
         Left            =   8625
         SmallChange     =   525
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3930
         Width           =   255
      End
      Begin VB.PictureBox pbcUpMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4155
         Picture         =   "Rptsets.frx":0000
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2760
         Width           =   180
      End
      Begin VB.PictureBox pbcDnMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4725
         Picture         =   "Rptsets.frx":00DA
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1710
         Width           =   180
      End
      Begin VB.TextBox edcName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   690
         MaxLength       =   60
         TabIndex        =   5
         Top             =   75
         Width           =   2775
      End
      Begin VB.CommandButton cmcMoveToSel 
         Appearance      =   0  'Flat
         Caption         =   "M&ove   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4035
         TabIndex        =   9
         Top             =   1650
         Width           =   945
      End
      Begin VB.CommandButton cmcMoveToUnsel 
         Appearance      =   0  'Flat
         Caption         =   "    Mo&ve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4035
         TabIndex        =   11
         Top             =   2700
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Selected Reports"
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
         Height          =   210
         Left            =   5190
         TabIndex        =   24
         Top             =   960
         Width           =   3690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Unselected Reports"
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
         Height          =   210
         Left            =   165
         TabIndex        =   23
         Top             =   945
         Width           =   3690
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   195
         X2              =   8820
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   195
         X2              =   8820
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label lacRptName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
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
         Height          =   210
         Left            =   165
         TabIndex        =   4
         Top             =   135
         Width           =   525
      End
      Begin VB.Label lacDescription 
         Appearance      =   0  'Flat
         Caption         =   "Description"
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
         Height          =   210
         Left            =   3600
         TabIndex        =   6
         Top             =   120
         Width           =   1050
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   135
      Top             =   5610
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptsets.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSets.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the RptSets message screen code
Option Explicit
Option Compare Text
'Report Set
Dim hmSnf As Integer 'Report Name file handle
Dim tmSnf As SNF        'SNF record image
Dim tmSnfSrchKey As INTKEY0    'SNF key record image
Dim imSnfRecLen As Integer        'SNF record length
'Report Set
Dim hmSrf As Integer 'Report Name file handle
Dim tmSrf As SRF        'SRF record image
Dim tmSrfSrchKey As LONGKEY0
Dim tmSrfSrchKey1 As INTKEY0    'SRF key record image
Dim imSrfRecLen As Integer        'SRF record length
'Report Name
Dim hmRnf As Integer 'Report Name file handle
Dim tmRnf As RNF        'RNF record image
Dim imRnfRecLen As Integer        'RNF record length
'Report Name
Dim hmRtf As Integer 'Report Name file handle
Dim tmRtf As RTF        'RTF record image
Dim imRtfRecLen As Integer        'RNF record length
'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imUnselIndex As Integer
Dim imSelIndex As Integer
Dim imSnfRptChg As Integer
Dim imSelectedIndex As Integer
Dim imFirstTimeSelect As Integer
Dim imInNew As Integer
Dim imFirstFocus As Integer
Dim tmSelSrf() As SRF
Dim tmUnselSrf() As SRF
Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        'mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    mInitRptList
    imFirstTimeSelect = True
    mSetCommands
    Screen.MousePointer = vbDefault
    imChgMode = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        slSvText = sgRptSetName
    Else
        slSvText = cbcSelect.Text
    End If
    If cbcSelect.ListCount <= 1 Then
        If imSelectedIndex = 0 Then
            cbcSelect.ListIndex = 0
            cbcSelect_Change
        Else
            cbcSelect.ListIndex = 0
        End If
        'If pbcSTab.Enabled Then
        '    pbcSTab.SetFocus
        'Else
        '    cmcCancel.SetFocus
        'End If
        'Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    DoEvents
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imFirstFocus Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
            End If
        Else
            cbcSelect.ListIndex = 0
            'mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
    imFirstFocus = False
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

    sgRptSetName = ""
    igRptSetReturn = 1
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    sgRptSetName = Trim$(edcName.Text)
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        edcName.SetFocus
        Exit Sub
    End If
    igRptSetReturn = 0
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcMoveToSel_Click()
    Dim slName As String
    Dim ilIndex As Integer
    For ilIndex = 0 To lbcUnselRpt.ListCount - 1 Step 1
        If lbcUnselRpt.Selected(ilIndex) Then
            slName = lbcUnselRpt.List(ilIndex)
            If (lbcSelRpt.List(0) = "") And (lbcSelRpt.ListCount = 1) Then
                lbcSelRpt.List(0) = slName
            Else
                lbcSelRpt.AddItem slName
            End If
        End If
    Next ilIndex
    For ilIndex = lbcUnselRpt.ListCount - 1 To 0 Step -1
        If lbcUnselRpt.Selected(ilIndex) Then
            lbcUnselRpt.RemoveItem ilIndex
        End If
    Next ilIndex
    edcRptDescription.Text = ""
    vbcRptSample.Value = vbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    vbcRptSample.Max = vbcRptSample.Min
    pbcRptSample(1).Picture = LoadPicture()
    imSnfRptChg = True
    mSetCommands
End Sub
Private Sub cmcMoveToUnsel_Click()
    Dim slName As String
    Dim ilIndex As Integer
    For ilIndex = 0 To lbcSelRpt.ListCount - 1 Step 1
        If lbcSelRpt.Selected(ilIndex) Then
            slName = lbcSelRpt.List(ilIndex)
            If (lbcUnselRpt.List(0) = "") And (lbcUnselRpt.ListCount = 1) Then
                lbcUnselRpt.List(0) = slName
            Else
                lbcUnselRpt.AddItem slName
            End If
        End If
    Next ilIndex
    For ilIndex = lbcSelRpt.ListCount - 1 To 0 Step -1
        If lbcSelRpt.Selected(ilIndex) Then
            lbcSelRpt.RemoveItem ilIndex
        End If
    Next ilIndex
    edcRptDescription.Text = ""
    vbcRptSample.Value = vbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    vbcRptSample.Max = vbcRptSample.Min
    pbcRptSample(1).Picture = LoadPicture()
    imSnfRptChg = True
    mSetCommands
End Sub
Private Sub cmcSave_Click()
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        edcName.SetFocus
        Exit Sub
    End If
    mPopulate
    cbcSelect.ListIndex = 0
End Sub
Private Sub cmcUndo_Click()
    Dim ilIndex As Integer
    Screen.MousePointer = vbHourglass  'Wait
    imSnfRptChg = False
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        If Not mReadRec(ilIndex) Then
            GoTo cmcUndoErr
        End If
        mInitRptList
        mSetCommands
        edcName.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDescription_Change()
    imSnfRptChg = True
    mSetCommands
End Sub
Private Sub edcDescription_GotFocus()
    imFirstTimeSelect = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcName_Change()
    imSnfRptChg = True
    mSetCommands
End Sub
Private Sub edcName_GotFocus()
    imFirstTimeSelect = False
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
Private Sub edcRptDescription_GotFocus()
    imFirstTimeSelect = False
    pbcClickFocus.SetFocus
End Sub
Private Sub edcRptDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Form_Activate()
    If imInNew Then
        Exit Sub
    End If
    Me.KeyPreview = True
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    RptSets.Refresh
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
        If (cbcSelect.Enabled) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmSelSrf
    Erase tmUnselSrf
    ilRet = btrClose(hmSnf)
    btrDestroy hmSnf
    ilRet = btrClose(hmSrf)
    btrDestroy hmSrf
    ilRet = btrClose(hmRnf)
    btrDestroy hmRnf
    ilRet = btrClose(hmRtf)
    btrDestroy hmRtf
    Set RptSets = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelRpt_Click()
    Dim slFromFile As String
    Dim slName As String
    Dim ilRnf As Integer
    Screen.MousePointer = vbHourglass
    vbcRptSample.Value = vbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    If lbcSelRpt.ListIndex = imSelIndex Then
        slName = Trim$(lbcSelRpt.List(lbcSelRpt.ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If (slName = Trim$(tgRnfList(ilRnf).tRnf.sName)) And (tgRnfList(ilRnf).tRnf.sType = "R") Then
                'edcRptDescription.Text = Left$(tgRnfList(ilRnf).tRnf.sDescription, tgRnfList(ilRnf).tRnf.iStrLen)
                edcRptDescription.Text = gStripChr0(tgRnfList(ilRnf).tRnf.sDescription)
                slFromFile = tgRnfList(ilRnf).tRnf.sRptSample
                On Error GoTo lbcSelRptErr:
                If gFileExist(sgRptPath & slFromFile) = 0 Then
                    pbcRptSample(1).Picture = LoadPicture(sgRptPath & slFromFile)
                End If
                vbcRptSample.Max = pbcRptSample(1).Height - pbcRptSample(0).Height
                vbcRptSample.Enabled = (pbcRptSample(0).Height < pbcRptSample(1).Height)
                If vbcRptSample.Enabled Then
                    vbcRptSample.SmallChange = pbcRptSample(0).Height
                    vbcRptSample.LargeChange = pbcRptSample(0).Height
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Next ilRnf
    End If
    edcRptDescription.Text = ""
    pbcRptSample(1).Picture = LoadPicture()
    vbcRptSample.Max = vbcRptSample.Min
    vbcRptSample.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
lbcSelRptErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcSelRpt_GotFocus()
    imFirstTimeSelect = False
End Sub
Private Sub lbcSelRpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imSelIndex = lbcSelRpt.TopIndex + Y \ fgListHtArial825 'fgListHtSerif825
End Sub
Private Sub lbcSelRpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imSelIndex = lbcSelRpt.TopIndex + Y \ fgListHtArial825 'fgListHtSerif825
    If (lbcUnselRpt.ListIndex <> imUnselIndex) Then
        edcRptDescription.Text = ""
        pbcRptSample(1).Picture = LoadPicture()
        vbcRptSample.Max = vbcRptSample.Min
        vbcRptSample.Enabled = False
    End If
End Sub
Private Sub lbcUnselRpt_Click()
    Dim slFromFile As String
    Dim slName As String
    Dim ilRnf As Integer
    Screen.MousePointer = vbHourglass
    vbcRptSample.Value = vbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    If lbcUnselRpt.ListIndex = imUnselIndex Then
        slName = Trim$(lbcUnselRpt.List(lbcUnselRpt.ListIndex))
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If (slName = Trim$(tgRnfList(ilRnf).tRnf.sName)) And (tgRnfList(ilRnf).tRnf.sType = "R") Then
                'edcRptDescription.Text = Left$(tgRnfList(ilRnf).tRnf.sDescription, tgRnfList(ilRnf).tRnf.iStrLen)
                edcRptDescription.Text = gStripChr0(tgRnfList(ilRnf).tRnf.sDescription)
                slFromFile = tgRnfList(ilRnf).tRnf.sRptSample
                On Error GoTo lbcUnselRptErr:
                If gFileExist(sgRptPath & slFromFile) = 0 Then
                    pbcRptSample(1).Picture = LoadPicture(sgRptPath & slFromFile)
                End If
                vbcRptSample.Max = pbcRptSample(1).Height - pbcRptSample(0).Height
                vbcRptSample.Enabled = (pbcRptSample(0).Height < pbcRptSample(1).Height)
                If vbcRptSample.Enabled Then
                    vbcRptSample.SmallChange = pbcRptSample(0).Height
                    vbcRptSample.LargeChange = pbcRptSample(0).Height
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Next ilRnf
    End If
    edcRptDescription.Text = ""
    pbcRptSample(1).Picture = LoadPicture()
    vbcRptSample.Max = vbcRptSample.Min
    vbcRptSample.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
lbcUnselRptErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub
Private Sub lbcUnselRpt_GotFocus()
    imFirstTimeSelect = False
End Sub
Private Sub lbcUnselRpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imUnselIndex = lbcUnselRpt.TopIndex + Y \ fgListHtArial825 'fgListHtSerif825
End Sub
Private Sub lbcUnselRpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imUnselIndex = lbcUnselRpt.TopIndex + Y \ fgListHtArial825 'fgListHtSerif825
    If (lbcUnselRpt.ListIndex <> imUnselIndex) Then
        edcRptDescription.Text = ""
        pbcRptSample(1).Picture = LoadPicture()
        vbcRptSample.Max = vbcRptSample.Min
        vbcRptSample.Enabled = False
    End If
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

    edcName.Text = ""
    edcDescription.Text = ""
    edcRptDescription.Text = ""
    mObtainSrf 0, hmSrf, tmSelSrf(), tmUnselSrf()
    edcRptDescription.Text = ""
    vbcRptSample.Value = vbcRptSample.Min
    pbcRptSample(1).Picture = LoadPicture()
    imSnfRptChg = False
    mSetCommands
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
    ReDim tmSelSrf(0 To 0) As SRF
    ReDim tmUnselSrf(0 To 0) As SRF
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    imFirstFocus = True
    imChgMode = False
    imBSMode = False
    imSnfRptChg = False
    imFirstTimeSelect = True
    imSelectedIndex = -1
    imInNew = False
    RptSets.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone RptSets

    hmSnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSnf, "", sgDBPath & "SNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SNF.Btr)", RptSets
    On Error GoTo 0
    imSnfRecLen = Len(tmSnf)

    hmSrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSrf, "", sgDBPath & "SRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: SRF.Btr)", RptSets
    On Error GoTo 0
    imSrfRecLen = Len(tmSrf)
    hmRnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRnf, "", sgDBPath & "RNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RNF.Btr)", RptSets
    On Error GoTo 0
    imRnfRecLen = Len(tmRnf)
    hmRtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRtf, "", sgDBPath & "RTF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: RTF.Btr)", RptSets
    On Error GoTo 0
    imRtfRecLen = Len(tmRtf)
    gObtainRNF hmRnf
    mPopulate
    'gCenterModalForm RptSets
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitRptList                    *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Unsel and Sel list    *
'*                      boxes                          *
'*                                                     *
'*******************************************************
Private Sub mInitRptList()
    Dim ilSrf As Integer
    Dim ilRnf As Integer
    lbcSelRpt.Clear
    lbcUnselRpt.Clear
    For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If tmSelSrf(ilSrf).iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                lbcSelRpt.AddItem Trim$(tgRnfList(ilRnf).tRnf.sName)
                Exit For
            End If
        Next ilRnf
    Next ilSrf
    'This code required to get list box to display
    If lbcSelRpt.ListCount <= 0 Then
        lbcSelRpt.AddItem ""
    End If
    For ilSrf = 0 To UBound(tmUnselSrf) - 1 Step 1
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If tmUnselSrf(ilSrf).iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                lbcUnselRpt.AddItem Trim$(tgRnfList(ilRnf).tRnf.sName)
                Exit For
            End If
        Next ilRnf
    Next ilSrf
    If lbcUnselRpt.ListCount <= 0 Then
        lbcUnselRpt.AddItem ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSetSrf                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain the selected reports of *
'*                      a set for a user               *
'*                                                     *
'*******************************************************
Private Sub mObtainSrf(ilSnfCode As Integer, hlSrf As Integer, tlSelSrf() As SRF, tlUnselSrf() As SRF)
'
'   gObtainSrf ilSnfCode, hlSrf, tmSelSrf(), tmUnselSrf()
'   Where:
'       gObtainRNF must be called prior to this call to load tgRNFLIST
'
    Dim ilSortCode As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim ilRnf As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ilSortCode = 0
    ReDim tlSelSrf(0 To 0) As SRF   'VB list box clear (list box used to retain code number so record can be found)
    ReDim tlUnselSrf(0 To 0) As SRF   'VB list box clear (list box used to retain code number so record can be found)
    imSrfRecLen = Len(tlSelSrf(0)) 'btrRecordLength(hlSrf)  'Get and save record length
    ilExtLen = Len(tlSelSrf(0))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSrf) 'Obtain number of records
    btrExtClear hlSrf   'Clear any previous extend operation
    tmSrfSrchKey1.iCode = ilSnfCode
    ilRet = btrGetGreaterOrEqual(hlSrf, tmSrf, imSrfRecLen, tmSrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    tlIntTypeBuff.iType = ilSnfCode
    ilOffSet = 4    'gFieldOffset("Prf", "PrfAdfCode")
    ilRet = btrExtAddLogicConst(hlSrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    Call btrExtSetBounds(hlSrf, llNoRec, -1, "UC", "SRF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlSrf, ilOffSet, ilExtLen)  'Extract First Name field
    If ilRet = BTRV_ERR_NONE Then
        'ilRet = btrExtGetNextExt(hlSrf)    'Extract record
        ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                ilExtLen = Len(tlSelSrf(ilSortCode))  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    If ilSortCode >= UBound(tlSelSrf) Then
                        ReDim Preserve tlSelSrf(0 To UBound(tlSelSrf) + 100) As SRF
                    End If
                    ilSortCode = ilSortCode + 1
                    ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
                    Loop
                Loop
                ReDim Preserve tlSelSrf(0 To ilSortCode) As SRF
            End If
        End If
    End If
    For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
        If (tgRnfList(ilRnf).tRnf.sType <> "C") And (tgRnfList(ilRnf).tRnf.sState <> "D") Then
            ilFound = False
            For ilLoop = 0 To UBound(tlSelSrf) - 1 Step 1
                If tgRnfList(ilRnf).tRnf.iCode = tlSelSrf(ilLoop).iRnfCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                tlUnselSrf(UBound(tlUnselSrf)).iSnfCode = ilSnfCode
                tlUnselSrf(UBound(tlUnselSrf)).iRnfCode = tgRnfList(ilRnf).tRnf.iCode
                tlUnselSrf(UBound(tlUnselSrf)).sViewMoney = "Y"
                ReDim Preserve tlUnselSrf(0 To UBound(tlUnselSrf) + 1) As SRF
            End If
        End If
    Next ilRnf
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
                    MsgBox "Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmSnf.sName) 'Reset text
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
'*      Procedure Name:mSnfPop                         *
'*                                                     *
'*             Created:5/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Report Set Name combo *
'*                      control                        *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
    Dim ilLoop As Integer
    gObtainSNF hmSnf, True
    cbcSelect.Clear
    For ilLoop = 0 To UBound(tgSnfCode) - 1 Step 1
        cbcSelect.AddItem Trim$(tgSnfCode(ilLoop).tSnf.sName)
    Next ilLoop
    cbcSelect.AddItem "[New]", 0
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
Private Function mReadRec(ilSelectIndex As Integer) As Integer
'
'   iRet = mReadRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    'slNameCode = tgSnfCode(ilSelectIndex - 1).sKey   'lbcTitleCode.List(ilSelectIndex - 1)
    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
    'On Error GoTo mReadRecErr
    'gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", RptSets
    'On Error GoTo 0
    'slCode = Trim$(slCode)
    tmSnfSrchKey.iCode = tgSnfCode(ilSelectIndex - 1).tSnf.iCode
    imSnfRecLen = Len(tmSnf)  'Get and save CmF record length (the read will change the length)
    ilRet = btrGetEqual(hmSnf, tmSnf, imSnfRecLen, tmSnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", RptSets
    On Error GoTo 0
    mObtainSrf tmSnf.iCode, hmSrf, tmSelSrf(), tmUnselSrf()
    edcName.Text = Trim$(tmSnf.sName)
    'edcDescription.Text = Trim$(Left$(tmSnf.sDescription, tmSnf.iStrLen))
    edcDescription.Text = gStripChr0(tmSnf.sDescription)
    If tmSnf.sState = "D" Then
        rbcRptState(1).Value = True
    Else
        rbcRptState(0).Value = True
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
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec()
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilSrf As Integer
    Dim ilRnf As Integer
    If Trim$(edcName.Text) = "" Then
        mSaveRec = False
        Exit Function
    End If
    'If Trim$(edcDescription.Text) = "" Then
    '    mSaveRec = False
    '    Exit Function
    'End If
    If (lbcSelRpt.ListCount <= 0) Or (lbcSelRpt.List(0) = "") Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        If imSelectedIndex <> 0 Then
            'Reread record in so latest is obtained
            tmSnfSrchKey.iCode = tmSnf.iCode
            imSnfRecLen = Len(tmSnf)  'Get and save CmF record length (the read will change the length)
            ilRet = btrGetEqual(hmSnf, tmSnf, imSnfRecLen, tmSnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        End If
        tmSnf.sName = Trim$(edcName.Text)
        'tmSnf.iStrLen = Len(edcDescription.Text)
        tmSnf.sDescription = Trim$(edcDescription.Text) & Chr$(0) '& Chr$(0)
        If rbcRptState(0).Value = True Then
            tmSnf.sState = "A"
        Else
            tmSnf.sState = "D"
        End If
        imSnfRecLen = Len(tmSnf) '- Len(tmSnf.sDescription) + Len(Trim$(tmSnf.sDescription)) ' + 2
        If imSelectedIndex = 0 Then
            tmSnf.iCode = 0 'Autoincrement
            ilRet = btrInsert(hmSnf, tmSnf, imSnfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Report Name Set)"
        Else
            ilRet = btrUpdate(hmSnf, tmSnf, imSnfRecLen)
            slMsg = "mSaveRec (btrUpdate: Report Name Set)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, RptSets
    On Error GoTo 0
    'Insert and delete records belonging to the set
    For ilLoop = 0 To lbcSelRpt.ListCount - 1 Step 1
    If lbcSelRpt.List(ilLoop) = "Budgets" Then
    ilRet = ilRet
    End If
        ilFound = False
        tmSrf.iRnfCode = -1
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If (Trim$(lbcSelRpt.List(ilLoop)) = Trim$(tgRnfList(ilRnf).tRnf.sName)) And (tgRnfList(ilRnf).tRnf.sType = "R") Then
                tmSrf.iRnfCode = tgRnfList(ilRnf).tRnf.iCode
                For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
                    If tmSelSrf(ilSrf).iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSrf
                Exit For
            End If
        Next ilRnf
        If (Not ilFound) And (tmSrf.iRnfCode > 0) Then
            tmSrf.lCode = 0
            tmSrf.iSnfCode = tmSnf.iCode
            tmSrf.sViewMoney = "Y"
            ilRet = btrInsert(hmSrf, tmSrf, imSrfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Report Set)"
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, RptSets
            On Error GoTo 0
        End If
    Next ilLoop
    For ilLoop = 0 To lbcUnselRpt.ListCount - 1 Step 1
        For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
            If (Trim$(lbcUnselRpt.List(ilLoop)) = Trim$(tgRnfList(ilRnf).tRnf.sName)) And (tgRnfList(ilRnf).tRnf.sType = "R") Then
                For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
                    If tmSelSrf(ilSrf).iRnfCode = tgRnfList(ilRnf).tRnf.iCode Then
                        'Delete Record
                        slMsg = "mSaveRec (btrDelete: Report Set)"
                        Do
                            tmSrfSrchKey.lCode = tmSelSrf(ilSrf).lCode
                            imSrfRecLen = Len(tmSrf)  'Get and save CmF record length (the read will change the length)
                            ilRet = btrGetEqual(hmSrf, tmSrf, imSrfRecLen, tmSrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            ilRet = btrDelete(hmSrf)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, RptSets
                        On Error GoTo 0
                        Exit For
                    End If
                Next ilSrf
                Exit For
            End If
        Next ilRnf
    Next ilLoop
    imFirstTimeSelect = True
    mSaveRec = True
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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
    If (edcName.Text <> "") And ((lbcSelRpt.ListCount > 0) And (lbcSelRpt.List(0) <> "")) Then
        If imSnfRptChg Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
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
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Buttons                     *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    Dim slStr As String
    If imSnfRptChg Then
        cmcUndo.Enabled = True
        slStr = Trim$(edcName.Text)
        If Len(slStr) = 0 Then
            cmcSave.Enabled = False
            Exit Sub
        End If
        If (lbcSelRpt.ListCount <= 0) Or (lbcSelRpt.List(0) = "") Then
            cmcSave.Enabled = False
            Exit Sub
        End If
        cmcSave.Enabled = True
    Else
        cmcUndo.Enabled = False
        cmcSave.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mStartNew                       *
'*                                                     *
'*             Created:7/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up a New rate card and     *
'*                      initiate RCTerms               *
'*                                                     *
'*******************************************************
Private Function mStartNew() As Integer
    Dim ilRet As Integer
    imInNew = True
    SetModel.Show vbModal
    DoEvents
    If igSetReturn = 0 Then    'Cancelled
        mStartNew = False
        imInNew = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass    '
    tmSnfSrchKey.iCode = igSnfModel
    imSnfRecLen = Len(tmSnf)  'Get and save CmF record length (the read will change the length)
    ilRet = btrGetEqual(hmSnf, tmSnf, imSnfRecLen, tmSnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mStartNew = False
        imInNew = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    mObtainSrf tmSnf.iCode, hmSrf, tmSelSrf(), tmUnselSrf()
    tmSnf.iCode = 0
    edcName.Text = ""   'Trim$(tmSnf.sName)
    edcDescription.Text = ""    'Trim$(Left$(tmSnf.sDescription, tmSnf.iStrLen))
    'If tmSnf.sState = "D" Then
    '    rbcRptState(1).Value = True
    'Else
        rbcRptState(0).Value = True
    'End If
    'Build program images from newest
    mInitRptList
    mObtainSrf 0, hmSrf, tmSelSrf(), tmUnselSrf()
    'Move selected to Unselect
    mStartNew = True
    mSetCommands
    Screen.MousePointer = vbDefault
    imInNew = False
    Exit Function

    On Error GoTo 0
    mStartNew = False
    Exit Function
End Function
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload RptSets
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_GotFocus()
    imFirstTimeSelect = False
End Sub
Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    If imInNew Then
        Exit Sub
    End If
    If (imSelectedIndex = 0) And (imFirstTimeSelect) Then
        imFirstTimeSelect = False
        ilRet = mStartNew()
        Screen.MousePointer = vbDefault
        'If Not ilRet Then
        '    If cbcSelect.ListCount <= 1 Then
        '        imTerminate = True
        '        mTerminate
        '        Exit Sub
        '    End If
        '    cbcSelect.SetFocus
        '    Exit Sub
        'End If
    End If
    mSetCommands
    edcName.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub vbcRptSample_Change()
    pbcRptSample(1).Top = -vbcRptSample.Value
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Report Sets"
End Sub
Private Sub plcRptState_Paint()
    plcRptState.CurrentX = 0
    plcRptState.CurrentY = 0
    plcRptState.Print "State"
End Sub
