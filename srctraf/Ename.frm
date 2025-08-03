VERSION 5.00
Begin VB.Form EName 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   1245
   ClientTop       =   2865
   ClientWidth     =   6660
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
   ScaleHeight     =   3630
   ScaleWidth      =   6660
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6045
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2940
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6045
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2595
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6105
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3255
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   105
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3420
      Width           =   105
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   660
      Top             =   2625
   End
   Begin VB.ListBox lbcGenre 
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
      Left            =   3420
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   750
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2145
      Visible         =   0   'False
      Width           =   1020
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
      Left            =   1770
      Picture         =   "Ename.frx":0000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2145
      Visible         =   0   'False
      Width           =   195
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
      Left            =   3435
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   1290
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
      Left            =   2325
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
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
      Left            =   510
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2325
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox edcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   5610
      MaxLength       =   3
      TabIndex        =   19
      Top             =   2355
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox edcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   5100
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2340
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox edcSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Left            =   4530
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2355
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox edcComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   540
      HelpContextID   =   8
      Left            =   495
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   510
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1260
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox plcSelect 
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   135
      ScaleHeight     =   705
      ScaleWidth      =   6300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   6360
      Begin VB.PictureBox pbcEType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1830
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   4
         Top             =   375
         Width           =   15
      End
      Begin VB.ComboBox cbcEType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   30
         Width           =   3045
      End
      Begin VB.ComboBox cbcVeh 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   3180
      End
      Begin VB.ComboBox cbcSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   30
         TabIndex        =   5
         Top             =   375
         Width           =   6255
      End
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
      Left            =   3930
      TabIndex        =   27
      Top             =   3180
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
      Left            =   6330
      TabIndex        =   26
      Top             =   2055
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
      Left            =   2805
      TabIndex        =   25
      Top             =   3180
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
      Left            =   1680
      TabIndex        =   24
      Top             =   3180
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
      Left            =   3930
      TabIndex        =   23
      Top             =   2805
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
      Left            =   2805
      TabIndex        =   22
      Top             =   2805
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
      Left            =   1680
      TabIndex        =   21
      Top             =   2805
      Width           =   1050
   End
   Begin VB.PictureBox pbcENm 
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
      Height          =   1410
      Left            =   480
      Picture         =   "Ename.frx":00FA
      ScaleHeight     =   1410
      ScaleWidth      =   5670
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1110
      Width           =   5670
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   765
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   6
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
      TabIndex        =   20
      Top             =   1695
      Width           =   45
   End
   Begin VB.PictureBox plcENm 
      ForeColor       =   &H00000000&
      Height          =   1530
      Left            =   420
      ScaleHeight     =   1470
      ScaleWidth      =   5730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1050
      Width           =   5790
   End
   Begin VB.Label plcScreen 
      Caption         =   "Event Names"
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
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3105
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "EName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Ename.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EName.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Event Names input screen code
Option Explicit
Option Compare Text
Dim tmETypeCode() As SORTCODE
Dim smETypeCodeTag As String
Dim tmGenreCode() As SORTCODE
Dim smGenreCodeTag As String
'Event name Field Areas
Dim tmCtrls(0 To 9)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current event name Box
Dim tmEnf As ENF        'Enf record image
Dim tmSrchKey As INTKEY0    'Enf key record image
Dim imRecLen As Integer        'Enf record length
Dim tmCef As CEF        'CEF record image
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim imCefRecLen As Integer        'CEF record length
Dim tmEtf As ETF        'Etf record image
Dim tmEtfSrchKey As INTKEY0    'Etf key record image
Dim imEtfRecLen As Integer        'Etf record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imVehSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imETypeSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmEnf As Integer    'Event name file handle
Dim hmCef As Integer    'Comment file handle
Dim hmEtf As Integer    'Event Type file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imComboBoxIndex As Integer
Dim imTimeFirst As Integer  'First time at field- set default if required
Dim imLenFirst As Integer   'First time at field-set default if required
Dim imProgFirst As Integer  'First time at field- set default if required
Dim imFacListIndex As Integer   'Retain Vehicle index when mPopulate called, so changes can force repopulation
Dim imEvtListIndex As Integer   'Retain Event type index when mPopulate called, so changes can force repopulation
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim smComment As String     'Save original value
Dim smGenre As String       'Save original value
Dim smInitVeh As String
Dim smInitEType As String
Dim smInitEName As String
Dim imFirstFocusVeh As Integer
Dim imFirstFocusEType As Integer
Dim imFirstFocusEName As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imPrgEType As Integer   'True=Prg Event Type; False=Not program event type
Dim imUpdateAllowed As Integer    'User can update records

Const NAMEINDEX = 1         'Name control/field
Const GENREINDEX = 2        'Genre control/field
Const COMMENTINDEX = 3      'Comment control/field
Const TIMEINDEX = 4         'Time format control/field
Const LENINDEX = 5          'Length format control/field
Const PROGINDEX = 6         'Program/Sponsor format control/field
Const SOURCEINDEX = 7       'Source control/field
Const TYPEINDEX = 8         'Primary/Secondary type control/field
Private Sub cbcEType_Change()
    Dim ilRet As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    If imChgMode = False Then
        imChgMode = True
        ilRet = gOptionLookAhead(cbcEType, imBSMode, slStr)
        If ilRet = 1 Then
'            If imFirstFocusEType Then 'Test if coming from sales source- if so, branch to first control
'                If cbcEType.ListCount > 1 Then
'                    cbcEType.ListIndex = 1
'                Else
'                    cbcEType.ListIndex = 0
'                End If
'            Else
                cbcEType.ListIndex = 0
'            End If
        End If
        imETypeSelectedIndex = cbcEType.ListIndex
        pbcENm.Cls
        mClearCtrlFields
        cbcSelect.Clear 'Force population
        sgNameCodeTag = ""
        If imETypeSelectedIndex > 0 Then
            slNameCode = tmETypeCode(imETypeSelectedIndex - 1).sKey 'lbcETypeCode.List(imETypeSelectedIndex - 1)
            ilRet = gParseItem(slNameCode, 3, "\", slCode)
            If Val(slCode) = 1 Then
                imPrgEType = True
            Else
                imPrgEType = False
            End If
        Else
            imPrgEType = False
        End If
        imChgMode = False
    End If
    mSetChg imBoxNo
End Sub
Private Sub cbcEType_Click()
    cbcEType_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcEType_DropDown()
    mETypePop
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcEType_GotFocus()
    Dim sSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocusEType Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusEType = False
        If igENameCallSource <> CALLNONE Then  'If from advertiser or contract- set name and branch to control
            If smInitEType = "" Then
                cbcEType.ListIndex = 0
            Else
                gFindMatch smInitEType, 0, cbcEType
                If gLastFound(cbcEType) >= 0 Then
                    cbcEType.ListIndex = gLastFound(cbcEType)
                    mSetCommands
                    cbcSelect.SetFocus
                    If cbcSelect.Enabled Then
                        cbcSelect.SetFocus
                        Exit Sub
                    End If
                Else
                    cbcEType.Text = smInitEType    'Name
                End If
            End If
'            cbcEType_Change
'            If smInitEType <> "" Then
'                mSetCommands
'                gFndFirst cbcEType,  smInitEType
'                If gLastFound(cbcEType) >= 0 Then
'                    cbcSelect.SetFocus
'                    Exit Sub
'                End If
'            End If
            Exit Sub
        End If
    End If
    sSvText = cbcEType.Text
    mETypePop
    If imTerminate Then
        Exit Sub
    End If
    If (sSvText = "") Or (sSvText = "[New]") Then
        If cbcEType.ListCount > 0 Then
            cbcEType.ListIndex = 1
        Else
            cbcEType.ListIndex = 0
        End If
    Else
        gFindMatch sSvText, 1, cbcEType
        If gLastFound(cbcEType) > 0 Then
            cbcEType.ListIndex = gLastFound(cbcEType)
        Else
            cbcEType.ListIndex = 0
        End If
    End If
    gCtrlGotFocus cbcEType
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcEType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcEType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcEType.SelLength <> 0 Then    'avoid deleting two characters
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
    pbcENm.Cls
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
    pbcENm_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
'    mSetCommands
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
    If imFirstFocusEName Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusEName = False
        If igENameCallSource <> CALLNONE Then  'If from advt or contract- set name and branch to control
            If smInitEName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = smInitEName    'New name
            End If
'            cbcSelect_Change
            If smInitEName <> "" Then
                mSetCommands
                gFindMatch smInitEName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cbcSelect.ListIndex = gLastFound(cbcSelect)
                    cmcDone.SetFocus
                    Exit Sub
                End If
                cbcSelect_Change
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
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
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcVeh_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcVeh.Text <> "" Then
            gManLookAhead cbcVeh, imBSMode, imComboBoxIndex
        End If
        imVehSelectedIndex = cbcVeh.ListIndex
        pbcENm.Cls
        mClearCtrlFields
        cbcSelect.Clear 'Force population
        sgNameCodeTag = ""
        imChgMode = False
    End If
    mSetChg imBoxNo
End Sub
Private Sub cbcVeh_Click()
    cbcVeh_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcVeh_GotFocus()
    Dim sSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocusVeh Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusVeh = False
        If igENameCallSource <> CALLNONE Then  'If from advertiser or contract- set name and branch to control
            If smInitVeh = "" Then
                cbcVeh.ListIndex = 0
            Else
                gFindMatch smInitVeh, 0, cbcVeh
                If gLastFound(cbcVeh) >= 0 Then
                    cbcVeh.ListIndex = gLastFound(cbcVeh)
                    mSetCommands
                    If cbcEType.Enabled Then
                        cbcEType.SetFocus
                        Exit Sub
                    End If
                Else
                    cbcVeh.Text = smInitVeh    'Name
                End If
            End If
'            cbcVeh_Change
'            If smInitVeh <> "" Then
'                mSetCommands
'                gFndFirst cbcVeh,  smInitVeh
'                If gLastFound(cbcVeh) >= 0 Then
'                    cbcEType.SetFocus
'                    Exit Sub
'                End If
'            End If
            Exit Sub
        End If
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    sSvText = cbcVeh.Text
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    If (sSvText = "") Or (sSvText = "[New]") Then
        If cbcVeh.ListCount = 1 Then
            cbcVeh.ListIndex = 0
        Else
            gFindMatch sgUserDefVehicleName, 1, cbcVeh
            If gLastFound(cbcVeh) > 0 Then
                cbcVeh.ListIndex = gLastFound(cbcVeh)
            Else
                cbcVeh.ListIndex = 0
            End If
        End If
        imComboBoxIndex = cbcVeh.ListIndex
        imVehSelectedIndex = cbcVeh.ListIndex
    Else
        gFindMatch sSvText, 1, cbcVeh
        If gLastFound(cbcVeh) > 0 Then
            cbcVeh.ListIndex = gLastFound(cbcVeh)
        Else
            cbcVeh.ListIndex = 0
        End If
        imComboBoxIndex = cbcVeh.ListIndex
        imVehSelectedIndex = cbcVeh.ListIndex
    End If
    imComboBoxIndex = imVehSelectedIndex
    gCtrlGotFocus cbcVeh
    Exit Sub
End Sub
Private Sub cbcVeh_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVeh_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVeh.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCancel_Click()
    If igENameCallSource <> CALLNONE Then
        igENameCallSource = CALLCANCELLED
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
    If igENameCallSource <> CALLNONE Then
        sgENameName = edcName 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgENameName = "[New]"
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
    If igENameCallSource <> CALLNONE Then
        If sgENameName = "[New]" Then
            igENameCallSource = CALLCANCELLED
        Else
            igENameCallSource = CALLDONE
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
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case GENREINDEX
            lbcGenre.Visible = Not lbcGenre.Visible
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
    Dim ilCode As Integer
    Dim slMsg As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        ilCode = tmEnf.iCode
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(EName, ilCode, "Cif.Btr", "CifEnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Inventory references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(EName, ilCode, "Crf.Btr", "CrfEnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Rotation references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gLICodeRefExist(EName, ilCode, "Lef.Btr", "LefEnfCode")   'lefenfCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Library Events references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(EName, ilCode, "Dlf.Btr", "DlfEnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Delivery Link references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(EName, ilCode, "Egf.Btr", "EgfEnfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Engineering Link references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmEnf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        If tmEnf.lCefCode <> 0 Then
            ilRet = btrDelete(hmCef)
            On Error GoTo cmcEraseErr
            gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete:Comment)", EName
            On Error GoTo 0
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
        End If
        slStamp = gFileDateTime(sgDBPath & "Enf.btr")
        ilRet = btrDelete(hmEnf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete: Event Name)", EName
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcNameCode.Tag <> "" Then
        '    If slStamp = lbcNameCode.Tag Then
        '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Enf.btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Enf.btr")
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
    pbcENm.Cls
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
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = EVENTNAMESLIST
    igRptType = 0
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "EName^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "EName^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'EName.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
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
        pbcENm.Cls
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcENm_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcENm.Cls
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
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
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
'    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
'        cbcSelect.Text = slName
'    Else
'        cbcSelect.ListIndex = 0
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmEnf.iCode
    cbcSelect.Clear
    sgNameCodeTag = ""
    mPopulate
    If imSvSelectedIndex <> 0 Then
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
    gCtrlGotFocus ActiveControl
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcComment_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcComment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcComment_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case GENREINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcGenre, imBSMode, slStr)
            If ilRet = 1 Then
                lbcGenre.ListIndex = 1
            End If
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
Private Sub edcDropDown_DblClick()
    If imBoxNo = GENREINDEX Then
        imDoubleClickName = True    'Double click event followed by mouse up
    End If
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case GENREINDEX
'            If lbcGenre.ListCount = 1 Then
'                lbcGenre.ListIndex = 0
'                If imTabDirection = -1 Then  'Right To Left
'                    pbcSTab.SetFocus
'                Else
'                    pbcTab.SetFocus
'                End If
'                Exit Sub
'            End If
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
            Case GENREINDEX
                gProcessArrowKey Shift, KeyCode, lbcGenre, imLbcArrowSetting
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
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case GENREINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
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
Private Sub edcSource_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcSource_GotFocus()
    gCtrlGotFocus edcSource
End Sub
Private Sub edcSource_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcType_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcType_GotFocus(Index As Integer)
    gCtrlGotFocus edcType(Index)
End Sub
Private Sub edcType_KeyPress(Index As Integer, KeyAscii As Integer)
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
    If (igWinStatus(EVENTNAMESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcENm.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcENm.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
    Me.KeyPreview = True
    EName.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        cbcVeh.Visible = False
        cbcVeh.Visible = True
        cbcEType.Visible = False
        cbcEType.Visible = True
        cbcSelect.Visible = False
        cbcSelect.Visible = True
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

    btrExtClear hmEtf   'Clear any previous extend operation
    ilRet = btrClose(hmEtf)
    btrDestroy hmEtf
    btrExtClear hmCef   'Clear any previous extend operation
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    btrExtClear hmEnf   'Clear any previous extend operation
    ilRet = btrClose(hmEnf)
    btrDestroy hmEnf

    Set EName = Nothing   'Remove data segment
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcGenre_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcGenre, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcGenre_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcGenre_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcGenre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcGenre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcGenre, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcLen_Click()
    gProcessLbcClick lbcLen, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcLen_GotFocus()
    gCtrlGotFocus ActiveControl
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

    edcName.Text = ""
    lbcGenre.ListIndex = -1
    edcComment.Text = ""
    lbcTime.ListIndex = -1
    lbcLen.ListIndex = -1
    lbcProg.ListIndex = -1
    edcSource.Text = ""
    edcType(0).Text = ""
    edcType(1).Text = ""
    tmEnf.iVefCode = 0  'Force this field to be reset in mMoveCtrlToRec
    tmEnf.iEtfCode = 0  'Force this field to be reset in mMoveCtrlToRec
    smComment = ""
    smGenre = ""
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
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            edcName.MaxLength = 30
            gMoveFormCtrl pbcENm, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case GENREINDEX   'Invoice sorting
            mGenrePop
            If imTerminate Then
                Exit Sub
            End If
            lbcGenre.Height = gListBoxHeight(lbcGenre.ListCount, 6)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcENm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcGenre.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcGenre.ListIndex < 0 Then
                If lbcGenre.ListCount <= 1 Then
                    lbcGenre.ListIndex = 0   '[New]
                Else
                    lbcGenre.ListIndex = 1
                End If
            End If
            If lbcGenre.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcGenre.List(lbcGenre.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case COMMENTINDEX 'Comment
            edcComment.Width = tmCtrls(ilBoxNo).fBoxW
            edcComment.MaxLength = 1000
            edcComment.Move pbcENm.Left + tmCtrls(ilBoxNo).fBoxX, pbcENm.Top + tmCtrls(ilBoxNo).fBoxY + fgOffset
            edcComment.Visible = True  'Set visibility
            edcComment.SetFocus
        Case TIMEINDEX   'Time format
            lbcTime.Height = gListBoxHeight(lbcTime.ListCount, 3)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW ' - cmcDropDown.Width
            edcDropDown.MaxLength = 17
            gMoveFormCtrl pbcENm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcTime.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcTime.ListIndex < 0 Then
                lbcTime.ListIndex = Val(tmEtf.sTimeForm) - 1
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
            gMoveFormCtrl pbcENm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcLen.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcLen.ListIndex < 0 Then
                lbcLen.ListIndex = Val(tmEtf.sLenForm) - 1
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
            gMoveFormCtrl pbcENm, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcProg.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcProg.ListIndex < 0 Then
                lbcProg.ListIndex = Val(tmEtf.sPgmForm) - 1
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
        Case SOURCEINDEX    'Source Index
            edcSource.Width = tmCtrls(ilBoxNo).fBoxW
            edcSource.MaxLength = 3
            gMoveFormCtrl pbcENm, edcSource, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcSource.Visible = True  'Set visibility
            edcSource.SetFocus
        Case TYPEINDEX    'Source Index
            edcType(0).Width = tmCtrls(ilBoxNo).fBoxW
            edcType(0).MaxLength = 3
            gMoveFormCtrl pbcENm, edcType(0), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcType(0).Visible = True  'Set visibility
            edcType(0).SetFocus
        Case TYPEINDEX + 1
            edcType(1).Width = tmCtrls(ilBoxNo).fBoxW
            edcType(1).MaxLength = 3
            gMoveFormCtrl pbcENm, edcType(1), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcType(1).Visible = True  'Set visibility
            edcType(1).SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mETypeBranch                    *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to event  *
'*                      type and process               *
'*                      communication back from event  *
'*                      type                           *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mETypeBranch() As Integer
'
'   ilRet = mETypeBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionLookAhead(cbcEType, imBSMode, slStr)
    If ilRet = 0 Then
        slNameCode = tmETypeCode(imETypeSelectedIndex - 1).sKey 'lbcETypeCode.List(imETypeSelectedIndex - 1)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        On Error GoTo mETypeBranchErr
        gCPErrorMsg ilRet, "mETypeBranch (gParseItem field 2)", EName
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmEtfSrchKey.iCode = CInt(slCode)
        ilRet = btrGetEqual(hmEtf, tmEtf, imEtfRecLen, tmEtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mETypeBranchErr
        gBtrvErrorMsg ilRet, "mETypeBranch (btrGetEqual)", EName
        On Error GoTo 0
        mETypeBranch = False
    Else
        'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
        '    mETypeBranch = True
        '    cbcEType.SetFocus
        '    Exit Function
        'End If
        'Screen.MousePointer = vbHourGlass  'Wait
        igETypeCallSource = CALLSOURCEENAME
        If cbcEType.Text = "[New]" Then
            sgETypeName = ""
        Else
            sgETypeName = slStr
        End If
        ilUpdateAllowed = imUpdateAllowed
        'igChildDone = False
        'edcLinkSrceDoneMsg.Text = ""
        'If (Not igStdAloneMode) And (imShowHelpMsg) Then
            If igTestSystem Then
                slStr = "EName^Test\" & sgUserName & "\" & Trim$(str$(igETypeCallSource)) & "\" & sgETypeName
            Else
                slStr = "EName^Prod\" & sgUserName & "\" & Trim$(str$(igETypeCallSource)) & "\" & sgETypeName
            End If
        'Else
        '    If igTestSystem Then
        '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igETypeCallSource)) & "\" & sgETypeName
        '    Else
        '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igETypeCallSource)) & "\" & sgETypeName
        '    End If
        'End If
        'lgShellRet = Shell(sgExePath & "EType.Exe " & slStr, 1)
        'EName.Enabled = False
        'Do While Not igChildDone
        '    DoEvents
        'Loop
        sgCommandStr = slStr
        EType.Show vbModal
        slStr = sgDoneMsg
        ilParse = gParseItem(slStr, 1, "\", sgETypeName)
        igETypeCallSource = Val(sgETypeName)
        ilParse = gParseItem(slStr, 2, "\", sgETypeName)
        'EName.Enabled = True
        'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
        'For ilLoop = 0 To 10
        '    DoEvents
        'Next ilLoop
        'Screen.MousePointer = vbDefault    'Default
        imUpdateAllowed = ilUpdateAllowed
'        gShowBranner
        'If imUpdateAllowed = False Then
        '    mSendHelpMsg "BF"
        'Else
        '    mSendHelpMsg "BT"
        'End If
        gShowBranner imUpdateAllowed
        mETypeBranch = False
        If igETypeCallSource = CALLDONE Then  'Done
            igETypeCallSource = CALLNONE
            smETypeCodeTag = ""
            cbcEType.Clear
            mETypePop
            If imTerminate Then
                mETypeBranch = False
                Exit Function
            End If
            gFindMatch sgETypeName, 1, cbcEType
            sgETypeName = ""
            If gLastFound(cbcEType) > 0 Then
                imChgMode = True
                cbcEType.ListIndex = gLastFound(cbcEType)
                imETypeSelectedIndex = cbcEType.ListIndex
                imChgMode = False
            Else
                imChgMode = True
                cbcEType.ListIndex = 0
                imChgMode = False
                cbcEType.SetFocus
                Exit Function
            End If
            slNameCode = tmETypeCode(imETypeSelectedIndex - 1).sKey 'lbcETypeCode.List(imETypeSelectedIndex - 1)
            ilRet = gParseItem(slNameCode, 3, "\", slCode)
            On Error GoTo mETypeBranchErr
            gCPErrorMsg ilRet, "mETypeBranch (gParseItem field 2)", EName
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmEtfSrchKey.iCode = CInt(slCode)
            ilRet = btrGetEqual(hmEtf, tmEtf, imEtfRecLen, tmEtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mETypeBranchErr
            gBtrvErrorMsg ilRet, "mETypeBranch (btrGetEqual)", EName
            On Error GoTo 0
        End If
        If igETypeCallSource = CALLCANCELLED Then  'Cancelled
            igETypeCallSource = CALLNONE
            sgETypeName = ""
            cbcEType.SetFocus
            Exit Function
        End If
        If igETypeCallSource = CALLTERMINATED Then
            igETypeCallSource = CALLNONE
            sgETypeName = ""
            cbcEType.SetFocus
            Exit Function
        End If
    End If
    If imEvtListIndex <> cbcEType.ListIndex Then
        cbcSelect.Clear
        sgNameCodeTag = ""
    End If
    imEvtListIndex = cbcEType.ListIndex
    If imVehSelectedIndex >= 0 Then
        cbcSelect.Enabled = True
        cbcSelect.SetFocus
    Else
        cbcVeh.SetFocus
    End If
    Exit Function
mETypeBranchErr:
    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mETypePop                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mETypePop()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = imETypeSelectedIndex
    If ilIndex > 0 Then
        slName = cbcEType.List(ilIndex)
    End If
    'ilRet = gPopEvtNmByTypeBox(EName, True, False, cbcEType, lbcETypeCode)
    ilRet = gPopEvtNmByTypeBox(EName, True, False, cbcEType, tmETypeCode(), smETypeCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mETypePopErr
        gCPErrorMsg ilRet, "mETypePop (gPopEvtNmByTypeBox: EType)", EName
        On Error GoTo 0
        cbcEType.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, cbcEType
            If gLastFound(cbcEType) > 0 Then
                cbcEType.ListIndex = gLastFound(cbcEType)
            Else
                cbcEType.ListIndex = -1
            End If
        Else
            cbcEType.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mETypePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenreBranch                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to genre  *
'*                      and process communication      *
'*                      back from genre                *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mGenreBranch() As Integer
'
'   ilRet = mGenreBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDropDown, lbcGenre, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        imDoubleClickName = False
        mGenreBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(GENRESLIST)) Then
    '    imDoubleClickName = False
    '    mGenreBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "E"
    igMNmCallSource = CALLSOURCEENAME
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "EName^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "EName^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'EName.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mGenreBranch = True
    imUpdateAllowed = ilUpdateAllowed
'    gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcGenre.Clear
        smGenreCodeTag = ""
        mGenrePop
        If imTerminate Then
            mGenreBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcGenre
        sgMNmName = ""
        If gLastFound(lbcGenre) > 0 Then
            imChgMode = True
            lbcGenre.ListIndex = gLastFound(lbcGenre)
            edcDropDown.Text = lbcGenre.List(lbcGenre.ListIndex)
            imChgMode = False
            mGenreBranch = False
            mSetChg imBoxNo
        Else
            imChgMode = True
            lbcGenre.ListIndex = 0
            edcDropDown.Text = lbcGenre.List(0)
            imChgMode = False
            mSetChg imBoxNo
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenrePop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Genre list            *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mGenrePop()
'
'   mGenrePop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcGenre.ListIndex
    If ilIndex > 0 Then
        slName = lbcGenre.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "E"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(EName, lbcGenre, lbcGenreCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(EName, lbcGenre, tmGenreCode(), smGenreCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mGenrePopErr
        gCPErrorMsg ilRet, "mGenrePop (gIMoveListBox)", EName
        On Error GoTo 0
        lbcGenre.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcGenre
            If gLastFound(lbcGenre) > 0 Then
                lbcGenre.ListIndex = gLastFound(lbcGenre)
            Else
                lbcGenre.ListIndex = -1
            End If
        Else
            lbcGenre.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mGenrePopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
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
    Dim ilRet As Integer    'Return Status
    imFirstActivate = True
    imTerminate = False
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    EName.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone EName
    'EName.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imRecLen = Len(tmEnf)  'Get and save ARF record length
    imEtfRecLen = Len(tmEtf)
    imFirstFocusVeh = True
    imFirstFocusEType = True
    imFirstFocusEName = True
    ilRet = gParseItem(sgENameName, 1, "\", smInitVeh)
    If ilRet <> CP_MSG_NONE Then
        smInitVeh = ""
    Else
        smInitVeh = Trim$(smInitVeh)
    End If
    ilRet = gParseItem(sgENameName, 2, "\", smInitEType)
    If ilRet <> CP_MSG_NONE Then
        smInitEType = ""
    Else
        smInitEType = Trim$(smInitEType)
    End If
    ilRet = gParseItem(sgENameName, 3, "\", smInitEName)
    If ilRet <> CP_MSG_NONE Then
        smInitEName = ""
    Else
        smInitEName = Trim$(smInitEName)
    End If
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmEtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmEtf, "", sgDBPath & "Etf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Etf.Btr)", EName
    On Error GoTo 0
    hmEnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Enf.btr)", EName
    On Error GoTo 0
    hmCef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCef, "", sgDBPath & "CEF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CEF.Btr)", EName
    On Error GoTo 0
    'Populate facilty and event type list boxes
    cbcVeh.Clear 'Force population
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    cbcEType.Clear 'Force population
    mETypePop
    If imTerminate Then
        Exit Sub
    End If
    lbcGenre.Clear 'Force population
    mGenrePop
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
'    gCenterModalForm EName
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list box to be populated
    mPopulate 'Only populate after user selected which Vehicle
    If Not imTerminate Then
        If cbcSelect.ListCount > 0 Then
            cbcSelect.ListIndex = 0 'This will generate a select_change event
        End If
        mSetCommands
    End If
    imTimeFirst = True
    imLenFirst = True
    imProgFirst = True
    imFacListIndex = 0   'Retain Vehicle index when mPopulate called, so changes can force repopulation
    imEvtListIndex = 0   'Retain Event type index when mPopulate called, so changes can force repopulation
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
'*             Created:5/13/93       By:D. LeVine      *
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
    flTextHeight = pbcENm.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcENm.Move 420, 1050, pbcENm.Width + fgPanelAdj, pbcENm.Height + fgPanelAdj
    pbcENm.Move plcENm.Left + fgBevelX, plcENm.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2895, fgBoxStH
    'Genre
    gSetCtrl tmCtrls(GENREINDEX), 2940, tmCtrls(NAMEINDEX).fBoxY, 2715, fgBoxStH
    'Comment
    gSetCtrl tmCtrls(COMMENTINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 5625, fgBoxStH
    tmCtrls(COMMENTINDEX).iReq = False
    'Time
    gSetCtrl tmCtrls(TIMEINDEX), 30, tmCtrls(COMMENTINDEX).fBoxY + 2 * fgStDeltaY, 1800, fgBoxStH
    'Length
    gSetCtrl tmCtrls(LENINDEX), 1845, tmCtrls(TIMEINDEX).fBoxY, 1080, fgBoxStH
    'Program/Sponsor
    gSetCtrl tmCtrls(PROGINDEX), 2940, tmCtrls(TIMEINDEX).fBoxY, 1095, fgBoxStH
    'Source
    gSetCtrl tmCtrls(SOURCEINDEX), 4050, tmCtrls(TIMEINDEX).fBoxY, 540, fgBoxStH
    tmCtrls(SOURCEINDEX).iReq = False
    '1st Program type
    gSetCtrl tmCtrls(TYPEINDEX), 4605, tmCtrls(TIMEINDEX).fBoxY, 525, fgBoxStH
    tmCtrls(TYPEINDEX).iReq = False
    '2nd Program type
    gSetCtrl tmCtrls(TYPEINDEX + 1), 5115, tmCtrls(TIMEINDEX).fBoxY, 525, fgBoxStH
    tmCtrls(TYPEINDEX + 1).iReq = False
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
    Dim ilLoop As Integer
    Dim slFacCode As String  'Vehicle name and code
    Dim slETypeCode As String   'Event type name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'code number
    Dim slNameCode As String
    'Vehicle can't be changed
    If (imVehSelectedIndex >= 0) And (tmEnf.iVefCode = 0) Then
        slFacCode = tgVehicle(imVehSelectedIndex).sKey 'Traffic!lbcVehicle.List(imVehSelectedIndex)
        ilRet = gParseItem(slFacCode, 2, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", EName
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmEnf.iVefCode = CInt(slCode)
    End If
    If (imETypeSelectedIndex > 0) And (tmEnf.iEtfCode = 0) Then
        slETypeCode = tmETypeCode(imETypeSelectedIndex - 1).sKey    'lbcETypeCode.List(imETypeSelectedIndex - 1)
        ilRet = gParseItem(slETypeCode, 3, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 3)", EName
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmEnf.iEtfCode = CInt(slCode)
        tmEtfSrchKey.iCode = CInt(slCode)
        ilRet = btrGetEqual(hmEtf, tmEtf, imEtfRecLen, tmEtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mMoveCtrlToRecErr
        gBtrvErrorMsg ilRet, "mMoveCtrlToRec (btrGetEqual: Event Type)", EName
        On Error GoTo 0
    End If
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmEnf.sName = edcName.Text
    End If
    If imPrgEType And (Not ilTestChg Or tmCtrls(GENREINDEX).iChg) Then
        If lbcGenre.ListIndex >= 1 Then
            slNameCode = tmGenreCode(lbcGenre.ListIndex - 1).sKey    'lbcGenreCode.List(lbcGenre.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", EName
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmEnf.iMnfGenre = CInt(slCode)
        Else
            tmEnf.iMnfGenre = 0
        End If
    Else
        tmEnf.iMnfGenre = 0
    End If
    If Not ilTestChg Or tmCtrls(COMMENTINDEX).iChg Then
        'tmCef.iStrLen = Len(edcComment.Text)
        tmCef.sComment = Trim$(edcComment.Text) & Chr$(0) '& Chr$(0) 'sgTB
    End If
    If Not ilTestChg Or tmCtrls(TIMEINDEX).iChg Then
        If (lbcTime.Text = "") And (lbcTime.ListIndex < 0) Then
            tmEnf.sTimeForm = ""
        Else
            tmEnf.sTimeForm = Trim$(str$(lbcTime.ListIndex + 1))
        End If
    End If
    If Not ilTestChg Or tmCtrls(LENINDEX).iChg Then
        If (lbcLen.Text = "") And (lbcLen.ListIndex < 0) Then
            tmEnf.sLenForm = ""
        Else
            tmEnf.sLenForm = Trim$(str$(lbcLen.ListIndex + 1))
        End If
    End If
    If Not ilTestChg Or tmCtrls(PROGINDEX).iChg Then
        If (lbcProg.Text = "") And (lbcProg.ListIndex < 0) Then
            tmEnf.sPgmForm = ""
        Else
            tmEnf.sPgmForm = Trim$(str$(lbcProg.ListIndex + 1))
        End If
    End If
    If Not ilTestChg Or tmCtrls(SOURCEINDEX).iChg Then
        tmEnf.sPgmSource = edcSource.Text
    End If
    For ilLoop = 0 To 1 Step 1
        If Not ilTestChg Or tmCtrls(TYPEINDEX + ilLoop).iChg Then
            tmEnf.sType(ilLoop) = edcType(ilLoop).Text
        End If
    Next ilLoop
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
    Dim slNameCode As String
    Dim slCode As String
    Dim slRecCode As String
    Dim ilRet As Integer
    edcName.Text = Trim$(tmEnf.sName)
    lbcGenre.ListIndex = 0
    smGenre = ""
    slRecCode = Trim$(str$(tmEnf.iMnfGenre))
    If imPrgEType Then
        For ilLoop = 0 To UBound(tmGenreCode) - 1 Step 1  'lbcGenreCode.ListCount - 1 Step 1
            slNameCode = tmGenreCode(ilLoop).sKey    'lbcGenreCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", EName
            On Error GoTo 0
            If slRecCode = slCode Then
                lbcGenre.ListIndex = ilLoop + 1
                smGenre = lbcGenre.List(ilLoop + 1)
                Exit For
            End If
        Next ilLoop
    End If
    'If tmCef.iStrLen > 0 Then
    '    edcComment.Text = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
    'Else
    '    edcComment.Text = ""
    'End If
    edcComment.Text = gStripChr0(tmCef.sComment)
    smComment = edcComment.Text
    If tmEnf.sTimeForm = "" Then
        lbcTime.ListIndex = -1
    Else
        lbcTime.ListIndex = Val(tmEnf.sTimeForm) - 1
    End If
    If tmEnf.sLenForm = "" Then
        lbcLen.ListIndex = -1
    Else
        lbcLen.ListIndex = Val(tmEnf.sLenForm) - 1
    End If
    If tmEnf.sPgmForm = "" Then
        lbcProg.ListIndex = -1
    Else
        lbcProg.ListIndex = Val(tmEnf.sPgmForm) - 1
    End If
    edcSource.Text = Trim$(tmEnf.sPgmSource)
    For ilLoop = 0 To 1 Step 1
        edcType(ilLoop).Text = Trim$(tmEnf.sType(ilLoop))
    Next ilLoop
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
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
                    MsgBox "Event name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmEnf.sName) 'Reset text
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
    'gInitStdAlone EName, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igENameCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igENameCallSource = CALLSOURCEPEVENT 'CALLNONE
    '    slCommand = "1\2\3\ACC\Contract Avails\Net Test"
    'End If
    If igENameCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgENameName = slStr
            ilRet = gParseItem(slCommand, 5, "\", slStr)
            If ilRet = CP_MSG_NONE Then
                sgENameName = sgENameName & "\" & slStr
                ilRet = gParseItem(slCommand, 6, "\", slStr)
                If ilRet = CP_MSG_NONE Then
                    sgENameName = sgENameName & "\" & slStr
                End If
            End If
        Else
            sgENameName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
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
    Dim slNameCode As String
    Dim slCode As String
    ReDim ilfilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    imPopReqd = False
    If (imVehSelectedIndex >= 0) And (imETypeSelectedIndex >= 1) Then
        slNameCode = tgVehicle(imVehSelectedIndex).sKey    'Traffic!lbcVehicle.List(imVehSelectedIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gParseItem field 2)", EName
        On Error GoTo 0
        ilfilter(0) = INTEGERFILTER
        slFilter(0) = slCode
        ilOffSet(0) = gFieldOffset("Enf", "EnfVefCode") '2
        slNameCode = tmETypeCode(imETypeSelectedIndex - 1).sKey 'lbcETypeCode.List(imETypeSelectedIndex - 1)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gParseItem field 3)", EName
        On Error GoTo 0
        ilfilter(1) = INTEGERFILTER
        slFilter(1) = slCode
        ilOffSet(1) = gFieldOffset("Enf", "EnfEtfCode") '4
        'ilRet = gIMoveListBox(EName, cbcSelect, lbcNameCode, "Enf.Btr", gFieldOffset("Enf", "EnfName"), 30, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(EName, cbcSelect, tgNameCode(), sgNameCodeTag, "Enf.Btr", gFieldOffset("Enf", "EnfName"), 30, ilfilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", EName
            On Error GoTo 0
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
            imPopReqd = True
        End If
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
'   iRet = mReadRec(ilSelectIndex, ilForUpdate)
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
    gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", EName
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmEnf, tmEnf, imRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Event Name)", EName
    On Error GoTo 0
    tmCefSrchKey.lCode = tmEnf.lCefCode
    If tmEnf.lCefCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mSaveRec (btrGetEqual:Comment)", EName
        On Error GoTo 0
    Else
        tmCef.lCode = 0
        'tmCef.iStrLen = 0
        tmCef.sComment = ""
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
    Dim tlEnf As ENF
    Dim ilEnfRet As Integer
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
        If imSelectedIndex <> 0 Then
            'Reread record in so latest is obtained
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        imCefRecLen = Len(tmCef)    '5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
        If imSelectedIndex = 0 Then 'New selected
            'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
            If gStripChr0(tmCef.sComment) <> "" Then
                tmCef.lCode = 0 'Autoincrement
                ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
            Else
                tmCef.lCode = 0
                ilRet = BTRV_ERR_NONE
            End If
            slMsg = "mSaveRec (btrInsert: Comment)"
        Else 'Old record-Update
            'If imCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
            If gStripChr0(tmCef.sComment) <> "" Then
                If tmCef.lCode = 0 Then
                    tmCef.lCode = 0 'Autoincrement
                    ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                Else
                    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                End If
            Else
                If tmEnf.lCefCode <> 0 Then
                    ilRet = btrDelete(hmCef)
                End If
                tmCef.lCode = 0
            End If
            slMsg = "mSaveRec (btrUpdate: Comment)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, EName
    On Error GoTo 0
    tmEnf.lCefCode = tmCef.lCode
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Enf.btr")
        'If Len(lbcNameCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcNameCode.Tag, Len(lbcNameCode.Tag) - Len(slStamp))
        'End If
        If Len(sgNameCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(sgNameCodeTag, Len(sgNameCodeTag) - Len(slStamp))
        End If
        If imSelectedIndex = 0 Then 'New selected
            tmEnf.iCode = 0  'Autoincrement
            tmEnf.iMerge = 0
            ilRet = btrInsert(hmEnf, tmEnf, imRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Event Name)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmEnf, tmEnf, imRecLen)
            slMsg = "mSaveRec (btrUpdate: Event Name)"
            If ilRet = BTRV_ERR_CONFLICT Then
                tmSrchKey.iCode = tmEnf.iCode
                ilEnfRet = btrGetEqual(hmEnf, tlEnf, imRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, EName
    On Error GoTo 0
'    'If lbcNameCode.Tag <> "" Then
'    '    If slStamp = lbcNameCode.Tag Then
'    '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Enf.btr")
'    '        If Len(slStamp) > Len(lbcNameCode.Tag) Then
'    '            lbcNameCode.Tag = lbcNameCode.Tag & Right$(slStamp, Len(slStamp) - Len(lbcNameCode.Tag))
'    '        End If
'    '    End If
'    'End If
'    If sgNameCodeTag <> "" Then
'        If slStamp = sgNameCodeTag Then
'            sgNameCodeTag = gFileDateTime(sgDBPath & "Enf.btr")
'            If Len(slStamp) > Len(sgNameCodeTag) Then
'                sgNameCodeTag = sgNameCodeTag & right$(slStamp, Len(slStamp) - Len(sgNameCodeTag))
'            End If
'        End If
'    End If
'    If imSelectedIndex <> 0 Then
'        'lbcNameCode.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
'        cbcSelect.RemoveItem imSelectedIndex
'    End If
'    cbcSelect.RemoveItem 0 'Remove [New]
'    slName = RTrim$(tmEnf.sName)
'    cbcSelect.AddItem slName
'    slName = tmEnf.sName + "\" + LTrim$(Str$(tmEnf.iCode)) 'slName + "\" + LTrim$(Str$(tmEnf.iCode))
'    'lbcNameCode.AddItem slName
'    gAddItemToSortCode slName, tgNameCode(), True
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
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcENm_Paint
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
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmEnf.sName, edcName, tmCtrls(ilBoxNo)
        Case GENREINDEX   'Genre
            gSetChgFlag smGenre, lbcGenre, tmCtrls(ilBoxNo)
        Case COMMENTINDEX   'Comment
            gSetChgFlag smComment, edcComment, tmCtrls(ilBoxNo)
        Case TIMEINDEX 'Time format
            If tmEnf.sTimeForm = "" Then
                gSetChgFlag tmEnf.sTimeForm, lbcTime, tmCtrls(ilBoxNo)
            Else
                slStr = lbcTime.List(Val(tmEnf.sTimeForm) - 1)
                gSetChgFlag slStr, lbcTime, tmCtrls(ilBoxNo)
            End If
        Case LENINDEX 'Length format
            If tmEnf.sLenForm = "" Then
                gSetChgFlag tmEnf.sLenForm, lbcLen, tmCtrls(ilBoxNo)
            Else
                slStr = lbcLen.List(Val(tmEnf.sLenForm) - 1)
                gSetChgFlag slStr, lbcLen, tmCtrls(ilBoxNo)
            End If
        Case PROGINDEX 'Selling or Airing or N/At
            If tmEnf.sPgmForm = "" Then
                gSetChgFlag tmEnf.sPgmForm, lbcProg, tmCtrls(ilBoxNo)
            Else
                slStr = lbcProg.List(Val(tmEnf.sPgmForm) - 1)
                gSetChgFlag slStr, lbcProg, tmCtrls(ilBoxNo)
            End If
        Case SOURCEINDEX 'Source
            gSetChgFlag tmEnf.sPgmSource, edcSource, tmCtrls(ilBoxNo)
        Case TYPEINDEX 'Type
            For ilLoop = 0 To 1 Step 1  'Set visibility
                gSetChgFlag tmEnf.sType(ilLoop), edcType(ilLoop), tmCtrls(ilBoxNo + ilLoop)
            Next ilLoop
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
    If (imVehSelectedIndex < 0) Or (imETypeSelectedIndex <= 0) Then
        cbcVeh.Enabled = True
        cbcEType.Enabled = True
        pbcEType.Enabled = True
        cbcSelect.Enabled = False
        pbcENm.Enabled = False  'Disallow mouse
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    Else
        pbcENm.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        If Not ilAltered Then
            cbcSelect.Enabled = True
            cbcVeh.Enabled = True
            cbcEType.Enabled = True
            pbcEType.Enabled = True
        Else
            cbcSelect.Enabled = False
            cbcVeh.Enabled = False
            cbcEType.Enabled = False
            pbcEType.Enabled = False
        End If
    End If
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
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case GENREINDEX   'Genre
            lbcGenre.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcGenre.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcGenre.List(lbcGenre.ListIndex)
            End If
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case COMMENTINDEX 'Comment
            edcComment.Visible = False  'Set visibility
            slStr = edcComment.Text
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case TIMEINDEX 'Time format
            lbcTime.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcTime.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcTime.List(lbcTime.ListIndex)
            End If
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case LENINDEX 'Length format
            lbcLen.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcLen.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcLen.List(lbcLen.ListIndex)
            End If
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case PROGINDEX 'Program
            lbcProg.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcProg.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcProg.List(lbcProg.ListIndex)
            End If
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case SOURCEINDEX   'Source
            edcSource.Visible = False  'Set visibility
            slStr = edcSource.Text
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case TYPEINDEX 'Type
            edcType(0).Visible = False
            slStr = edcType(0).Text
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
        Case TYPEINDEX + 1 'Type
            edcType(1).Visible = False
            slStr = edcType(1).Text
            gSetShow pbcENm, slStr, tmCtrls(ilBoxNo)
    End Select
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
'   mTerminate
'   Where:
'

    sgDoneMsg = Trim$(str$(igENameCallSource)) & "\" & sgENameName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload EName
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
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If imPrgEType And ((ilCtrlNo = GENREINDEX) Or (ilCtrlNo = TESTALLCTRLS)) Then
        If lbcGenre.ListCount <= 1 Then
            tmCtrls(GENREINDEX).iReq = False
        Else
            tmCtrls(GENREINDEX).iReq = True
        End If
        If gFieldDefinedCtrl(lbcGenre, "", "Genre must be specified", tmCtrls(GENREINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = GENREINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = COMMENTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcComment, "", "Comment must be specified", tmCtrls(COMMENTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = COMMENTINDEX
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
        If gFieldDefinedCtrl(lbcProg, "", "Name format must be specified", tmCtrls(PROGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PROGINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SOURCEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcSource, "", "Market Rank must be specified", tmCtrls(SOURCEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SOURCEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcType(0), "", "Type must be specified", tmCtrls(TYPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TYPEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If gFieldDefinedCtrl(edcType(1), "", "Type must be specified", tmCtrls(TYPEINDEX + 1).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TYPEINDEX + 1
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    mTestFields = YES
End Function
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
    'ilRet = gPopUserVehicleBox(EName, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, cbcVeh, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(EName, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, cbcVeh, tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox)", EName
        On Error GoTo 0
'        cbcVeh.AddItem "[New]", 0  'Force as first item on list
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
Private Sub pbcENm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (ilBox = GENREINDEX) And (Not imPrgEType) Then
                    Beep
                    Exit Sub
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcENm_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcENm.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcENm.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcENm.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcEType_GotFocus()
    If GetFocus() <> pbcEType.hwnd Then
        Exit Sub
    End If
    If mETypeBranch() Then
        Exit Sub
    End If
    plcSelect.Visible = False
    plcSelect.Visible = True
    EName.Refresh
    mSetCommands
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
    If imBoxNo = GENREINDEX Then
        If mGenreBranch() Then
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
                    If (lbcGenre.ListCount > 1) And imPrgEType Then
                        ilBox = 2
                    Else
                        ilBox = 3
                    End If
                End If
            Case 1 'Name (first control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
            Case COMMENTINDEX
                If (lbcGenre.ListCount = 1) Or (Not imPrgEType) Then
                    ilFound = False
                End If
                ilBox = GENREINDEX
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
    If imBoxNo = GENREINDEX Then
        If mGenreBranch() Then
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
            Case NAMEINDEX
                If (lbcGenre.ListCount = 1) Or (Not imPrgEType) Then
                    ilFound = False
                End If
                ilBox = GENREINDEX
            Case TYPEINDEX + 1 'last control
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igENameCallSource = CALLNONE) Then
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
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case GENREINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcGenre, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
