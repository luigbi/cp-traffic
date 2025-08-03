VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Persnnel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   1380
   ClientTop       =   2895
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
   ScaleHeight     =   4320
   ScaleWidth      =   6570
   Begin VB.TextBox edcEMail 
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
      Left            =   495
      TabIndex        =   16
      Top             =   3030
      Visible         =   0   'False
      Width           =   5520
   End
   Begin VB.TextBox edcBAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   915
      MaxLength       =   40
      TabIndex        =   15
      Top             =   2580
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcBAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   720
      MaxLength       =   40
      TabIndex        =   14
      Top             =   2445
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcBAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   585
      MaxLength       =   40
      TabIndex        =   13
      Top             =   2325
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   720
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1875
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   615
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1740
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox edcCAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   495
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1620
      Visible         =   0   'False
      Width           =   4215
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
      Left            =   2745
      TabIndex        =   1
      Top             =   315
      Width           =   3390
   End
   Begin VB.TextBox edcComment 
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
      Height          =   540
      HelpContextID   =   8
      Left            =   855
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3255
      Visible         =   0   'False
      Width           =   5625
   End
   Begin MSMask.MaskEdBox mkcFax 
      Height          =   210
      Left            =   3270
      TabIndex        =   8
      Tag             =   "The number and extension of the buyer."
      Top             =   1335
      Visible         =   0   'False
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      ForeColor       =   0
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   480
      TabIndex        =   7
      Tag             =   "The number and extension of the buyer."
      Top             =   1335
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      ForeColor       =   0
      MaxLength       =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA Ext(AAAA)"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5010
      ScaleHeight     =   210
      ScaleWidth      =   1050
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox edcTitle 
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
      Left            =   3285
      MaxLength       =   20
      TabIndex        =   6
      Top             =   975
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   6
      Left            =   6090
      Top             =   3795
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
      Left            =   5760
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3825
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
      Left            =   5760
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3690
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
      Left            =   5835
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3495
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1770
      Width           =   75
   End
   Begin VB.TextBox edcName 
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
      Left            =   420
      MaxLength       =   30
      TabIndex        =   5
      Top             =   945
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
      Left            =   225
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
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
      Left            =   195
      ScaleHeight     =   90
      ScaleWidth      =   75
      TabIndex        =   18
      Top             =   1980
      Width           =   75
   End
   Begin VB.PictureBox pbcPersonnel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   435
      Picture         =   "Persnnel.frx":0000
      ScaleHeight     =   2790
      ScaleWidth      =   5670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   810
      Width           =   5670
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   2655
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   2655
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1065
      TabIndex        =   19
      Top             =   3915
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Height          =   285
      Left            =   4350
      TabIndex        =   22
      Top             =   3915
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3255
      TabIndex        =   21
      Top             =   3915
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   3915
      Width           =   945
   End
   Begin VB.PictureBox plcPersonnel 
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
      Height          =   2910
      Left            =   390
      ScaleHeight     =   2850
      ScaleWidth      =   5700
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   765
      Width           =   5760
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   3810
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Persnnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Persnnel.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Persnnel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Advertiser Product input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 13)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imAdvtOrAgyFlag As Integer   '0=Advertiser; 1=Agency
Dim imPassedAdvtOrAgyCode As Integer    'Advertiser or agency code
Dim smBuyerOrPayableFlag As String
Dim smInitPersonnel As String
Dim tmPnf As PNF        'Pnf record image
Dim tmPnfSrchKey As INTKEY0    'Pnf key record image
Dim hmPnf As Integer    'Personnel file handle
Dim imPnfRecLen As Integer
'E-mail
Dim tmCef As CEF            'CEF record image
Dim hmCef As Integer        'CEF Handle
Dim imCefRecLen As Integer      'CEF record length
Dim tmCefSrchKey As LONGKEY0    'CEF key record image
Dim smEMail As String
'Comment
Dim tmCxf As CXF        'Cxf record image
Dim tmCxfSrchKey As LONGKEY0    'Cxf key record image
Dim hmCxf As Integer    'Personnel file handle
Dim imCxfRecLen As Integer
Dim smComment As String     'Save original value
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
Dim imBoxNo As Integer   'Current Media Box
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imUpdateAllowed As Integer    'User can update records
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visibly
                                'False= Make list box invisible
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imFirstFocusPersonnel As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imRecLen As Integer        'ADF record length
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim smPhoneImage As String  'Blank phone image- obtained from mkcPhone.text before input
Dim smFaxImage As String    'Blank fax image
Dim imState As Integer  '0=Active; 1=Dormant
Dim smNewFlag As String 'Send to calling program: N=No; Y=Yes
Dim imMaxIndex As Integer
Dim smReturnCode As String

Const NAMEINDEX = 1     'Name control/field
Const TITLEINDEX = 2
Const PHONEINDEX = 3
Const FAXINDEX = 4
Const STATEINDEX = 5
Const EMAILINDEX = 6
Const COMMENTINDEX = 7
Const CADDRINDEX = 8
Const BADDRINDEX = 11

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
    pbcPersonnel.Cls
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
    pbcPersonnel_Paint
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
    If imFirstFocusPersonnel Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocusPersonnel = False
        If igPersonnelCallSource <> CALLNONE Then  'If from advt or contract- set name and branch to control
            If smInitPersonnel = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = smInitPersonnel    'New name
            End If
            cbcSelect_Change
            If smInitPersonnel <> "" Then
                'mSetCommands
                gFindMatch smInitPersonnel, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cbcSelect.ListIndex = gLastFound(cbcSelect)
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
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        mSetCommands
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
    smNewFlag = "N"
    If igPersonnelCallSource <> CALLNONE Then
        igPersonnelCallSource = CALLCANCELLED
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
    smNewFlag = "N"
    If igPersonnelCallSource <> CALLNONE Then
        sgPersonnelName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgPersonnelName = "[New]"
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
    If igPersonnelCallSource <> CALLNONE Then
        If sgPersonnelName = "[New]" Then
            igPersonnelCallSource = CALLCANCELLED
        Else
            igPersonnelCallSource = CALLDONE
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
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        Screen.MousePointer = vbHourglass
        If smBuyerOrPayableFlag = "B" Then
            ilRet = gIICodeRefExist(Persnnel, tmPnf.iCode, "Agf.Btr", "AgfPnfBuyer")
            If ilRet Then
                Screen.MousePointer = vbDefault
                slMsg = "Cannot erase - a Agency references name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            ilRet = gIICodeRefExist(Persnnel, tmPnf.iCode, "Adf.Btr", "AdfPnfBuyer")
            If ilRet Then
                Screen.MousePointer = vbDefault
                slMsg = "Cannot erase - a Advertiser references name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        Else
            ilRet = gIICodeRefExist(Persnnel, tmPnf.iCode, "Agf.Btr", "AgfPnfPay")
            If ilRet Then
                Screen.MousePointer = vbDefault
                slMsg = "Cannot erase - a Agency references name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            ilRet = gIICodeRefExist(Persnnel, tmPnf.iCode, "Adf.Btr", "AdfPnfPay")
            If ilRet Then
                Screen.MousePointer = vbDefault
                slMsg = "Cannot erase - a Advertiser references name"
                ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmPnf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        gGetSyncDateTime slSyncDate, slSyncTime
        slStamp = gFileDateTime(sgDBPath & "Pnf.Btr")
        If tmPnf.lCxfCode > 0 Then
            ilRet = btrDelete(hmCxf)
'            If tgSpf.sRemoteUsers = "Y" Then
'                tmDsf.lCode = 0
'                tmDsf.sFileName = "CXF"
'                gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                tmDsf.iRemoteID = tmCxf.iRemoteID
'                tmDsf.lAutoCode = tmCxf.lAutoCode
'                tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'                tmDsf.lCntrNo = 0
'                ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'            End If
        End If
        ilRet = btrDelete(hmPnf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Persnnel
        On Error GoTo 0
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "PNF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmPnf.iRemoteID
'            tmDsf.lAutoCode = tmPnf.iAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcPersonnelCode.Tag <> "" Then
        '    If slStamp = lbcPersonnelCode.Tag Then
        '        lbcPersonnelCode.Tag = FileDateTime(sgDBPath & "Pnf.Btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Pnf.Btr")
            End If
        End If
        'lbcPersonnelCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcPersonnel.Cls
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
    Dim imSvSelectedIndex As Integer
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
'    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
'        cbcSelect.Text = slName
''        cbcSelect_Change    'Call change so picture area repainted
'    Else
'        cbcSelect.ListIndex = 0
''        mClearCtrlFields 'This is required as select_change will not be generated
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmPnf.iCode
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

Private Sub edcEMail_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcEMail_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEMail_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
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
End Sub

Private Sub edcTitle_Change()
    mSetChg TITLEINDEX
End Sub
Private Sub edcTitle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTitle_KeyPress(KeyAscii As Integer)
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
    If imAdvtOrAgyFlag = 0 Then
        If (igWinStatus(ADVERTISERSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcPersonnel.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcPersonnel.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    Else
        If (igWinStatus(AGENCIESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            pbcPersonnel.Enabled = False
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
            imUpdateAllowed = False
        Else
            pbcPersonnel.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
            imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Persnnel.Refresh
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
    
    Erase tgNameCode

    btrExtClear hmPnf   'Clear any previous extend operation
    ilRet = btrClose(hmPnf)
    btrDestroy hmPnf
    btrExtClear hmCxf   'Clear any previous extend operation
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    btrExtClear hmCef   'Clear any previous extend operation
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    
    Set Persnnel = Nothing   'Remove data segment

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
    edcTitle.Text = ""
    mkcPhone.Text = smPhoneImage
    mkcFax.Text = smFaxImage
    imState = -1
    For ilLoop = 0 To 2 Step 1
        edcCAddr(ilLoop).Text = ""
        edcBAddr(ilLoop).Text = ""
    Next ilLoop
    edcEMail.Text = ""
    smEMail = ""
    edcComment.Text = ""
    smComment = ""
    tmPnf.iAdfCode = 0
    tmPnf.iAgfCode = 0
    smReturnCode = ""
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
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            If smBuyerOrPayableFlag = "B" Then
                edcName.MaxLength = 30
            Else
                edcName.MaxLength = 30
            End If
            gMoveFormCtrl pbcPersonnel, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case TITLEINDEX 'Title
            edcTitle.Width = tmCtrls(ilBoxNo).fBoxW
            edcTitle.MaxLength = 20
            gMoveFormCtrl pbcPersonnel, edcTitle, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcTitle.Visible = True  'Set visibility
            edcTitle.SetFocus
        Case PHONEINDEX 'Phone and extension
            mkcPhone.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcPersonnel, mkcPhone, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcPhone.Visible = True  'Set visibility
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax
            mkcFax.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcPersonnel, mkcFax, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcFax.Visible = True  'Set visibility
            mkcFax.SetFocus
        Case STATEINDEX   'Active/Dormant
            If imState < 0 Then
                imState = 0    'Active
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcPersonnel, pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcState_Paint
            pbcState.Visible = True
            pbcState.SetFocus
        Case CADDRINDEX 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcPersonnel, edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcPersonnel, edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Width = tmCtrls(CADDRINDEX).fBoxW
            edcCAddr(ilBoxNo - CADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcPersonnel, edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = True  'Set visibility
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case BADDRINDEX 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcPersonnel, edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcPersonnel, edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Width = tmCtrls(BADDRINDEX).fBoxW
            edcBAddr(ilBoxNo - BADDRINDEX).MaxLength = 40
            gMoveFormCtrl pbcPersonnel, edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = True  'Set visibility
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case EMAILINDEX 'Address ID
            edcEMail.Width = tmCtrls(ilBoxNo).fBoxW
            edcEMail.MaxLength = 0
            gMoveFormCtrl pbcPersonnel, edcEMail, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcEMail.Visible = True  'Set visibility
            edcEMail.SetFocus
        Case COMMENTINDEX 'Comment
            edcComment.Width = tmCtrls(ilBoxNo).fBoxW
            edcComment.MaxLength = 1000
            edcComment.Move pbcPersonnel.Left + tmCtrls(ilBoxNo).fBoxX, pbcPersonnel.Top + tmCtrls(ilBoxNo).fBoxY + fgOffset
            edcComment.Visible = True  'Set visibility
            edcComment.SetFocus
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
    If smBuyerOrPayableFlag = "B" Then
        imMaxIndex = BADDRINDEX + 2
    Else
        imMaxIndex = COMMENTINDEX
    End If
    mInitBox
    Persnnel.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone Persnnel
    'Persnnel.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture

    imPopReqd = False
    imRecLen = Len(tmPnf)  'Get and save PRF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imDoubleClickName = False
    imFirstFocusPersonnel = True
    imSelectedIndex = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imBypassSetting = False
    smPhoneImage = mkcPhone.Text
    smFaxImage = mkcFax.Text
    hmPnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pnf.Btr)", Persnnel
    On Error GoTo 0
    imPnfRecLen = Len(tmPnf)
    hmCef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.Btr)", Persnnel
    On Error GoTo 0
    imCefRecLen = Len(tmCef)
    hmCxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cxf.Btr)", Persnnel
    On Error GoTo 0
    imCxfRecLen = Len(tmCxf)
'    hmDsf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dsf.Btr)", Persnnel
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm Persnnel
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
    flTextHeight = pbcPersonnel.TextHeight("1") - 35
    'Position panel and picture areas with panel
    If smBuyerOrPayableFlag = "B" Then
        pbcPersonnel.Height = 2790
    Else
        pbcPersonnel.Height = 1410
        cmcDone.Top = cmcDone.Top - 1380
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        cmcErase.Top = cmcDone.Top
        Persnnel.Height = Persnnel.Height - 1380
    End If
    plcPersonnel.Move 390, 765, pbcPersonnel.Width + fgPanelAdj, pbcPersonnel.Height + fgPanelAdj
    pbcPersonnel.Move plcPersonnel.Left + fgBevelX, plcPersonnel.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
    'Title
    gSetCtrl tmCtrls(TITLEINDEX), 2850, tmCtrls(NAMEINDEX).fBoxY, 2805, fgBoxStH
    tmCtrls(TITLEINDEX).iReq = False
    'Phone
    gSetCtrl tmCtrls(PHONEINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    tmCtrls(PHONEINDEX).iReq = False
    'Fax
    gSetCtrl tmCtrls(FAXINDEX), 2850, tmCtrls(PHONEINDEX).fBoxY, 1725, fgBoxStH
    tmCtrls(FAXINDEX).iReq = False
    'State
    gSetCtrl tmCtrls(STATEINDEX), 4590, tmCtrls(PHONEINDEX).fBoxY, 1065, fgBoxStH
    'E-Mail
    gSetCtrl tmCtrls(EMAILINDEX), 30, tmCtrls(PHONEINDEX).fBoxY + fgStDeltaY, 5625, fgBoxStH
    tmCtrls(EMAILINDEX).iReq = False
    'Comment
    gSetCtrl tmCtrls(COMMENTINDEX), 30, tmCtrls(EMAILINDEX).fBoxY + fgStDeltaY, 5625, fgBoxStH
    tmCtrls(COMMENTINDEX).iReq = False
    'Contract Address
    gSetCtrl tmCtrls(CADDRINDEX), 30, tmCtrls(COMMENTINDEX).fBoxY + fgStDeltaY, 5625, fgBoxStH
    tmCtrls(CADDRINDEX).iReq = False
    gSetCtrl tmCtrls(CADDRINDEX + 1), 30, tmCtrls(CADDRINDEX).fBoxY + flTextHeight, tmCtrls(CADDRINDEX).fBoxW, flTextHeight
    tmCtrls(CADDRINDEX + 1).iReq = False
    gSetCtrl tmCtrls(CADDRINDEX + 2), 30, tmCtrls(CADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(CADDRINDEX).fBoxW, flTextHeight
    tmCtrls(CADDRINDEX + 2).iReq = False
    'Billing Address
    gSetCtrl tmCtrls(BADDRINDEX), 30, tmCtrls(CADDRINDEX).fBoxY + fgAddDeltaY, 5625, fgBoxStH
    tmCtrls(BADDRINDEX).iReq = False
    gSetCtrl tmCtrls(BADDRINDEX + 1), 30, tmCtrls(BADDRINDEX).fBoxY + flTextHeight, tmCtrls(BADDRINDEX).fBoxW, flTextHeight
    tmCtrls(BADDRINDEX + 1).iReq = False
    gSetCtrl tmCtrls(BADDRINDEX + 2), 30, tmCtrls(BADDRINDEX + 1).fBoxY + flTextHeight, tmCtrls(BADDRINDEX).fBoxW, flTextHeight
    tmCtrls(BADDRINDEX + 2).iReq = False
End Sub
Private Sub mkcFax_Change()
    mSetChg FAXINDEX
End Sub
Private Sub mkcFax_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub mkcPhone_Change()
    mSetChg PHONEINDEX
End Sub
Private Sub mkcPhone_GotFocus()
    gCtrlGotFocus ActiveControl
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

    If imAdvtOrAgyFlag = 0 Then
        tmPnf.iAdfCode = imPassedAdvtOrAgyCode
        tmPnf.iAgfCode = 0
    Else
        tmPnf.iAgfCode = imPassedAdvtOrAgyCode
        tmPnf.iAdfCode = 0
    End If
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmPnf.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(TITLEINDEX).iChg Then
        tmPnf.sTitle = edcTitle.Text
    End If
    If Not ilTestChg Or tmCtrls(PHONEINDEX).iChg Then
        gGetPhoneNo mkcPhone, tmPnf.sPhone
    End If
    If Not ilTestChg Or tmCtrls(FAXINDEX).iChg Then
        gGetPhoneNo mkcFax, tmPnf.sFax
    End If
    If smBuyerOrPayableFlag = "B" Then
        For ilLoop = 0 To 2 Step 1
            If Not ilTestChg Or tmCtrls(CADDRINDEX + ilLoop).iChg Then
                tmPnf.sCntrAddr(ilLoop) = edcCAddr(ilLoop).Text
            End If
        Next ilLoop
        For ilLoop = 0 To 2 Step 1
            If Not ilTestChg Or tmCtrls(BADDRINDEX + ilLoop).iChg Then
                tmPnf.sBillAddr(ilLoop) = edcBAddr(ilLoop).Text
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To 2 Step 1
            tmPnf.sCntrAddr(ilLoop) = ""
        Next ilLoop
        For ilLoop = 0 To 2 Step 1
            tmPnf.sBillAddr(ilLoop) = ""
        Next ilLoop
    End If
    If Not ilTestChg Or tmCtrls(EMAILINDEX).iChg Then
        'tmCef.iStrLen = Len(edcEMail.Text)
        tmCef.sComment = Trim$(edcEMail.Text) & Chr$(0) '& Chr$(0) 'sgTB
    End If
    If Not ilTestChg Or tmCtrls(COMMENTINDEX).iChg Then
        'tmCxf.iStrLen = Len(edcComment.Text)
        tmCxf.sComment = Trim$(edcComment.Text) & Chr$(0) '& Chr$(0) 'sgTB
    End If
    If smBuyerOrPayableFlag = "B" Then
        tmPnf.sType = "B"
    Else
        tmPnf.sType = "P"
    End If
    If imState = 1 Then
        tmPnf.sState = "D"
    Else
        tmPnf.sState = "A"
    End If
    Exit Sub

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
    edcName.Text = Trim$(tmPnf.sName)
    edcTitle.Text = Trim$(tmPnf.sTitle)
    gSetPhoneNo tmPnf.sPhone, mkcPhone
    gSetPhoneNo tmPnf.sFax, mkcFax
    Select Case tmPnf.sState
        Case "A"
            imState = 0
        Case "D"
            imState = 1
        Case Else
            imState = -1
    End Select
    If smBuyerOrPayableFlag = "B" Then
        For ilLoop = 0 To 2 Step 1
            edcCAddr(ilLoop).Text = Trim$(tmPnf.sCntrAddr(ilLoop))
        Next ilLoop
        For ilLoop = 0 To 2 Step 1
            edcBAddr(ilLoop).Text = Trim$(tmPnf.sBillAddr(ilLoop))
        Next ilLoop
    Else
        For ilLoop = 0 To 2 Step 1
            edcCAddr(ilLoop).Text = ""
        Next ilLoop
        For ilLoop = 0 To 2 Step 1
            edcBAddr(ilLoop).Text = ""
        Next ilLoop
    End If
    'If tmCef.iStrLen > 0 Then
    '    edcEMail.Text = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
    'Else
    '    edcEMail.Text = ""
    'End If
    edcEMail.Text = gStripChr0(tmCef.sComment)
    smEMail = edcEMail.Text
    'If tmCxf.iStrLen > 0 Then
    '    edcComment.Text = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
    'Else
    '    edcComment.Text = ""
    'End If
    edcComment.Text = gStripChr0(tmCxf.sComment)
    smComment = edcComment.Text
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
                    MsgBox "Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmPnf.sName) 'Reset text
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
    'gInitStdAlone Persnnel, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igPersonnelCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igPersonnelCallSource = 21 'CALLNONE
    '    'slCommand = "Advt^Prod\Guide\22\Advt\29\B\Kathy Gilbert"
    '    'slCommand = "Agency^Prod^NOHELP\Guide\21\Agy\0\B\"
    '    slCommand = "Traffic^Test\Counterpoint\2\Agy\7\B\Bruce Heim"
    'End If
    If igPersonnelCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            If StrComp("Advt", Trim$(slStr), 1) = 0 Then
                imAdvtOrAgyFlag = 0
            Else
                imAdvtOrAgyFlag = 1
            End If
            ilRet = gParseItem(slCommand, 5, "\", slStr)
            If ilRet = CP_MSG_NONE Then
                imPassedAdvtOrAgyCode = Val(slStr)
                ilRet = gParseItem(slCommand, 6, "\", slStr)
                If ilRet = CP_MSG_NONE Then
                    smBuyerOrPayableFlag = slStr
                    ilRet = gParseItem(slCommand, 7, "\", slStr)
                    If ilRet = CP_MSG_NONE Then
                        smInitPersonnel = slStr
                    Else
                        smInitPersonnel = ""
                    End If
                Else
                    smBuyerOrPayableFlag = ""
                    smInitPersonnel = ""
                End If
            Else
                imPassedAdvtOrAgyCode = -1
                smBuyerOrPayableFlag = ""
                smInitPersonnel = ""
            End If
        Else
            imAdvtOrAgyFlag = -1
            imPassedAdvtOrAgyCode = -1
            smBuyerOrPayableFlag = ""
            smInitPersonnel = ""
        End If
    Else
        imAdvtOrAgyFlag = 1
        smBuyerOrPayableFlag = "B"
        imPassedAdvtOrAgyCode = -1
        smInitPersonnel = ""
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
    Dim ilCode As Integer
    Dim ilFlag As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    imPopReqd = False
    ilFlag = imAdvtOrAgyFlag
    ilCode = imPassedAdvtOrAgyCode
    'If ilCode > 0 Then
        'ilRet = gPopPersonnelBox(Persnnel, ilFlag, ilCode, smBuyerOrPayableFlag, True, 0, cbcSelect, lbcPersonnelCode)
        ilRet = gPopPersonnelBox(Persnnel, ilFlag, ilCode, smBuyerOrPayableFlag, True, 0, cbcSelect, tgNameCode(), sgNameCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", Persnnel
            On Error GoTo 0
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
            imPopReqd = True
        End If
    'Else
    '    If cbcSelect.List(0) <> "[New]" Then
    '        cbcSelect.AddItem "[New]", 0  'Force as first item on list
    '        ReDim tgNameCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    '    End If
    '    imPopReqd = True
    'End If
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
    slNameCode = tgNameCode(ilSelectIndex - 1).sKey    'lbcPersonnelCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", Persnnel
    On Error GoTo 0
    smReturnCode = slCode
    tmPnfSrchKey.iCode = CLng(slCode)
    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Personnel)", Persnnel
    On Error GoTo 0
    tmCefSrchKey.lCode = tmPnf.lEMailCefCode
    If tmPnf.lEMailCefCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '5013
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mSaveRec (btrGetEqual:E-Mail)", Persnnel
        On Error GoTo 0
    Else
        tmCef.lCode = 0
        'tmCef.iStrLen = 0
        tmCef.sComment = ""
    End If
    tmCxfSrchKey.lCode = tmPnf.lCxfCode
    If tmPnf.lCxfCode <> 0 Then
        tmCxf.sComment = ""
        imCxfRecLen = Len(tmCxf)    '5013
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mSaveRec (btrGetEqual:Comment)", Persnnel
        On Error GoTo 0
    Else
        tmCxf.lCode = 0
        'tmCxf.iStrLen = 0
        tmCxf.sComment = ""
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCefState                                                                            *
'******************************************************************************************

'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim llRecPos As Long
    Dim ilRetC As Integer
    Dim ilCxfState As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlPnf As PNF
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
    ilCxfState = 0
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Pnf.Btr")
        'If Len(lbcPersonnelCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcPersonnelCode.Tag, Len(lbcPersonnelCode.Tag) - Len(slStamp))
        'End If
        If Len(sgNameCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(sgNameCodeTag, Len(sgNameCodeTag) - Len(slStamp))
        End If
        If imSelectedIndex <> 0 Then
            'Reread record in so lastest is obtained
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        imCefRecLen = Len(tmCef) '- Len(tmCef.sComment) + Len(Trim$(tmCef.sComment)) ' + 2 '5 = fixed record length; 2 is the length of the record which is part of the variable record
        If imSelectedIndex = 0 Then 'New selected
            'If Len(Trim$(tmCef.sComment)) > 2 Then '-2 so the control character at the end is not counted
            If gStripChr0(tmCef.sComment) <> "" Then
                tmCef.lCode = 0 'Autoincrement
                ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
            Else
                tmCef.lCode = 0
                ilRet = BTRV_ERR_NONE
            End If
            slMsg = "mSaveRec (btrInsert: E-Mail)"
        Else 'Old record-Update
            'If Len(Trim$(tmCef.sComment)) > 2 Then '-2 so the control character at the end is not counted
            If gStripChr0(tmCef.sComment) <> "" Then
                If tmCef.lCode = 0 Then
                    tmCef.lCode = 0 'Autoincrement
                    ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                Else
                    ilRet = btrUpdate(hmCef, tmCef, imCefRecLen)
                End If
            Else
                If tmPnf.lEMailCefCode <> 0 Then
                    ilRet = btrDelete(hmCef)
                End If
                tmCef.lCode = 0
            End If
            slMsg = "mSaveRec (btrUpdate: E-Mail)"
        End If
        If (ilRet <> BTRV_ERR_CONFLICT) And (ilRet <> BTRV_ERR_NONE) Then
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Persnnel
            On Error GoTo 0
        End If
        imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment)) ' + 2 '5 = fixed record length; 2 is the length of the record which is part of the variable record
        If imSelectedIndex = 0 Then 'New selected
            If Len(Trim$(tmCxf.sComment)) > 2 Then '-2 so the control character at the end is not counted
                ilCxfState = 1
                tmCxf.lCode = 0 'Autoincrement
                tmCxf.sComType = "N"    'Personnel
                tmCxf.sShProp = "N"
                tmCxf.sShOrder = "N"
                tmCxf.sShSpot = "N"
                tmCxf.sShInv = "N"
                tmCxf.sShInsertion = "N"
                tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                tmCxf.lAutoCode = tmCxf.lCode
                ilRet = btrInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
            Else
                tmCxf.lCode = 0
                ilRet = BTRV_ERR_NONE
            End If
            slMsg = "mSaveRec (btrInsert: Comment)"
        Else 'Old record-Update
            If Len(Trim$(tmCxf.sComment)) > 2 Then '-2 so the control character at the end is not counted
                If tmCxf.lCode = 0 Then
                    ilCxfState = 1
                    tmCxf.lCode = 0 'Autoincrement
                    tmCxf.sComType = "N"    'Personnel
                    tmCxf.sShProp = "N"
                    tmCxf.sShOrder = "N"
                    tmCxf.sShSpot = "N"
                    tmCxf.sShInv = "N"
                    tmCxf.sShInsertion = "N"
                    tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmCxf.lAutoCode = tmCxf.lCode
                    ilRet = btrInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                Else
                    ilCxfState = 2
                    tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                    gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                    gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                    ilRet = btrUpdate(hmCxf, tmCxf, imCxfRecLen)
                End If
            Else
                If tmPnf.lCxfCode <> 0 Then
                    ilCxfState = 3
                    ilRet = btrDelete(hmCxf)
                End If
                tmCxf.lCode = 0
            End If
            slMsg = "mSaveRec (btrUpdate: Comment)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Persnnel
    On Error GoTo 0
    tmPnf.lEMailCefCode = tmCef.lCode
    tmPnf.lCxfCode = tmCxf.lCode
    If ilCxfState = 1 Then
        Do
            'tmCxfSrchKey.lCode = tmCxf.lCode
            'imCxfRecLen = Len(tmCxf)
            'ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Personnel)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, Persnnel
            'On Error GoTo 0
            tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
            tmCxf.lAutoCode = tmCxf.lCode
            tmCxf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
            gPackTime slSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
            imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
            ilRet = btrUpdate(hmCxf, tmCxf, imCxfRecLen)
            slMsg = "mSaveRec (btrUpdate:Personnel)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Persnnel
        On Error GoTo 0
    ElseIf ilCxfState = 3 Then
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "CXF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmCxf.iRemoteID
'            tmDsf.lAutoCode = tmCxf.lAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'        End If
    End If
    Do  'Loop until record updated or added
        If imSelectedIndex = 0 Then 'New selected
            smNewFlag = "Y"
            tmPnf.iCode = 0
            tmPnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmPnf.iAutoCode = tmPnf.iCode
            ilRet = btrInsert(hmPnf, tmPnf, imRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Personnel)"
        Else 'Old record-Update
            tmPnf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmPnf.iSyncDate(0), tmPnf.iSyncDate(1)
            gPackTime slSyncTime, tmPnf.iSyncTime(0), tmPnf.iSyncTime(1)
            ilRet = btrUpdate(hmPnf, tmPnf, imRecLen)
            slMsg = "mSaveRec (btrUpdate: Personnel)"
            If ilRet = BTRV_ERR_CONFLICT Then
                tmPnfSrchKey.iCode = tmPnf.iCode
                ilRetC = btrGetEqual(hmPnf, tlPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Persnnel
    On Error GoTo 0
    ilRet = btrGetPosition(hmPnf, llRecPos)
    If imSelectedIndex = 0 Then 'New selected
        Do
            'tmPnfSrchKey.iCode = tmPnf.iCode
            'ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Personnel)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, Persnnel
            'On Error GoTo 0
            tmPnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmPnf.iAutoCode = tmPnf.iCode
            tmPnf.iSourceID = tgUrf(0).iRemoteUserID
            gPackDate slSyncDate, tmPnf.iSyncDate(0), tmPnf.iSyncDate(1)
            gPackTime slSyncTime, tmPnf.iSyncTime(0), tmPnf.iSyncTime(1)
            ilRet = btrUpdate(hmPnf, tmPnf, imPnfRecLen)
            slMsg = "mSaveRec (btrUpdate:Personnel)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Persnnel
        On Error GoTo 0
    End If
    smReturnCode = Trim$(str$(tmPnf.iCode))
'    'If lbcPersonnelCode.Tag <> "" Then
'    '    If slStamp = lbcPersonnelCode.Tag Then
'    '        lbcPersonnelCode.Tag = FileDateTime(sgDBPath & "Pnf.Btr")
'    '        If Len(slStamp) > Len(lbcPersonnelCode.Tag) Then
'    '            lbcPersonnelCode.Tag = lbcPersonnelCode.Tag & Right$(slStamp, Len(slStamp) - Len(lbcPersonnelCode.Tag))
'    '        End If
'    '    End If
'    'End If
'    If sgNameCodeTag <> "" Then
'        If slStamp = sgNameCodeTag Then
'            sgNameCodeTag = FileDateTime(sgDBPath & "Pnf.Btr")
'            If Len(slStamp) > Len(sgNameCodeTag) Then
'                sgNameCodeTag = sgNameCodeTag & right$(slStamp, Len(slStamp) - Len(sgNameCodeTag))
'            End If
'        End If
'    End If
'    If imSelectedIndex <> 0 Then
'        'lbcPersonnelCode.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
'        cbcSelect.RemoveItem imSelectedIndex
'    End If
'    cbcSelect.RemoveItem 0 'Remove [New]
'    slName = RTrim$(tmPnf.sName)
'    cbcSelect.AddItem slName
'    slName = tmPnf.sName + "\" + LTrim$(Str$(tmPnf.iCode)) 'slName + "\" + LTrim$(Str$(tmPnf.lCode))
'    'lbcPersonnelCode.AddItem slName
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
                    pbcPersonnel_Paint
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
            gSetChgFlag tmPnf.sName, edcName, tmCtrls(ilBoxNo)
        Case TITLEINDEX 'Name
            gSetChgFlag tmPnf.sTitle, edcTitle, tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            gSetChgFlag tmPnf.sPhone, mkcPhone, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            gSetChgFlag tmPnf.sFax, mkcFax, tmCtrls(ilBoxNo)
        Case STATEINDEX
        Case CADDRINDEX 'Contract address
            gSetChgFlag tmPnf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case CADDRINDEX + 1 'Contract address
            gSetChgFlag tmPnf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case CADDRINDEX + 2 'Contract address
            gSetChgFlag tmPnf.sCntrAddr(ilBoxNo - CADDRINDEX), edcCAddr(ilBoxNo - CADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX 'Billing address
            gSetChgFlag tmPnf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX + 1 'Billing address
            gSetChgFlag tmPnf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case BADDRINDEX + 2 'Billing address
            gSetChgFlag tmPnf.sBillAddr(ilBoxNo - BADDRINDEX), edcBAddr(ilBoxNo - BADDRINDEX), tmCtrls(ilBoxNo)
        Case EMAILINDEX
            gSetChgFlag smEMail, edcEMail, tmCtrls(ilBoxNo)
        Case COMMENTINDEX   'Comment
            gSetChgFlag smComment, edcComment, tmCtrls(ilBoxNo)
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
    If (cbcSelect.Text = "") Then
        cbcSelect.Enabled = True
        pbcPersonnel.Enabled = False  'Disallow mouse
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    Else
        If imAdvtOrAgyFlag = 0 Then
            If (igWinStatus(ADVERTISERSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcPersonnel.Enabled = False  'Disallow mouse
                pbcSTab.Enabled = False
                pbcTab.Enabled = False
            Else
                pbcPersonnel.Enabled = True
                pbcSTab.Enabled = True
                pbcTab.Enabled = True
            End If
        Else
            If (igWinStatus(AGENCIESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcPersonnel.Enabled = False  'Disallow mouse
                pbcSTab.Enabled = False
                pbcTab.Enabled = False
            Else
                pbcPersonnel.Enabled = True
                pbcSTab.Enabled = True
                pbcTab.Enabled = True
            End If
        End If
        If Not ilAltered Then
            cbcSelect.Enabled = True
        Else
            cbcSelect.Enabled = False
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    'If imPassedAdvtOrAgyCode <> 0 Then
    If (igPersonnelCallSource = CALLNONE) Or (imPassedAdvtOrAgyCode > 0) Then
        If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
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
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus                      *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case TITLEINDEX 'Name
            edcTitle.SetFocus
        Case PHONEINDEX 'Phone number plus extension
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax number
            mkcFax.SetFocus
        Case STATEINDEX   'Print style
            pbcState.SetFocus
        Case CADDRINDEX 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).SetFocus
        Case BADDRINDEX 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).SetFocus
        Case EMAILINDEX 'E-Mail
            edcEMail.SetFocus
        Case COMMENTINDEX 'Comment
            edcComment.SetFocus
    End Select
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
    Dim ilPos As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
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
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case TITLEINDEX 'Name
            edcTitle.Visible = False  'Set visibility
            slStr = edcTitle.Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            mkcPhone.Visible = False  'Set visibility
            If mkcPhone.Text = smPhoneImage Then
                slStr = ""
            Else
                slStr = mkcPhone.Text
            End If
            If slStr <> "" Then
                If InStr(slStr, "(____)") <> 0 Then
                    ilPos = InStr(slStr, "Ext(")
                    slStr = Left$(slStr, ilPos - 1)
                End If
            End If
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            mkcFax.Visible = False  'Set visibility
            If mkcFax.Text = smFaxImage Then
                slStr = ""
            Else
                slStr = mkcFax.Text
            End If
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case STATEINDEX   'Print style
            pbcState.Visible = False  'Set visibility
            If imState = 0 Then
                slStr = "Active"
            ElseIf imState = 1 Then
                slStr = "Dormant"
            Else
                slStr = ""
            End If
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX + 1 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case CADDRINDEX + 2 'Address
            edcCAddr(ilBoxNo - CADDRINDEX).Visible = False
            slStr = edcCAddr(ilBoxNo - CADDRINDEX).Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX + 1 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case BADDRINDEX + 2 'Address
            edcBAddr(ilBoxNo - BADDRINDEX).Visible = False
            slStr = edcBAddr(ilBoxNo - BADDRINDEX).Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case EMAILINDEX 'Address ID
            edcEMail.Visible = False  'Set visibility
            slStr = edcEMail.Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
        Case COMMENTINDEX 'Comment
            edcComment.Visible = False  'Set visibility
            slStr = edcComment.Text
            gSetShow pbcPersonnel, slStr, tmCtrls(ilBoxNo)
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

    sgDoneMsg = Trim$(str$(igPersonnelCallSource)) & "\" & sgPersonnelName & "\" & smNewFlag & "\" & smReturnCode
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload Persnnel
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
    If (ilCtrlNo = TITLEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcTitle, "", "Title must be specified", tmCtrls(TITLEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TITLEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PHONEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(mkcPhone, smPhoneImage, "Phone # must be specified", tmCtrls(PHONEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PHONEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = FAXINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(mkcFax, smFaxImage, "Fax # must be specified", tmCtrls(FAXINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = FAXINDEX
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
        If gFieldDefinedStr(slStr, "", "Active/Dormant must be specified", tmCtrls(STATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcCAddr(0), "", "Contract Address must be specified", tmCtrls(CADDRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CADDRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = BADDRINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcBAddr(0), "", "Billing Address must be specified", tmCtrls(BADDRINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = BADDRINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = EMAILINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcComment, "", "E-Mail must be specified", tmCtrls(EMAILINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = EMAILINDEX
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
Private Sub pbcPersonnel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If (cbcSelect.ListIndex <> imSelectedIndex) Then
        If cbcSelect.Enabled Then
            cbcSelect_Change
            cbcSelect.SetFocus
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If ilBox <= imMaxIndex Then
                    mSetShow imBoxNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcPersonnel_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcPersonnel.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcPersonnel.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcPersonnel.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    'Handle double names- if drop down selected the index is changed to the
    'first name without any events- forces back so change occurs
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If (cbcSelect.ListIndex <> imSelectedIndex) And (imSelectedIndex <> 0) Then
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
            ilBox = NAMEINDEX
            mSetCommands
        Case NAMEINDEX 'Name
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
Private Sub pbcState_GotFocus()
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
        If imState = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imState = 1
            pbcState_Paint
        ElseIf imState = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imState = 0
            pbcState_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imState = 0 Then
        tmCtrls(imBoxNo).iChg = True
        imState = 1
    ElseIf imState = 1 Then
        tmCtrls(imBoxNo).iChg = True
        imState = 0
    End If
    pbcState_Paint
    mSetCommands
End Sub
Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    If imState = 0 Then
        pbcState.Print "Active"
    ElseIf imState = 1 Then
        pbcState.Print "Dormant"
    Else
        pbcState.Print "   "
    End If
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
        Case imMaxIndex 'COMMENTINDEX
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igPersonnelCallSource = CALLNONE) Then
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
Private Sub plcPersonnel_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If cbcSelect.ListIndex <> imSelectedIndex Then
        cbcSelect_Change
        'cbcSelect.SetFocus
        Exit Sub
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Personnel"
End Sub
Private Sub edcCAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcCAddr_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcBAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcBAddr_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
    If Index > 0 Then
        If (edcBAddr(0).Text = "") Then
            edcBAddr(1).Text = ""
            edcBAddr(2).Text = ""
            pbcTab.SetFocus
        End If
    End If
End Sub

