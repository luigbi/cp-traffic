VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form SOffice 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   840
   ClientTop       =   1635
   ClientWidth     =   5325
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
   ScaleHeight     =   3765
   ScaleWidth      =   5325
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   210
      Width           =   3180
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3300
      ScaleHeight     =   210
      ScaleWidth      =   1470
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2595
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4665
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2970
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3600
      Width           =   105
   End
   Begin VB.ListBox lbcSRegion 
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
      Left            =   495
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4365
      Top             =   2820
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   495
      TabIndex        =   12
      Tag             =   "The number and extension of the buyer."
      Top             =   2295
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
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
   Begin MSMask.MaskEdBox mkcFax 
      Height          =   210
      Left            =   540
      TabIndex        =   13
      Tag             =   "The number and extension of the buyer."
      Top             =   2580
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3300
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   885
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
      Left            =   4320
      Picture         =   "Soffice.frx":0000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcSSource 
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
      Left            =   1665
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   900
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1815
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   795
      MaxLength       =   25
      TabIndex        =   10
      Top             =   1725
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   675
      MaxLength       =   25
      TabIndex        =   11
      Top             =   1650
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcMktRank 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3345
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1245
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   525
      MaxLength       =   20
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   2715
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
      Left            =   3255
      TabIndex        =   22
      Top             =   3345
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
      Left            =   4635
      TabIndex        =   21
      Top             =   2895
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
      Left            =   2130
      TabIndex        =   20
      Top             =   3345
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
      Left            =   1005
      TabIndex        =   19
      Top             =   3345
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
      Left            =   3255
      TabIndex        =   18
      Top             =   3000
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
      Left            =   2130
      TabIndex        =   17
      Top             =   3000
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
      Left            =   1005
      TabIndex        =   16
      Top             =   3000
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
      Height          =   2100
      Left            =   465
      Picture         =   "Soffice.frx":00FA
      ScaleHeight     =   2100
      ScaleWidth      =   4335
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   735
      Width           =   4335
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   765
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
      TabIndex        =   15
      Top             =   1695
      Width           =   45
   End
   Begin VB.PictureBox plcSoff 
      ForeColor       =   &H00000000&
      Height          =   2220
      Left            =   420
      ScaleHeight     =   2160
      ScaleWidth      =   4380
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   675
      Width           =   4440
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4710
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4710
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label plcScreen 
      Caption         =   "Sales Offices"
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
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   1395
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3195
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "SOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Soffice.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SOffice.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Sales Office input screen code
Option Explicit
Option Compare Text
'Sales Office Field Areas
Dim tmCtrls(0 To 10)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmSRCode() As SORTCODE
Dim smSRCodeTag As String
Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String
Dim tmSSCode() As SORTCODE
Dim smSSCodeTag As String
Dim imBoxNo As Integer   'Current Sales Office Box
Dim imState As Integer  '0=Active; 1=Dormant
Dim tmSof As SOF        'SOF record image
Dim tmSofSrchKey As INTKEY0    'SOF key record image
Dim imSofRecLen As Integer        'SOF record length
Dim imUpdateAllowed As Integer    'User can update records
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmSof As Integer 'Sales Office file handle
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smPhoneImage As String  'Blank phone image- obtained from mkcPhone.text before input
Dim smFaxImage As String    'Blank fax image
Dim smSSource As String     'Sales source name, saved to determine if same changed
Dim smSRegion As String     'Sales region name, saved to determine if same changed
Dim imFirstActivate As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Const NAMEINDEX = 1     'Name control/field
Const SSOURCEINDEX = 2  'Sales Source control/field
Const SREGIONINDEX = 3  'Sales Region control/field
Const MKTRANKINDEX = 4  'Market Rank control/field
Const ADDRESSINDEX = 5  'Address control/field
Const PHONEINDEX = 8    'Phone/extension control/field
Const FAXINDEX = 9      'Fax control/field
Const STATEINDEX = 10
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
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcSoff_Paint
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
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If igSofCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgSofName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgSofName   'New name
            End If
            cbcSelect_Change
            If sgSofName <> "" Then
                mSetCommands
                gFindMatch sgSofName, 1, cbcSelect
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
    If igSofCallSource <> CALLNONE Then
        igSofCallSource = CALLCANCELLED
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
    If igSofCallSource <> CALLNONE Then
        sgSofName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgSofName = "[New]"
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
    If igSofCallSource <> CALLNONE Then
        If sgSofName = "[New]" Then
            igSofCallSource = CALLCANCELLED
        Else
            igSofCallSource = CALLDONE
            If lbcSSource.ListIndex >= 1 Then
                sgSofName = Trim$(edcName.Text) & "/" & Trim$(lbcSSource.List(lbcSSource.ListIndex))
            Else
                sgSofName = Trim$(edcName.Text) 'Save name for returning
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
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case SSOURCEINDEX
            lbcSSource.Visible = Not lbcSSource.Visible
        Case SREGIONINDEX
            lbcSRegion.Visible = Not lbcSRegion.Visible
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
    If imSelectedIndex > 0 Then
        'Check that record is not referenced-Code missing
        Screen.MousePointer = vbHourglass
        ilRet = gIICodeRefExist(SOffice, tmSof.iCode, "Bvf.Btr", "BvfSofCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Budget by Veghicle references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SOffice, tmSof.iCode, "Slf.Btr", "SlfSofCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Salesperson references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(SOffice, tmSof.iCode, "Pjf.Btr", "PjfSofCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract Projection references this name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmSof.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        slStamp = gFileDateTime(sgDBPath & "Sof.btr")
        ilRet = btrDelete(hmSof)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", SOffice
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcNameCode.Tag <> "" Then
        '    If slStamp = lbcNameCode.Tag Then
        '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Sof.btr")
        '    End If
        'End If
        If smNameCodeTag <> "" Then
            If slStamp = smNameCodeTag Then
                smNameCodeTag = gFileDateTime(sgDBPath & "Sof.btr")
            End If
        End If
        'lbcNameCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tmNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcSoff.Cls
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
Private Sub cmcReport_Click()
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = SALESOFFICESLIST
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SOffice^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "SOffice^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SOffice^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "SOffice^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptNoSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'SOffice.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'SOffice.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    ''Screen.MousePointer = vbDefault    'Default
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
        pbcSoff.Cls
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcSoff_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcSoff.Cls
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
    'Must reset display so altered flag is cleared and setcommand will turn select on
'    If imSvSelectedIndex <> 0 Then
'        cbcSelect.Text = slName
'    Else
'        cbcSelect.ListIndex = 0
'    End If
'    cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmSof.iCode
    cbcSelect.Clear
    smNameCodeTag = ""
    mPopulate
    If imSvSelectedIndex <> 0 Then
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
Private Sub edcAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcAddr_GotFocus(Index As Integer)
    gCtrlGotFocus edcAddr(Index)
End Sub
Private Sub edcAddr_KeyPress(Index As Integer, KeyAscii As Integer)
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
        Case SSOURCEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSSource, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSSource.ListIndex = 0
            End If
        Case SREGIONINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcSRegion, imBSMode, slStr)
            If ilRet = 1 Then
                lbcSRegion.ListIndex = 0
            End If
    End Select
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case SSOURCEINDEX
            If lbcSSource.ListCount = 1 Then
                lbcSSource.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case SREGIONINDEX
'            If lbcSSource.ListCount = 1 Then
'                lbcSSource.ListIndex = 0
'                If imTabDirection = -1 Then  'Right To Left
'                    pbcSTab.SetFocus
'                Else
'                    pbcTab.SetFocus
'                End If
'                Exit Sub
'            End If
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
            Case SSOURCEINDEX
                gProcessArrowKey Shift, KeyCode, lbcSSource, imLbcArrowSetting
            Case SREGIONINDEX
                gProcessArrowKey Shift, KeyCode, lbcSRegion, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
            Case SSOURCEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case SREGIONINDEX
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
Private Sub edcMktRank_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcMktRank_GotFocus()
    gCtrlGotFocus edcMktRank
End Sub
Private Sub edcMktRank_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
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
    slStr = edcMktRank.Text
    slStr = Left$(slStr, edcMktRank.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcMktRank.SelStart - edcMktRank.SelLength)
    If gCompNumberStr(slStr, "9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
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
    If (igWinStatus(SALESOFFICESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
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
    SOffice.Refresh
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
    Erase tmSRCode
    Erase tmNameCode
    Erase tmSSCode
    btrExtClear hmSof   'Clear any previous extend operation
    ilRet = btrClose(hmSof)
    btrDestroy hmSof
    
    Set SOffice = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSRegion_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSRegion, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSRegion_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcSRegion_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSRegion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSRegion_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSRegion, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcSSource_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSSource, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcSSource_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcSSource_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcSSource_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcSSource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSSource, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
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
    lbcSSource.ListIndex = -1
    lbcSRegion.ListIndex = -1
    smSSource = ""
    smSRegion = ""
    edcMktRank.Text = ""
    For ilLoop = 0 To 2 Step 1
        edcAddr(ilLoop).Text = ""
    Next ilLoop
    mkcPhone.Text = smPhoneImage
    mkcFax.Text = smFaxImage
    imState = -1
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
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
            edcName.MaxLength = 20
            gMoveFormCtrl pbcSoff, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case SSOURCEINDEX   'Sales Source
            mSSourcePop
            If imTerminate Then
                Exit Sub
            End If
            lbcSSource.Height = gListBoxHeight(lbcSSource.ListCount, 4)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcSoff, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcSSource.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcSSource.ListIndex < 0 Then
                If lbcSSource.ListCount <= 1 Then
                    lbcSSource.ListIndex = 0   '[New]
                Else
                    lbcSSource.ListIndex = 1
                End If
            End If
            If lbcSSource.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcSSource.List(lbcSSource.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case SREGIONINDEX   'Sales Region
            mSRegionPop
            If imTerminate Then
                Exit Sub
            End If
            lbcSRegion.Height = gListBoxHeight(lbcSRegion.ListCount, 4)
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 20
            gMoveFormCtrl pbcSoff, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcSRegion.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            imChgMode = True
            If lbcSRegion.ListIndex < 0 Then
                lbcSRegion.ListIndex = 1   '[None]
            End If
            If lbcSRegion.ListIndex < 0 Then
                edcDropDown.Text = ""
            Else
                edcDropDown.Text = lbcSRegion.List(lbcSRegion.ListIndex)
            End If
            imChgMode = False
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MKTRANKINDEX 'Market Rank
            edcMktRank.Width = tmCtrls(ilBoxNo).fBoxW
            edcMktRank.MaxLength = 4
            gMoveFormCtrl pbcSoff, edcMktRank, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcMktRank.Visible = True  'Set visibility
            edcMktRank.SetFocus
        Case ADDRESSINDEX 'Address
            edcAddr(0).Width = tmCtrls(ilBoxNo).fBoxW
            edcAddr(0).MaxLength = 25
            gMoveFormCtrl pbcSoff, edcAddr(0), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(0).Visible = True  'Set visibility
            edcAddr(0).SetFocus
        Case ADDRESSINDEX + 1 'Address
            edcAddr(1).Width = tmCtrls(ilBoxNo).fBoxW
            edcAddr(1).MaxLength = 25
            gMoveFormCtrl pbcSoff, edcAddr(1), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(1).Visible = True  'Set visibility
            edcAddr(1).SetFocus
        Case ADDRESSINDEX + 2 'Address
            edcAddr(2).Width = tmCtrls(ilBoxNo).fBoxW
            edcAddr(2).MaxLength = 25
            gMoveFormCtrl pbcSoff, edcAddr(2), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(2).Visible = True  'Set visibility
            edcAddr(2).SetFocus
        Case PHONEINDEX 'Phone and extension
            mkcPhone.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSoff, mkcPhone, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcPhone.Visible = True  'Set visibility
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax
            mkcFax.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcSoff, mkcFax, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcFax.Visible = True  'Set visibility
            mkcFax.SetFocus
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
'*             Created:4/22/93       By:D. LeVine      *
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
'    mInitDDE
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    SOffice.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    gCenterStdAlone SOffice
    'sOffice.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imPopReqd = False
    imFirstFocus = True
    imSofRecLen = Len(tmSof)  'Get and save SOFF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    smPhoneImage = mkcPhone.Text
    smFaxImage = mkcFax.Text
    hmSof = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "SOF.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: sof.btr)", SOffice
    On Error GoTo 0
    lbcSSource.Clear 'Force list box to be populated
    mSSourcePop
    If imTerminate Then
        Exit Sub
    End If
    lbcSRegion.Clear 'Force list box to be populated
    mSRegionPop
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm SOffice
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
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
    imLBCtrls = 1
    flTextHeight = pbcSoff.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSoff.Move 420, 675, pbcSoff.Width + fgPanelAdj, pbcSoff.Height + fgPanelAdj
    pbcSoff.Move plcSoff.Left + fgBevelX, plcSoff.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2145, fgBoxStH
    'Sales Source
    gSetCtrl tmCtrls(SSOURCEINDEX), 2190, tmCtrls(NAMEINDEX).fBoxY, 2130, fgBoxStH
    'Sales Region
    gSetCtrl tmCtrls(SREGIONINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2145, fgBoxStH
    tmCtrls(SREGIONINDEX).iReq = False
    'Market Rank
    gSetCtrl tmCtrls(MKTRANKINDEX), 2190, tmCtrls(SREGIONINDEX).fBoxY, 2130, fgBoxStH
    'Address
    gSetCtrl tmCtrls(ADDRESSINDEX), 30, tmCtrls(SREGIONINDEX).fBoxY + fgStDeltaY, 4290, fgBoxStH
    tmCtrls(ADDRESSINDEX).iReq = False
    gSetCtrl tmCtrls(ADDRESSINDEX + 1), 30, tmCtrls(ADDRESSINDEX).fBoxY + flTextHeight, tmCtrls(ADDRESSINDEX).fBoxW, flTextHeight
    tmCtrls(ADDRESSINDEX + 1).iReq = False
    gSetCtrl tmCtrls(ADDRESSINDEX + 2), 30, tmCtrls(ADDRESSINDEX + 1).fBoxY + flTextHeight, tmCtrls(ADDRESSINDEX).fBoxW, flTextHeight
    tmCtrls(ADDRESSINDEX + 2).iReq = False
    'Phone
    gSetCtrl tmCtrls(PHONEINDEX), 30, tmCtrls(ADDRESSINDEX).fBoxY + fgAddDeltaY, 2805, fgBoxStH
    tmCtrls(PHONEINDEX).iReq = False
    'Fax
    gSetCtrl tmCtrls(FAXINDEX), 30, tmCtrls(PHONEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    tmCtrls(FAXINDEX).iReq = False
    'State
    gSetCtrl tmCtrls(STATEINDEX), 2850, tmCtrls(FAXINDEX).fBoxY, 1470, fgBoxStH
End Sub
Private Sub mkcFax_Change()
    mSetChg imBoxNo
End Sub
Private Sub mkcFax_GotFocus()
    gCtrlGotFocus mkcFax
End Sub
Private Sub mkcPhone_Change()
    mSetChg imBoxNo
End Sub
Private Sub mkcPhone_GotFocus()
    gCtrlGotFocus mkcPhone
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
    Dim slSSCode As String  'Sales source name and code
    Dim slSRCode As String  'Sales region name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Sales source code number
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmSof.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(SSOURCEINDEX).iChg Then
        If lbcSSource.ListIndex >= 1 Then
            slSSCode = tmSSCode(lbcSSource.ListIndex - 1).sKey 'lbcSSCode.List(lbcSSource.ListIndex - 1)
            ilRet = gParseItem(slSSCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", SOffice
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmSof.iMnfSSCode = CInt(slCode)
        Else
            tmSof.iMnfSSCode = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(SREGIONINDEX).iChg Then
        If lbcSRegion.ListIndex >= 2 Then
            slSRCode = tmSRCode(lbcSRegion.ListIndex - 2).sKey 'lbcSRCode.List(lbcSRegion.ListIndex - 2)
            ilRet = gParseItem(slSRCode, 2, "\", slCode)
            On Error GoTo mMoveCtrlToRecErr
            gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", SOffice
            On Error GoTo 0
            slCode = Trim$(slCode)
            tmSof.iMnfRegion = CInt(slCode)
        Else
            tmSof.iMnfRegion = 0
        End If
    End If
    If Not ilTestChg Or tmCtrls(MKTRANKINDEX).iChg Then
        tmSof.iMktRank = Val(edcMktRank.Text)
    End If
    For ilLoop = 0 To 2 Step 1
        If Not ilTestChg Or tmCtrls(ADDRESSINDEX + ilLoop).iChg Then
            tmSof.sAddr(ilLoop) = edcAddr(ilLoop).Text
        End If
    Next ilLoop
    If Not ilTestChg Or tmCtrls(PHONEINDEX).iChg Then
        gGetPhoneNo mkcPhone, tmSof.sPhone
    End If
    If Not ilTestChg Or tmCtrls(FAXINDEX).iChg Then
        gGetPhoneNo mkcFax, tmSof.sFax
    End If
    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
        Select Case imState
            Case 0  'Active
                tmSof.sState = "A"
            Case 1  'Dormant
                tmSof.sState = "D"
        End Select
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
    Dim slRecCode As String
    Dim slSSCode As String  'Sales source name and code
    Dim slSRCode As String  'Sales region name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Sales source code number
    edcName.Text = Trim$(tmSof.sName)
    'look up sales source name from code number
    lbcSSource.ListIndex = 0
    smSSource = ""
    slRecCode = Trim$(str$(tmSof.iMnfSSCode))
    For ilLoop = 0 To UBound(tmSSCode) - 1 Step 1 'lbcSSCode.ListCount - 1 Step 1
        slSSCode = tmSSCode(ilLoop).sKey   'lbcSSCode.List(ilLoop)
        ilRet = gParseItem(slSSCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", SOffice
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcSSource.ListIndex = ilLoop + 1
            smSSource = lbcSSource.List(ilLoop + 1)
            Exit For
        End If
    Next ilLoop
    lbcSRegion.ListIndex = 1
    smSRegion = ""
    slRecCode = Trim$(str$(tmSof.iMnfRegion))
    For ilLoop = 0 To UBound(tmSRCode) - 1 Step 1 'lbcSRCode.ListCount - 1 Step 1
        slSRCode = tmSRCode(ilLoop).sKey   'lbcSRCode.List(ilLoop)
        ilRet = gParseItem(slSRCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", SOffice
        On Error GoTo 0
        If slRecCode = slCode Then
            lbcSRegion.ListIndex = ilLoop + 2
            smSRegion = lbcSRegion.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    edcMktRank.Text = Trim$(str$(tmSof.iMktRank))
    For ilLoop = 0 To 2 Step 1
        edcAddr(ilLoop).Text = Trim$(tmSof.sAddr(ilLoop))
    Next ilLoop
    gSetPhoneNo tmSof.sPhone, mkcPhone
    gSetPhoneNo tmSof.sFax, mkcFax
    Select Case tmSof.sState
        Case "A"
            imState = 0
        Case "D"
            imState = 1
        Case Else
            imState = -1
    End Select
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
                    MsgBox "Sales Office already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmSof.sName) 'Reset text
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
    'gInitStdAlone sOffice, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igSofCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igSofCallSource = CALLNONE
    'End If
    If igSofCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgSofName = slStr
        Else
            sgSofName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
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

    imPopReqd = False
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(SOffice, cbcSelect, lbcNameCode, "Sof.Btr", gFieldOffset("Sof", "SofName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(SOffice, cbcSelect, tmNameCode(), smNameCodeTag, "Sof.Btr", gFieldOffset("Sof", "SofName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", SOffice
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
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer)
'
'   iRet = mReadRec()
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
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", SOffice
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmSofSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", SOffice
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
'*             Created:4/22/93       By:D. LeVine      *
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
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Sof.btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmSof.iCode = 0  'Autoincrement
            'tmSof.iMerge = 0
            tmSof.iRemoteID = tgUrf(0).iRemoteUserID
            tmSof.iAutoCode = tmSof.iCode
            ilRet = btrInsert(hmSof, tmSof, imSofRecLen, INDEXKEY0)
            slMsg = "Save Not Completed (btrInsert)"
        Else 'Old record-Update
            ilRet = btrUpdate(hmSof, tmSof, imSofRecLen)
            slMsg = "Save Not Completed (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, SOffice
    On Error GoTo 0
    If imSelectedIndex = 0 Then 'New selected
        Do
            'tmSofSrchKey.iCode = tmSof.iCode
            'ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Sales Office)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, SOffice
            'On Error GoTo 0
            tmSof.iRemoteID = tgUrf(0).iRemoteUserID
            tmSof.iAutoCode = tmSof.iCode
            ilRet = btrUpdate(hmSof, tmSof, imSofRecLen)
            slMsg = "mSaveRec (btrUpdate:Sales Office)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, SOffice
        On Error GoTo 0
    End If
'    'If lbcNameCode.Tag <> "" Then
'    '    If slStamp = lbcNameCode.Tag Then
'    '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Sof.btr")
'    '    End If
'    'End If
'    If smNameCodeTag <> "" Then
'        If slStamp = smNameCodeTag Then
'            smNameCodeTag = gFileDateTime(sgDBPath & "Sof.btr")
'        End If
'    End If
'    If imSelectedIndex <> 0 Then
'        'lbcNameCode.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tmNameCode()
'        cbcSelect.RemoveItem imSelectedIndex
'    End If
'    cbcSelect.RemoveItem 0 'Remove [New]
'    slName = RTrim$(tmSof.sName)
'    cbcSelect.AddItem slName
'    slName = tmSof.sName + "\" + LTrim$(Str$(tmSof.iCode)) 'slName + "\" + LTrim$(Str$(tmSof.iCode))
'    'lbcNameCode.AddItem slName
'    gAddItemToSortCode slName, tmNameCode(), True
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
                    pbcSoff_Paint
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
            gSetChgFlag tmSof.sName, edcName, tmCtrls(ilBoxNo)
        Case SSOURCEINDEX   'Sales Source
            gSetChgFlag smSSource, lbcSSource, tmCtrls(ilBoxNo)
        Case SREGIONINDEX   'Sales Region
            gSetChgFlag smSRegion, lbcSRegion, tmCtrls(ilBoxNo)
        Case MKTRANKINDEX 'Market Name
            gSetChgFlag str$(tmSof.iMktRank), edcMktRank, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX 'Address
            gSetChgFlag tmSof.sAddr(ilBoxNo - ADDRESSINDEX), edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 1 'Address
            gSetChgFlag tmSof.sAddr(ilBoxNo - ADDRESSINDEX), edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 2 'Address
            gSetChgFlag tmSof.sAddr(ilBoxNo - ADDRESSINDEX), edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            gSetChgFlag tmSof.sPhone, mkcPhone, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            gSetChgFlag tmSof.sFax, mkcFax, tmCtrls(ilBoxNo)
        Case STATEINDEX
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
    'Merge set only if change mode
    If (imSelectedIndex > 0) And (tgUrf(0).sMerge = "I") And (imUpdateAllowed) Then
        cmcMerge.Enabled = True
    Else
        cmcMerge.Enabled = False
    End If
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
'*             Created:4/22/93       By:D. LeVine      *
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
        Case SSOURCEINDEX   'Sales Source
            lbcSSource.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcSSource.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcSSource.List(lbcSSource.ListIndex)
            End If
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case SREGIONINDEX   'Sales Source
            lbcSRegion.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcSRegion.ListIndex <= 0 Then
                slStr = ""
            Else
                slStr = lbcSRegion.List(lbcSRegion.ListIndex)
            End If
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case MKTRANKINDEX 'Market Name
            edcMktRank.Visible = False  'Set visibility
            slStr = edcMktRank.Text
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX 'Address
            edcAddr(0).Visible = False
            slStr = edcAddr(0).Text
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 1 'Address
            edcAddr(1).Visible = False
            slStr = edcAddr(1).Text
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 2 'Address
            edcAddr(2).Visible = False
            slStr = edcAddr(2).Text
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
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
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            mkcFax.Visible = False  'Set visibility
            If mkcFax.Text = smFaxImage Then
                slStr = ""
            Else
                slStr = mkcFax.Text
            End If
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
            gSetShow pbcSoff, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSRegionBranch                  *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      source and process             *
'*                      communication back from sales  *
'*                      region                         *
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
Private Function mSRegionBranch() As Integer
'
'   ilRet = mSRegionBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDropDown, lbcSRegion, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mSRegionBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(SALESREGIONSLIST)) Then
    '    imDoubleClickName = False
    '    mSRegionBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "G"
    igMNmCallSource = CALLSOURCESALESOFFICE
    If lbcSRegion.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SOffice^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "SOffice^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SOffice^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "SOffice^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'SOffice.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'SOffice.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSRegionBranch = True
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
        lbcSRegion.Clear
        smSRCodeTag = ""
        mSRegionPop
        If imTerminate Then
            mSRegionBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcSRegion
        sgMNmName = ""
        If gLastFound(lbcSRegion) > 0 Then
            imChgMode = True
            lbcSRegion.ListIndex = gLastFound(lbcSRegion)
            edcDropDown.Text = lbcSRegion.List(lbcSRegion.ListIndex)
            imChgMode = False
            mSRegionBranch = False
            mSetChg SREGIONINDEX
        Else
            imChgMode = True
            lbcSRegion.ListIndex = 0
            edcDropDown.Text = lbcSRegion.List(0)
            imChgMode = False
            mSetChg SREGIONINDEX
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
'*      Procedure Name:mSRegionPop                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales Region list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSRegionPop()
'
'   mSRegionPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcSRegion.ListIndex
    If ilIndex > 1 Then
        slName = lbcSRegion.List(ilIndex)
    End If
    'Repopulate if required- if sales Region changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "G"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(SOffice, lbcSRegion, lbcSRCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(SOffice, lbcSRegion, tmSRCode(), smSRCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSRegionPopErr
        gCPErrorMsg ilRet, "mSRegionPop (gIMoveListBox)", SOffice
        On Error GoTo 0
        lbcSRegion.AddItem "[None]", 0  'Force as first item on list
        lbcSRegion.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcSRegion
            If gLastFound(lbcSRegion) > 1 Then
                lbcSRegion.ListIndex = gLastFound(lbcSRegion)
            Else
                lbcSRegion.ListIndex = -1
            End If
        Else
            lbcSRegion.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mSRegionPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSSourceBranch                  *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to sales  *
'*                      source and process             *
'*                      communication back from sales  *
'*                      source                         *
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
Private Function mSSourceBranch() As Integer
'
'   ilRet = mSSourceBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcSSource, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mSSourceBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(SALESSOURCESLIST)) Then
    '    imDoubleClickName = False
    '    mSSourceBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "S"
    igMNmCallSource = CALLSOURCESALESOFFICE
    If lbcSSource.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "SOffice^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "SOffice^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "SOffice^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "SOffice^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'SOffice.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'SOffice.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSSourceBranch = True
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
        lbcSSource.Clear
        smSSCodeTag = ""
        mSSourcePop
        If imTerminate Then
            mSSourceBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcSSource
        sgMNmName = ""
        If gLastFound(lbcSSource) > 0 Then
            imChgMode = True
            lbcSSource.ListIndex = gLastFound(lbcSSource)
            edcDropDown.Text = lbcSSource.List(lbcSSource.ListIndex)
            imChgMode = False
            mSSourceBranch = False
            mSetChg SSOURCEINDEX
        Else
            imChgMode = True
            lbcSSource.ListIndex = 0
            edcDropDown.Text = lbcSSource.List(0)
            imChgMode = False
            mSetChg SSOURCEINDEX
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
'*      Procedure Name:mSSourcePop                     *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Sales source list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSSourcePop()
'
'   mSSourcePop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcSSource.ListIndex
    If ilIndex > 0 Then
        slName = lbcSSource.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "S"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(SOffice, lbcSSource, lbcSSCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(SOffice, lbcSSource, tmSSCode(), smSSCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSSourcePopErr
        gCPErrorMsg ilRet, "mSSourcePop (gIMoveListBox)", SOffice
        On Error GoTo 0
        lbcSSource.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcSSource
            If gLastFound(lbcSSource) > 0 Then
                lbcSSource.ListIndex = gLastFound(lbcSSource)
            Else
                lbcSSource.ListIndex = -1
            End If
        Else
            lbcSSource.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mSSourcePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    sgDoneMsg = Trim$(str$(igSofCallSource)) & "\" & sgSofName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload SOffice
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
    If (ilCtrlNo = SSOURCEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcSSource, "", "Sales source must be specified", tmCtrls(SSOURCEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SSOURCEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SREGIONINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(lbcSRegion, "", "Sales region must be specified", tmCtrls(SREGIONINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SREGIONINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = MKTRANKINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcMktRank, "", "Market Rank must be specified", tmCtrls(MKTRANKINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = MKTRANKINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ADDRESSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcAddr(0), "", "Address must be specified", tmCtrls(ADDRESSINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ADDRESSINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ADDRESSINDEX + 1) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcAddr(1), "", "Address must be specified", tmCtrls(ADDRESSINDEX + 1).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ADDRESSINDEX + 1
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = ADDRESSINDEX + 2) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcAddr(2), "", "Address must be specified", tmCtrls(ADDRESSINDEX + 2).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ADDRESSINDEX + 2
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
            If (ilBox = ADDRESSINDEX + 1) Or (ilBox = ADDRESSINDEX + 2) Then
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
End Sub
Private Sub pbcSoff_Paint()
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        pbcSoff.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSoff.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSoff.Print tmCtrls(ilBox).sShow
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
    If imBoxNo = SSOURCEINDEX Then
        If mSSourceBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SREGIONINDEX Then
        If mSRegionBranch() Then
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
                    ilBox = 2
                End If
            Case NAMEINDEX 'Name (first control within header)
                mSetShow imBoxNo
                imBoxNo = -1
                If cbcSelect.Enabled Then
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
            Case MKTRANKINDEX
                If lbcSRegion.ListCount = 2 Then
                    imChgMode = True
                    lbcSRegion.ListIndex = 1
                    imChgMode = False
                    ilFound = False
                End If
                ilBox = SREGIONINDEX
            Case Else
                ilBox = ilBox - 1
        End Select
    Loop While Not ilFound
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
    If imState = 0 Then  'Active
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
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SSOURCEINDEX Then
        If mSSourceBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = SREGIONINDEX Then
        If mSRegionBranch() Then
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
            Case SSOURCEINDEX
                If lbcSRegion.ListCount = 2 Then
                    imChgMode = True
                    lbcSRegion.ListIndex = 1
                    imChgMode = False
                    ilFound = False
                End If
                ilBox = SREGIONINDEX
            Case UBound(tmCtrls) 'last control
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igSofCallSource = CALLNONE) Then
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
Private Sub plcSoff_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case SSOURCEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSSource, edcDropDown, imChgMode, imLbcArrowSetting
        Case SREGIONINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSRegion, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
