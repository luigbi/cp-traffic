VERSION 5.00
Begin VB.Form CPMPkg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   240
   ClientTop       =   2985
   ClientWidth     =   10440
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
   ScaleHeight     =   5700
   ScaleWidth      =   10440
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   890
      Left            =   255
      Picture         =   "CPMPkg.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2145
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmcSpecDropDown 
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
      Left            =   1575
      Picture         =   "CPMPkg.frx":631E
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcSpecDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2205
      MaxLength       =   20
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   1605
      Picture         =   "CPMPkg.frx":6418
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   285
      MaxLength       =   10
      TabIndex        =   18
      Top             =   2580
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.ListBox lbcPDFName 
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
      Left            =   6180
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1455
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.FileListBox lbcPDFFile 
      Height          =   285
      Left            =   8205
      Pattern         =   "*.PDF"
      TabIndex        =   10
      Top             =   5325
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox pbcProgrammatic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3375
      ScaleHeight     =   210
      ScaleWidth      =   1020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer tmcInit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1830
      Top             =   5250
   End
   Begin VB.ListBox lbcVehGp3 
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
      Left            =   5490
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1005
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   10350
      ScaleHeight     =   120
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   450
      Width           =   45
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
      HelpContextID   =   1
      Left            =   5340
      TabIndex        =   25
      Top             =   5325
      Width           =   945
   End
   Begin VB.PictureBox pbcInvTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4575
      ScaleHeight     =   210
      ScaleWidth      =   1170
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1080
      Top             =   5235
   End
   Begin VB.PictureBox pbcAlter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4395
      ScaleHeight     =   210
      ScaleWidth      =   1020
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3465
      ScaleHeight     =   210
      ScaleWidth      =   1035
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox pbcPkgTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   60
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   21
      Top             =   5280
      Width           =   135
   End
   Begin VB.PictureBox pbcPkgSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -45
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   15
      Top             =   285
      Width           =   105
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
      Left            =   4125
      TabIndex        =   24
      Top             =   5325
      Width           =   945
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
      Left            =   2895
      TabIndex        =   23
      Top             =   5325
      Width           =   945
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   60
      Picture         =   "CPMPkg.frx":6512
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   390
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5385
      Width           =   75
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   -60
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   14
      Top             =   1020
      Width           =   135
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   1320
      Width           =   105
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
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   1650
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1650
   End
   Begin VB.PictureBox pbcSpec 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Left            =   240
      Picture         =   "CPMPkg.frx":681C
      ScaleHeight     =   375
      ScaleWidth      =   4575
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   4575
   End
   Begin VB.PictureBox plcSpec 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   195
      ScaleHeight     =   435
      ScaleWidth      =   4605
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   4665
   End
   Begin VB.PictureBox pbcPkg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   3975
      Left            =   240
      Picture         =   "CPMPkg.frx":BDDE
      ScaleHeight     =   3975
      ScaleWidth      =   8370
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8370
      Begin VB.Label lacCover 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0FFFF&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3825
         TabIndex        =   31
         Top             =   3750
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.Label lacPkgFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   30
         Top             =   270
         Visible         =   0   'False
         Width           =   8385
      End
   End
   Begin VB.PictureBox plcPkg 
      ForeColor       =   &H00000000&
      Height          =   4095
      Left            =   225
      ScaleHeight     =   4035
      ScaleWidth      =   8670
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1110
      Width           =   8730
      Begin VB.VScrollBar vbcPkg 
         Height          =   3975
         LargeChange     =   17
         Left            =   8400
         Min             =   1
         TabIndex        =   22
         Top             =   30
         Value           =   1
         Width           =   240
      End
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4905
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2955
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5415
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5295
      TabIndex        =   27
      Top             =   2850
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6075
      TabIndex        =   1
      Top             =   75
      Width           =   4200
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   300
      Picture         =   "CPMPkg.frx":11D978
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "CPMPkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CPMPkg.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the CPM Package input screen code
'
Option Explicit
Option Compare Text
'Dim hmDrf As Integer
Dim hmMnf As Integer
'Vehicle
Dim imAddPkg As Integer     'Force vef and vpf to be re-read
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim tmPkgVehicle() As SORTCODE
Dim smPkgVehicleTag As String
Dim smOrigDemo As String
'Vehicle Options
Dim hmVpf As Integer        'Vehicle file handle
Dim tmVpf As VPF            'VPF record image
Dim tmVpfSrchKey As VPFKEY0 'VPF key record image
Dim imVpfRecLen As Integer  'VPF record length
'Vehicle Features
Dim hmVff As Integer        'Vehicle file handle
Dim tmVff As VFF            'VFF record image
Dim tmVffSrchKey As INTKEY0 'VFF key record image
Dim tmVffSrchKey1 As INTKEY0
Dim imVffRecLen As Integer  'VFF record length
'CPM Package Vechcle
Dim hmPvf As Integer        'CPM Vehicle file handle
Dim tmPvf() As PVF            'PVF record image
Dim tmTPvf As PVF
Dim tmPvfSrchKey As LONGKEY0 'PVF key record image
Dim imPvfRecLen As Integer  'PVF record length

'Specification Area
Dim tmSpecCtrls(0 To 1) As FIELDAREA
Dim imLBSpecCtrls As Integer
Dim imSpecBoxNo As Integer
Dim smSpecSave(0 To 1) As String  'Values saved (1=Name)
'Package Vehicle Areas
Dim tmPkgCtrls(0 To 3)  As FIELDAREA    'Time/Days
Dim imLBPkgCtrls As Integer
Dim imPkgBoxNo As Integer
Dim imPkgRowNo As Integer
Dim imPkgChg As Integer
Dim smTShow(0 To 3) As String
Dim tmPBDP() As RCPBDPGEN
Dim tmRdf As RDF
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imUpdateAllowed As Integer
Dim imLbcArrowSetting As Integer
Dim imTabDirection As Integer
Dim imDirProcess As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imFirstFocus As Integer
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imComboBoxIndex As Integer
Dim imSettingValue As Integer
Dim imFirstTimeSelect As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor
Const LBONE = 1
Const AVGINDEX = RCAVGRATEINDEX 'Also in RateCard.Frm and StdPkg screens

'HEADER Column
Const NAMEINDEX = 1          'Name control/field

'GRID Columns
Const PKGVEHINDEX = 1       'Vehicle control/field
Const PKGDPINDEX = 2        'Ad Location control/field
Const PKGPERCENTINDEX = 3   'Percent
Dim hmDrf As Integer
Dim hmDpf As Integer
Dim hmDef As Integer
Dim hmRaf As Integer

Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then  'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    mClearCtrlFields
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
    pbcSpec.Cls
    pbcPkg.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
        mInitShow
        mGetTotals
    Else
        imSelectedIndex = 0
        If slStr <> "[New]" Then
            edcSpecDropDown.MaxLength = 20
            edcSpecDropDown.Text = slStr
            mSpecSetShow NAMEINDEX
        End If
    End If
    imFirstTimeSelect = True
    pbcSpec_Paint
    pbcPkg_Paint
    imChgMode = False
    imBypassSetting = False
    mSetCommands
    Screen.MousePointer = vbDefault
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
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
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'mInitDDE
        If igDPNameCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgDPName <> "" Then
                mSetCommands
                gFindMatch sgDPName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    Exit Sub
                End If
            End If
            vbcPkg.Visible = False
            DoEvents
            vbcPkg.Visible = True
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        imFirstTimeSelect = True
        pbcStartNew.SetFocus
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
Private Sub cmcCancel_Click()
    If igDPNameCallSource <> CALLNONE Then
        igDPNameCallSource = CALLCANCELLED
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus cmcCancel
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igDPNameCallSource <> CALLNONE Then
        sgDPName = smSpecSave(1) 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgDPName = "[New]"
            If Not imTerminate Then
                mSpecEnableBox imSpecBoxNo
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
            mSpecEnableBox imSpecBoxNo
            Exit Sub
        End If
    End If
    If igDPNameCallSource <> CALLNONE Then
        If sgDPName = "[New]" Then
            igDPNameCallSource = CALLCANCELLED
        Else
            igDPNameCallSource = CALLDONE
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imPkgBoxNo
        Case PKGPERCENTINDEX
    End Select
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcProof_Click()
    Dim hlProof As Integer
    Dim ilRet As Integer
    Dim slToFile As String
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim ilLoop As Integer
    Dim slDateTime As String
    'Check for illegal characters in name
    slToFile = sgExportPath & gFileNameFilter(Trim$(smSpecSave(1))) & ".csv"
    ilRet = 0
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    ilRet = 0
    'hlProof = FreeFile
    'Open slToFile For Output As hlProof
    ilRet = gFileOpen(slToFile, "Output", hlProof)
    If ilRet <> 0 Then
        MsgBox "Open " & slToFile & ", Error #" & Str(err.Number), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        Exit Sub
    End If
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
                ilVef = gBinarySearchVef(tmPBDP(ilLoop).iVefCode)
                If ilVef <> -1 Then
                    ilRdf = gBinarySearchRdf(tmPBDP(ilLoop).iRdfCode)
                    If ilRdf <> -1 Then
                        Print #hlProof, Trim$(tgMVef(ilVef).sName) & "," & Trim$(tgMRdf(ilRdf).sName)
                    End If
                End If
            End If
        End If
    Next ilLoop
    Close #hlProof
    MsgBox "Create file: " & slToFile, vbOKOnly + vbApplicationModal, "File Saved"
    Exit Sub
End Sub

Private Sub cmcSpecDropDown_Click()
    edcSpecDropDown.SelStart = 0
    edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    edcSpecDropDown.SetFocus
End Sub
Private Sub cmcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = cbcSelect.Text   'Save name
    If slName = "" Then Exit Sub
    
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mSpecEnableBox imSpecBoxNo
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imSpecBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    ilCode = tmVef.iCode
    cbcSelect.Clear
    smPkgVehicleTag = ""
    mPopulate
    For ilLoop = 0 To UBound(tmPkgVehicle) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tmPkgVehicle(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
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
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub
Private Sub edcDropDown_Change()
    Select Case imPkgBoxNo
        Case PKGPERCENTINDEX
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imPkgBoxNo
        Case PKGPERCENTINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imPkgBoxNo
        Case PKGPERCENTINDEX
            ilPos = InStr(edcDropDown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
                If ilPos > 0 Then
                    If KeyAscii = KEYDECPOINT Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "100.00") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcSpecDropDown_Change()
    imLbcArrowSetting = False
End Sub
Private Sub edcSpecDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcSpecDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcSpecDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim slComp As String
    
        If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
            If edcSpecDropDown.SelLength <> 0 Then    'avoid deleting two characters
                imBSMode = True 'Force deletion of character prior to selected text
            End If
        End If
        ilKey = KeyAscii
        If Not gCheckKeyAscii(ilKey) Then
            KeyAscii = 0
            Exit Sub
        End If
    
End Sub
Private Sub edcSpecDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcSpecTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        edcSpecDropDown.SelStart = 0
        edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
    End If
End Sub
Private Sub edcSpecDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        'do Doubleclick here
        imDoubleClickName = False
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
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSpec.Enabled = False
        pbcPkg.Enabled = False
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        pbcPkgSTab.Enabled = False
        pbcPkgTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcSpec.Enabled = True
        pbcPkg.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        pbcPkgSTab.Enabled = True
        pbcPkgTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    'Me.Visible = False
    DoEvents    'Process events so pending keys are not sent to this
    'Me.Visible = True
    Me.KeyPreview = True
    CPMPkg.Refresh
    vbcPkg.Visible = False
    DoEvents
    vbcPkg.Visible = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And ((imSpecBoxNo > 0) Or (imPkgBoxNo > 0)) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imSpecBoxNo > 0 Then
            mSpecEnableBox imSpecBoxNo
        ElseIf imPkgBoxNo > 0 Then
            mPkgEnableBox imPkgBoxNo
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
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = (((lgPercentAdjW - 10) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = ((lgPercentAdjW - 10) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
    
    pbcPkgTab.Left = -pbcPkgTab.Width - 200
    pbcPkgSTab.Left = -pbcPkgSTab.Width - 200
    pbcSpecTab.Left = -pbcSpecTab.Width - 200
    pbcSpecSTab.Left = -pbcSpecSTab.Width - 200
    pbcStartNew.Left = -pbcStartNew.Width - 200
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    If Not igManUnload Then
        mSpecSetShow imSpecBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mSpecEnableBox imSpecBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Cancel = 0
    Erase tmPvf
    Erase tmPBDP
    Erase smPkgShow
    Erase smPkgSave
    Erase tmPkgVehicle
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    ilRet = btrClose(hmPvf)
    btrDestroy hmPvf
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    ilRet = btrClose(hmDef)
    btrDestroy hmDef
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    
    Set CPMPkg = Nothing   'Remove data segment
End Sub

Private Sub imcKey_Click()
    pbcKey.Visible = Not pbcKey.Visible
End Sub

Private Sub lbcDemo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcPDFName_Click()
    gProcessLbcClick lbcPDFName, edcSpecDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcPDFName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim slStr As String
    imPkgChg = False
    For ilLoop = LBound(smSpecSave) To UBound(smSpecSave) Step 1
        smSpecSave(ilLoop) = ""
    Next ilLoop
    'For ilLoop = LBound(imSpecSave) To UBound(imSpecSave) Step 1
    '    imSpecSave(ilLoop) = -1
    'Next ilLoop
    For ilLoop = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilLoop).sShow = ""
        tmSpecCtrls(ilLoop).iChg = False
    Next ilLoop
    tmVef.sName = ""
    tmVef.sStdPrice = ""
    tmVef.sStdInvTime = ""
    tmVef.sStdAlter = ""
    tmVef.sStdAlterName = ""
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        tmPBDP(ilLoop).sKey = "1" & tmPBDP(ilLoop).sSvKey
        tmPBDP(ilLoop).lAvgPrice = tmPBDP(ilLoop).lSvAvgPrice
        tmPBDP(ilLoop).lAvgAud = 0
        tmPBDP(ilLoop).iAvgRating = 0
        tmPBDP(ilLoop).lCPP = 0
        tmPBDP(ilLoop).lCPM = 0
        tmPBDP(ilLoop).lPop = 0
    Next ilLoop
    'smOrigProgrammaticAllowed = ""
    'smOrigSalesBrochure = ""
    mInitShowFields
    For ilLoop = LBONE To UBound(smPkgSave, 2) - 1 Step 1
        slStr = ""
        smPkgSave(1, ilLoop) = slStr
        slStr = ""
        smPkgSave(2, ilLoop) = slStr
    Next ilLoop
    For ilLoop = LBound(smTShow) To UBound(smTShow) Step 1
        smTShow(ilLoop) = ""
    Next ilLoop
    vbcPkg.Value = vbcPkg.Min
    ReDim tmPvf(0 To 0) As PVF
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetTotals                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine totals               *
'*                                                     *
'*******************************************************
Private Sub mGetTotals()
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim llPop As Long
    Dim llTCost As Long
    Dim slStr As String
    Dim slTotalPct As String
    Dim llLnSpots As Long
    ilCount = 0
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(2, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(2, ilLoop))) > 0 Then
                ilCount = ilCount + 1
            End If
        End If
    Next ilLoop
    slTotalPct = "0.0"
    If ilCount <= 0 Then
        For ilLoop = LBound(smTShow) To UBound(smTShow) Step 1
            smTShow(ilLoop) = ""
        Next ilLoop
    Else
        ilCount = ilCount - 1
        ilCount = 0
        For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
            If Trim$(smPkgSave(2, ilLoop)) <> "" Then
                If Val(Trim$(smPkgSave(2, ilLoop))) > 0 Then
                    slTotalPct = gAddStr(Trim$(smPkgSave(2, ilLoop)), slTotalPct)
                    ilCount = ilCount + 1
                End If
            End If
        Next ilLoop
        llLnSpots = 1   'ilTNoSpots
        If tgSpf.sSAudData = "H" Then
            'slStr = gLongToStrDec(llTGrImp, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            'slStr = gLongToStrDec(llTGrImp, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            'slStr = gLongToStrDec(llTGrImp, 3)
        Else
            'slStr = Trim$(str$(llTGrImp))
        End If
        
        slStr = slTotalPct
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPERCENTINDEX)
        smTShow(1) = tmPkgCtrls(PKGPERCENTINDEX).sShow
        
    End If
    lacCover.Visible = True
    lacCover.Visible = False
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
    Screen.MousePointer = vbHourglass
    imLBSpecCtrls = 1
    imLBPkgCtrls = 1
    imAddPkg = False
    imFirstActivate = True
    imcKey.Picture = IconTraf!imcKey.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    igDPAltered = False
    imTerminate = False
    imBypassSetting = False
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imPopReqd = False
    imFirstFocus = True
    imSelectedIndex = -1
    imPkgBoxNo = -1
    imSpecBoxNo = -1
    imPkgChg = False
    imSettingValue = False
    imFirstTimeSelect = True
    hmVef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", CPMPkg
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vpf.Btr)", CPMPkg
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)
    hmVff = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vff.Btr)", CPMPkg
    On Error GoTo 0
    imVffRecLen = Len(tmVff)
    ReDim tmPvf(0 To 0) As PVF
    hmPvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pvf.Btr)", CPMPkg
    On Error GoTo 0
    imPvfRecLen = Len(tmPvf(0))
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "DRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: DRF.Btr)", CPMPkg
    On Error GoTo 0
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "MNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MNF.Btr)", CPMPkg
    On Error GoTo 0
    hmDpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", CPMPkg
    On Error GoTo 0
    ' setup global variable for Demo Plus file (to see if any exists)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    hmDef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", CPMPkg
    On Error GoTo 0
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", CPMPkg

    On Error GoTo 0
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    If imTerminate Then
        Exit Sub
    End If
    mInitBox
    mInitPkg
    'lbcVehGp3.Clear
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If Not imTerminate Then
        '5/30/19: moved here from cbcSelect_GotFocus to aviod multi-calls to create screen
        'cbcSelect.ListIndex = 0 'This will generate a select_change event
        If igDPNameCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgDPName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgDPName    'New name
            End If
        End If
        mSetCommands
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gCenterStdAlone CPMPkg
    tmcInit.Enabled = True
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
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long
    Dim ilLoop As Integer
    Dim ilBox As Integer

    flTextHeight = pbcSpec.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcSpec.Move 240, 540, pbcSpec.Width + fgPanelAdj, pbcSpec.Height + fgPanelAdj
    pbcSpec.Move plcSpec.Left + fgBevelX, plcSpec.Top + fgBevelY
    'Package Name
    gSetCtrl tmSpecCtrls(NAMEINDEX), 30, 30, 4500, fgBoxStH
    
    '10/25/14: One pixel removed from top and left side when using macromedia fireworks
    For ilBox = LBound(tmSpecCtrls) To UBound(tmSpecCtrls) Step 1
        tmSpecCtrls(ilBox).fBoxX = tmSpecCtrls(ilBox).fBoxX - 15
        tmSpecCtrls(ilBox).fBoxY = tmSpecCtrls(ilBox).fBoxY - 15
    Next ilBox
    
    plcPkg.Move 225, 1110, pbcPkg.Width + vbcPkg.Width + fgPanelAdj, pbcPkg.Height + fgPanelAdj
    pbcPkg.Move plcPkg.Left + fgBevelX, plcPkg.Top + fgBevelY
    pbcArrow.Move plcPkg.Left - pbcArrow.Width - 15
    'Vehicle
    gSetCtrl tmPkgCtrls(PKGVEHINDEX), 30, 225, 1710, fgBoxGridH
    'Ad Location
    gSetCtrl tmPkgCtrls(PKGDPINDEX), 1755, tmPkgCtrls(PKGVEHINDEX).fBoxY, 1545, fgBoxGridH
    'Percent
    gSetCtrl tmPkgCtrls(PKGPERCENTINDEX), 3315, tmPkgCtrls(PKGVEHINDEX).fBoxY, 450, fgBoxGridH
    tmPkgCtrls(PKGPERCENTINDEX).iReq = True

    llMax = 0
    For ilLoop = imLBPkgCtrls To UBound(tmPkgCtrls) Step 1
        tmPkgCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmPkgCtrls(ilLoop).fBoxW)
        Do While (tmPkgCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmPkgCtrls(ilLoop).fBoxW = tmPkgCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmPkgCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmPkgCtrls(ilLoop).fBoxX)
            Do While (tmPkgCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmPkgCtrls(ilLoop).fBoxX = tmPkgCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmPkgCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmPkgCtrls(ilLoop - 1).fBoxX + tmPkgCtrls(ilLoop - 1).fBoxW + 15 < tmPkgCtrls(ilLoop).fBoxX Then
                        tmPkgCtrls(ilLoop - 1).fBoxW = tmPkgCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmPkgCtrls(ilLoop - 1).fBoxX + tmPkgCtrls(ilLoop - 1).fBoxW + 15 > tmPkgCtrls(ilLoop).fBoxX Then
                        tmPkgCtrls(ilLoop - 1).fBoxW = tmPkgCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmPkgCtrls(ilLoop).fBoxX + tmPkgCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmPkgCtrls(ilLoop).fBoxX + tmPkgCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop

    pbcPkg.Picture = LoadPicture("")
    pbcPkg.Width = llMax
    plcPkg.Width = llMax + vbcPkg.Width + 2 * fgBevelX + 15
    Me.Width = plcPkg.Width + 2 * plcPkg.Left
    cbcSelect.Left = Me.Width - cbcSelect.Width - 120
    lacPkgFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (CPMPkg.Width - 4 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcDone.Top = CPMPkg.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top

    llAdjTop = cmcDone.Top - plcSpec.Top - plcSpec.Height - 120 - tmPkgCtrls(1).fBoxH
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcPkg.Top + llAdjTop + 2 * fgBevelY + 240 < cmcDone.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcPkg.Height = llAdjTop + 2 * fgBevelY
    pbcPkg.Left = plcPkg.Left + fgBevelX
    pbcPkg.Top = plcPkg.Top + fgBevelY
    pbcPkg.Height = plcPkg.Height - 2 * fgBevelY
    vbcPkg.Left = plcPkg.Width - vbcPkg.Width - fgBevelX - 30
    vbcPkg.Top = fgBevelY
    vbcPkg.Height = pbcPkg.Height
    lacCover.Top = pbcPkg.Height - lacCover.Height - 15
    lacCover.Width = tmPkgCtrls(PKGPERCENTINDEX).fBoxX + tmPkgCtrls(PKGPERCENTINDEX).fBoxW '- tmPkgCtrls(PKGPRICEINDEX).fBoxX
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitPkg                        *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize Rate card Items for  *
'*                     Package                         *
'*                                                     *
'*******************************************************

'**********************************************************
'Rules to use for active Podcast Medium Vehicles (1/20/22)
'**********************************************************
'Rate card Screen:
'**********************************************************
'  Always show active podcast medium vehicles
'    Vpf.sGMedium = "P" (PodCast)
'**********************************************************
'Std Pkg Screen:
'**********************************************************
'  Show podcast medium vehicles when vehicle has programming
'    Vpf.sGMedium = "P" (PodCast)
'    LTF_Lbrary_Title WHERE LtfVefCode
'**********************************************************
'CPM Pkg Screen: (you are here)
'**********************************************************
'  Show podcast medium vehicles when it has an ad server..
'  vendor defined in vehicle options
'    Vpf.sGMedium = "P" (PodCast)
'
'    pvfType="C" =Podcast Ad Server (CPM only)
'    Vff.iAvfCode <> 0 (has Ad Server)
'
'    CpmPkg button visible when sFeatures8=PODADSERVER
'**********************************************************
Private Sub mInitPkg()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slVehName As String
    Dim ilVefCode As Integer
    Dim slDPName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilBypass As Integer
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilVff As Integer
    Dim ilRdfCode As Integer
    Dim slStr As String
    Dim llUpper As Long
    Dim slVpfMedium As String
    ReDim tmPBDP(0 To 1) As RCPBDPGEN
    ReDim smPkgShow(0 To 10, 0 To UBound(tmPBDP)) As String * 30
    ReDim smPkgSave(0 To 2, 0 To UBound(tmPBDP)) As String * 10
    For ilLoop = LBound(smPkgShow, 1) To UBound(smPkgShow, 1) Step 1
        For ilIndex = LBound(smPkgShow, 2) To UBound(smPkgShow, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(smPkgSave, 1) To UBound(smPkgSave, 1) Step 1
        For ilIndex = LBound(smPkgSave, 2) To UBound(smPkgSave, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    ilRet = 0
    On Error GoTo mInitPkgErr
    llUpper = LBound(tmRifRec)
    On Error GoTo 0
    If ilRet = 0 Then
        For ilLoop = LBONE To UBound(tmRifRec) - 1 Step 1
            'Vehicle
            gFindMatch Trim$(smRCSave(1, ilLoop)), 0, RateCard!lbcVehicle
            ilIndex = gLastFound(RateCard!lbcVehicle)
            If ilIndex >= 0 Then
                slNameCode = tgRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
                ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                ilRet = gParseItem(slVehName, 3, "|", slVehName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                'Test if Package vehicle- if so, bypass
                ilBypass = False
                ilVef = gBinarySearchVef(ilVefCode)
                If ilVef <> -1 Then
                    If (tgMVef(ilVef).sType = "P") Then
                        ilBypass = True
                    End If
                    
                    '2/11/21 - Show only Podcast Vehicles in Grid
                    If ilBypass = False Then
                        ilVpf = gBinarySearchVpf(ilVefCode)
                        slVpfMedium = tgVpf(ilVpf).sGMedium
                        If ilVpf <> -1 Then
                            If tgVpf(ilVpf).sGMedium <> "P" Then
                                ilBypass = True
                            End If
                        End If
                    End If
                    
                    '5/19/21 - Prevent Podcast vehicles with No AdServer
                    If ilBypass = False Then
                        ilVff = gBinarySearchVff(ilVefCode)
                        If ilVff <> -1 Then
                            If tgVff(ilVff).iAvfCode = 0 Then
                                ilBypass = True
                            End If
                        End If
                    End If
                End If
            Else
                ilVefCode = -1
                slVehName = "Missing"
                ilBypass = False
            End If
            If Not ilBypass Then
                'Ad Location
                gFindMatch Trim$(smRCSave(2, ilLoop)), 0, RateCard!lbcDPName
                ilIndex = gLastFound(RateCard!lbcDPName)
                If ilIndex >= 0 Then
                    slNameCode = RateCard!lbcDPNameCode.List(ilIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slDPName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilRdfCode = Val(slCode)
                Else
                    ilRdfCode = -1
                    slDPName = "Missing"
                End If
                
                'Build record into tmPBDP
                tmPBDP(UBound(tmPBDP)).sSvKey = tmRifRec(ilLoop).sKey
                tmPBDP(UBound(tmPBDP)).iRdfCode = ilRdfCode
                tmPBDP(UBound(tmPBDP)).sVehName = slVehName
                tmPBDP(UBound(tmPBDP)).sDPName = slDPName   'mMakePrgName(ilRdfCode)
                tmPBDP(UBound(tmPBDP)).iVefCode = ilVefCode
                slStr = Trim$(smRCShow(AVGINDEX, ilLoop))
                gUnformatStr slStr, UNFMTDEFAULT, slStr
                tmPBDP(UBound(tmPBDP)).lAvgPrice = gStrDecToLong(slStr, 0)
                tmPBDP(UBound(tmPBDP)).lSvAvgPrice = gStrDecToLong(slStr, 0)
                tmPBDP(UBound(tmPBDP)).lAvgAud = 0
                tmPBDP(UBound(tmPBDP)).iAvgRating = 0
                tmPBDP(UBound(tmPBDP)).lGrImp = 0
                tmPBDP(UBound(tmPBDP)).lGRP = 0
                tmPBDP(UBound(tmPBDP)).lCPP = 0
                tmPBDP(UBound(tmPBDP)).lCPM = 0
                tmPBDP(UBound(tmPBDP)).lPop = 0
                tmPBDP(UBound(tmPBDP)).iVehDormant = imRCSave(9, ilLoop)
                tmPBDP(UBound(tmPBDP)).iDPDormant = imRCSave(10, ilLoop)
                tmPBDP(UBound(tmPBDP)).iPkgVeh = imRCSave(11, ilLoop)
                tmPBDP(UBound(tmPBDP)).sMedium = slVpfMedium
                ReDim Preserve tmPBDP(0 To UBound(tmPBDP) + 1) As RCPBDPGEN
            End If
        Next ilLoop
    End If
    ReDim smPkgShow(0 To 10, 0 To UBound(tmPBDP)) As String * 30
    ReDim smPkgSave(0 To 2, 0 To UBound(tmPBDP)) As String * 10
    For ilLoop = LBound(smPkgShow, 1) To UBound(smPkgShow, 1) Step 1
        For ilIndex = LBound(smPkgShow, 2) To UBound(smPkgShow, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    For ilLoop = LBound(smPkgSave, 1) To UBound(smPkgSave, 1) Step 1
        For ilIndex = LBound(smPkgSave, 2) To UBound(smPkgSave, 2) Step 1
            smPkgShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    imSettingValue = True
    vbcPkg.Min = LBONE  'LBound(tmPBDP)
    imSettingValue = True
    If UBound(tmPBDP) - 1 <= vbcPkg.LargeChange + 1 Then ' + 1 Then
        vbcPkg.Max = LBONE  'LBound(tmPBDP)
    Else
        vbcPkg.Max = UBound(tmPBDP) - vbcPkg.LargeChange
    End If
    imSettingValue = True
    vbcPkg.Value = vbcPkg.Min
    pbcPkg_Paint
    Exit Sub
mInitPkgErr:
    ilRet = 1
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShow                       *
'*                                                     *
'*             Created:7/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitShow()
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilLoop As Integer

    slStr = smSpecSave(1)
    gSetShow pbcSpec, slStr, tmSpecCtrls(NAMEINDEX)

    For ilLoop = LBONE To UBound(smPkgSave, 2) - 1 Step 1
        slStr = Trim$(smPkgSave(2, ilLoop))
        gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPERCENTINDEX)
        smPkgShow(PKGPERCENTINDEX, ilLoop) = tmPkgCtrls(PKGPERCENTINDEX).sShow
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShowFields                 *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show for rate card  *
'*                      fields                         *
'*                                                     *
'*******************************************************
Private Sub mInitShowFields()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llLoop As Long
    Dim ilVef As Integer
    Dim ilRet As Integer
    
    If UBound(tmPBDP) > 1 Then
        'ArraySortTyp fnAV(tmPBDP(), 1), UBound(tmPBDP) - 1, 0, LenB(tmPBDP(1)), 0, LenB(tmPBDP(1).sKey), 0
        For llLoop = LBound(tmPBDP) To UBound(tmPBDP) - 1 Step 1
            tmPBDP(llLoop) = tmPBDP(llLoop + 1)
        Next llLoop
        ReDim Preserve tmPBDP(0 To UBound(tmPBDP) - 1) As RCPBDPGEN
        
        ArraySortTyp fnAV(tmPBDP(), 0), UBound(tmPBDP), 0, LenB(tmPBDP(0)), 0, LenB(tmPBDP(0).sKey), 0
        
        ReDim Preserve tmPBDP(0 To UBound(tmPBDP) + 1) As RCPBDPGEN
        For llLoop = UBound(tmPBDP) - 1 To LBound(tmPBDP) Step -1
            tmPBDP(llLoop + 1) = tmPBDP(llLoop)
            
            Debug.Print tmPBDP(llLoop).sVehName
            
        Next llLoop
    End If
    
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        'Vehicle Name
        slStr = tmPBDP(ilLoop).sVehName
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGVEHINDEX)
        smPkgShow(PKGVEHINDEX, ilLoop) = tmPkgCtrls(PKGVEHINDEX).sShow
        'Ad Location name
        slStr = tmPBDP(ilLoop).sDPName
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGDPINDEX)
        smPkgShow(PKGDPINDEX, ilLoop) = tmPkgCtrls(PKGDPINDEX).sShow
        'Impression %
        slStr = ""
        smPkgSave(2, ilLoop) = slStr
        gSetShow pbcPkg, slStr, tmPkgCtrls(PKGPERCENTINDEX)
        smPkgShow(PKGPERCENTINDEX, ilLoop) = tmPkgCtrls(PKGPERCENTINDEX).sShow
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitVef                        *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Inialize vehicle               *
'*                                                     *
'*******************************************************
Private Sub mInitVef()
    gInitVef tmVef
    tmVef.sType = "P"
    tmVef.sState = "A"
    tmVef.sExportRAB = "N"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec iTest
'   Where:
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    tmVef.sName = smSpecSave(1)    'Name
    ReDim tmPvf(0 To 0) As PVF
    ilIndex = LBound(tmPvf(0).iVefCode)
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(2, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(2, ilLoop))) > 0 Then
                tmPvf(UBound(tmPvf)).iVefCode(ilIndex) = tmPBDP(ilLoop).iVefCode
                tmPvf(UBound(tmPvf)).iRdfCode(ilIndex) = tmPBDP(ilLoop).iRdfCode
                tmPvf(UBound(tmPvf)).iNoSpot(ilIndex) = Val(Trim$(smPkgSave(1, ilLoop)))
                tmPvf(UBound(tmPvf)).iPctRate(ilIndex) = gStrDecToLong(Trim$(smPkgSave(2, ilLoop)), 2)
                tmPvf(UBound(tmPvf)).sType = "C" 'pvfType Set to "C" = Podcast Ad Server (CPM only)
                ilIndex = ilIndex + 1
                If ilIndex > UBound(tmPvf(0).iVefCode) Then
                    ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
                    ilIndex = LBound(tmPvf(0).iVefCode)
                End If
            End If
        End If
    Next ilLoop
    If ilIndex > LBound(tmPvf(0).iVefCode) Then
        ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
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
'*             Created:7/01/93       By:D. LeVine      *
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
    Dim ilTest As Integer
    Dim ilPvf As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    
    smSpecSave(1) = Trim$(tmVef.sName)
    For ilTest = LBONE To UBound(tmPBDP) - 1 Step 1
        'Debug.Print tmPBDP(ilTest).sSvKey
        tmPBDP(ilTest).sKey = "1" & tmPBDP(ilTest).sSvKey
        If tmPBDP(0).lGrImp > 0 Then
            tmPBDP(ilTest).sKey = "0" & tmPBDP(ilTest).sSvKey
        End If
    Next ilTest
    For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
        For ilLoop = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
            If tmPvf(ilPvf).iVefCode(ilLoop) > 0 Then
                For ilTest = LBONE To UBound(tmPBDP) - 1 Step 1
                    If (tmPvf(ilPvf).iVefCode(ilLoop) = tmPBDP(ilTest).iVefCode) And (tmPvf(ilPvf).iRdfCode(ilLoop) = tmPBDP(ilTest).iRdfCode) Then
                        If tmPvf(ilPvf).iPctRate(0) > 0 Then
                            'Sort Veh's with Impressions to the top of the list
                            tmPBDP(ilTest).sKey = "0" & tmPBDP(ilTest).sSvKey
                        End If
                        Exit For
                    End If
                Next ilTest
            End If
        Next ilLoop
    Next ilPvf
    mInitShowFields
    For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
        For ilLoop = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
            If tmPvf(ilPvf).iVefCode(ilLoop) > 0 Then
                For ilTest = LBONE To UBound(tmPBDP) - 1 Step 1
                    If (tmPvf(ilPvf).iVefCode(ilLoop) = tmPBDP(ilTest).iVefCode) And (tmPvf(ilPvf).iRdfCode(ilLoop) = tmPBDP(ilTest).iRdfCode) Then
                        smPkgSave(1, ilTest) = Trim$(Str$(tmPvf(ilPvf).iNoSpot(ilLoop)))
                        smPkgSave(2, ilTest) = gIntToStrDec(tmPvf(ilPvf).iPctRate(ilLoop), 2)
                        Exit For
                    End If
                Next ilTest
            End If
        Next ilLoop
    Next ilPvf

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
    Dim ilVef As Integer
    If smSpecSave(1) <> "" Then    'Test name
        slStr = Trim$(smSpecSave(1))
        'gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(slStr, Trim$(tgMVef(ilVef).sName), 1) = 0 Then
                If (imSelectedIndex = 0) Or (tgMVef(ilVef).iCode <> tmVef.iCode) Then
                    Beep
                    MsgBox "CPM Package Name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    imSpecBoxNo = NAMEINDEX
                    mSpecEnableBox imSpecBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        Next ilVef
    End If
    mOKName = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPkgEnableBox                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mPkgEnableBox(ilBoxNo As Integer)
'
'   mPkgEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBPkgCtrls Or ilBoxNo > UBound(tmPkgCtrls) Then
        Exit Sub
    End If
    If (imPkgRowNo < vbcPkg.Value) Or (imPkgRowNo >= vbcPkg.Value + vbcPkg.LargeChange + 1) Then
        pbcArrow.Visible = False
        lacPkgFrame.Visible = False
        Exit Sub
    End If
    lacPkgFrame.Move 0, tmPkgCtrls(PKGVEHINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15) - 30
    lacPkgFrame.Visible = True

    pbcArrow.Visible = False
    pbcArrow.Move plcPkg.Left - pbcArrow.Width - 15, plcPkg.Top + tmPkgCtrls(PKGVEHINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case PKGPERCENTINDEX 'Start/End
            edcDropDown.Width = tmPkgCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 7
            gMoveTableCtrl pbcPkg, edcDropDown, tmPkgCtrls(PKGPERCENTINDEX).fBoxX, tmPkgCtrls(PKGPERCENTINDEX).fBoxY + (imPkgRowNo - vbcPkg.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = Trim$(smPkgSave(2, imPkgRowNo))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPkgSetFocus                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mPkgSetFocus(ilBoxNo As Integer)
'
'   mPkgSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBPkgCtrls Or ilBoxNo > UBound(tmPkgCtrls) Then
        Exit Sub
    End If
    If (imPkgRowNo < vbcPkg.Value) Or (imPkgRowNo >= vbcPkg.Value + vbcPkg.LargeChange + 1) Then
        pbcArrow.Visible = False
        lacPkgFrame.Visible = False
        Exit Sub
    End If

    pbcArrow.Visible = False
    pbcArrow.Move plcPkg.Left - pbcArrow.Width - 15, plcPkg.Top + tmPkgCtrls(1).fBoxY + (imPkgRowNo - 1) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case PKGPERCENTINDEX 'Start/End
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPkgSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mPkgSetShow(ilBoxNo As Integer)
'
'   mPkgSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    pbcArrow.Visible = False
    lacPkgFrame.Visible = False
    If ilBoxNo < imLBPkgCtrls Or ilBoxNo > UBound(tmPkgCtrls) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case PKGPERCENTINDEX 'Vehicle
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
            gSetShow pbcPkg, slStr, tmPkgCtrls(ilBoxNo)
            smPkgShow(PKGPERCENTINDEX, imPkgRowNo) = tmPkgCtrls(ilBoxNo).sShow
            If Trim$(smPkgSave(2, imPkgRowNo)) <> edcDropDown.Text Then
                imPkgChg = True
                smPkgSave(2, imPkgRowNo) = edcDropDown.Text
                'mGetPkgAud imPkgRowNo
                mGetTotals
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(CPMPkg, VEHCPMPKG + ACTIVEVEH + DORMANTVEH, cbcSelect, tmPkgVehicle(), smPkgVehicleTag)
    
    
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox: CPMPkg)", CPMPkg
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
'*             Created:6/04/93       By:D. LeVine      *
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
    Dim llPvfCode As Long

    ilRet = 0
    slNameCode = tmPkgVehicle(ilSelectIndex - 1).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", CPMPkg
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmVefSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual-VEF)", CPMPkg
    On Error GoTo 0
    llPvfCode = tmVef.lPvfCode
    ReDim tmPvf(0 To 0) As PVF
    Do While llPvfCode > 0
        tmPvfSrchKey.lCode = llPvfCode
        ilRet = btrGetEqual(hmPvf, tmPvf(UBound(tmPvf)), imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", CPMPkg
        On Error GoTo 0
        llPvfCode = tmPvf(UBound(tmPvf)).lLkPvfCode
        ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
    Loop
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
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim llPvfCode As Long
    Dim ilLen As Integer
    Dim ilPvf As Integer
    Dim ilVef As Integer
    Dim ilLenTest As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilFirstFd As Integer
    Dim ilPvf1 As Integer
    Dim ilVef1 As Integer
    Dim ilLoop1 As Integer
    mSpecSetShow imSpecBoxNo
    mPkgSetShow imPkgBoxNo
    If mTestSaveFields(SHOWMSG, True) = NO Then
        mSaveRec = False
        Exit Function
    End If
    ilFound = False
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(2, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(2, ilLoop))) > 0 Then
                ilFound = True
                Exit For
            End If
        End If
    Next ilLoop
    If Not ilFound Then
        ilRet = MsgBox("At least one Vehicle must be associated with the Package", vbOKOnly + vbExclamation, "Incomplete")
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    ilRet = btrBeginTrans(hmVef, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "VEF.Btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        Else
            mInitVef
        End If
        'Delete PVF, then Create PVF
        If imSelectedIndex <> 0 Then 'NOT New selected
            For ilLoop = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
                Do
                    tmPvfSrchKey.lCode = tmPvf(ilLoop).lCode
                    ilRet = btrGetEqual(hmPvf, tmTPvf, imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilCRet = btrAbortTrans(hmPvf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
                        mSaveRec = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hmPvf)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrAbortTrans(hmPvf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
                    mSaveRec = False
                    Exit Function
                End If
            Next ilLoop
        End If
        
        mMoveCtrlToRec
        llPvfCode = 0
        For ilLoop = UBound(tmPvf) - 1 To LBound(tmPvf) Step -1
            tmPvf(ilLoop).lCode = 0
            tmPvf(ilLoop).sName = smSpecSave(1)
            tmPvf(ilLoop).lLkPvfCode = llPvfCode
            ilRet = btrInsert(hmPvf, tmPvf(ilLoop), imPvfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                ilCRet = btrAbortTrans(hmPvf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
                mSaveRec = False
                Exit Function
            End If
            llPvfCode = tmPvf(ilLoop).lCode
        Next ilLoop
        If imSelectedIndex = 0 Then 'New selected
            imAddPkg = True
            tmVef.iCode = 0  'Autoincrement
            tmVef.sType = "P" 'Package Vehicle
            tmVef.lPvfCode = llPvfCode
            tmVef.iRemoteID = tgUrf(0).iRemoteUserID
            tmVef.iAutoCode = tmVef.iCode
            ilRet = btrInsert(hmVef, tmVef, imVefRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:CPMPkg)"
        Else 'Old record-Update
            tmVef.lPvfCode = llPvfCode
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            slMsg = "mSaveRec (btrUpdate:CPMPkg)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        ilCRet = btrAbortTrans(hmPvf)
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
        mSaveRec = False
        Exit Function
    End If
    If imSelectedIndex = 0 Then 'New selected
        Do
            tmVef.iRemoteID = tgUrf(0).iRemoteUserID
            tmVef.iAutoCode = tmVef.iCode
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            slMsg = "mSaveRec (btrUpdate:CPMPkg)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPvf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later (" & Str$(ilRet) & ")", vbOKOnly + vbExclamation, "Rate Card")
            mSaveRec = False
            Exit Function
        End If
        'tgMVef(UBound(tgMVef)) = tmVef
        'ReDim Preserve tgMVef(0 To UBound(tgMVef) + 1) As VEF
        'If UBound(tgMVef) > 1 Then
        '    'ArraySortTyp fnAV(tgMVef(), 1), UBound(tgMVef) - 1, 0, LenB(tgMVef(1)), 0, -1, 0
        '    ArraySortTyp fnAV(tgMVef(), 0), UBound(tgMVef), 0, LenB(tgMVef(0)), 0, -1, 0
        'End If
        sgMVefStamp = ""
        ilRet = gObtainVef()
    Else
        ilRet = gBinarySearchVef(tmVef.iCode)
        If ilRet <> -1 Then
            tgMVef(ilRet) = tmVef
        End If
    End If
    ilRet = gVpfFind(CPMPkg, tmVef.iCode)
    'Update lengths
    ilFirstFd = False
    tmVpfSrchKey.iVefKCode = tmVef.iCode
    ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        For ilLen = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
            tmVpf.iSLen(ilLen) = 0
        Next ilLen
        ilIndex = LBound(tmVpf.iSLen)
        For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
            For ilVef = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
                ilLoop = gBinarySearchVpf(tmPvf(ilPvf).iVefCode(ilVef))
                If ilLoop <> -1 Then
                    ilFirstFd = True
                    For ilLen = LBound(tgVpf(ilLoop).iSLen) To UBound(tgVpf(ilLoop).iSLen) Step 1
                        If tgVpf(ilLoop).iSLen(ilLen) <> 0 Then
                            ilFound = False
                            For ilLenTest = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
                                If tmVpf.iSLen(ilLenTest) = tgVpf(ilLoop).iSLen(ilLen) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLenTest
                            If Not ilFound Then
                                'Test if in all other vehicles- if not don't add
                                ilFound = True
                                For ilPvf1 = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
                                    For ilVef1 = LBound(tmPvf(ilPvf1).iVefCode) To UBound(tmPvf(ilPvf1).iVefCode) Step 1
                                            ilLoop1 = gBinarySearchVpf(tmPvf(ilPvf1).iVefCode(ilVef1))
                                            If ilLoop1 <> -1 Then
                                                ilFound = False
                                                For ilLenTest = LBound(tmVpf.iSLen) To UBound(tmVpf.iSLen) Step 1
                                                    If tgVpf(ilLoop).iSLen(ilLen) = tgVpf(ilLoop1).iSLen(ilLenTest) Then
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                Next ilLenTest
                                            End If
                                        If Not ilFound Then
                                            Exit For
                                        End If
                                    Next ilVef1
                                    If Not ilFound Then
                                        Exit For
                                    End If
                                Next ilPvf1
                                If Not ilFound Then
                                    ilFound = True
                                Else
                                    ilFound = False
                                End If
                            End If
                            If Not ilFound Then
                                tmVpf.iSLen(ilIndex) = tgVpf(ilLoop).iSLen(ilLen)
                                If ilIndex = LBound(tmVpf.iSLen) Then
                                    tmVpf.iSDLen = tgVpf(ilLoop).iSDLen
                                End If
                                ilIndex = ilIndex + 1
                                If ilIndex > UBound(tmVpf.iSLen) Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilLen
                End If
                If ilFirstFd Then
                    Exit For
                End If
            Next ilVef
            If ilFirstFd Then
                Exit For
            End If
        Next ilPvf
        
        tmVpf.sGMedium = "P" 'Podcast Vehicle
        'Save PVF_Package_Vehicle
        ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
        ilRet = gBinarySearchVpf(tmVpf.iVefKCode)
        If ilRet <> -1 Then
            tgVpf(ilRet) = tmVpf
        End If
    End If
    ilFirstFd = False
    ilRet = mVffReadRec(tmVef.iCode)
    If ilRet Then
        ilRet = btrUpdate(hmVff, tmVff, imVffRecLen)
        sgVffStamp = "~"
        ilRet = gVffRead()
    End If
    'Update lengths
    ilRet = btrEndTrans(hmPvf)
    gFileChgdUpdate "vef.btr", False
    gFileChgdUpdate "vpf.btr", False
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
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
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmSpecCtrls(), TESTALLCTRLS)
    If mTestSaveFields(NOMSG, True) = YES Then  'No Then
        If (ilAltered = YES) Or (imPkgChg = True) Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & smSpecSave(1)
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
'*             Created:6/30/93       By:D. LeVine      *
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
    'Update button set if all mandatory fields have data and any field altered
    Dim ilAltered As Integer
    Dim ilLoop As Integer
    
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmSpecCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestSaveFields(NOMSG, False) = YES) And ((ilAltered = YES) Or (imPkgChg = True)) Then
        If imUpdateAllowed Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
    'cmcProof.Enabled = False
    For ilLoop = LBONE To UBound(tmPBDP) - 1 Step 1
        If Trim$(smPkgSave(1, ilLoop)) <> "" Then
            If Val(Trim$(smPkgSave(1, ilLoop))) > 0 Then
'                cmcProof.Enabled = True
                Exit For
            End If
        End If
    Next ilLoop
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecEnableBox                  *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecEnableBox(ilBoxNo As Integer)
'
'   mSpecEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Width = tmSpecCtrls(ilBoxNo).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcSpecDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcSpecDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcSpec, edcSpecDropDown, tmSpecCtrls(ilBoxNo).fBoxX, tmSpecCtrls(ilBoxNo).fBoxY
            imChgMode = True
            edcSpecDropDown.Text = smSpecSave(1)
            imChgMode = False
            edcSpecDropDown.SelStart = 0
            edcSpecDropDown.SelLength = Len(edcSpecDropDown.Text)
            edcSpecDropDown.Visible = True
            edcSpecDropDown.SetFocus
        
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetChg                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSpecSetChg(ilBoxNo As Integer, ilUseSave As Integer)
'
'   mSpecSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    Dim slInitStr As String
    If ilBoxNo < imLBSpecCtrls Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            slStr = edcSpecDropDown.Text
            gSetChgFlagStr tmVef.sName, slStr, tmSpecCtrls(ilBoxNo)
        
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetFocus                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSpecSetFocus(ilBoxNo As Integer)
'
'   mSpecSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecSetShow                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSpecSetShow(ilBoxNo As Integer)
'
'   mSpecSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBSpecCtrls) Or (ilBoxNo > UBound(tmSpecCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX
            edcSpecDropDown.Visible = False
            slStr = Trim$(edcSpecDropDown.Text)
            smSpecSave(1) = slStr
            gSetShow pbcSpec, slStr, tmSpecCtrls(ilBoxNo)
        
    End Select
    mSpecSetChg imSpecBoxNo, False
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
    Screen.MousePointer = vbDefault
    If imAddPkg Then
'        sgMVefStamp = "~'"
'        ilRet = gObtainVef()
'        sgVpfStamp = "~"    'Force read
'        ilRet = gVpfRead()
    End If

    igManUnload = YES
    Unload CPMPkg
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilMsg As Integer, ilSetBox As Integer) As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim slTotalPct As String
    If smSpecSave(1) = "" Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Name must be specified", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = NAMEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If

    slTotalPct = "0.0"
    For ilLoop = LBONE To UBound(smPkgSave, 2) - 1 Step 1
        If Trim$(smPkgSave(2, ilLoop)) <> "" Then
            slTotalPct = gAddStr(Trim$(smPkgSave(2, ilLoop)), slTotalPct)
        End If
    Next ilLoop
    If gCompNumberStr(slTotalPct, "100.00") <> 0 Then
        If ilMsg = SHOWMSG Then
            ilRes = MsgBox("Impression Percent Not Equal to 100", vbOKOnly + vbExclamation, "Incomplete")
        End If
        If ilSetBox Then
            imSpecBoxNo = NAMEINDEX
        End If
        mTestSaveFields = NO
        Exit Function
    End If

    mTestSaveFields = YES
End Function
Private Sub pbcClickFocus_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    pbcArrow.Visible = False
End Sub
Private Sub pbcPkg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilCompRow As Integer
    ilCompRow = vbcPkg.LargeChange + 1
    If UBound(smPkgSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smPkgSave, 2) + 1  'UBound(tgBvfRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = PKGPERCENTINDEX To PKGPERCENTINDEX Step 1
            If (X >= tmPkgCtrls(ilBox).fBoxX) And (X <= (tmPkgCtrls(ilBox).fBoxX + tmPkgCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmPkgCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmPkgCtrls(ilBox).fBoxY + tmPkgCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPkg.Value - 1
                    If ilRowNo >= UBound(smPkgSave, 2) Then
                        Beep
                        mPkgSetFocus imPkgBoxNo
                        Exit Sub
                    End If
                    mPkgSetShow imPkgBoxNo
                    imPkgRowNo = ilRow + vbcPkg.Value - 1
                    imPkgBoxNo = ilBox
                    mPkgEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mPkgSetFocus imPkgBoxNo
End Sub
Private Sub pbcPkg_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long
    Dim slStr As String

    mPaintPkgTitle
    ilStartRow = vbcPkg.Value '+ 1  'Top location
    ilEndRow = vbcPkg.Value + vbcPkg.LargeChange ' + 1
    If ilEndRow > UBound(smPkgSave, 2) Then
        ilEndRow = UBound(smPkgSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcPkg.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBPkgCtrls To UBound(tmPkgCtrls) Step 1
            If (ilBox = PKGVEHINDEX) And (tmPBDP(ilRow).iPkgVeh > 0) Then
                pbcPkg.ForeColor = BLUE
            End If
            If (ilBox = PKGVEHINDEX) And (tmPBDP(ilRow).iVehDormant > 0) Then
                pbcPkg.ForeColor = Red
            End If
            If (ilBox = PKGDPINDEX) And (tmPBDP(ilRow).iDPDormant > 0) Then
                pbcPkg.ForeColor = Red
            End If
            If (ilBox = PKGVEHINDEX) And (tmPBDP(ilRow).sMedium = "P") Then
                'PODCAST CPM, Show Vehicle Name in Italic font
                pbcPkg.FontItalic = True
            End If
            pbcPkg.CurrentX = tmPkgCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPkg.CurrentY = tmPkgCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(smPkgShow(ilBox, ilRow))
            pbcPkg.Print slStr
            pbcPkg.ForeColor = llColor
            pbcPkg.FontItalic = False
        Next ilBox
    Next ilRow
    For ilBox = LBONE To UBound(smTShow) Step 1
        If ilBox = LBONE Then
            pbcPkg.CurrentX = tmPkgCtrls(ilBox + PKGPERCENTINDEX - LBONE).fBoxX + fgBoxInsetX
        Else
           ' pbcPkg.CurrentX = tmPkgCtrls(ilBox + PKGPERCENTINDEX - LBONE + 1).fBoxX + fgBoxInsetX
        End If
        
        pbcPkg.CurrentY = tmPkgCtrls(1).fBoxY + (vbcPkg.LargeChange + 1) * (fgBoxGridH + 15) + 15
        If smTShow(ilBox) <> "" Then
            If Val(smTShow(ilBox)) = 100 Then
                pbcPkg.ForeColor = DARKGREEN
            Else
                pbcPkg.ForeColor = Red
            End If
        End If
        pbcPkg.Print smTShow(ilBox)
        pbcPkg.ForeColor = llColor
    Next ilBox
End Sub
Private Sub pbcPkgSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcPkgSTab.HWnd Then
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    imTabDirection = -1 'Set- Right to left
    Select Case imPkgBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imSettingValue = True
            vbcPkg.Value = 1
            imSettingValue = False
            imPkgRowNo = 1
            ilBox = PKGPERCENTINDEX
            imPkgBoxNo = ilBox
            mPkgEnableBox ilBox
            Exit Sub
        Case PKGPERCENTINDEX 'Name (first control within header)
            mPkgSetShow imPkgBoxNo
            If imPkgRowNo <= 1 Then
                If cbcSelect.Enabled Then
                    imPkgBoxNo = -1
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
                mPkgSetShow imPkgBoxNo
                pbcSpecSTab.SetFocus
                Exit Sub
            Else
'                If imSpecSave(1) <> 2 Then
'                    ilBox = PKGPERCENTINDEX
'                Else
                    ilBox = PKGPERCENTINDEX
'                End If
                imPkgRowNo = imPkgRowNo - 1
                If imPkgRowNo < vbcPkg.Value Then
                    imSettingValue = True
                    vbcPkg.Value = vbcPkg.Value - 1
                    imSettingValue = False
                End If
                imPkgBoxNo = ilBox
                mPkgEnableBox ilBox
                Exit Sub
            End If
        Case Else
            ilBox = imPkgBoxNo - 1
    End Select
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = ilBox
    mPkgEnableBox ilBox
End Sub
Private Sub pbcPkgTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcPkgTab.HWnd Then
        Exit Sub
    End If
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
    imTabDirection = 0 'Set- Left to right
    Select Case imPkgBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imPkgRowNo = UBound(tmPBDP) - 1
            imSettingValue = True
            If imPkgRowNo <= vbcPkg.LargeChange + 1 Then
                vbcPkg.Value = 1
            Else
                vbcPkg.Value = imPkgRowNo - vbcPkg.LargeChange - 1
            End If
            imSettingValue = False
            'If imSpecSave(1) <> 2 Then
            '    ilBox = PKGPERCENTINDEX
            'Else
                ilBox = PKGPERCENTINDEX
            'End If
        Case 0
            ilBox = PKGPERCENTINDEX
        Case PKGPERCENTINDEX
            mPkgSetShow imPkgBoxNo
            If imPkgRowNo + 1 >= UBound(tmPBDP) Then
                cmcDone.SetFocus
                Exit Sub
            End If
            imPkgRowNo = imPkgRowNo + 1
            If imPkgRowNo > vbcPkg.Value + vbcPkg.LargeChange Then
                imSettingValue = True
                vbcPkg.Value = vbcPkg.Value + 1
                imSettingValue = False
            End If
            ilBox = PKGPERCENTINDEX
            imPkgBoxNo = ilBox
            mPkgEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imPkgBoxNo + 1
    End Select
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = ilBox
    mPkgEnableBox ilBox
End Sub
Private Sub pbcSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = NAMEINDEX To NAMEINDEX Step 1
        If (X >= tmSpecCtrls(ilBox).fBoxX) And (X <= (tmSpecCtrls(ilBox).fBoxX + tmSpecCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmSpecCtrls(ilBox).fBoxY)) And (Y <= (tmSpecCtrls(ilBox).fBoxY + tmSpecCtrls(ilBox).fBoxH)) Then
                mSpecSetShow imSpecBoxNo
                imSpecBoxNo = ilBox
                mSpecEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSpecSetFocus imSpecBoxNo
End Sub
Private Sub pbcSpec_Paint()
    Dim ilBox As Integer
    Dim slStr As String

    For ilBox = imLBSpecCtrls To UBound(tmSpecCtrls) Step 1
        pbcSpec.CurrentX = tmSpecCtrls(ilBox).fBoxX + fgBoxInsetX
        slStr = tmSpecCtrls(ilBox).sShow
        pbcSpec.CurrentY = tmSpecCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSpec.Print tmSpecCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSpecSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-Right to left
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    Select Case imSpecBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                imPkgRowNo = 0
                imPkgBoxNo = 3
                pbcPkgTab.SetFocus
            End If
        Case NAMEINDEX
            mSpecSetShow imSpecBoxNo
            imSpecBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            If (cmcUpdate.Enabled) And (igDPNameCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub
Private Sub pbcSpecTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSpecTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    mPkgSetShow imPkgBoxNo
    imPkgBoxNo = -1
    imPkgRowNo = -1
    Select Case imSpecBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            ilBox = NAMEINDEX
        Case 0
            ilBox = NAMEINDEX
        
        Case NAMEINDEX
            imPkgRowNo = 0
            imPkgBoxNo = 3
            pbcPkgTab.SetFocus
            Exit Sub
        Case Else
            ilBox = imSpecBoxNo + 1
    End Select
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = ilBox
    mSpecEnableBox ilBox
End Sub
Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    If (Not imFirstTimeSelect) Then
        If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
            pbcSpecTab.SetFocus
        Else
            imPkgRowNo = 0
            imPkgBoxNo = 3
            pbcPkgTab.SetFocus
        End If
        Exit Sub
    End If
    imFirstTimeSelect = False
    If (imSelectedIndex = 0) And (cbcSelect.ListCount > 1) Then
        igStdPkgReturn = 0
        igStdPkgModel = 0
        sgTmpSortTag = "C" 'Show CPM Package Vehicles
        SPModel.Show vbModal
        If (igStdPkgReturn = 1) And (igStdPkgModel > 0) Then    'Done
            For ilLoop = LBound(tmPkgVehicle) To UBound(tmPkgVehicle) - 1 Step 1
                slNameCode = tmPkgVehicle(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If CInt(slCode) = igStdPkgModel Then
                    pbcPkg.Cls
                    ilRet = mReadRec(ilLoop + 1, SETFORREADONLY)
                    tmVef.sName = ""
                    tmVef.iCode = 0
                    mMoveRecToCtrl
                    mInitShow
                    mGetTotals
                    ReDim tmPvf(0 To 0) As PVF
                    pbcSpec_Paint
                    pbcPkg_Paint
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    '2/7/09: Added to handle case where focus can't be set
    On Error Resume Next
    pbcSpecSTab.SetFocus
    On Error GoTo 0
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
End Sub

Private Sub tmcInit_Timer()
    tmcInit.Enabled = False
    vbcPkg.Visible = False
    DoEvents
    vbcPkg.Visible = True
End Sub

Private Sub vbcPkg_Change()
    If imSettingValue Then
        pbcPkg.Cls
        pbcPkg_Paint
        imSettingValue = False
    Else
        mPkgSetShow imPkgBoxNo
        pbcPkg.Cls
        pbcPkg_Paint
        If (igWinStatus(RATECARDSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            mPkgEnableBox imPkgBoxNo
        End If
    End If
End Sub
Private Sub vbcPkg_GotFocus()
    mSpecSetShow imSpecBoxNo
    imSpecBoxNo = -1
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "CPM Packages"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLnTitle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintPkgTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer

    llColor = pbcPkg.ForeColor
    slFontName = pbcPkg.FontName
    flFontSize = pbcPkg.FontSize
    ilFillStyle = pbcPkg.FillStyle
    llFillColor = pbcPkg.FillColor
    pbcPkg.ForeColor = BLUE
    pbcPkg.FontBold = False
    pbcPkg.FontSize = 7
    pbcPkg.FontName = "Arial"
    pbcPkg.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    pbcPkg.Line (tmPkgCtrls(PKGVEHINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGVEHINDEX).fBoxW + 15, tmPkgCtrls(PKGVEHINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGVEHINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGVEHINDEX).fBoxW - 15, tmPkgCtrls(PKGVEHINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Vehicle"
    pbcPkg.Line (tmPkgCtrls(PKGDPINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGDPINDEX).fBoxW + 15, tmPkgCtrls(PKGDPINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGDPINDEX).fBoxX, 30)-Step(tmPkgCtrls(PKGDPINDEX).fBoxW - 15, tmPkgCtrls(PKGDPINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcPkg.CurrentX = tmPkgCtrls(PKGDPINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Ad Location"
    pbcPkg.Line (tmPkgCtrls(PKGPERCENTINDEX).fBoxX - 15, 15)-Step(tmPkgCtrls(PKGPERCENTINDEX).fBoxW + 15, tmPkgCtrls(PKGPERCENTINDEX).fBoxH + 15), BLUE, B
    pbcPkg.CurrentX = tmPkgCtrls(PKGPERCENTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = 15
    pbcPkg.Print "Impression %"

    ilLineCount = 0
    llTop = tmPkgCtrls(1).fBoxY
    Do
        For ilLoop = imLBPkgCtrls To UBound(tmPkgCtrls) Step 1
            pbcPkg.Line (tmPkgCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmPkgCtrls(ilLoop).fBoxW + 15, tmPkgCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop < PKGPERCENTINDEX) Then
                pbcPkg.Line (tmPkgCtrls(ilLoop).fBoxX, llTop)-Step(tmPkgCtrls(ilLoop).fBoxW - 15, tmPkgCtrls(ilLoop).fBoxH - 15), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmPkgCtrls(1).fBoxH + 15
    Loop While llTop + tmPkgCtrls(1).fBoxH + tmPkgCtrls(1).fBoxH + 30 < pbcPkg.Height
    vbcPkg.LargeChange = ilLineCount - 1
    llTop = llTop + 30
    
    pbcPkg.Line (tmPkgCtrls(PKGPERCENTINDEX).fBoxX - 15, llTop)-Step(tmPkgCtrls(PKGPERCENTINDEX).fBoxW + 15, tmPkgCtrls(PKGPERCENTINDEX).fBoxH + 15), BLUE, B
    pbcPkg.Line (tmPkgCtrls(PKGPERCENTINDEX).fBoxX, llTop + 15)-Step(tmPkgCtrls(PKGPERCENTINDEX).fBoxW - 15, tmPkgCtrls(PKGPERCENTINDEX).fBoxH - 15), LIGHTYELLOW, BF

    pbcPkg.FontSize = flFontSize
    pbcPkg.FontName = slFontName
    pbcPkg.FontSize = flFontSize
    pbcPkg.ForeColor = llColor
    pbcPkg.FontBold = True

    pbcPkg.CurrentX = tmPkgCtrls(1).fBoxX + 15  'fgBoxInsetX
    pbcPkg.CurrentY = llTop '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPkg.Print "Total:"
    lacCover.Top = llTop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVffReadRec                     *
'*                                                     *
'*             Created:6/23/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVffReadRec(ilVefCode As Integer) As Integer
'
'   iRet = mVffReadRec()
'   Where:
'       ilVefCode(I) - Vehicle Code
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    If ilVefCode > 0 Then
        tmVffSrchKey1.iCode = ilVefCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        mVffReadRec = False
        Exit Function
    End If
    If ilRet <> BTRV_ERR_NONE Then
        'Add Record
        tmVff.iCode = 0
        tmVff.iVefCode = ilVefCode  'tmVef.iCode
        tmVff.sGroupName = ""
        tmVff.sWegenerExportID = ""
        tmVff.sOLAExportID = ""
        tmVff.iLiveCompliantAdj = 5
        tmVff.iUstCode = 0
        tmVff.iUrfCode = tgUrf(0).iCode
        'tmVff.sXDXMLForm = "S"
        If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
            tmVff.sXDXMLForm = "S"
        Else
            If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
                tmVff.sXDXMLForm = "P"
            Else
                tmVff.sXDXMLForm = ""
            End If
        End If
        tmVff.sXDISCIPrefix = ""
        tmVff.sXDProgCodeID = ""
        tmVff.sXDSaveCF = "Y"
        tmVff.sXDSaveHDD = "N"
        tmVff.sXDSaveNAS = "N"
        'tmVff.sUnused = ""
        tmVff.iCwfCode = 0
        tmVff.sAirWavePrgID = ""
        tmVff.sExportAirWave = ""
        tmVff.sExportNYESPN = ""
        tmVff.sPledgeVsAir = "N"
        tmVff.sFedDelivery(0) = ""
        tmVff.sFedDelivery(1) = ""
        tmVff.sFedDelivery(2) = ""
        tmVff.sFedDelivery(3) = ""
        tmVff.sFedDelivery(4) = ""
        'tmVff.sFedDelivery(5) = ""
        gPackDate "1/1/1990", tmVff.iLastAffExptDate(0), tmVff.iLastAffExptDate(1)
        tmVff.sMoveSportToNon = "N"
        tmVff.sMoveSportToSport = "N"
        tmVff.sMoveNonToSport = "N"
        tmVff.sMergeTraffic = "S"
        tmVff.sMergeAffiliate = "S"
        tmVff.sMergeWeb = "S"
        tmVff.sPledgeByEvent = "N"
        tmVff.lPledgeHdVtfCode = 0
        tmVff.lPledgeFtVtfCode = 0
        tmVff.iPledgeClearance = 0
        tmVff.sExportEncoESPN = "N"
        tmVff.sWebName = ""
        tmVff.lSeasonGhfCode = 0
        tmVff.iMcfCode = 0
        tmVff.sExportAudio = "N"
        tmVff.sExportMP2 = "N"
        tmVff.sExportCnCSpot = "N"
        tmVff.sExportEnco = "N"
        tmVff.sExportCnCNetInv = "N"
        tmVff.sIPumpEventTypeOV = ""
        tmVff.sExportIPump = "N"
        tmVff.sAddr4 = ""
        tmVff.lBBOpenCefCode = 0
        tmVff.lBBCloseCefCode = 0
        tmVff.lBBBothCefCode = 0
        tmVff.sXDSISCIPrefix = ""
        tmVff.sXDSSaveCF = "Y"
        tmVff.sXDSSaveHDD = "N"
        tmVff.sXDSSaveNAS = "N"
        tmVff.sMGsOnWeb = "N"   '"Y"
        tmVff.sReplacementOnWeb = "N"   '"Y"
        tmVff.sExportMatrix = "N"
        tmVff.sSentToXDSStatus = "N"
        tmVff.sStationComp = "N"
        tmVff.sExportSalesForce = "N"
        tmVff.sExportEfficio = "N"
        tmVff.sExportJelli = "N"
        tmVff.sOnXMLInsertion = "N"
        tmVff.sOnInsertions = "N"
        tmVff.sPostLogSource = "N"
        tmVff.sExportTableau = "N"
        tmVff.sStationPassword = ""
        tmVff.sHonorZeroUnits = "N"
        tmVff.sHideCommOnLog = "N"
        tmVff.sHideCommOnWeb = "N"
        tmVff.iConflictWinLen = 0
        tmVff.sACT1LineupCode = ""
        tmVff.sPrgmmaticAllow = "N"
        tmVff.sSalesBrochure = ""
        tmVff.sCartOnWeb = "N"
        tmVff.sDefaultAudioType = "R"
        tmVff.iLogExptArfCode = 0
        'tmVff.sUnused = ""
        tmVff.sASICallLetters = ""
        tmVff.sASIBand = ""
        tmVff.sExportCustom = "" 'TTP 9992
        ilRet = btrInsert(hmVff, tmVff, imVffRecLen, INDEXKEY0)
        On Error GoTo mVffReadRecErr
        gBtrvErrorMsg ilRet, "mVffReadRec (btrInsert)", CPMPkg
        On Error GoTo 0
    End If
    mVffReadRec = True
    Exit Function
mVffReadRecErr:
    On Error GoTo 0
    mVffReadRec = False
    Exit Function
End Function
