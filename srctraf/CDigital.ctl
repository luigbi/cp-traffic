VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl CDigital 
   Appearance      =   0  'Flat
   ClientHeight    =   5895
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   9345
   ClipControls    =   0   'False
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   9345
   Begin VB.PictureBox pbcComingSoon 
      Height          =   1125
      Left            =   405
      Picture         =   "CDigital.ctx":0000
      ScaleHeight     =   1065
      ScaleWidth      =   5445
      TabIndex        =   10
      Top             =   1770
      Width           =   5505
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5370
      Width           =   45
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8010
      Top             =   5385
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   1125
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   330
      Width           =   45
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   990
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2595
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "CDigital.ctx":129C2
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   90
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5640
      Width           =   75
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
      Left            =   4935
      TabIndex        =   7
      Top             =   5460
      Visible         =   0   'False
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
      Left            =   3360
      TabIndex        =   6
      Top             =   5460
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDigital 
      Height          =   5130
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   9049
      _Version        =   393216
      Rows            =   31
      Cols            =   28
      FixedRows       =   5
      FixedCols       =   2
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   28
   End
   Begin VB.Label plcScreen 
      Caption         =   "Digital Content"
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
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "CDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CDigital.ctl on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imPopReqd                     imSelectedIndex               imComboBoxIndex           *
'*  imBypassSetting               imTypeRowNo                                             *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mPopulate                                                                             *
'*                                                                                        *
'* Public Property Procedures (Marked)                                                    *
'*  Enabled(Let)                  Verify(Get)                                             *
'*                                                                                        *
'* Public User-Defined Events (Marked)                                                    *
'*  SetSave                                                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CDigital.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

Public Event SetSave(ilStatus As Integer) 'VBC NR

'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer

Dim smNowDate As String
Dim lmNowDate As Long
Dim lmFirstAllowedChgDate As Long

Dim imCtrlVisible As Integer
Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim lmTopRow As Long
Dim imInitNoRows As Integer


Const SITEINDEX = 2   '1
Const POSITIONINDEX = 4   '2
Const FILETYPEINDEX = 6 '3
Const ADTYPEINDEX = 8
Const SIZEINDEX = 10 '7
Const STARTDATEINDEX = 12 '8
Const ENDDATEINDEX = 14
Const NOWEEKSINDEX = 16
Const CPMRATEINDEX = 18
Const TOTALIMPINDEX = 20
Const GROSSTOTALINDEX = 22
Const COMMENTINDEX = 24
Const FILECODEINDEX = 26
Const STATUSINDEX = 27







Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************


    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub






Private Sub edcDropDown_Change()
    grdDigital.CellForeColor = vbBlack
End Sub


Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                                                   *
'******************************************************************************************

    Dim ilKey As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case lmEnableCol
    End Select
End Sub


Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Exit Sub
    End If
    imFirstActivate = False
    imUpdateAllowed = igUpdateAllowed
    'If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    If Not imUpdateAllowed Then
        grdDigital.Enabled = False
    Else
        grdDigital.Enabled = True
    End If
    gShowBranner imUpdateAllowed
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If lmEnableCol > 0 Then
            mEnableBox
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub


Private Sub grdDigital_EnterCell()
    mSetShow
End Sub

Private Sub grdDigital_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
'    If y < grdDigital.RowHeight(0) Then
'        mSortCol grdDigital.Col
'        Exit Sub
'    End If
    'Determine row and col mouse up onto
    On Error GoTo grdDigitalErr
    pbcArrow.Visible = False
    ilCol = grdDigital.MouseCol
    ilRow = grdDigital.MouseRow
    If ilCol < grdDigital.FixedCols Then
        grdDigital.Redraw = True
        Exit Sub
    End If
    If ilRow < grdDigital.FixedRows Then
        grdDigital.Redraw = True
        Exit Sub
    End If
    If ilRow Mod 2 = 0 Then
        ilRow = ilRow + 1
    End If
    If grdDigital.ColWidth(ilCol) <= 15 Then
        grdDigital.Redraw = True
        Exit Sub
    End If
    If grdDigital.RowHeight(ilRow) <= 15 Then
        grdDigital.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdDigital.TopRow
    DoEvents
    If grdDigital.TextMatrix(ilRow, SITEINDEX) = "" Then
        grdDigital.Redraw = True
        Exit Sub
    End If
    grdDigital.Col = ilCol
    grdDigital.Row = ilRow
    If Not mColOk() Then
        grdDigital.Redraw = True
        Exit Sub
    End If
    grdDigital.Redraw = True
    mEnableBox
    On Error GoTo 0
    Exit Sub
grdDigitalErr:
    On Error GoTo 0
    If (lmEnableRow >= grdDigital.FixedRows) And (lmEnableRow < grdDigital.Rows) Then
        grdDigital.Row = lmEnableRow
        grdDigital.Col = lmEnableCol
        mSetFocus
    End If
    grdDigital.Redraw = False
    grdDigital.Redraw = True
    Exit Sub
End Sub

Private Sub grdDigital_Scroll()
    pbcArrow.Visible = False
    If grdDigital.RowIsVisible(grdDigital.Row) Then
        pbcArrow.Move grdDigital.Left - pbcArrow.Width - 30, grdDigital.Top + grdDigital.RowPos(grdDigital.Row) + (grdDigital.RowHeight(grdDigital.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
    End If
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         slNameCode                    slName                    *
'*  slCode                        ilLoop                        slDaypart                 *
'*  slLineNo                      slStr                                                   *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdDigital, grdDigital, vbHourglass
    imFirstActivate = True
    imTerminate = False
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    imBypassFocus = False
    imSettingValue = False
    imStartMode = True
    imChgMode = False
    imBSMode = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imCtrlVisible = False
    imCtrlVisible = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate + 1
    mInitBox

    Screen.MousePointer = vbDefault
    gSetMousePointer grdDigital, grdDigital, vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdDigital, grdDigital, vbDefault
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilLoop                        ilCol                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim llRow As Long
    'flTextHeight = pbcDates.TextHeight("1") - 35

    grdDigital.Move 180, 120, Width - pbcArrow.Width - 120
    grdDigital.Height = Height - grdDigital.Top - 120
    grdDigital.Redraw = False
    grdDigital.RowHeight(0) = 2 * fgBoxGridH + 15
    grdDigital.Rows = grdDigital.FixedRows + 2
    llRow = grdDigital.FixedRows
    Do
        If llRow + 1 > grdDigital.Rows Then
            grdDigital.AddItem ""
            grdDigital.RowHeight(grdDigital.Rows - 1) = fgBoxGridH
            grdDigital.AddItem ""
            grdDigital.RowHeight(grdDigital.Rows - 1) = 15
            mInitNew llRow
        End If
        llRow = llRow + 2
    Loop While grdDigital.RowIsVisible(llRow - 2)
    imInitNoRows = grdDigital.Rows
    mGridDigitalLayout
    mGridDigitalColumnWidths
    mGridDigitalColumns
    mClearCtrlFields
    pbcComingSoon.Left = Width / 2 - pbcComingSoon.Width / 2
    'pbcComingSoon.Top = Height / 3 - pbcComingSoon.Height / 2
    pbcComingSoon.Top = grdDigital.Top + pbcComingSoon.Height / 2
    'gGrid_IntegralHeight grdDigital, CInt(fgBoxGridH) + 15
    'grdDigital.Height = grdDigital.Height + 45
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mTerminate
'   Where:
'


    Screen.MousePointer = vbDefault
    gSetMousePointer grdDigital, grdDigital, vbDefault
    igManUnload = YES
    'Unload CDigital
    'Set CDigital = Nothing   'Remove data segment
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub








Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        Do
            ilNext = False
            Select Case grdDigital.Col
                Case SITEINDEX
                    mSetShow
                    Exit Sub
                Case Else
                    grdDigital.Col = grdDigital.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdDigital.Row = grdDigital.FixedRows
        grdDigital.Col = grdDigital.FixedCols
        Do
            If mColOk() Then
                Exit Do
            Else
                grdDigital.Col = grdDigital.Col + 1
            End If
        Loop
    End If
    mEnableBox
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        Do
            ilNext = False
            Select Case grdDigital.Col
                Case COMMENTINDEX
                    mSetShow
                    Exit Sub
                Case Else
                    grdDigital.Col = grdDigital.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdDigital.Row = grdDigital.FixedRows
        grdDigital.Col = grdDigital.FixedCols
        Do
            If mColOk() Then
                Exit Do
            Else
                grdDigital.Col = grdDigital.Col + 1
            End If
        Loop
    End If
    mEnableBox
End Sub





Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub







Private Sub mClearCtrlFields()
    Dim ilCol As Integer
    Dim llRow As Long


    lmEnableRow = -1
    grdDigital.Redraw = False
    If grdDigital.Rows > imInitNoRows Then
        For llRow = grdDigital.Rows - 1 To imInitNoRows Step -1
            grdDigital.RemoveItem llRow
        Next llRow
    End If
    For llRow = grdDigital.FixedRows To grdDigital.Rows - 1 Step 2
        For ilCol = 0 To grdDigital.Cols - 1 Step 1
            grdDigital.TextMatrix(llRow, ilCol) = ""
        Next ilCol
    Next llRow
    grdDigital.Redraw = True

End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
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

    'RaiseEvent SetSave(True)

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
Private Sub mEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilLang                        slNameCode                *
'*  slCode                        ilCode                        ilRet                     *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (grdDigital.Row < grdDigital.FixedRows) Or (grdDigital.Row >= grdDigital.Rows) Or (grdDigital.Col < grdDigital.FixedCols) Or (grdDigital.Col >= grdDigital.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdDigital.Row
    lmEnableCol = grdDigital.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdDigital.Left - pbcArrow.Width - 30, grdDigital.Top + grdDigital.RowPos(grdDigital.Row) + (grdDigital.RowHeight(grdDigital.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True

    Select Case grdDigital.Col
        Case SITEINDEX
        Case POSITIONINDEX
        Case FILETYPEINDEX
        Case ADTYPEINDEX
        Case SIZEINDEX
        Case STARTDATEINDEX
        Case ENDDATEINDEX
        Case NOWEEKSINDEX
        Case CPMRATEINDEX
        Case TOTALIMPINDEX
        Case GROSSTOTALINDEX
        Case COMMENTINDEX
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilNoGames                     slStr                         ilOrigUpper               *
'*  ilLoop                        llRow                         llSvRow                   *
'*  llSvCol                                                                               *
'******************************************************************************************


    pbcArrow.Visible = False
    If (lmEnableRow >= grdDigital.FixedRows) And (lmEnableRow < grdDigital.Rows) Then
        Select Case lmEnableCol
            Case SITEINDEX
            Case POSITIONINDEX
            Case FILETYPEINDEX
            Case ADTYPEINDEX
            Case SIZEINDEX
            Case STARTDATEINDEX
            Case ENDDATEINDEX
            Case NOWEEKSINDEX
            Case CPMRATEINDEX
            Case TOTALIMPINDEX
            Case GROSSTOTALINDEX
            Case COMMENTINDEX
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdDigital.Row < grdDigital.FixedRows) Or (grdDigital.Row >= grdDigital.Rows) Or (grdDigital.Col < grdDigital.FixedCols) Or (grdDigital.Col >= grdDigital.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdDigital.Col - 1 Step 1
        llColPos = llColPos + grdDigital.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdDigital.ColWidth(grdDigital.Col)
    ilCol = grdDigital.Col
    Do While ilCol < grdDigital.Cols - 1
        If (Trim$(grdDigital.TextMatrix(grdDigital.Row - 1, grdDigital.Col)) <> "") And (Trim$(grdDigital.TextMatrix(grdDigital.Row - 1, grdDigital.Col)) = Trim$(grdDigital.TextMatrix(grdDigital.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdDigital.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdDigital.Col
        Case SITEINDEX
        Case POSITIONINDEX
        Case FILETYPEINDEX
        Case ADTYPEINDEX
        Case SIZEINDEX
        Case STARTDATEINDEX
        Case ENDDATEINDEX
        Case NOWEEKSINDEX
        Case CPMRATEINDEX
        Case TOTALIMPINDEX
        Case GROSSTOTALINDEX
        Case COMMENTINDEX
    End Select
End Sub



Private Sub mGridDigitalLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    'Layout Fixed Rows:0=>Edge; 1=>Blue border; 2=>Column Title 1; 3=Column Title 2; 4=>Blue border
    '       Rows: 5=>input; 6=>blue row line; 7=>Input; 8=>blue row line
    'Layout Fixed Columns: 0=>Edge; 1=Blue border;
    '       Columns: 2=>Input; 3=>Blue column line; 4=>Input; 5=>Blue Column;....
    grdDigital.RowHeight(0) = 15
    grdDigital.RowHeight(1) = 15
    grdDigital.RowHeight(2) = 180
    grdDigital.RowHeight(3) = 180
    grdDigital.RowHeight(4) = 15
    For ilRow = grdDigital.FixedRows To grdDigital.Rows - 1 Step 2
        grdDigital.RowHeight(ilRow) = fgBoxGridH
        grdDigital.Row = ilRow
        For ilCol = 0 To grdDigital.Cols - 1 Step 1
            grdDigital.ColAlignment(ilCol) = flexAlignLeftCenter
            grdDigital.Col = ilCol
            grdDigital.CellBackColor = vbWhite
        Next ilCol
        grdDigital.RowHeight(ilRow + 1) = 15
    Next ilRow

    'For ilCol = 0 To grdDigital.Cols - 1 Step 1
    '    grdDigital.ColAlignment(ilCol) = flexAlignLeftCenter
    'Next ilCol
    grdDigital.ColWidth(0) = 15
    grdDigital.ColWidth(1) = 15
    For ilCol = grdDigital.FixedCols + 1 To grdDigital.Cols - 1 Step 2
        grdDigital.ColWidth(ilCol) = 15
    Next ilCol
    'Horizontal Blue Border Lines
    grdDigital.Row = 1
    For ilCol = 1 To grdDigital.Cols - 1 Step 1
        grdDigital.Col = ilCol
        grdDigital.CellBackColor = vbBlue
    Next ilCol
    grdDigital.Row = 4
    For ilCol = 1 To grdDigital.Cols - 1 Step 1
        grdDigital.Col = ilCol
        grdDigital.CellBackColor = vbBlue
    Next ilCol
    'Horizontal Blue lines
    For ilRow = grdDigital.FixedRows + 1 To grdDigital.Rows - 1 Step 2
        grdDigital.Row = ilRow
        For ilCol = 1 To grdDigital.Cols - 1 Step 1
            grdDigital.Col = ilCol
            grdDigital.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Border Lines
    grdDigital.Col = 1
    For ilRow = 1 To grdDigital.Rows - 1 Step 1
        grdDigital.Row = ilRow
        grdDigital.CellBackColor = vbBlue
    Next ilRow
    grdDigital.Col = 1
    For ilRow = 1 To grdDigital.Rows - 1 Step 1
        grdDigital.Row = ilRow
        grdDigital.CellBackColor = vbBlue
    Next ilRow
    ''Set color in fix area to white
    'grdDigital.Col = 2
    'For ilRow = grdDigital.FixedRows To grdDigital.Rows - 1 Step 2
    '    grdDigital.Row = ilRow
    '    grdDigital.CellBackColor = vbWhite
    'Next ilRow

    'Vertical Blue Lines
    For ilCol = grdDigital.FixedCols + 1 To grdDigital.Cols - 1 Step 2
        grdDigital.Col = ilCol
        For ilRow = 1 To grdDigital.Rows - 1 Step 1
            grdDigital.Row = ilRow
            grdDigital.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
End Sub



Private Sub mGridDigitalColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdDigital.Row = 2
    grdDigital.Col = SITEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Site or"
    grdDigital.Row = 3
    grdDigital.Col = SITEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Channel"
    grdDigital.Row = 2
    grdDigital.Col = POSITIONINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Position"
    grdDigital.Row = 3
    grdDigital.Col = POSITIONINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = ""
    grdDigital.Row = 2
    grdDigital.Col = FILETYPEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "File"
    grdDigital.Row = 3
    grdDigital.Col = FILETYPEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Type"
    grdDigital.Row = 2
    grdDigital.Col = ADTYPEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Ad"
    grdDigital.Row = 3
    grdDigital.Col = ADTYPEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Type"
    grdDigital.Row = 2
    grdDigital.Col = SIZEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Size or"
    grdDigital.Row = 3
    grdDigital.Col = SIZEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Format"
    grdDigital.Row = 2
    grdDigital.Col = STARTDATEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Start"
    grdDigital.Row = 3
    grdDigital.Col = STARTDATEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Date"
    grdDigital.Row = 2
    grdDigital.Col = ENDDATEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "End"
    grdDigital.Row = 3
    grdDigital.Col = ENDDATEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Date"
    grdDigital.Row = 2
    grdDigital.Col = NOWEEKSINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "# Weeks"
    grdDigital.Row = 3
    grdDigital.Col = NOWEEKSINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = ""
    grdDigital.Row = 2
    grdDigital.Col = CPMRATEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "CPM"
    grdDigital.Row = 3
    grdDigital.Col = CPMRATEINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Rate"
    grdDigital.Row = 2
    grdDigital.Col = TOTALIMPINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Total"
    grdDigital.Row = 3
    grdDigital.Col = TOTALIMPINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Impressions"
    grdDigital.Row = 2
    grdDigital.Col = GROSSTOTALINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Gross"
    grdDigital.Row = 3
    grdDigital.Col = GROSSTOTALINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "Total"
    grdDigital.Row = 2
    grdDigital.Col = COMMENTINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = "C"
    grdDigital.Row = 3
    grdDigital.Col = COMMENTINDEX
    grdDigital.CellFontBold = False
    grdDigital.CellFontName = "Arial"
    grdDigital.CellFontSize = 6.75
    grdDigital.CellForeColor = vbBlue
    grdDigital.CellBackColor = vbWhite
    grdDigital.TextMatrix(grdDigital.Row, grdDigital.Col) = ""

End Sub

Private Sub mGridDigitalColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdDigital.ColWidth(FILECODEINDEX) = 0
    grdDigital.ColWidth(STATUSINDEX) = 0
    grdDigital.ColWidth(SITEINDEX) = 0.15 * grdDigital.Width
    grdDigital.ColWidth(POSITIONINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(FILETYPEINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(ADTYPEINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(SIZEINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(STARTDATEINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(ENDDATEINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(NOWEEKSINDEX) = 0.06 * grdDigital.Width
    grdDigital.ColWidth(CPMRATEINDEX) = 0.05 * grdDigital.Width
    grdDigital.ColWidth(TOTALIMPINDEX) = 0.1 * grdDigital.Width
    grdDigital.ColWidth(GROSSTOTALINDEX) = 0.1 * grdDigital.Width
    grdDigital.ColWidth(COMMENTINDEX) = 0.015 * grdDigital.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdDigital.Width
    For ilCol = 0 To grdDigital.Cols - 1 Step 1
        llWidth = llWidth + grdDigital.ColWidth(ilCol)
        If (grdDigital.ColWidth(ilCol) > 15) And (grdDigital.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdDigital.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdDigital.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdDigital.Width
            For ilCol = 0 To grdDigital.Cols - 1 Step 1
                If (grdDigital.ColWidth(ilCol) > 15) And (grdDigital.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdDigital.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdDigital.FixedCols To grdDigital.Cols - 1 Step 1
                If grdDigital.ColWidth(ilCol) > 15 Then
                    ilColInc = grdDigital.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdDigital.ColWidth(ilCol) = grdDigital.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub









Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    mColOk = True
    If grdDigital.ColWidth(grdDigital.Col) <= 15 Then
        mColOk = False
        Exit Function
    End If
    If grdDigital.RowHeight(grdDigital.Row) <= 15 Then
        mColOk = False
        Exit Function
    End If
    If grdDigital.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    If grdDigital.CellForeColor = vbRed Then
        mColOk = False
        Exit Function
    End If

End Function
























Public Sub Action(ilType As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Select Case ilType
        Case 1  'Clear Focus
            mSetShow
            pbcArrow.Visible = False
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            If ilRet = 0 Then
                Form_Load
                Form_Activate
                Select Case Contract.tscLine.SelectedItem.Index
                    Case imTabMap(TABMULTIMEDIA)    '1  'Multi-Media
                    'Case 2  'Digital
                    '    mInit
                    Case imTabMap(TABNTR)    '3  'NTR
                    Case imTabMap(TABAIRTIME)    '4  'Air Time
                    Case imTabMap(TABPODCASTCPM)    'Podcast CPM
                    Case imTabMap(TABMERCH)    '  'Merchandising
                    Case imTabMap(TABPROMO)    '6  'Promotional
                    Case imTabMap(TABINSTALL)    '7  'Installment
                End Select
            End If
        Case 3  'terminate function
            mSetShow
            pbcArrow.Visible = False
            cmcCancel_Click
        Case 4  'Clear
            If imInitNoRows > 0 Then
                mClearCtrlFields
            End If
            Screen.MousePointer = vbDefault
            gSetMousePointer grdDigital, grdDigital, vbDefault
        Case 5  'Save
            mSetShow
            pbcArrow.Visible = False
            mSaveRec
    End Select
    Exit Sub
UserControlErr:
    ilRet = 1
    Resume Next
End Sub
Public Property Let Enabled(ilState As Integer) 'VBC NR
    UserControl.Enabled = ilState 'VBC NR
    PropertyChanged "Enabled" 'VBC NR
End Property 'VBC NR

Private Sub UserControl_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Form_MouseUp Button, Shift, X, Y
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mInitNew(llRowNo As Long)
    Dim ilCol As Integer
    Dim llRow As Long

    For ilCol = grdDigital.FixedCols To grdDigital.Cols - 1 Step 1
        grdDigital.TextMatrix(llRowNo, ilCol) = ""
    Next ilCol
    'grdDigital.Row = llRowNo
    'grdDigital.Col = 1
    'grdDigital.CellBackColor = vbWhite
    'Horizontal Line
    grdDigital.Row = llRowNo + 1
    For ilCol = 1 To grdDigital.Cols - 1 Step 1
        grdDigital.Col = ilCol
        grdDigital.CellBackColor = vbBlue
    Next ilCol
    'Vertical Lines
    grdDigital.Col = 1
    For llRow = llRowNo To llRowNo + 1 Step 1
        grdDigital.Row = llRow
        grdDigital.CellBackColor = vbBlue
    Next llRow
    grdDigital.Col = 3
    For llRow = llRowNo To llRowNo + 1 Step 1
        grdDigital.Row = llRow
        grdDigital.CellBackColor = vbBlue
    Next llRow
    For ilCol = grdDigital.FixedCols + 1 To grdDigital.Cols - 1 Step 2
        grdDigital.Col = ilCol
        For llRow = llRowNo To llRowNo + 1 Step 1
            grdDigital.Row = llRow
            grdDigital.CellBackColor = vbBlue
        Next llRow
    Next ilCol
    'Set Fix area Column to white
    grdDigital.Col = 2
    grdDigital.Row = llRowNo
    grdDigital.CellBackColor = vbWhite
    If grdDigital.RowHeight(grdDigital.TopRow) <= 15 Then
        grdDigital.TopRow = grdDigital.TopRow + 1
    End If
End Sub

Private Sub mSaveRec()

End Sub

Public Property Get Verify() As Integer 'VBC NR
    pbcArrow.Visible = False 'VBC NR
    If imUpdateAllowed Then 'VBC NR
        'Add call to mTestFields
        Verify = True 'VBC NR
    Else 'VBC NR
        Verify = True 'VBC NR
    End If 'VBC NR
End Property 'VBC NR

