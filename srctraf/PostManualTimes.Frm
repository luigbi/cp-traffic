VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PostManualTimes 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   10230
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
   ScaleHeight     =   5040
   ScaleWidth      =   10230
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5520
      TabIndex        =   20
      Top             =   4665
      Width           =   945
   End
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "PostManualTimes.frx":0000
      Left            =   165
      List            =   "PostManualTimes.frx":0002
      TabIndex        =   19
      Top             =   510
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.ListBox lbcCopyNm 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "PostManualTimes.frx":0004
      Left            =   1080
      List            =   "PostManualTimes.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3255
      Visible         =   0   'False
      Width           =   4425
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
      Left            =   2160
      Picture         =   "PostManualTimes.frx":0008
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5685
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "PostManualTimes.frx":0102
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "PostManualTimes.frx":0DC0
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox plcCalendar 
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   3105
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   30
         Picture         =   "PostManualTimes.frx":10CA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   240
            Left            =   510
            TabIndex        =   13
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   16
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   510
      Top             =   4635
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "PostManualTimes.frx":3EE4
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   10485
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   345
      Width           =   30
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1245
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   930
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
      Left            =   11115
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   15
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3810
      TabIndex        =   8
      Top             =   4665
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTimes 
      Height          =   3810
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   675
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   6720
      _Version        =   393216
      Rows            =   10
      Cols            =   10
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
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
      _Band(0).Cols   =   10
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   90
      Picture         =   "PostManualTimes.frx":41EE
      Top             =   255
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Post Times"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   9810
   End
End
Attribute VB_Name = "PostManualTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PostManualTimes.Frm on Wed 6/17/09 @ 12:
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSbfSrchKey2                 smMissedCodeTag                                         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mMoveCtrlToRec                mInitCif                                                *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostManualTimes.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas

Dim hmCif As Integer
Dim tmCif As CIF        'Rvf record image
Dim tmCifSrchKey As LONGKEY0
Dim imCifRecLen As Integer        'RvF record length

Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length


Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim imChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer
Dim imUpdateAllowed As Integer
Dim imChgMode As Integer
Dim imLbcMouseDown As Integer
Dim imLbcArrowSetting As Integer
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imBypassFocus As Integer
Dim smMissedTime As String
Dim imStatus As Integer     '0=Paid; 1=Bonus
Dim imFirstActivate As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

Dim tmCopyNmCode() As SORTCODE
Dim smCopyNmCodeTag As String

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Const WKDATEINDEX = 0
Const DAYPARTTIMESINDEX = 1
Const DAYPARTDAYSINDEX = 2
Const LENGTHINDEX = 3
Const STATUSINDEX = 4
Const DATEINDEX = 5
Const TIMEINDEX = 6
Const ISCIINDEX = 7
Const SPOTINDEX = 8
Const SORTINDEX = 9


Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub

Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub

Private Sub cmcCancel_Click()
    igPostManualTimes = 0
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imChg Then
        'If MsgBox("Retain all changes (Dates/Times not saved Until press Save on Manual Posting screen)?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
            igPostManualTimes = 1
        'End If
    Else
        igPostManualTimes = 0
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDropDown_Click()
    Select Case lmEnableCol
        Case DATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case TIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case ISCIINDEX
            lbcCopyNm.Visible = Not lbcCopyNm.Visible
    End Select
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer

    Select Case lmEnableCol
        Case DATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case TIMEINDEX
        Case ISCIINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcCopyNm, imBSMode, slStr)
    End Select
    grdTimes.CellForeColor = vbBlack

End Sub

Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case DATEINDEX
        Case TIMEINDEX
        Case ISCIINDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                         slStr                                                   *
'******************************************************************************************

    Dim ilKey As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer

    ilKey = KeyAscii

    Select Case lmEnableCol
        Case DATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case TIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case ISCIINDEX
    End Select

End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String

    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmEnableCol
            Case DATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case TIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case ISCIINDEX
                gProcessArrowKey Shift, KeyCode, lbcCopyNm, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmEnableCol
            Case DATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case TIMEINDEX
            Case ISCIINDEX
        End Select
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imDoubleClickName = False
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    PostManualTimes.Refresh
    Me.KeyPreview = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_Initialize()
    Me.Width = (CLng(50) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(75) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone PostManualTimes
    DoEvents
    mSetControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
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
    tmcStart.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmCopyNmCode

    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    
    Set PostManualTimes = Nothing   'Remove data segment
End Sub


Private Sub grdTimes_EnterCell()
    mSetShow
End Sub

Private Sub grdTimes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdTimes.TopRow
    grdTimes.Redraw = False
End Sub


Private Sub grdTimes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imIgnoreScroll = False
    'If y < grdTimes.RowHeight(0) Then
    '    grdTimes.Col = grdTimes.MouseCol
    '    mSortCol grdTimes.Col
    '    Exit Sub
    'End If
    pbcArrow.Visible = False
    ilCol = grdTimes.MouseCol
    ilRow = grdTimes.MouseRow
    If ilCol < grdTimes.FixedCols Then
        grdTimes.Redraw = True
        Exit Sub
    End If
    If ilRow < grdTimes.FixedRows Then
        grdTimes.Redraw = True
        Exit Sub
    End If
    grdTimes.Redraw = True
    lmTopRow = grdTimes.TopRow
    If Not mColOk() Then
        pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.RowHeight(grdTimes.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Exit Sub
    End If
    DoEvents
    grdTimes.Col = ilCol
    grdTimes.Row = ilRow
    mEnableBox
End Sub

Private Sub grdTimes_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdTimes.Redraw = False Then
        grdTimes.Redraw = True
        If lmTopRow < grdTimes.FixedRows Then
            grdTimes.TopRow = grdTimes.FixedRows
        Else
            grdTimes.TopRow = lmTopRow
        End If
        grdTimes.Refresh
        grdTimes.Redraw = False
    End If
    If (imCtrlVisible) And (grdTimes.Row >= grdTimes.FixedRows) And (grdTimes.Col >= 0) And (grdTimes.Col < grdTimes.Cols - 1) Then
        If grdTimes.RowIsVisible(grdTimes.Row) Then
            pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.RowHeight(grdTimes.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If

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
'*  ilLoop                        slNameCode                    slCode                    *
'*                                                                                        *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilAdf As Integer

    imFirstActivate = True
    imTerminate = False
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False
    imChgMode = False
    imCalType = 0   'Standard
    imDoubleClickName = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imBypassFocus = False
    lmEnableRow = -1
    lmEnableCol = -1

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    'cmcDone.Top = PostManualTimes.Height - (3 * cmcDone.Height) / 2
    'cmcCancel.Top = cmcDone.Top
    'grdTimes.Height = cmcDone.Top - grdTimes.Top - 120
    'gCenterStdAlone PostManualTimes
    Screen.MousePointer = vbHourglass
    imChg = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", PostManualTimes
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save CHF record length
    hmCpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", PostManualTimes
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save CHF record length

    mInitBox
    mPopListKey
    mCopyPop
    mPopulate
    If imTerminate Then
        Exit Sub
    End If

    lacScreen.Caption = "Post Times: " & sgCntrLineInfo
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload PostManualTimes
    igManUnload = NO
End Sub





Private Sub imcKey_Click()
    lbcKey.Visible = Not lbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = False
End Sub

Private Sub lacScreen_Click()
    pbcClickFocus.SetFocus
End Sub



Private Sub lbcCopyNm_Click()
    If imLbcMouseDown Then
        imLbcArrowSetting = False
    Else
        imLbcArrowSetting = True
    End If
    gProcessLbcClick lbcCopyNm, edcDropDown, imChgMode, imLbcArrowSetting
    imLbcMouseDown = False
End Sub

Private Sub lbcCopyNm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcCopyNm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub


Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropDown.SetFocus
End Sub

Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub pbcClickFocus_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    If imFirstFocus Then
        imFirstFocus = False
    End If
    If grdTimes.Visible Then
        mSetShow
        grdTimes.Row = 0
        grdTimes.Col = SPOTINDEX
        mSetCommands
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub




Private Sub mSetCommands()
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
    Dim ilLoop As Integer
    'flTextHeight = pbcDates.TextHeight("1") - 35

    'mGridLayout
    'mGridColumnWidths
    'mGridColumns
    'grdTimes.Move 180, lacScreen.Top + lacScreen.Height + 120, grdTimes.Width, cmcDone.Top - (lacScreen.Top + lacScreen.Height) - 240
    'grdTimes.Height = grdTimes.RowPos(0) + 14 * grdTimes.RowHeight(0) + fgPanelAdj - 15
    'imInitNoRows = (cmcDone.Top - 120 - grdTimes.Top) \ fgFlexGridRowH
    imInitNoRows = (grdTimes.Height - 120) \ fgFlexGridRowH
    'grdTimes.Height = grdTimes.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
    mClearGrid
    mGridLayout
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdTimes.Rows - 1 Step 1
        grdTimes.RowHeight(ilRow) = fgFlexGridRowH ' + 15
    Next ilRow
    For ilCol = 0 To grdTimes.Cols - 1 Step 1
        grdTimes.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mSetGridTitles()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdTimes.Row = grdTimes.FixedRows - 1
    grdTimes.Col = WKDATEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Week Date"
    grdTimes.Col = DAYPARTTIMESINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Times"
    grdTimes.Col = DAYPARTDAYSINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Days"
    grdTimes.Col = LENGTHINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Length"
    grdTimes.Col = STATUSINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Status"
    grdTimes.Col = DATEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Aired Date"
    grdTimes.Col = TIMEINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Aired Time"
    grdTimes.Col = ISCIINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "ISCI"
    grdTimes.Col = SPOTINDEX
    grdTimes.CellFontBold = False
    grdTimes.CellFontName = "Arial"
    grdTimes.CellFontSize = 6.75
    grdTimes.CellForeColor = vbBlue
    'grdTimes.CellBackColor = LIGHTBLUE
    grdTimes.TextMatrix(grdTimes.Row, grdTimes.Col) = "Spot Index"

End Sub

Private Sub mSetGridColumns()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdTimes.ColWidth(SPOTINDEX) = 0
    grdTimes.ColWidth(SORTINDEX) = 0
    grdTimes.ColWidth(WKDATEINDEX) = 0.08 * grdTimes.Width
    grdTimes.ColWidth(DAYPARTTIMESINDEX) = 0.12 * grdTimes.Width
    grdTimes.ColWidth(DAYPARTDAYSINDEX) = 0.08 * grdTimes.Width
    grdTimes.ColWidth(LENGTHINDEX) = 0.05 * grdTimes.Width
    grdTimes.ColWidth(STATUSINDEX) = 0.08 * grdTimes.Width
    grdTimes.ColWidth(DATEINDEX) = 0.08 * grdTimes.Width
    grdTimes.ColWidth(TIMEINDEX) = 0.1 * grdTimes.Width
    grdTimes.ColWidth(ISCIINDEX) = 0.25 * grdTimes.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdTimes.Width
    For ilCol = 0 To grdTimes.Cols - 2 Step 1
        llWidth = llWidth + grdTimes.ColWidth(ilCol)
        If (grdTimes.ColWidth(ilCol) > 15) And (grdTimes.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdTimes.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdTimes.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdTimes.Width
            For ilCol = 0 To grdTimes.Cols - 2 Step 1
                If (grdTimes.ColWidth(ilCol) > 15) And (grdTimes.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdTimes.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdTimes.FixedCols To grdTimes.Cols - 2 Step 1
                If grdTimes.ColWidth(ilCol) > 15 Then
                    ilColInc = grdTimes.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdTimes.ColWidth(ilCol) = grdTimes.ColWidth(ilCol) + 15
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


Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim llIndex As Long

    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        slStr = Trim$(grdTimes.TextMatrix(llRow, WKDATEINDEX))
        If slStr <> "" Then
            If ilCol = DATEINDEX Then
                llIndex = Trim$(grdTimes.TextMatrix(llRow, SPOTINDEX))
                slSort = Trim$(Str$(gDateValue(tgPostManualTimesInfo(llIndex).sDate)))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = TIMEINDEX Then
                llIndex = Trim$(grdTimes.TextMatrix(llRow, SPOTINDEX))
                slSort = Trim$(Str$(gTimeToLong(tgPostManualTimesInfo(llIndex).sTime, False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = ISCIINDEX) Then
                slSort = grdTimes.TextMatrix(llRow, ISCIINDEX)
                Do While Len(slSort) < 20
                    slSort = "0" & slSort
                Loop
             ElseIf (ilCol = LENGTHINDEX) Then
                slSort = grdTimes.TextMatrix(llRow, LENGTHINDEX)
                Do While Len(slSort) < 3
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = DAYPARTTIMESINDEX) Then
                slSort = grdTimes.TextMatrix(llRow, DAYPARTTIMESINDEX) & grdTimes.TextMatrix(llRow, DAYPARTDAYSINDEX)
            Else
                slSort = grdTimes.TextMatrix(llRow, ilCol)
            End If
            slStr = grdTimes.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdTimes.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdTimes.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdTimes, DATEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdTimes.Row < grdTimes.FixedRows) Or (grdTimes.Row >= grdTimes.Rows) Or (grdTimes.Col < grdTimes.FixedCols) Or (grdTimes.Col >= grdTimes.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdTimes.Row
    lmEnableCol = grdTimes.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.RowHeight(grdTimes.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    slStr = grdTimes.TextMatrix(lmEnableRow, lmEnableCol)
    Select Case grdTimes.Col
        Case DATEINDEX
            edcDropDown.MaxLength = 10
            slStr = grdTimes.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If (grdTimes.Row > grdTimes.FixedRows) And (gDateValue(grdTimes.TextMatrix(grdTimes.Row, WKDATEINDEX)) <= gDateValue(grdTimes.TextMatrix(grdTimes.Row - 1, WKDATEINDEX)) + 6) Then
                    slStr = grdTimes.TextMatrix(grdTimes.Row - 1, grdTimes.Col)
                End If
                If slStr = "" Then
                    slStr = grdTimes.TextMatrix(grdTimes.Row, WKDATEINDEX)
                End If
            End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            End If
            edcDropDown.Text = slStr
        Case TIMEINDEX
            edcDropDown.MaxLength = 11
            slStr = grdTimes.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case ISCIINDEX
            lbcCopyNm.Height = gListBoxHeight(lbcCopyNm.ListCount, 10)
            edcDropDown.MaxLength = 20
            imChgMode = True
            gFindMatch slStr, 0, lbcCopyNm
            If gLastFound(lbcCopyNm) >= 0 Then
                lbcCopyNm.ListIndex = gLastFound(lbcCopyNm)
                edcDropDown.Text = Trim$(lbcCopyNm.List(lbcCopyNm.ListIndex))
            Else
                edcDropDown.Text = Trim$(slStr)
            End If
            imChgMode = False
    End Select
    mSetFocus
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

    If (grdTimes.Row < grdTimes.FixedRows) Or (grdTimes.Row >= grdTimes.Rows) Or (grdTimes.Col < grdTimes.FixedCols) Or (grdTimes.Col >= grdTimes.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.RowHeight(grdTimes.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdTimes.Col - 1 Step 1
        llColPos = llColPos + grdTimes.ColWidth(ilCol)
    Next ilCol
    Select Case grdTimes.Col
        Case DATEINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - cmcDropDown.Width - 30, grdTimes.RowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TIMEINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - cmcDropDown.Width - 30, grdTimes.RowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ISCIINDEX
            edcDropDown.Move grdTimes.Left + llColPos + 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30, grdTimes.ColWidth(grdTimes.Col) - cmcDropDown.Width - 30, grdTimes.RowHeight(grdTimes.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + lbcCopyNm.Height < cmcDone.Top Then
                lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top - lbcCopyNm.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetCommands
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
'*  slAudience                    slPopulation                  slRating                  *
'*                                                                                        *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String

    pbcArrow.Visible = False
    If (lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
             Case DATEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gDateValue(grdTimes.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                    imChg = True
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                If slStr = "" Then
                    grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Missed"
                    grdTimes.TextMatrix(lmEnableRow, TIMEINDEX) = ""
                Else
                    grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Aired"
                End If
            Case TIMEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcTme.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gTimeToLong(grdTimes.TextMatrix(lmEnableRow, lmEnableCol), False) <> gTimeToLong(slStr, False) Then
                    imChg = True
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                'If slStr = "" Then
                '    grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Missed"
                'ElseIf grdTimes.TextMatrix(lmEnableRow, DATEINDEX) <> "" Then
                '    grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Aired"
                'End If
                If grdTimes.TextMatrix(lmEnableRow, DATEINDEX) <> "" Then
                    grdTimes.TextMatrix(lmEnableRow, STATUSINDEX) = "Aired"
                End If
            Case ISCIINDEX
                edcDropDown.Visible = False  'Set visibility
                cmcDropDown.Visible = False
                lbcCopyNm.Visible = False
                slStr = Trim$(edcDropDown.Text)
                If StrComp(grdTimes.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                End If
                grdTimes.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                If slStr <> "" Then
                    gFindMatch Trim$(slStr), 0, lbcCopyNm
                    If gLastFound(lbcCopyNm) < 0 Then
                        lbcCopyNm.RemoveItem 0
                        lbcCopyNm.AddItem slStr
                        lbcCopyNm.ItemData(lbcCopyNm.NewIndex) = 0
                        lbcCopyNm.AddItem "[None]", 0
                    End If
                End If
        End Select
    End If
    pbcArrow.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub



Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        If (lmEnableCol = DATEINDEX) And ((lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows)) Then
            If Trim$(edcDropDown.Text) = "" Then
                'imcTrash_Click
                Exit Sub
            End If
        End If
        mSetShow
        Do
            ilPrev = False
            If (grdTimes.Col = DATEINDEX) Then
                If grdTimes.Row > grdTimes.FixedRows Then
                    lmTopRow = -1
                    grdTimes.Row = grdTimes.Row - 1
                    If grdTimes.TextMatrix(grdTimes.Row, SPOTINDEX) = "" Then
                        grdTimes.Row = grdTimes.Row - 1
                    End If
                    If Not grdTimes.RowIsVisible(grdTimes.Row) Then
                        grdTimes.TopRow = grdTimes.TopRow - 1
                    End If
                    grdTimes.Col = ISCIINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdTimes.Col = grdTimes.Col - 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdTimes.TopRow = grdTimes.FixedRows
        grdTimes.Col = DATEINDEX
        grdTimes.Row = grdTimes.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        If (lmEnableCol = DATEINDEX) And ((lmEnableRow >= grdTimes.FixedRows) And (lmEnableRow < grdTimes.Rows)) Then
            If Trim$(edcDropDown.Text) = "" Then
                'imcTrash_Click
                Exit Sub
            End If
        End If
        mSetShow
        Do
            ilNext = False
            If (grdTimes.Col = ISCIINDEX) Then
                llRow = grdTimes.Rows
                Do
                    llRow = llRow - 1
                Loop While grdTimes.TextMatrix(llRow, WKDATEINDEX) = ""
                llRow = llRow + 1
                If (grdTimes.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdTimes.Row = grdTimes.Row + 1
                    'If Not grdTimes.RowIsVisible(grdTimes.Row) Or (grdTimes.Row - (grdTimes.TopRow - grdTimes.FixedRows) >= imInitNoRows) Then
                    If Not grdTimes.RowIsVisible(grdTimes.Row) Then 'Or (grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + 30 + edcDropDown.Height > grdTimes.Height) Then
                        imIgnoreScroll = True
                        grdTimes.TopRow = grdTimes.TopRow + 1
                    End If
                    grdTimes.Col = DATEINDEX
                    'grdTimes.TextMatrix(grdTimes.Row, CODEINDEX) = 0
                    If Trim$(grdTimes.TextMatrix(grdTimes.Row, WKDATEINDEX)) <> "" Then
                        If mColOk() Then
                            mEnableBox
                        Else
                            cmcCancel.SetFocus
                        End If
                    Else
                        If grdTimes.TextMatrix(grdTimes.Row, SPOTINDEX) = "" Then
                            grdTimes.Row = grdTimes.Row + 1
                        End If
                        imFromArrow = True
                        pbcArrow.Move grdTimes.Left - pbcArrow.Width - 30, grdTimes.Top + grdTimes.RowPos(grdTimes.Row) + (grdTimes.RowHeight(grdTimes.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdTimes.Col = grdTimes.Col + 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdTimes.TopRow = grdTimes.FixedRows
        grdTimes.Col = ISCIINDEX
        grdTimes.Row = grdTimes.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slMsg                         llLoop                                                  *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mSaveRecErr                                                                           *
'******************************************************************************************

    Dim llRow As Long
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim slStr As String
    Dim llIndex As Long
    Dim slDate As String

    If Not imUpdateAllowed Then
        mSaveRec = False
        Exit Function
    End If
    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdTimes, grdTimes, vbHourglass
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        If mGridFieldsOk(llRow) = False Then
            ilError = True
        End If
    Next llRow
    If (ilError) Then
        gSetMousePointer grdTimes, grdTimes, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        MsgBox "Check input fields marked in Red as in error", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        mSaveRec = False
        Exit Function
    End If
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        If Trim$(grdTimes.TextMatrix(llRow, SPOTINDEX)) <> "" Then
            llIndex = Trim$(grdTimes.TextMatrix(llRow, SPOTINDEX))
            If Trim$(grdTimes.TextMatrix(llRow, DATEINDEX)) <> "" Then
                tgPostManualTimesInfo(llIndex).sDate = Trim$(grdTimes.TextMatrix(llRow, DATEINDEX))
                tgPostManualTimesInfo(llIndex).sTime = Trim$(grdTimes.TextMatrix(llRow, TIMEINDEX))
            Else
                tgPostManualTimesInfo(llIndex).sDate = ""
                tgPostManualTimesInfo(llIndex).sTime = ""
            End If
            tgPostManualTimesInfo(llIndex).sISCI = Trim$(grdTimes.TextMatrix(llRow, ISCIINDEX))
            tgPostManualTimesInfo(llIndex).lCifCode = mAddCopy(llRow)
        End If
    Next llRow
    imChg = False
    mSaveRec = True
    gSetMousePointer grdTimes, grdTimes, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr: 'VBC NR
    On Error GoTo 0
    gSetMousePointer grdTimes, grdTimes, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mPopulate()

    Dim llRow As Long
    Dim ilCol As Integer
    Dim llLoop As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilVeh As Integer
    Dim llColor As Long
    Dim llCntr As Long
    Dim slDaypartYellow As String

    grdTimes.Redraw = False
    llRow = grdTimes.FixedRows
    llColor = grdTimes.ForeColor
    For llCntr = 0 To UBound(tgPostManualCntrInfo) - 1 Step 1
        For llLoop = 0 To UBound(tgPostManualTimesInfo) - 1 Step 1
            If tgPostManualCntrInfo(llCntr).lRow = tgPostManualTimesInfo(llLoop).lRow Then
                If llRow <> grdTimes.FixedRows Then
                    If slDaypartYellow <> tgPostManualCntrInfo(llCntr).sDaypartTimes & tgPostManualCntrInfo(llCntr).sDaypartDays & tgPostManualCntrInfo(llCntr).sLength Then
                        If llRow >= grdTimes.Rows Then
                            grdTimes.AddItem ""
                            grdTimes.RowHeight(llRow) = fgFlexGridRowH
                        End If
                        grdTimes.Row = llRow
                        For ilCol = 0 To grdTimes.Cols - 1 Step 1
                            grdTimes.Col = ilCol
                            grdTimes.CellBackColor = LIGHTYELLOW
                        Next ilCol
                        llRow = llRow + 1
                        slDaypartYellow = tgPostManualCntrInfo(llCntr).sDaypartTimes & tgPostManualCntrInfo(llCntr).sDaypartDays & tgPostManualCntrInfo(llCntr).sLength
                    End If
                Else
                    slDaypartYellow = tgPostManualCntrInfo(llCntr).sDaypartTimes & tgPostManualCntrInfo(llCntr).sDaypartDays & tgPostManualCntrInfo(llCntr).sLength
                End If
                If llRow >= grdTimes.Rows Then
                    grdTimes.AddItem ""
                    grdTimes.RowHeight(llRow) = fgFlexGridRowH
                End If
                grdTimes.Row = llRow
                grdTimes.Col = WKDATEINDEX
                grdTimes.CellBackColor = LIGHTYELLOW
                grdTimes.TextMatrix(llRow, WKDATEINDEX) = tgPostManualTimesInfo(llLoop).sWkDate
                grdTimes.Col = DAYPARTTIMESINDEX
                grdTimes.CellBackColor = LIGHTYELLOW
                grdTimes.TextMatrix(llRow, DAYPARTTIMESINDEX) = tgPostManualCntrInfo(llCntr).sDaypartTimes
                grdTimes.Col = DAYPARTDAYSINDEX
                grdTimes.CellBackColor = LIGHTYELLOW
                grdTimes.TextMatrix(llRow, DAYPARTDAYSINDEX) = tgPostManualCntrInfo(llCntr).sDaypartDays
                grdTimes.Col = LENGTHINDEX
                grdTimes.CellBackColor = LIGHTYELLOW
                grdTimes.TextMatrix(llRow, LENGTHINDEX) = tgPostManualCntrInfo(llCntr).sLength
                grdTimes.Col = STATUSINDEX
                grdTimes.CellBackColor = LIGHTYELLOW
                If tgPostManualTimesInfo(llLoop).sSchStatus <> "M" Then
                    grdTimes.TextMatrix(llRow, STATUSINDEX) = "Aired"
                    grdTimes.TextMatrix(llRow, DATEINDEX) = tgPostManualTimesInfo(llLoop).sDate
                    grdTimes.TextMatrix(llRow, TIMEINDEX) = tgPostManualTimesInfo(llLoop).sTime
                    grdTimes.TextMatrix(llRow, ISCIINDEX) = tgPostManualTimesInfo(llLoop).sISCI
                Else
                    grdTimes.TextMatrix(llRow, STATUSINDEX) = "Missed"
                End If
                grdTimes.TextMatrix(llRow, SPOTINDEX) = Trim$(Str$(llLoop))
                llRow = llRow + 1
            End If
        Next llLoop
    Next llCntr
    For llRow = grdTimes.Rows - 1 To grdTimes.FixedRows Step -1
        If grdTimes.TextMatrix(llRow, WKDATEINDEX) <> "" Then
            Exit For
        End If
        grdTimes.RemoveItem llRow
    Next llRow
    imLastColSorted = -1
    'mSortCol TIMEINDEX
    'mSortCol DATEINDEX
    'mSortCol LENGTHINDEX
    'mSortCol DAYPARTTIMESINDEX
    grdTimes.Row = 0
    grdTimes.Col = SPOTINDEX
    grdTimes.Redraw = True
    mSetCommands

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk(llRowNo As Long) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
    slStr = Trim$(grdTimes.TextMatrix(llRowNo, WKDATEINDEX))
    If slStr <> "" Then
        slStr = Trim$(Trim$(grdTimes.TextMatrix(llRowNo, DATEINDEX)))
        If slStr <> "" Then
            If Not gValidDate(slStr) Then
                ilError = True
                grdTimes.Row = llRowNo
                grdTimes.Col = DATEINDEX
                grdTimes.CellForeColor = vbRed
            End If
            slStr = Trim$(grdTimes.TextMatrix(llRowNo, TIMEINDEX))
            If slStr = "" Then
                grdTimes.TextMatrix(llRowNo, TIMEINDEX) = "Missing"
                ilError = True
                grdTimes.Row = llRowNo
                grdTimes.Col = TIMEINDEX
                grdTimes.CellForeColor = vbRed
            Else
                If Not gValidTime(slStr) Then
                    ilError = True
                    grdTimes.Row = llRowNo
                    grdTimes.Col = TIMEINDEX
                    grdTimes.CellForeColor = vbRed
                End If
            End If
        End If
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function


Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    If Not imUpdateAllowed Then
        mColOk = False
    ElseIf igRepSpotBilled Then
        mColOk = False
    Else
        mColOk = True
        If grdTimes.CellBackColor = LIGHTYELLOW Then
            mColOk = False
            Exit Function
        End If
        If grdTimes.ForeColor = DARKGREEN Then
            mColOk = False
            Exit Function
        End If
    End If
End Function



Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    imBypassFocus = True    'Don't change select text
                    edcDropDown.SetFocus
                    'SendKeys slKey
                    gSendKeys edcDropDown, slKey
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub tmcStart_Timer()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                         slNameCode                    slCode                    *
'*  ilRet                                                                                 *
'******************************************************************************************


    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
    End If
    'Check if allowed to defined input
    'One vehicle only
    imUpdateAllowed = True
    If (igWinStatus(POSTLOGSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    End If
    If Not imUpdateAllowed Then
        grdTimes.Enabled = False
    End If
End Sub



Private Sub mClearGrid()
    Dim llRow As Long
    Dim ilCol As Integer

    grdTimes.Redraw = False
    grdTimes.Rows = imInitNoRows
    For llRow = grdTimes.FixedRows To grdTimes.Rows - 1 Step 1
        grdTimes.RowHeight(llRow) = fgFlexGridRowH
        For ilCol = 0 To grdTimes.Cols - 1 Step 1
            grdTimes.Row = llRow
            grdTimes.Col = ilCol
            grdTimes.TextMatrix(llRow, ilCol) = ""
            If (ilCol = WKDATEINDEX) Or (ilCol = DAYPARTTIMESINDEX) Or (ilCol = DAYPARTDAYSINDEX) Or (ilCol = LENGTHINDEX) Or (ilCol = STATUSINDEX) Then
                grdTimes.CellBackColor = LIGHTYELLOW
            End If
        Next ilCol
    Next llRow
    grdTimes.Redraw = True
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCopyPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mCopyPopErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilAdvtCode As Integer
    Dim ilISCIProd As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String

    ilISCIProd = 4
    ilAdvtCode = igPostManualAdfCode
    ilRet = gPopCopyForAdvtBox(PostManualTimes, ilAdvtCode, ilISCIProd, 1, lbcCopyNm, tmCopyNmCode(), smCopyNmCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        lbcCopyNm.Clear
        For ilLoop = 0 To UBound(tmCopyNmCode) - 1 Step 1
            slNameCode = tmCopyNmCode(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            lbcCopyNm.AddItem Trim$(slName)
            lbcCopyNm.ItemData(lbcCopyNm.NewIndex) = slCode
        Next ilLoop
        lbcCopyNm.AddItem "[None]", 0   ' Add this at the top of the list
    End If
    Exit Sub
mCopyPopErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcDropDown.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub




Private Function mAddCopy(llRowNo As Long) As Long
    Dim slISCI As String
    Dim ilRet As Integer
    Dim slDate As String

    slISCI = Trim$(grdTimes.TextMatrix(llRowNo, ISCIINDEX))
    gFindMatch slISCI, 0, lbcCopyNm
    If gLastFound(lbcCopyNm) >= 0 Then
        If Val(lbcCopyNm.ItemData(gLastFound(lbcCopyNm))) > 0 Then
            mAddCopy = Val(lbcCopyNm.ItemData(gLastFound(lbcCopyNm)))
            Exit Function
        End If
    Else
        mAddCopy = 0
        Exit Function
    End If
    'Add Product/ISCI
    tmCpf.lCode = 0
    tmCpf.sName = ""
    tmCpf.sISCI = slISCI
    tmCpf.sCreative = ""
    tmCpf.iRotEndDate(0) = 0
    tmCpf.iRotEndDate(1) = 0
    tmCpf.lSifCode = 0
    ilRet = btrInsert(hmCpf, tmCpf, imCpfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mAddCopy = 0
        Exit Function
    End If
    'Add copy inventory
    tmCif.lCode = 0  'Autoincrement
    tmCif.iMcfCode = 0
    tmCif.sName = ""
    tmCif.sCut = ""
    tmCif.sReel = ""
    tmCif.iLen = igPostManualLength
    tmCif.iEtfCode = 0
    tmCif.iEnfCode = 0
    tmCif.iAdfCode = igPostManualAdfCode
    tmCif.lcpfCode = tmCpf.lCode
    tmCif.iMnfComp(0) = 0
    tmCif.iMnfComp(1) = 0
    tmCif.iMnfAnn = 0
    tmCif.sHouse = "Y"
    tmCif.sCleared = "Y"
    tmCif.lCsfCode = 0
    tmCif.iNoTimesAir = 0
    tmCif.sCartDisp = "N"
    tmCif.sTapeDisp = "N"
    tmCif.sPurged = "A"
    tmCif.iPurgeDate(0) = 0
    tmCif.iPurgeDate(1) = 0
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
    tmCif.iUsedDate(0) = 0
    tmCif.iUsedDate(1) = 0
    tmCif.iRotStartDate(0) = 0
    tmCif.iRotStartDate(1) = 0
    tmCif.iRotEndDate(0) = 0
    tmCif.iRotEndDate(1) = 0
    tmCif.iUrfCode = tgUrf(0).iCode
    tmCif.sPrint = "N"
    tmCif.iLangMnfCode = 0
    tmCif.sUnused = ""
    ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        mAddCopy = tmCif.lCode
    Else
        mAddCopy = 0
    End If
End Function

Private Sub mSetControls()
    Dim ilGap As Integer
    Dim llRow As Long

    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcCancel.Top = cmcDone.Top
    grdTimes.Move 180, lacScreen.Top + 2 * lacScreen.Height + 60, PostManualTimes.Width - 360, cmcDone.Top - (lacScreen.Top + 2 * lacScreen.Height + 60)
    lacScreen.Left = 120
    imcKey.Move lacScreen.Left, lacScreen.Top + lacScreen.Height
    lbcKey.Move lacScreen.Left, imcKey.Top + lacScreen.Height + 60
    cmcDone.Left = PostManualTimes.Width / 2 - cmcDone.Width - ilGap / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilGap
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdTimes, CInt(fgFlexGridRowH)
    grdTimes.Height = grdTimes.Height + 30
End Sub
Private Sub mPopListKey()
    Dim llMaxWidth As Long
    lbcKey.Clear
    lbcKey.AddItem "To Set as Aired:"
    lbcKey.AddItem "     Enter Date and Time"
    lbcKey.AddItem "To Set as Missed:"
    lbcKey.AddItem "     Blank Out Date"
    
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontName = "Arial"
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontSize = 8
    llMaxWidth = (Traffic.pbcArial.TextWidth("     Enter Date and Time")) + 180
    lbcKey.Width = llMaxWidth
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.Height = (lbcKey.ListCount) * 225
    lbcKey.Height = gListBoxHeight(lbcKey.ListCount, 4)
    lbcKey.Move imcKey.Left, imcKey.Top + imcKey.Height
End Sub
