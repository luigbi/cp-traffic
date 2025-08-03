VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl AffCommentGrid 
   Appearance      =   0  'Flat
   ClientHeight    =   2940
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   8865
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
   ScaleHeight     =   2940
   ScaleWidth      =   8865
   Begin VB.ListBox lbcVehicleView 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCommentGrid.ctx":0000
      Left            =   2970
      List            =   "AffCommentGrid.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmcVehicleView 
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
      Left            =   1935
      Picture         =   "AffCommentGrid.ctx":0004
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2250
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   6630
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   945
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   30
         Picture         =   "AffCommentGrid.ctx":00FE
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   15
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   300
         TabIndex        =   18
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcTextWidth 
      Height          =   180
      Left            =   825
      ScaleHeight     =   120
      ScaleWidth      =   405
      TabIndex        =   12
      Top             =   2010
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ListBox lbcSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCommentGrid.ctx":2F18
      Left            =   5475
      List            =   "AffCommentGrid.ctx":2F1A
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcStation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCommentGrid.ctx":2F1C
      Left            =   4500
      List            =   "AffCommentGrid.ctx":2F1E
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1785
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox ckcOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   1455
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCommentGrid.ctx":2F20
      Left            =   2520
      List            =   "AffCommentGrid.ctx":2F22
      Sorted          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2895
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   945
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
      Left            =   3840
      Picture         =   "AffCommentGrid.ctx":2F24
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5370
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   8
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
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "AffCommentGrid.ctx":301E
      ScaleHeight     =   180
      ScaleWidth      =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   90
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5640
      Width           =   75
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdComment 
      Height          =   1005
      Left            =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "AffCommentGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AffCommentGrid.ctl on Wed 6/17/09 @ 12:56 P
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
' File Name: AffCommentGrid.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

Event Tip(slTip As String)
Event CommentStatus(slStatus As String)
Event StationSelected(ilShttCode As Integer, slVehicleName As String)
Event GetCommentType()
Event CommentFocus()

Private rst_cct As ADODB.Recordset
Private rst_Ust As ADODB.Recordset
Private rst_dnt As ADODB.Recordset
Private rst_cst As ADODB.Recordset

Public Event SetSave(ilStatus As Integer) 'VBC NR

Private smPrevTip As String
'Program library dates Field Areas
Private imChgMode As Integer    'Change mode status (so change not entered when in change)
Private imBSMode As Integer     'Backspace flag
Private imFirstActivate As Integer
Private imTerminate As Integer  'True = terminating task, False= OK
Private imSettingValue As Integer
Private imLbcArrowSetting As Integer
Private imBypassFocus As Integer
Private imDoubleClickName As Integer
Private imLbcMouseDown As Integer  'True=List box mouse down
Private imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Private imUpdateAllowed As Integer
Private imStartMode As Integer
Private imLastColSorted As Integer
Private imLastSort As Integer
Private imFromArrow As Integer
Private imSourceDefaultCode As Integer
Private imDoubleClick As Integer
Private imLoadVehicleView As Integer

Private smUserName As String
Private lmUserColor As Long

Private smNowDate As String
Private lmNowDate As Long
Private lmFirstAllowedChgDate As Long

Private imCtrlVisible As Integer
Private lmEnableRow As Long
Private lmEnableCol As Long
Private lmTopRow As Long

Private imShttCode As Integer
Private imVefCode As Integer
Private imGridForm As Integer   '0=Without Stations column; 1=With Station columns
Private imFilterStatus As Integer   'True = On, False = Off
Private smWhichComment As String * 1   'A=All; M=Mine; D=Department

'Calendar
'Private tmCDCtrls(1 To 7) As FIELDAREA
Private tmCDCtrls(0 To 6) As FIELDAREA
Private imCalYear As Integer    'Month of displayed calendar
Private imCalMonth As Integer   'Year of displayed calendar
Private lmCalStartDate As Long  'Start date of displayed calendar
Private lmCalEndDate As Long    'End date of displayed calendar
Private imCalType As Integer
Private fmBoxGridH As Single

'Comment Grid- grdComment
Const CPOSTEDDATEINDEX = 0
Const CBYINDEX = 1
Const CCALLETTERINDEX = 2
Const CVEHICLEINDEX = 3
Const CFOLLOWUPINDEX = 4
Const COKINDEX = 5
Const CSOURCEINDEX = 6
Const CCOMMENTINDEX = 7
Const CDELETEINDEX = 8
Const CPOSTEDTIMEINDEX = 9
Const CUSERNAMEINDEX = 10
Const CCCTCODEINDEX = 11
Const CORIGOKINDEX = 12
Const CSORTINDEX = 13
Const CCHGDINDEX = 14









Private Sub ckcOK_Click()
    Dim ilRow As Integer
    Dim ilCurRow As Integer
    
    ilCurRow = grdComment.Row
    If ckcOK.Value = vbChecked Then
        grdComment.Col = COKINDEX
        grdComment.CellFontName = "Monotype Sorts"
        grdComment.TextMatrix(ilCurRow, COKINDEX) = "4"
        grdComment.TextMatrix(ilCurRow, CCHGDINDEX) = 1
    Else
        grdComment.TextMatrix(ilCurRow, COKINDEX) = " "
        grdComment.TextMatrix(ilCurRow, CCHGDINDEX) = 1
    End If
    mSetCommands
End Sub


Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropdown.SelStart = 0
    edcDropdown.SelLength = Len(edcDropdown.Text)
    edcDropdown.SetFocus
End Sub

Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropdown.SelStart = 0
    edcDropdown.SelLength = Len(edcDropdown.Text)
    edcDropdown.SetFocus
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdComment.Col
        Case CPOSTEDDATEINDEX
        Case CBYINDEX
        Case CCALLETTERINDEX
            lbcStation.Visible = Not lbcStation.Visible
        Case CVEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case CFOLLOWUPINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case COKINDEX
        Case CSOURCEINDEX
            lbcSource.Visible = Not lbcSource.Visible
        Case CCOMMENTINDEX
    End Select

End Sub

Private Sub cmcVehicleView_Click()
    lbcVehicleView.Visible = Not lbcVehicleView.Visible
End Sub

Private Sub cmcVehicleView_GotFocus()
    mSetShow
End Sub

Private Sub edcDropdown_Change()
    Dim slStr As String
    
    Select Case lmEnableCol
        Case CPOSTEDDATEINDEX
        Case CBYINDEX
        Case CCALLETTERINDEX
            mDropdownChangeEvent lbcStation
        Case CVEHICLEINDEX
            mDropdownChangeEvent lbcVehicle
        Case CFOLLOWUPINDEX
            slStr = edcDropdown.Text
            If Not gIsDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case COKINDEX
        Case CSOURCEINDEX
            mDropdownChangeEvent lbcSource
        Case CCOMMENTINDEX
    End Select
End Sub


Private Sub edcDropdown_DblClick()
    Select Case lmEnableCol
        Case CPOSTEDDATEINDEX
        Case CBYINDEX
        Case CCALLETTERINDEX
        Case CVEHICLEINDEX
        Case CFOLLOWUPINDEX
        Case COKINDEX
        Case CSOURCEINDEX
            imDoubleClick = True
        Case CCOMMENTINDEX
    End Select
End Sub

Private Sub edcDropdown_GotFocus()
    RaiseEvent Tip("")
    'gCtrlGotFocus ActiveControl
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
    If lmEnableCol = CFOLLOWUPINDEX Then
        'Filter characters (allow only BackSpace, numbers 0 thru 9
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub


Private Sub Form_Activate()
    If imFirstActivate Then
    End If
    imFirstActivate = False
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub Form_Load()
    'mInit
End Sub


Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmEnableCol
            Case CPOSTEDDATEINDEX
            Case CBYINDEX
            Case CCALLETTERINDEX
                gProcessArrowKey Shift, KeyCode, lbcStation, True
            Case CVEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, True
            Case CFOLLOWUPINDEX
                If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                    If (Shift And vbAltMask) > 0 Then
                        plcCalendar.Visible = Not plcCalendar.Visible
                    Else
                        slDate = edcDropdown.Text
                        If gIsDate(slDate) Then
                            If KeyCode = KEYUP Then 'Up arrow
                                slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                            Else
                                slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                            End If
                            gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                            edcDropdown.Text = slDate
                        End If
                    End If
                    edcDropdown.SelStart = 0
                    edcDropdown.SelLength = Len(edcDropdown.Text)
                End If
            Case COKINDEX
            Case CSOURCEINDEX
                gProcessArrowKey Shift, KeyCode, lbcSource, True
            Case CCOMMENTINDEX
        End Select
    End If
    If lmEnableCol = CFOLLOWUPINDEX Then
        If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
            If (Shift And vbAltMask) > 0 Then
            Else
                slDate = edcDropdown.Text
                If gIsDate(slDate) Then
                    If KeyCode = KEYLEFT Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcDropdown.Text = slDate
                End If
            End If
            edcDropdown.SelStart = 0
            edcDropdown.SelLength = Len(edcDropdown.Text)
        End If
    End If
End Sub

Private Sub edcDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClick Then
        Select Case lmEnableCol
            Case CPOSTEDDATEINDEX
            Case CBYINDEX
            Case CCALLETTERINDEX
            Case CVEHICLEINDEX
            Case CFOLLOWUPINDEX
            Case COKINDEX
            Case CSOURCEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case CCOMMENTINDEX
        End Select
        imDoubleClick = False
    End If
End Sub

Private Sub grdComment_EnterCell()
    If lmEnableRow <> grdComment.MouseRow Then
        mSetShow
        mSaveRec
    Else
        mSetShow
    End If
End Sub

Private Sub grdComment_GotFocus()
    RaiseEvent CommentFocus
End Sub

Private Sub grdComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    slStr = ""
    If Not imCtrlVisible Then
        If (grdComment.MouseRow >= grdComment.FixedRows) And (grdComment.TextMatrix(grdComment.MouseRow, grdComment.MouseCol)) <> "" Then
            If (grdComment.MouseCol >= CPOSTEDDATEINDEX) And (grdComment.MouseCol <= CCOMMENTINDEX) Then
                'grdComment.ToolTipText = grdComment.TextMatrix(grdComment.MouseRow, grdComment.MouseCol)
                If grdComment.MouseCol = CBYINDEX Then
                    slStr = Trim$(grdComment.TextMatrix(grdComment.MouseRow, CUSERNAMEINDEX))
                Else
                    slStr = Trim$(grdComment.TextMatrix(grdComment.MouseRow, grdComment.MouseCol))
                    If grdComment.MouseCol = CPOSTEDDATEINDEX Then
                        slStr = slStr & " " & Trim$(grdComment.TextMatrix(grdComment.MouseRow, CPOSTEDTIMEINDEX))
                    End If
                End If
            End If
        End If
    End If
    If smPrevTip <> slStr Then
        If grdComment.MouseCol = CCOMMENTINDEX Then
            RaiseEvent Tip(slStr)
            grdComment.ToolTipText = ""
        Else
            RaiseEvent Tip("")
            grdComment.ToolTipText = slStr
        End If
    End If
    smPrevTip = slStr
End Sub

Private Sub grdComment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim llCode As Long
    Dim ilRet As Integer

    'Determine if in header
'    If y < grdComment.RowHeight(0) Then
'        mSortCol grdComment.Col
'        Exit Sub
'    End If
    RaiseEvent Tip("")
    If Y < grdComment.RowHeight(0) Then
        grdComment.Row = grdComment.MouseRow
        grdComment.Col = grdComment.MouseCol
        If grdComment.CellBackColor = LIGHTBLUE Then
            mMousePointer vbHourglass
            mSetShow
            mCommentSortCol grdComment.Col
            grdComment.Row = 0
            grdComment.Col = CCCTCODEINDEX
            mMousePointer vbDefault
        End If
        Exit Sub
    End If
    'Determine row and col mouse up onto
    On Error GoTo grdCommentErr
    pbcArrow.Visible = False
    ilCol = grdComment.MouseCol
    ilRow = grdComment.MouseRow
    If ilCol < grdComment.FixedCols Then
        grdComment.Redraw = True
        Exit Sub
    End If
    If ilRow < grdComment.FixedRows Then
        grdComment.Redraw = True
        Exit Sub
    End If
    If ilCol = CCALLETTERINDEX Then
        SQLQuery = "SELECT * FROM cct WHERE cctCode = " & grdComment.TextMatrix(ilRow, CCCTCODEINDEX)
        Set rst_cct = gSQLSelectCall(SQLQuery)
        If Not rst_cct.EOF Then
            RaiseEvent StationSelected(rst_cct!cctShfCode, Trim$(grdComment.TextMatrix(ilRow, CVEHICLEINDEX)))
        End If
        grdComment.Redraw = True
        Exit Sub
    End If
    If ilCol = CDELETEINDEX Then
        If sgUstAllowCmmtDelete = "Y" Then
            llCode = Val(grdComment.TextMatrix(ilRow, CCCTCODEINDEX))
            If llCode > 0 Then
                ilRet = MsgBox("This will permanently remove the Comment, are you sure", vbYesNo + vbQuestion, "Remove")
                If ilRet = vbYes Then
                    SQLQuery = "DELETE FROM cct WHERE cctCode = " & llCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "AffCommentGrid-grdComment_MouseUp"
                        Exit Sub
                    End If
                    grdComment.RemoveItem ilRow
                End If
            Else
                grdComment.RemoveItem ilRow
            End If
        End If
        grdComment.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdComment.TopRow
    DoEvents
    If grdComment.TextMatrix(ilRow, CVEHICLEINDEX) = "" Then
        grdComment.Redraw = False
        Do
            ilRow = ilRow - 1
            If ilRow < grdComment.FixedRows Then
                Exit Do
            End If
        Loop While Trim(grdComment.TextMatrix(ilRow, CVEHICLEINDEX)) = ""
        ilRow = ilRow + 1
        ilCol = CVEHICLEINDEX
    End If
    grdComment.Col = ilCol
    grdComment.Row = ilRow
    If Not mColOk() Then
        grdComment.Redraw = True
        Exit Sub
    End If
    grdComment.Redraw = True
    mEnableBox
    On Error GoTo 0
    Exit Sub
grdCommentErr:
    On Error GoTo 0
    If (lmEnableRow >= grdComment.FixedRows) And (lmEnableRow < grdComment.Rows) Then
        grdComment.Row = lmEnableRow
        grdComment.Col = lmEnableCol
        mSetFocus
    End If
    grdComment.Redraw = False
    grdComment.Redraw = True
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffCommentGrid-grdComment"
    Exit Sub
End Sub

Private Sub grdComment_Scroll()
    RaiseEvent Tip("")
    mSetShow
    pbcArrow.Visible = False
    If grdComment.RowIsVisible(grdComment.Row) Then
        pbcArrow.Move grdComment.Left - pbcArrow.Width, grdComment.Top + grdComment.RowPos(grdComment.Row) + (grdComment.RowHeight(grdComment.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
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
    Dim ilLoop As Integer
    Dim slStr As String

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdComment, grdComment, vbHourglass
    imFirstActivate = True
    imTerminate = False
    'pbcArrow.Picture = IconTraf!imcArrow.Picture
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
    imLastColSorted = -1
    imLastSort = -1
    imFromArrow = False
    lmEnableRow = -1
    imLoadVehicleView = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate + 1
    mInitBox
    mPopVehicle
    mPopSource
    mGetUserInfo
    'Calendar setup
    imCalType = 0
    imBypassFocus = False
    fmBoxGridH = 180      'Height of grid area (distance from bottom of form letter to bottom of form box)
    'For ilLoop = 1 To 7 Step 1
    For ilLoop = 0 To 6 Step 1
        'gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fmBoxGridH
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop), 225, 240, fmBoxGridH
    Next ilLoop
    slStr = gObtainNextMonday(smNowDate)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint

    cmcVehicleView.Width = cmcCalDn.Width
    cmcVehicleView.Left = grdComment.Left + grdComment.ColPos(CFOLLOWUPINDEX) - cmcVehicleView.Width
    cmcVehicleView.Top = grdComment.Top + 15
    cmcVehicleView.Height = grdComment.RowHeight(0) - 15
    cmcVehicleView.Font = "Monotype Sorts"
    cmcVehicleView.Visible = True
    lbcVehicleView.Width = grdComment.ColPos(CFOLLOWUPINDEX) - grdComment.ColPos(CBYINDEX)
    lbcVehicleView.Left = cmcVehicleView.Left + cmcVehicleView.Width - lbcVehicleView.Width
    lbcVehicleView.Top = cmcVehicleView.Top + cmcVehicleView.Height + 15
    'gSetListBoxHeight lbcVehicleView, 6

    Screen.MousePointer = vbDefault
    gSetMousePointer grdComment, grdComment, vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdComment, grdComment, vbDefault
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

    'grdComment.Move 180, 120, Width - pbcArrow.Width - 120
    'grdComment.Height = Height - grdComment.Top - 120
    'grdComment.Redraw = False
    pbcSTab.Move -100, -100
    pbcTab.Move -100, -100
    pbcClickFocus.Move -100, -100
    mSetGridColumns
    mSetGridTitles
    mClearGrid grdComment
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
    gSetMousePointer grdComment, grdComment, vbDefault
End Sub

Private Sub lbcSource_Click()
    edcDropdown.Text = lbcSource.List(lbcSource.ListIndex)
End Sub

Private Sub lbcSource_DblClick()
    imDoubleClick = True
End Sub

Private Sub lbcSource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClick Then
        imLbcArrowSetting = False
        'gProcessLbcClick lbcSource, edcDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcStation_Click()
    edcDropdown.Text = lbcStation.List(lbcStation.ListIndex)
End Sub

Private Sub lbcVehicle_Click()
    edcDropdown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
End Sub

Private Sub lbcVehicleView_Click()
    If imChgMode Then
        Exit Sub
    End If
    imLoadVehicleView = False
    mPopCommentGrid
    pbcClickFocus.SetFocus
    lbcVehicleView.Visible = False
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
        'If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
        '    If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
        If (X >= tmCDCtrls(ilWkDay).fBoxX) And (X <= (tmCDCtrls(ilWkDay).fBoxX + tmCDCtrls(ilWkDay).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay).fBoxH + 15) + tmCDCtrls(ilWkDay).fBoxH) Then
                edcDropdown.Text = Format$(llDate, "m/d/yy")
                edcDropdown.SelStart = 0
                edcDropdown.SelLength = Len(edcDropdown.Text)
                imBypassFocus = True
                edcDropdown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropdown.SetFocus
End Sub

Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer
    Dim ilIndex As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        If imGridForm = 1 Then
            grdComment.Col = CCALLETTERINDEX
        Else
            grdComment.Col = CVEHICLEINDEX
        End If
        mEnableBox
        Exit Sub
    End If
    imTabDirection = -1
    If imCtrlVisible Then
        Do
            ilNext = False
            Select Case grdComment.Col
                Case CCALLETTERINDEX
                    If grdComment.Row = grdComment.FixedRows Then
                        mSetShow
                        pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdComment.Row = grdComment.Row - 1
                    If Not grdComment.RowIsVisible(grdComment.Row) Then
                        grdComment.TopRow = grdComment.TopRow - 1
                    End If
                    grdComment.Col = CCOMMENTINDEX
                Case CVEHICLEINDEX
                    If imGridForm <> 1 Then
                        If grdComment.Row = grdComment.FixedRows Then
                            mSetShow
                            pbcClickFocus.SetFocus
                            Exit Sub
                        End If
                        lmTopRow = -1
                        grdComment.Row = grdComment.Row - 1
                        If Not grdComment.RowIsVisible(grdComment.Row) Then
                            grdComment.TopRow = grdComment.TopRow - 1
                        End If
                        grdComment.Col = CCOMMENTINDEX
                    Else
                        grdComment.Col = grdComment.Col - 1
                    End If
                Case CSOURCEINDEX
                    grdComment.Col = CFOLLOWUPINDEX
                Case Else
                    grdComment.Col = grdComment.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdComment.Row = grdComment.FixedRows
        grdComment.Col = grdComment.FixedCols
        Do
            If mColOk() Then
                mSetShow
                Exit Do
            Else
                If grdComment.Col < CCOMMENTINDEX Then
                    grdComment.Col = grdComment.Col + 1
                Else
                    mSetShow
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
            End If
        Loop
    End If
    mEnableBox
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilNext As Integer
    Dim ilIndex As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim ilLoop As Integer

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = 0
    If imCtrlVisible Then
        If lmEnableCol = CSOURCEINDEX Then
            If Not mBranchSource() Then
                mEnableBox
                Exit Sub
            End If
        End If
        llEnableRow = lmEnableRow
        llEnableCol = lmEnableCol
        mSetShow
        grdComment.Row = llEnableRow
        grdComment.Col = llEnableCol
        Do
            ilNext = False
            Select Case grdComment.Col
                Case CCOMMENTINDEX
                    pbcClickFocus.SetFocus
                    Exit Sub
                Case CFOLLOWUPINDEX
                    grdComment.Col = CSOURCEINDEX
                Case Else
                    grdComment.Col = grdComment.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        grdComment.Row = grdComment.FixedRows
        grdComment.Col = grdComment.FixedCols
        Do
            If mColOk() Then
                mSetShow
                Exit Do
            Else
                If grdComment.Col < CCOMMENTINDEX Then
                    grdComment.Col = grdComment.Col + 1
                Else
                    mSetShow
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
            End If
        Loop
    End If
    mEnableBox
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
    Dim llRow As Long
    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        If grdComment.TextMatrix(llRow, CCHGDINDEX) = "1" Then
            cmcVehicleView.Visible = False
            Exit Sub
        End If
    Next llRow
    cmcVehicleView.Visible = True
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

    If (grdComment.Row < grdComment.FixedRows) Or (grdComment.Row >= grdComment.Rows) Or (grdComment.Col < grdComment.FixedCols) Or (grdComment.Col >= CPOSTEDTIMEINDEX) Then
        Exit Sub
    End If
    lmEnableRow = grdComment.Row
    lmEnableCol = grdComment.Col
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdComment.Left - pbcArrow.Width, grdComment.Top + grdComment.RowPos(grdComment.Row) + (grdComment.RowHeight(grdComment.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    cmcVehicleView.Visible = False

    Select Case grdComment.Col
        Case CCALLETTERINDEX
            mSetLbcGridControl lbcStation
            edcDropdown.MaxLength = 0
        Case CVEHICLEINDEX
            mSetLbcGridControl lbcVehicle
            edcDropdown.MaxLength = 0
        Case CFOLLOWUPINDEX
            mSetCalendarGridControl
        Case COKINDEX
            ckcOK.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30
            grdComment.Col = COKINDEX
            grdComment.CellFontName = "Monotype Sorts"
            'If ckcOK.Height > grdComment.RowHeight(grdComment.Row) - 15 Then
                ckcOK.FontName = "Arial"
                ckcOK.Height = grdComment.RowHeight(grdComment.Row) - 15
            'End If
            If grdComment.TextMatrix(grdComment.Row, COKINDEX) = "4" Then
                ckcOK.Value = vbChecked
            Else
                ckcOK.Value = vbUnchecked
            End If
            
            ckcOK.Visible = True
            ckcOK.SetFocus
        Case CSOURCEINDEX
            mSetLbcGridControl lbcSource
            edcDropdown.MaxLength = 0
        Case CCOMMENTINDEX
            mSetEdcGridControl
            edcDropdown.MaxLength = 255
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
    Dim slStr As String

    If (lmEnableRow >= grdComment.FixedRows) And (lmEnableRow < grdComment.Rows) Then
        Select Case lmEnableCol
            Case CCALLETTERINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                End If
                grdComment.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case CVEHICLEINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                End If
                grdComment.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case CFOLLOWUPINDEX
                slStr = edcDropdown.Text 'cbcDate.Text
                If slStr <> "" Then
                    If Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol)) <> "" Then
                        If gDateValue(Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol))) <> gDateValue(Trim$(slStr)) Then
                            grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                        End If
                    Else
                        grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                    End If
                Else
                    If Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol)) <> "" Then
                        grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                    End If
                End If
                grdComment.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case CSOURCEINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                End If
                grdComment.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case CCOMMENTINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdComment.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdComment.TextMatrix(lmEnableRow, CCHGDINDEX) = "1"
                End If
                grdComment.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case COKINDEX
        End Select
    End If
    lbcVehicleView.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    pbcArrow.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcVehicle.Visible = False
    lbcStation.Visible = False
    lbcSource.Visible = False
    ckcOK.Visible = False
    'cbcDate.Visible = False
    plcCalendar.Visible = False
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

    If (grdComment.Row < grdComment.FixedRows) Or (grdComment.Row >= grdComment.Rows) Or (grdComment.Col < grdComment.FixedCols) Or (grdComment.Col >= grdComment.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdComment.Col - 1 Step 1
        llColPos = llColPos + grdComment.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdComment.ColWidth(grdComment.Col)
    ilCol = grdComment.Col
    Do While ilCol < grdComment.Cols - 1
        If (Trim$(grdComment.TextMatrix(grdComment.Row - 1, grdComment.Col)) <> "") And (Trim$(grdComment.TextMatrix(grdComment.Row - 1, grdComment.Col)) = Trim$(grdComment.TextMatrix(grdComment.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdComment.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdComment.Col
        Case CCALLETTERINDEX
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcStation.Visible = True
            edcDropdown.SetFocus
        Case CVEHICLEINDEX
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcVehicle.Visible = True
            edcDropdown.SetFocus
        Case CFOLLOWUPINDEX
            'cbcDate.Visible = True
            'cbcDate.SetFocus
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case COKINDEX
            ckcOK.Visible = True
            ckcOK.SetFocus
        Case CSOURCEINDEX
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcSource.Visible = True
            edcDropdown.SetFocus
        Case CCOMMENTINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
    End Select
End Sub






Private Sub mSetGridColumns()
    Dim ilCol As Integer
    grdComment.Width = Width - pbcArrow.Width  'grdStations.Width
    grdComment.Height = Height
    gGrid_IntegralHeight grdComment
    grdComment.Height = grdComment.Height + 30
    'grdComment.Move grdStations.Left, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1)
    grdComment.Move pbcArrow.Width, 0
    grdComment.ColWidth(CPOSTEDTIMEINDEX) = 0
    grdComment.ColWidth(CUSERNAMEINDEX) = 0
    grdComment.ColWidth(CCCTCODEINDEX) = 0
    grdComment.ColWidth(CORIGOKINDEX) = 0
    grdComment.ColWidth(CSORTINDEX) = 0
    grdComment.ColWidth(CCHGDINDEX) = 0
    grdComment.ColWidth(CPOSTEDDATEINDEX) = grdComment.Width * 0.06
    grdComment.ColWidth(CBYINDEX) = grdComment.Width * 0.05
    If imGridForm = 0 Then
        grdComment.ColWidth(CCALLETTERINDEX) = 0
    Else
        grdComment.ColWidth(CCALLETTERINDEX) = grdComment.Width * 0.12
    End If
    grdComment.ColWidth(CVEHICLEINDEX) = grdComment.Width * 0.17
    grdComment.ColWidth(CFOLLOWUPINDEX) = grdComment.Width * 0.08
    grdComment.ColWidth(COKINDEX) = grdComment.Width * 0.05
    grdComment.ColWidth(CSOURCEINDEX) = grdComment.Width * 0.04
    grdComment.ColWidth(CDELETEINDEX) = grdComment.Width * 0.06
           
    grdComment.ColWidth(CCOMMENTINDEX) = grdComment.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To CDELETEINDEX Step 1
        If ilCol <> CCOMMENTINDEX Then
            grdComment.ColWidth(CCOMMENTINDEX) = grdComment.ColWidth(CCOMMENTINDEX) - grdComment.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdComment
    grdComment.ColAlignment(CPOSTEDDATEINDEX) = flexAlignRightCenter
    grdComment.ColAlignment(CPOSTEDTIMEINDEX) = flexAlignRightCenter
    grdComment.ColAlignment(CFOLLOWUPINDEX) = flexAlignRightCenter
End Sub

Private Sub mSetGridTitles()
    grdComment.TextMatrix(0, CPOSTEDDATEINDEX) = "Posted"
    grdComment.TextMatrix(0, CBYINDEX) = "By"
    grdComment.TextMatrix(0, CCALLETTERINDEX) = "Call Letters"
    grdComment.TextMatrix(0, CVEHICLEINDEX) = "Vehicle"
    grdComment.TextMatrix(0, CFOLLOWUPINDEX) = "Follow-up"
    grdComment.TextMatrix(0, COKINDEX) = "Done"
    grdComment.TextMatrix(0, CSOURCEINDEX) = "Src"
    grdComment.TextMatrix(0, CCOMMENTINDEX) = "Comment"
    grdComment.TextMatrix(0, CDELETEINDEX) = "Delete"
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
    If grdComment.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    If grdComment.CellForeColor = vbRed Then
        mColOk = False
        Exit Function
    End If
    If grdComment.Col = CBYINDEX Then
        mColOk = False
        Exit Function
    End If
    If grdComment.ColWidth(grdComment.Col) <= 0 Then
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
    Dim ilIndex As Integer
    Dim llVef As Long
    Dim ilShtt As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilCode As Integer
    
    Select Case ilType
        Case 1  'Clear Focus
            mSetShow
            pbcArrow.Visible = False
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            Form_Load
            Form_Activate
            mInit
        Case 3  'Populate
            imLoadVehicleView = True
            mPopCommentGrid
            imLoadVehicleView = False
        Case 4  'Clear
            mClearGrid grdComment
            Screen.MousePointer = vbDefault
            gSetMousePointer grdComment, grdComment, vbDefault
        Case 5  'Save
            mSetShow
            pbcArrow.Visible = False
            ilRet = mSaveRec()
        Case 6  'Add new comment
            grdComment.AddItem "", grdComment.FixedRows
            grdComment.Row = grdComment.FixedRows
            grdComment.TextMatrix(grdComment.Row, CPOSTEDDATEINDEX) = Format$(gNow(), "m/d/yy")
            grdComment.TextMatrix(grdComment.Row, CPOSTEDTIMEINDEX) = Format$(gNow(), "h:mm:ssAM/PM")
            grdComment.Col = CPOSTEDDATEINDEX
            grdComment.CellBackColor = LIGHTYELLOW
            grdComment.TextMatrix(grdComment.Row, CBYINDEX) = smUserName    'sgUserName
            grdComment.Col = CBYINDEX
            grdComment.CellBackColor = lmUserColor  'LIGHTYELLOW
            grdComment.TextMatrix(grdComment.Row, CCALLETTERINDEX) = ""
            For ilShtt = 0 To lbcStation.ListCount - 1 Step 1
                If imShttCode = lbcStation.ItemData(ilShtt) Then
                    grdComment.TextMatrix(grdComment.Row, CCALLETTERINDEX) = Trim$(lbcStation.List(ilShtt))
                    Exit For
                End If
            Next ilShtt
            grdComment.TextMatrix(grdComment.Row, CVEHICLEINDEX) = ""
            If imVefCode > 0 Then
                llVef = gBinarySearchVef(CLng(imVefCode))
                If llVef <> -1 Then
                    grdComment.TextMatrix(grdComment.Row, CVEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
                    ilFound = False
                    For ilLoop = 0 To lbcVehicleView.ListCount - 1 Step 1
                        If lbcVehicleView.ItemData(ilLoop) = imVefCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        If lbcVehicleView.ListIndex <= 0 Then
                            ilCode = -1
                        ElseIf lbcVehicleView.ListIndex = 1 Then
                            ilCode = 0
                        Else
                            ilCode = lbcVehicleView.ItemData(lbcVehicleView.ListIndex)
                        End If
                        lbcVehicleView.RemoveItem 1
                        lbcVehicleView.RemoveItem 0
                        lbcVehicleView.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                        lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = imVefCode
                        lbcVehicleView.AddItem "[All Vehicles]", 0
                        lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = 0
                        lbcVehicleView.AddItem "[All Comments]", 0
                        lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = 0
                        gSetListBoxHeight lbcVehicleView, 6
                        imChgMode = True
                        If ilCode = -1 Then
                            lbcVehicleView.ListIndex = 0
                        ElseIf ilCode = 0 Then
                            lbcVehicleView.ListIndex = 1
                        Else
                            For ilLoop = 2 To lbcVehicleView.ListCount - 1 Step 1
                                If lbcVehicleView.ItemData(ilLoop) = ilCode Then
                                    lbcVehicleView.ListIndex = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                        imChgMode = False
                    End If
                End If
            End If
            grdComment.TextMatrix(grdComment.Row, CFOLLOWUPINDEX) = ""
            grdComment.TextMatrix(grdComment.Row, COKINDEX) = ""
            grdComment.TextMatrix(grdComment.Row, CSOURCEINDEX) = ""
            grdComment.TextMatrix(grdComment.Row, CCOMMENTINDEX) = ""
            grdComment.TextMatrix(grdComment.Row, CCCTCODEINDEX) = "0"
            grdComment.TextMatrix(grdComment.Row, CCHGDINDEX) = "0"
            grdComment.TextMatrix(grdComment.Row, CDELETEINDEX) = "Delete"
            grdComment.Col = CDELETEINDEX
            grdComment.CellBackColor = GRAY
            If imGridForm = 1 Then
                grdComment.Col = CCALLETTERINDEX
            Else
                grdComment.Col = CVEHICLEINDEX
            End If
            grdComment.TopRow = grdComment.FixedRows
            mEnableBox
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

Private Sub UserControl_GotFocus()
    RaiseEvent CommentFocus
End Sub

Private Sub UserControl_Initialize()
    mSetFonts
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Form_MouseUp Button, Shift, X, Y
End Sub


Private Function mSaveRec() As Integer
    Dim llRow As Long
    Dim ilLen As Integer
    Dim ilPos As Integer
    Dim slVehicle As String
    Dim slDate As String
    Dim slOK As String
    Dim slChgd As String
    Dim llCode As Long
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    Dim slComment As String
    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim ilSource As Integer
    Dim ilSrc As Integer

    On Error GoTo ErrHand

    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        slVehicle = Trim(grdComment.TextMatrix(llRow, CVEHICLEINDEX))
        slDate = grdComment.TextMatrix(llRow, CFOLLOWUPINDEX)
        If slDate = "" Then
            slDate = "12/31/2069"
        End If
        slComment = gFixQuote(Trim$(grdComment.TextMatrix(llRow, CCOMMENTINDEX))) '& Chr(0)
        If imGridForm = 1 Then
            For ilShtt = 0 To lbcStation.ListCount - 1 Step 1
                If UCase$(Trim$(grdComment.TextMatrix(llRow, CCALLETTERINDEX))) = UCase(Trim$(lbcStation.List(ilShtt))) Then
                    ilShttCode = lbcStation.ItemData(ilShtt)
                    Exit For
                End If
            Next ilShtt
        Else
            ilShttCode = imShttCode
        End If
        If (slVehicle <> "") And (slComment <> "") And ((imGridForm <> 1) Or ((imGridForm = 1) And (ilShttCode > 0))) Then
            slOK = "N"
            If grdComment.TextMatrix(llRow, COKINDEX) = "4" Then
                slOK = "Y"
            End If
            ilVefCode = 0
            If slVehicle <> "[All Vehicles]" Then
                ilVef = SendMessageByString(lbcVehicle.hwnd, LB_FINDSTRING, -1, slVehicle)
                If ilVef >= 0 Then
                    ilVefCode = Val(lbcVehicle.ItemData(ilVef))
                End If
            End If
            ilSource = 0
            ilSrc = SendMessageByString(lbcSource.hwnd, LB_FINDSTRING, -1, grdComment.TextMatrix(llRow, CSOURCEINDEX))
            If ilSrc >= 0 Then
                ilSource = lbcSource.ItemData(ilSrc)
            End If
            llCode = Val(grdComment.TextMatrix(llRow, CCCTCODEINDEX))
            slChgd = grdComment.TextMatrix(llRow, CCHGDINDEX)
            If (slChgd = "1") And (imGridForm = 1) And (llCode > 0) Then
                'Key 1 (shttCode) can not be modified, so delete record and insert a new one
                SQLQuery = "SELECT * FROM cct WHERE cctCode = " & llCode
                Set rst_cct = gSQLSelectCall(SQLQuery)
                If Not rst_cct.EOF Then
                    If rst_cct!cctShfCode <> ilShttCode Then
                        SQLQuery = "DELETE FROM cct WHERE cctCode = " & llCode
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/13/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "AffCommentGrid-mSaveRec"
                            mSaveRec = False
                            Exit Function
                        End If
                        llCode = 0
                    End If
                End If
            End If
            If slChgd = "1" Then
                If llCode > 0 Then
                    SQLQuery = "Update cct Set "
                    'SQLQuery = SQLQuery & "cctShfCode = " & imShttCode & ", "
                    SQLQuery = SQLQuery & "cctVefCode = " & ilVefCode & ", "
                    SQLQuery = SQLQuery & "cctActionDate = '" & Format$(slDate, sgSQLDateForm) & "', "
                    If slOK <> grdComment.TextMatrix(llRow, CORIGOKINDEX) Then
                        SQLQuery = SQLQuery & "cctDone = '" & slOK & "', "
                        SQLQuery = SQLQuery & "cctDoneUstCode = " & igUstCode & ", "
                        If slOK = "Y" Then
                            SQLQuery = SQLQuery & "cctDoneDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "cctDoneTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                        Else
                            SQLQuery = SQLQuery & "cctDoneDate = '" & Format$("1/1/1970", sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "cctDoneTime = '" & Format$("12AM", sgSQLTimeForm) & "', "
                        End If
                    End If
                    SQLQuery = SQLQuery & "cctChgdUstCode = " & igUstCode & ", "
                    SQLQuery = SQLQuery & "cctChgdDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "cctChgdTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                    SQLQuery = SQLQuery & "cctCstCode = " & ilSource & ", "
                    SQLQuery = SQLQuery & "cctComment = '" & slComment & "' "
                    'SQLQuery = SQLQuery & "cctUstCode = " & igUstCode & ", "
                    'SQLQuery = SQLQuery & "cctEnteredDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "cctEnteredTime = '" & Format$(gNow(), sgSQLTimeForm) & "' "
                    SQLQuery = SQLQuery & " WHERE cctCode = " & llCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "AffCommentGrid-mSaveRec"
                        mSaveRec = False
                        Exit Function
                    End If
                Else
                    Do
                        SQLQuery = "SELECT MAX(cctCode) from cct"
                        Set rst_cct = gSQLSelectCall(SQLQuery)
                        If IsNull(rst_cct(0).Value) Then
                            llCode = 1
                        Else
                            If Not rst_cct.EOF Then
                                llCode = rst_cct(0).Value + 1
                            Else
                                llCode = 1
                            End If
                        End If
                        ilRet = 0
                        SQLQuery = "INSERT INTO cct (cctCode, cctShfCode, cctVefCode, cctActionDate, cctToEMailUstCode, cctCstCode, cctDone, cctDoneUstCode, cctDoneDate, cctDoneTime, cctChgdUstCode, cctChgdDate, cctChgdTime, cctComment, cctUstCode, cctEnteredDate, cctEnteredTime, cctUnused)"
                        SQLQuery = SQLQuery & " VALUES ("
                        SQLQuery = SQLQuery & llCode & ", "
                        SQLQuery = SQLQuery & ilShttCode & ", "
                        SQLQuery = SQLQuery & ilVefCode & ", "
                        SQLQuery = SQLQuery & "'" & Format$(slDate, sgSQLDateForm) & "', "
                        SQLQuery = SQLQuery & "'" & 0 & "', "
                        SQLQuery = SQLQuery & ilSource & ", "
                        SQLQuery = SQLQuery & "'" & slOK & "', "
                        If slOK = "Y" Then
                            SQLQuery = SQLQuery & igUstCode & ", "
                            SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
                        Else
                            SQLQuery = SQLQuery & "'" & 0 & "', "
                            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
                        End If
                        SQLQuery = SQLQuery & "'" & 0 & "', "
                        SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
                        SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
                        SQLQuery = SQLQuery & "'" & slComment & "', "
                        SQLQuery = SQLQuery & igUstCode & ", "
                        SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
                        SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
                        SQLQuery = SQLQuery & "'" & "" & "'" & ")"
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/13/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "AffCommentGrid-mSaveRec"
                            mSaveRec = False
                            Exit Function
                        End If
                    Loop While ilRet <> 0
                    grdComment.TextMatrix(llRow, CCCTCODEINDEX) = llCode
                    RaiseEvent CommentStatus("C")
                End If
                grdComment.TextMatrix(llRow, CCHGDINDEX) = "0"
            End If
        End If
    Next llRow
    mSetCommands
    mSaveRec = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffCommentGrid-mSaveRec"
    mSaveRec = False
End Function

Public Property Get Verify() As Integer 'VBC NR
    pbcArrow.Visible = False 'VBC NR
    If imUpdateAllowed Then 'VBC NR
        'Add call to mTestFields
        Verify = True 'VBC NR
    Else 'VBC NR
        Verify = True 'VBC NR
    End If 'VBC NR
End Property 'VBC NR

Private Function mPopCommentGrid() As Integer
    Dim llRow As Long
    Dim llVef As Long
    Dim ilCol As Integer
    Dim llYellowRow As Long
    Dim llCol As Long
    Dim ilShtt As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilIncludeComment As Integer
    Dim ilIndex As Integer
    Dim ilDntCode As Integer
    Dim ilUstCode As Integer
    Dim ilSrc As Integer
    Dim ilVefCode As Integer
    Dim ilVehicleIndex As Integer
    Dim ilFound As Integer
    
    mPopCommentGrid = False
    
    If imLoadVehicleView Then
        ilVehicleIndex = -1
        lbcVehicleView.Clear
    Else
        ilVehicleIndex = lbcVehicleView.ListIndex
    End If
    On Error GoTo ErrHand:
    Screen.MousePointer = vbHourglass
    RaiseEvent GetCommentType
    gSetMousePointer grdComment, grdComment, vbHourglass
    slDate = Format$(gNow(), "m/d/yyyy")
    If Weekday(slDate, vbMonday) <= vbThursday Then
        slDate = gObtainNextSunday(slDate)
    Else
        slDate = gObtainNextSunday(slDate)
        slDate = DateAdd("d", 7, slDate)
    End If
    'If Follow-up only allow Mine
    If imGridForm = 1 Then
        smWhichComment = "M"
    End If
    Select Case smWhichComment
        Case "A"    'All
            ilDntCode = -1
            ilUstCode = -1
        Case "M"    'Mine
            ilDntCode = -1
            ilUstCode = igUstCode
        Case "D"    'Department
            ilUstCode = -1
            SQLQuery = "SELECT ustDntCode FROM Ust Where ustCode = " & igUstCode
            Set rst_Ust = gSQLSelectCall(SQLQuery)
            If Not rst_Ust.EOF Then
                If rst_Ust!ustDntCode > 0 Then
                    ilDntCode = rst_Ust!ustDntCode
                Else
                    ilDntCode = -1
                End If
            Else
                ilDntCode = -1
            End If
        Case Else
            ilDntCode = -1
            ilUstCode = -1
    End Select
    llDate = gDateValue(slDate)
    lbcStation.Clear
    grdComment.Rows = 2
    mClearGrid grdComment
    gGrid_FillWithRows grdComment
    grdComment.Redraw = False
    grdComment.Row = 0
    grdComment.Col = CPOSTEDDATEINDEX
    grdComment.CellBackColor = LIGHTBLUE
    grdComment.Col = CBYINDEX
    grdComment.CellBackColor = LIGHTBLUE
    grdComment.Col = CCALLETTERINDEX
    grdComment.CellBackColor = LIGHTBLUE
    grdComment.Col = CVEHICLEINDEX
    grdComment.CellBackColor = LIGHTBLUE
    grdComment.Col = CFOLLOWUPINDEX
    grdComment.CellBackColor = LIGHTBLUE
    llRow = grdComment.FixedRows
    ilLoop = 0
    'imShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
    Do While ((imGridForm = 1) And ((ilLoop < UBound(igCommentShttCode)) Or (Not imFilterStatus))) Or (imGridForm = 0)
        If imGridForm = 0 Then
            SQLQuery = "SELECT * FROM cct"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " cctShfCode = " & imShttCode & ")"
        Else
            'Missing OK test
            If imFilterStatus Then
                SQLQuery = "SELECT * FROM cct"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " cctShfCode = " & igCommentShttCode(ilLoop) & ")"
            Else
                SQLQuery = "SELECT * FROM cct"
            End If
        End If
        SQLQuery = SQLQuery & " ORDER BY cctEnteredDate Desc, cctEnteredTime Desc"
        Set rst_cct = gSQLSelectCall(SQLQuery)
        Do While Not rst_cct.EOF
            If imGridForm = 1 Then
                'Follow-up date must be defined and less then next sunday and Done not Checked
                ilIncludeComment = False
                If rst_cct!cctDone <> "Y" Then
                    If Not IsNull(rst_cct!cctActionDate) Then
                        If gDateValue(rst_cct!cctActionDate) <> gDateValue("12/31/2069") Then
                            If gDateValue(rst_cct!cctActionDate) <= llDate Then
                                ilIncludeComment = True
                            End If
                        End If
                    End If
                End If
            Else
                ilIncludeComment = True
            End If
            If ilIncludeComment And ilDntCode >= 0 Then
                If rst_cct!cctUstCode > 0 Then
                    SQLQuery = "SELECT ustDntCode FROM Ust Where ustCode = " & rst_cct!cctUstCode
                    Set rst_Ust = gSQLSelectCall(SQLQuery)
                    If Not rst_Ust.EOF Then
                        If rst_Ust!ustDntCode <> ilDntCode Then
                            ilIncludeComment = False
                        End If
                    Else
                        ilIncludeComment = False
                    End If
                Else
                    ilIncludeComment = False
                End If
                If ilIncludeComment = False Then
                    If rst_cct!cctToEMailUstCode > 0 Then
                        SQLQuery = "SELECT ustDntCode FROM Ust Where ustCode = " & rst_cct!cctToEMailUstCode
                        Set rst_Ust = gSQLSelectCall(SQLQuery)
                        If Not rst_Ust.EOF Then
                            If rst_Ust!ustDntCode = ilDntCode Then
                                ilIncludeComment = True
                            End If
                        End If
                    End If
                End If
            End If
            If ilIncludeComment And ilUstCode >= 0 Then
                If (rst_cct!cctUstCode <> ilUstCode) And (rst_cct!cctToEMailUstCode <> ilUstCode) Then
                    ilIncludeComment = False
                End If
            End If
            If (ilIncludeComment) And (imLoadVehicleView) Then
                ilFound = False
                For ilLoop = 0 To lbcVehicleView.ListCount - 1 Step 1
                    If lbcVehicleView.ItemData(ilLoop) = rst_cct!cctVefCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    llVef = gBinarySearchVef(CLng(rst_cct!cctVefCode))
                    If llVef <> -1 Then
                        lbcVehicleView.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                        lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = rst_cct!cctVefCode
                    End If
                End If
            End If
            If (ilIncludeComment) And (ilVehicleIndex > 0) Then
                ilVefCode = lbcVehicleView.ItemData(ilVehicleIndex)
                If rst_cct!cctVefCode > 0 Then
                    If rst_cct!cctVefCode <> ilVefCode Then
                        ilIncludeComment = False
                    End If
                End If
            End If
            If ilIncludeComment Then
                If llRow >= grdComment.Rows Then
                    grdComment.AddItem ""
                End If
                grdComment.Row = llRow
                For ilCol = CPOSTEDDATEINDEX To CCOMMENTINDEX Step 1
                    If ilCol <> COKINDEX Then
                        grdComment.Col = ilCol
                        If (ilCol = CCOMMENTINDEX) And ((StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) Or (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0)) Then
                        ElseIf ilCol = CCALLETTERINDEX Then
                            grdComment.CellBackColor = LIGHTGREENCOLOR
                        Else
                            If (sgUstAllowCmmtChg = "Y") And ((ilCol = CVEHICLEINDEX) Or (ilCol = CFOLLOWUPINDEX) Or (ilCol = CSOURCEINDEX) Or (ilCol = CCOMMENTINDEX)) Then
                            Else
                                grdComment.CellBackColor = LIGHTYELLOW
                            End If
                        End If
                    End If
                Next ilCol
                grdComment.Col = CDELETEINDEX
                If sgUstAllowCmmtDelete = "Y" Then
                    grdComment.TextMatrix(llRow, CDELETEINDEX) = "Delete"
                    grdComment.CellBackColor = GRAY
                Else
                    grdComment.CellBackColor = LIGHTYELLOW
                End If
                If Not IsNull(rst_cct!cctEnteredDate) Then
                    grdComment.TextMatrix(llRow, CPOSTEDDATEINDEX) = Format$(Trim$(rst_cct!cctEnteredDate), "m/d/yy")
                Else
                    grdComment.TextMatrix(llRow, CPOSTEDDATEINDEX) = ""
                End If
                If Not IsNull(rst_cct!cctEnteredTime) Then
                    grdComment.TextMatrix(llRow, CPOSTEDTIMEINDEX) = Format$(Trim$(rst_cct!cctEnteredTime), "h:mm:ssAM/PM")
                Else
                    grdComment.TextMatrix(llRow, CPOSTEDTIMEINDEX) = ""
                End If
                grdComment.TextMatrix(llRow, CBYINDEX) = ""
                If rst_cct!cctUstCode > 0 Then
                    SQLQuery = "SELECT ustname, ustReportName, ustUserInitials, ustDntCode FROM Ust Where ustCode = " & rst_cct!cctUstCode
                    Set rst_Ust = gSQLSelectCall(SQLQuery)
                    If Not rst_Ust.EOF Then
                        If rst_Ust!ustDntCode > 0 Then
                            SQLQuery = "SELECT dntColor FROM Dnt Where dntCode = " & rst_Ust!ustDntCode
                            Set rst_dnt = gSQLSelectCall(SQLQuery)
                            If Not rst_dnt.EOF Then
                                grdComment.Col = CBYINDEX
                                grdComment.CellBackColor = rst_dnt!dntColor
                            End If
                        End If
                        If Trim$(rst_Ust!ustUserInitials) <> "" Then
                            grdComment.TextMatrix(llRow, CBYINDEX) = Trim$(rst_Ust!ustUserInitials)
                        Else
                            If Trim$(rst_Ust!ustReportName) <> "" Then
                                grdComment.TextMatrix(llRow, CBYINDEX) = Trim$(rst_Ust!ustReportName)
                            Else
                                grdComment.TextMatrix(llRow, CBYINDEX) = Trim$(rst_Ust!ustname)
                            End If
                        End If
                        If Trim$(rst_Ust!ustReportName) <> "" Then
                            grdComment.TextMatrix(llRow, CUSERNAMEINDEX) = Trim$(rst_Ust!ustReportName)
                        Else
                            grdComment.TextMatrix(llRow, CUSERNAMEINDEX) = Trim$(rst_Ust!ustname)
                        End If
                    End If
                End If
                If rst_cct!cctShfCode > 0 Then
                    ilShtt = gBinarySearchStationInfoByCode(CLng(rst_cct!cctShfCode))
                    If ilShtt <> -1 Then
                        grdComment.TextMatrix(llRow, CCALLETTERINDEX) = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                        ilIndex = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, grdComment.TextMatrix(llRow, CCALLETTERINDEX))
                        If ilIndex < 0 Then
                            lbcStation.AddItem grdComment.TextMatrix(llRow, CCALLETTERINDEX)
                            lbcStation.ItemData(lbcStation.NewIndex) = rst_cct!cctShfCode
                        End If
                    Else
                        grdComment.TextMatrix(llRow, CCALLETTERINDEX) = ""
                    End If
                Else
                    grdComment.TextMatrix(llRow, CCALLETTERINDEX) = ""
                End If
                If rst_cct!cctVefCode > 0 Then
                    llVef = gBinarySearchVef(CLng(rst_cct!cctVefCode))
                    If llVef <> -1 Then
                        grdComment.TextMatrix(llRow, CVEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
                    Else
                        grdComment.TextMatrix(llRow, CVEHICLEINDEX) = ""
                    End If
                Else
                    grdComment.TextMatrix(llRow, CVEHICLEINDEX) = "[All Vehicles]"
                End If
                If Not IsNull(rst_cct!cctActionDate) Then
                    If gDateValue(rst_cct!cctActionDate) <> gDateValue("12/31/2069") Then
                        grdComment.TextMatrix(llRow, CFOLLOWUPINDEX) = Format$(Trim$(rst_cct!cctActionDate), "m/d/yy")
                    Else
                        grdComment.TextMatrix(llRow, CFOLLOWUPINDEX) = ""
                    End If
                Else
                    grdComment.TextMatrix(llRow, CFOLLOWUPINDEX) = ""
                End If
                grdComment.Col = COKINDEX
                grdComment.CellFontName = "Monotype Sorts"
                grdComment.TextMatrix(llRow, COKINDEX) = ""
                grdComment.TextMatrix(llRow, CORIGOKINDEX) = ""
                If rst_cct!cctDone = "Y" Then
                    grdComment.TextMatrix(llRow, COKINDEX) = "4"
                    grdComment.TextMatrix(llRow, CORIGOKINDEX) = "Y"
                End If
                grdComment.TextMatrix(llRow, CSOURCEINDEX) = ""
                For ilSrc = 0 To lbcSource.ListCount - 1 Step 1
                    If rst_cct!cctCstCode = lbcSource.ItemData(ilSrc) Then
                        lbcSource.ListIndex = ilSrc
                        grdComment.TextMatrix(llRow, CSOURCEINDEX) = lbcSource.List(ilSrc)
                        Exit For
                    End If
                Next ilSrc
                grdComment.TextMatrix(llRow, CCOMMENTINDEX) = Trim$(rst_cct!cctComment)
                grdComment.TextMatrix(llRow, CCCTCODEINDEX) = rst_cct!cctCode
                llRow = llRow + 1
            End If
            rst_cct.MoveNext
        Loop
        If imGridForm = 0 Then
            Exit Do
        End If
        If Not imFilterStatus Then
            Exit Do
        End If
        ilLoop = ilLoop + 1
    Loop
    grdComment.Rows = grdComment.Rows + (grdComment.Height \ grdComment.RowHeight(1))
    For llYellowRow = llRow To grdComment.Rows - 1 Step 1
        grdComment.Row = llYellowRow
        For llCol = CPOSTEDDATEINDEX To CDELETEINDEX Step 1
            grdComment.Col = llCol
            grdComment.CellBackColor = LIGHTYELLOW
            grdComment.TextMatrix(llYellowRow, CCCTCODEINDEX) = 0
            grdComment.TextMatrix(llYellowRow, CCHGDINDEX) = "0"
        Next llCol
    Next llYellowRow
    If imGridForm = 1 Then
        imLastColSorted = -1
        imLastSort = -1
        mCommentSortCol CPOSTEDDATEINDEX
        mCommentSortCol CPOSTEDDATEINDEX
    End If
    grdComment.Row = 0
    grdComment.Col = CCCTCODEINDEX
    grdComment.Redraw = True
    If imLoadVehicleView Then
        lbcVehicleView.AddItem "[All Vehicles]", 0
        lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = 0
        lbcVehicleView.AddItem "[All Comments]", 0
        lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = 0
        gSetListBoxHeight lbcVehicleView, 6
        imChgMode = True
        lbcVehicleView.ListIndex = 0
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    gSetMousePointer grdComment, grdComment, vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffCommentGrid-mPopCommentGrid"
    grdComment.Redraw = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdComment, grdComment, vbDefault
End Function

Private Sub mClearGrid(grdCtrl As MSHFlexGrid)
    Dim llRow As Long
    Dim llCol As Long
    
    'Set color within cells
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        For llCol = 0 To grdCtrl.Cols - 1 Step 1
            grdCtrl.Row = llRow
            grdCtrl.Col = llCol
            grdCtrl.Text = ""
            grdCtrl.CellBackColor = vbWhite
        Next llCol
    Next llRow
End Sub


Public Property Let StationCode(ilShttCode As Integer)
    'UserControl.Enabled = ilState
    imShttCode = ilShttCode
    PropertyChanged "StationCode"
End Property
Public Property Get StationCode() As Integer
   StationCode = imShttCode
End Property

Public Property Get ColumnStartLocation(ilColumnNumber As Integer) As Long
   ColumnStartLocation = grdComment.ColPos(ilColumnNumber)
End Property

Public Property Get FollowUpAddAllowed() As Boolean
    If lbcStation.ListCount > 0 Then
        FollowUpAddAllowed = True
    Else
        FollowUpAddAllowed = False
    End If
End Property


Public Property Let VehicleCode(ilVefCode As Integer)
    'UserControl.Enabled = ilState
    imVefCode = ilVefCode
    PropertyChanged "VehicleCode"
End Property
Public Property Get VehicleCode() As Integer
   VehicleCode = imVefCode
End Property
Public Property Let CommentGridForm(ilGridForm As Integer)
    'UserControl.Enabled = ilState
    imGridForm = ilGridForm
    PropertyChanged "CommentGridForm"
    mSetGridColumns
    'mPopCommentGrid
End Property

Public Property Get CommentGridForm() As Integer
   CommentGridForm = imGridForm
End Property

Public Property Let FilterStatus(ilFilterStatus As Integer)
    'UserControl.Enabled = ilState
    imFilterStatus = ilFilterStatus
    PropertyChanged "FilterStatus"
    mSetGridColumns
    'mPopCommentGrid
End Property


Public Property Let WhichComment(slWhichComment As String)
    'UserControl.Enabled = ilState
    smWhichComment = slWhichComment
    PropertyChanged "WhichComment"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'imShttCode = PropBag.ReadProperty("StationCode", 0)
    imGridForm = PropBag.ReadProperty("CommentGridForm", 0)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Call PropBag.WriteProperty("StationCode", 0)
    Call PropBag.WriteProperty("CommentGridForm", imGridForm)
End Sub


Private Sub UserControl_Resize()
    pbcArrow.Width = 90
    grdComment.Width = Width - pbcArrow.Width
    grdComment.Height = Height
    grdComment.Move pbcArrow.Width, 0
    gGrid_IntegralHeight grdComment
    grdComment.Height = grdComment.Height + 30
    UserControl.Height = grdComment.Height
    mSetGridColumns

    cmcVehicleView.Width = cmcCalDn.Width
    cmcVehicleView.Left = grdComment.Left + grdComment.ColPos(CFOLLOWUPINDEX) - cmcVehicleView.Width
    cmcVehicleView.Top = grdComment.Top + 15
    cmcVehicleView.Height = grdComment.RowHeight(0) - 15
    cmcVehicleView.Font = "Monotype Sorts"
    cmcVehicleView.Visible = True
    lbcVehicleView.Width = grdComment.ColPos(CFOLLOWUPINDEX) - grdComment.ColPos(CBYINDEX)
    lbcVehicleView.Left = cmcVehicleView.Left + cmcVehicleView.Width - lbcVehicleView.Width
    lbcVehicleView.Top = cmcVehicleView.Top + cmcVehicleView.Height + 15

End Sub

Private Sub mCommentSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        slStr = Trim$(grdComment.TextMatrix(llRow, CVEHICLEINDEX))
        If slStr <> "" Then
            If ilCol = CPOSTEDDATEINDEX Then
                If Trim$(grdComment.TextMatrix(llRow, CPOSTEDDATEINDEX)) <> "" Then
                    slSort = Trim$(Str$(gDateValue(grdComment.TextMatrix(llRow, CPOSTEDDATEINDEX))))
                Else
                    slSort = ""
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
                If Trim$(grdComment.TextMatrix(llRow, CPOSTEDTIMEINDEX)) <> "" Then
                    slStr = Trim$(Str$(gTimeToLong(grdComment.TextMatrix(llRow, CPOSTEDTIMEINDEX), False)))
                Else
                    slStr = ""
                End If
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
                slSort = slSort & slStr
            ElseIf ilCol = CFOLLOWUPINDEX Then
                If Trim$(grdComment.TextMatrix(llRow, CFOLLOWUPINDEX)) <> "" Then
                    slSort = Trim$(Str$(gDateValue(grdComment.TextMatrix(llRow, CFOLLOWUPINDEX))))
                Else
                    slSort = ""
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdComment.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slStr = grdComment.TextMatrix(llRow, CSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdComment.TextMatrix(llRow, CSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdComment.TextMatrix(llRow, CSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = CSORTINDEX
    Else
        imLastColSorted = -1
        imLastColSorted = -1
    End If
    gGrid_SortByCol grdComment, CVEHICLEINDEX, CSORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Sub mSetLbcGridControl(lbcCtrl As ListBox)
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    
    If (grdComment.Row < grdComment.FixedRows) Or (grdComment.Row >= grdComment.Rows) Or (grdComment.Col < grdComment.FixedCols) Or (grdComment.Col >= CPOSTEDTIMEINDEX) Then
        Exit Sub
    End If
    If lmEnableCol = CCALLETTERINDEX Then
        edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - cmcDropDown.Width - 30, grdComment.RowHeight(grdComment.Row) - 15
    ElseIf lmEnableCol = CSOURCEINDEX Then
        edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, 4 * grdComment.ColWidth(grdComment.Col) - cmcDropDown.Width - 30, grdComment.RowHeight(grdComment.Row) - 15
    Else
        'edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, 3 * grdComment.ColWidth(grdComment.Col) - cmcDropDown.Width - 30, grdComment.RowHeight(grdComment.Row) - 15
        edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, 4 * pbcTextWidth.TextWidth("nnnnnnnnnn"), grdComment.RowHeight(grdComment.Row) - 15
    End If
    cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
    lbcCtrl.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
    gSetListBoxHeight lbcCtrl, 6
    If lbcCtrl.Top + lbcCtrl.Height > grdComment.Height Then
        lbcCtrl.Top = edcDropdown.Top - lbcCtrl.Height
        If lbcCtrl.Top <= 0 Then
            lbcCtrl.Move cmcDropDown.Left + cmcDropDown.Width, (grdComment.Height - lbcCtrl.Height) / 2, edcDropdown.Width + cmcDropDown.Width
        End If
    Else
        If lbcCtrl.Top + lbcCtrl.Height > grdComment.Height Then
            lbcCtrl.Move cmcDropDown.Left + cmcDropDown.Width, (grdComment.Height - lbcCtrl.Height) / 2, edcDropdown.Width + cmcDropDown.Width
        End If
    End If
    slStr = grdComment.Text
    ilIndex = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRING, -1, slStr)
    If ilIndex >= 0 Then
        lbcCtrl.ListIndex = ilIndex
        edcDropdown.Text = lbcCtrl.List(lbcCtrl.ListIndex)
    Else
        If (lmEnableCol = CSOURCEINDEX) And (imSourceDefaultCode <> -1) Then
            For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
                If lbcSource.ItemData(ilLoop) = imSourceDefaultCode Then
                    lbcCtrl.ListIndex = ilLoop
                    edcDropdown.Text = lbcCtrl.List(ilLoop)
                End If
            Next ilLoop
        Else
            lbcCtrl.ListIndex = -1
            edcDropdown.Text = ""
        End If
    End If
    If edcDropdown.Height > grdComment.RowHeight(grdComment.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdComment.RowHeight(grdComment.Row) - 15
    End If
    edcDropdown.Visible = True
    cmcDropDown.Visible = True
    lbcCtrl.Visible = True
    edcDropdown.SetFocus
End Sub


Private Sub mSetEdcGridControl()
    If (grdComment.Row < grdComment.FixedRows) Or (grdComment.Row >= grdComment.Rows) Or (grdComment.Col < grdComment.FixedCols) Or (grdComment.Col >= CPOSTEDTIMEINDEX) Then
        Exit Sub
    End If
    edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
    edcDropdown.Text = grdComment.Text
    If edcDropdown.Height > grdComment.RowHeight(grdComment.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdComment.RowHeight(grdComment.Row) - 15
    End If
    edcDropdown.Visible = True
    edcDropdown.SetFocus
End Sub
Private Sub mSetCalendarGridControl()
    If (grdComment.Row < grdComment.FixedRows) Or (grdComment.Row >= grdComment.Rows) Or (grdComment.Col < grdComment.FixedCols) Or (grdComment.Col >= CPOSTEDTIMEINDEX) Then
        Exit Sub
    End If
'    edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
'    edcDropdown.Text = grdComment.Text
'    If edcDropdown.Height > grdComment.RowHeight(grdComment.Row) - 15 Then
'        edcDropdown.FontName = "Arial"
'        edcDropdown.Height = grdComment.RowHeight(grdComment.Row) - 15
'    End If
'    edcDropdown.Visible = True
'    edcDropdown.SetFocus
    edcDropdown.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
    cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
    plcCalendar.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height
    If plcCalendar.Top + plcCalendar.Height > grdComment.Height Then
        plcCalendar.Top = edcDropdown.Top - plcCalendar.Height
        If plcCalendar.Top <= 0 Then
            plcCalendar.Move cmcDropDown.Left + cmcDropDown.Width, (grdComment.Height - plcCalendar.Height) / 2
        End If
    Else
        If plcCalendar.Top + plcCalendar.Height > grdComment.Height Then
            plcCalendar.Move cmcDropDown.Left + cmcDropDown.Width, (grdComment.Height - plcCalendar.Height) / 2
        End If
    End If
    lacCalName.FontBold = True
    lacCalName.FontName = "Arial"
    lacCalName.FontSize = 8
    lacCalName.FontBold = True
    plcCalendar.FontBold = True
    pbcCalendar.FontName = "Arial"
    pbcCalendar.FontSize = 8
    pbcCalendar.FontBold = True
    edcDropdown.Text = grdComment.Text
    If edcDropdown.Height > grdComment.RowHeight(grdComment.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdComment.RowHeight(grdComment.Row) - 15
    End If
    edcDropdown.Visible = True
    cmcDropDown.Visible = True
    edcDropdown.SetFocus
End Sub

Private Sub mPopVehicle()
    Dim ilVef As Integer

    lbcVehicle.Clear
    For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicle.AddItem (Trim(tgVehicleInfo(ilVef).sVehicle))
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgVehicleInfo(ilVef).iCode
        'lbcVehicleView.AddItem (Trim(tgVehicleInfo(ilVef).sVehicle))
        'lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = tgVehicleInfo(ilVef).iCode
    Next ilVef
    lbcVehicle.AddItem "[All Vehicles]", 0
    lbcVehicle.ItemData(lbcVehicle.NewIndex) = 0
    'lbcVehicleView.AddItem "[All Vehicles]", 0
    'lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = 0
    'lbcVehicleView.AddItem "[None]", 0
    'lbcVehicleView.ItemData(lbcVehicleView.NewIndex) = 0
    'imChgMode = True
    'lbcVehicleView.ListIndex = 0
    'imChgMode = False
    Exit Sub
End Sub

Private Sub mDropdownChangeEvent(lbcCtrl As ListBox)
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer

    slStr = edcDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        lbcCtrl.ListIndex = llRow
        edcDropdown.Text = lbcCtrl.List(lbcCtrl.ListIndex)
        edcDropdown.SelStart = ilLen
        edcDropdown.SelLength = Len(edcDropdown.Text)
    End If
End Sub

Public Sub mSetFonts()
    Dim Ctrl As control
    Dim ilFontSize As Integer
    Dim ilColorFontSize As Integer
    Dim ilBold As Integer
    Dim ilChg As Integer
    Dim slStr As String
    Dim slFontName As String
    
    
    'On Error Resume Next
    ilFontSize = 14
    ilBold = True
    ilColorFontSize = 10
    slFontName = "Arial"
    If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
        ilFontSize = 10
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 800 Then
        ilFontSize = 10
        ilBold = True
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 1024 Then
        ilFontSize = 12
        ilBold = True
    End If
    For Each Ctrl In UserControl.Controls
        If TypeOf Ctrl Is MSHFlexGrid Then
            Ctrl.Font.Name = slFontName
            Ctrl.FontFixed.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.FontFixed.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
            Ctrl.FontFixed.Bold = ilBold
        ElseIf TypeOf Ctrl Is TabStrip Then
            Ctrl.Font.Name = slFontName
            Ctrl.Font.Size = ilFontSize
            Ctrl.Font.Bold = ilBold
        ''ElseIf TypeOf Ctrl Is Resize Then
        ''ElseIf TypeOf Ctrl Is Timer Then
        ''ElseIf TypeOf Ctrl Is Image Then
        ''ElseIf TypeOf Ctrl Is ImageList Then
        ''ElseIf TypeOf Ctrl Is CommonDialog Then
        ''ElseIf TypeOf Ctrl Is AffExportCriteria Then
        ''ElseIf TypeOf Ctrl Is AffCommentGrid Then
        ''ElseIf TypeOf Ctrl Is AffContactGrid Then
        ''Else
        'ElseIf (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Label) Then
        ElseIf (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is ListBox) _
               Or (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is PictureBox) _
               Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Label) _
               Or (TypeOf Ctrl Is CSI_Calendar) Or (TypeOf Ctrl Is CSI_Calendar_UP) Or (TypeOf Ctrl Is CSI_ComboBoxList) Or (TypeOf Ctrl Is CSI_DayPicker) Then
            ilChg = 0
            If TypeOf Ctrl Is CommandButton Then
               ilChg = 1
            Else
                If (Ctrl.ForeColor = vbBlack) Or (Ctrl.ForeColor = &H80000008) Or (Ctrl.ForeColor = &H80000012) Or (Ctrl.ForeColor = &H8000000F) Then
                    ilChg = 1
                Else
                    ilChg = 2
                End If
            End If
            slStr = Ctrl.Name
            If (InStr(1, slStr, "Arrow", vbTextCompare) > 0) Or ((InStr(1, slStr, "Dropdown", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Or ((InStr(1, slStr, "cmcCal", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Then
                ilChg = 0
            End If
            If ilChg = 1 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilFontSize
                Ctrl.FontBold = ilBold
            ElseIf ilChg = 2 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilColorFontSize
                Ctrl.FontBold = False
            End If
        End If
    Next Ctrl
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    rst_cct.Close
    rst_Ust.Close
    rst_dnt.Close
    rst_cst.Close
End Sub
Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
    gSetMousePointer grdComment, grdComment, ilMousepointer
End Sub


Private Sub mGetUserInfo()
    On Error GoTo ErrHand:
    smUserName = sgUserName
    lmUserColor = LIGHTYELLOW
    SQLQuery = "SELECT ustname, ustReportName, ustUserInitials, ustDntCode FROM Ust Where ustCode = " & igUstCode
    Set rst_Ust = gSQLSelectCall(SQLQuery)
    If Not rst_Ust.EOF Then
        If rst_Ust!ustDntCode > 0 Then
            SQLQuery = "SELECT dntColor FROM Dnt Where dntCode = " & rst_Ust!ustDntCode
            Set rst_dnt = gSQLSelectCall(SQLQuery)
            If Not rst_dnt.EOF Then
                lmUserColor = rst_dnt!dntColor
            End If
        End If
        If Trim$(rst_Ust!ustUserInitials) <> "" Then
            smUserName = Trim$(rst_Ust!ustUserInitials)
        Else
            If Trim$(rst_Ust!ustReportName) <> "" Then
                smUserName = Trim$(rst_Ust!ustReportName)
            Else
                smUserName = Trim$(rst_Ust!ustname)
            End If
        End If
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffCommentGrid-mGetUserInfo"
End Sub

Private Sub mPopSource()
    imSourceDefaultCode = -1
    lbcSource.Clear
    SQLQuery = "SELECT * FROM Cst Order by cstSortCode"
    Set rst_cst = gSQLSelectCall(SQLQuery)
    Do While Not rst_cst.EOF
        lbcSource.AddItem Trim$(rst_cst!cstName)
        lbcSource.ItemData(lbcSource.NewIndex) = rst_cst!cstCode
        If rst_cst!cstDefault = "Y" Then
            imSourceDefaultCode = rst_cst!cstCode
        End If
        rst_cst.MoveNext
    Loop
    lbcSource.AddItem "[New]", 0
    lbcSource.ItemData(lbcSource.NewIndex) = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffCommentGrid-mPopSource"
End Sub

Private Function mBranchSource() As Integer
    Dim ilLoop As Integer
    
    If (edcDropdown.Text = "[New]") Or (imDoubleClick) Then
        sgCmmtSrcName = ""
        If imDoubleClick Then
            sgCmmtSrcName = Trim$(edcDropdown.Text)
        End If
        frmCmmtSrc.Show vbModal
        mPopSource
        If igCmmtSrcReturn Then
            lbcSource.ListIndex = -1
            For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
                If lbcSource.ItemData(ilLoop) = igCmmtSrcCode Then
                    imDoubleClick = False
                    lbcSource.ListIndex = ilLoop
                    edcDropdown.Text = lbcSource.List(ilLoop)
                    mBranchSource = True
                    Exit Function
                End If
            Next ilLoop
            mBranchSource = False
        Else
            mBranchSource = False
        End If
    Else
        mBranchSource = True
    End If
    imDoubleClick = False
End Function

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
    slStr = edcDropdown.Text
    If gIsDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    'lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Move tmCDCtrls(ilWkDay).fBoxX - 30, tmCDCtrls(ilWkDay).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay).fBoxH + 15) - 30
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
