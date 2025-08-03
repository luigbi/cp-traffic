VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl AffContactGrid 
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
   Begin VB.ListBox lbcEMailRights 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffContactGrid.ctx":0000
      Left            =   6390
      List            =   "AffContactGrid.ctx":0010
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CheckBox ckcAffEMail 
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
      Left            =   6165
      TabIndex        =   9
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   1935
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CheckBox ckcAffLabel 
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
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   1455
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CheckBox ckcISCI2Contact 
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
      Left            =   5340
      TabIndex        =   8
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   1485
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ListBox lbcTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffContactGrid.ctx":004D
      Left            =   2550
      List            =   "AffContactGrid.ctx":004F
      Sorted          =   -1  'True
      TabIndex        =   3
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
      Picture         =   "AffContactGrid.ctx":0051
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
      TabIndex        =   12
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
      TabIndex        =   10
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
      Picture         =   "AffContactGrid.ctx":014B
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5640
      Width           =   75
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdContact 
      Height          =   1005
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "AffContactGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AffContactGrid.ctl on Wed 6/17/09 @ 12:56 P
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
' File Name: AffContactGrid.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

Private rst_Shtt As ADODB.Recordset
Private rst_artt As ADODB.Recordset
Private rst_vef As ADODB.Recordset
    
Event SetSave(ilStatus As Integer) 'VBC NR
Event ContactFocus()
Event PhoneChanged(slPhone As String)
Event FaxChanged(slFax As String)

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
Private imFromArrow As Integer

Dim smNowDate As String
Dim lmNowDate As Long
Dim lmFirstAllowedChgDate As Long

Dim imCtrlVisible As Integer
Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim lmTopRow As Long

Dim imShttCode As Integer
Dim bmStationIsVehicle As Boolean

Dim smSource As String  'S=Station; A=Agreement; M=Management
'ttp 5352
Dim myEmail As CEmail

Dim saf_rst As ADODB.Recordset
Private smEMailDistribution As String


'Personnel Contact Grid- grdContact
Const PCNAMEINDEX = 0
Const PCTITLEINDEX = 1
Const PCPHONEINDEX = 2
Const PCFAXINDEX = 3
Const PCEMAILINDEX = 4
Const PCEMAILRIGHTSINDEX = 5
Const PCAFFLABELINDEX = 6
Const PCISCIINDEX = 7
Const PCAFFEMAILINDEX = 8
Const PCDELETEINDEX = 9
Const PCARTTCODEINDEX = 10
Const PCSORTINDEX = 11
Const PCCHGDINDEX = 12


Private Sub ckcAffEMail_Click()
    Dim ilRow As Integer
    Dim ilCurRow As Integer
    
    ilCurRow = grdContact.Row
    If ckcAffEMail.Value = vbChecked Then
        grdContact.Col = PCAFFEMAILINDEX
        grdContact.CellFontName = "Monotype Sorts"
        grdContact.TextMatrix(ilCurRow, PCAFFEMAILINDEX) = "4"
        grdContact.TextMatrix(ilCurRow, PCCHGDINDEX) = 1
        '' If the check mark has just been checked, turn off all other check marks.
        '' Only one is allowed to be checked.
        'For ilRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
        '    If ilRow <> ilCurRow Then
        '        If grdContact.TextMatrix(ilRow, PCAFFEMAILINDEX) = "4" Then
        '            grdContact.TextMatrix(ilRow, PCCHGDINDEX) = 1
        '            grdContact.TextMatrix(ilRow, PCAFFEMAILINDEX) = " "
        '        End If
        '    End If
        'Next
    Else
        grdContact.TextMatrix(ilCurRow, PCAFFEMAILINDEX) = " "
        grdContact.TextMatrix(ilCurRow, PCCHGDINDEX) = 1
    End If

End Sub

Private Sub ckcAffLabel_Click()
    Dim ilRow As Integer
    Dim ilCurRow As Integer
    
    ilCurRow = grdContact.Row
    If ckcAffLabel.Value = vbChecked Then
        grdContact.Col = PCAFFLABELINDEX
        grdContact.CellFontName = "Monotype Sorts"
        grdContact.TextMatrix(ilCurRow, PCAFFLABELINDEX) = "4"
        grdContact.TextMatrix(ilCurRow, PCCHGDINDEX) = 1
        ' If the check mark has just been checked, turn off all other check marks.
        ' Only one is allowed to be checked.
        For ilRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
            If ilRow <> ilCurRow Then
                If grdContact.TextMatrix(ilRow, PCAFFLABELINDEX) = "4" Then
                    grdContact.TextMatrix(ilRow, PCCHGDINDEX) = 1
                    grdContact.TextMatrix(ilRow, PCAFFLABELINDEX) = " "
                End If
            End If
        Next
    Else
        grdContact.TextMatrix(ilCurRow, PCAFFLABELINDEX) = " "
        grdContact.TextMatrix(ilCurRow, PCCHGDINDEX) = 1
    End If

End Sub


Private Sub ckcISCI2Contact_Click()
    Dim ilCurRow As Integer
    
    ilCurRow = grdContact.Row
    If ckcISCI2Contact.Value = vbChecked Then
        grdContact.Col = PCISCIINDEX
        grdContact.CellFontName = "Monotype Sorts"
        grdContact.TextMatrix(ilCurRow, PCISCIINDEX) = "4"
        grdContact.TextMatrix(ilCurRow, PCCHGDINDEX) = 1
    Else
        grdContact.TextMatrix(ilCurRow, PCISCIINDEX) = " "
        grdContact.TextMatrix(ilCurRow, PCCHGDINDEX) = 1
    End If
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdContact.Col
        Case PCNAMEINDEX
        Case PCTITLEINDEX
            lbcTitle.Visible = Not lbcTitle.Visible
        Case PCPHONEINDEX
        Case PCFAXINDEX
        Case PCEMAILINDEX
        Case PCEMAILRIGHTSINDEX
            lbcEMailRights.Visible = Not lbcEMailRights.Visible
        Case PCAFFLABELINDEX
        Case PCISCIINDEX
        Case PCAFFEMAILINDEX
    End Select

End Sub

Private Sub edcDropdown_Change()
    Select Case lmEnableCol
        Case PCNAMEINDEX
        Case PCTITLEINDEX
            mDropdownChangeEvent lbcTitle
        Case PCPHONEINDEX
        Case PCFAXINDEX
        Case PCEMAILINDEX
        Case PCEMAILRIGHTSINDEX
            mDropdownChangeEvent lbcEMailRights
        Case PCAFFLABELINDEX
        Case PCISCIINDEX
        Case PCAFFEMAILINDEX
    End Select
End Sub


Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
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
    mInit
End Sub


Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmEnableCol
            Case PCNAMEINDEX
            Case PCTITLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcTitle, True
            Case PCPHONEINDEX
            Case PCFAXINDEX
            Case PCEMAILINDEX
            Case PCEMAILRIGHTSINDEX
                gProcessArrowKey Shift, KeyCode, lbcEMailRights, True
            Case PCAFFLABELINDEX
            Case PCISCIINDEX
            Case PCAFFEMAILINDEX
        End Select
    End If
End Sub

Private Sub grdContact_EnterCell()
    Dim ilRet As Integer
    
    If lmEnableRow <> grdContact.MouseRow Then
        mSetShow
        '9/5/11: Moved to mSetShow and test if source is Management
        ''Called here so that Affiliate Management changes are saved without the user pressing a Save button
        'ilRet = mSaveRec()
    Else
        mSetShow
    End If
End Sub

Private Sub grdContact_GotFocus()
    RaiseEvent ContactFocus
End Sub

Private Sub grdContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    'grdContact.ToolTipText = ""
    slStr = ""
    If (grdContact.MouseRow >= grdContact.FixedRows) And (grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol)) <> "" Then
        If grdContact.MouseCol = PCAFFLABELINDEX Then
            If grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol) = "4" Then
                'grdContact.ToolTipText = "Checked"
                slStr = "Checked"
            End If
        ElseIf grdContact.MouseCol = PCISCIINDEX Then
            If grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol) = "4" Then
                'grdContact.ToolTipText = "Checked"
                slStr = "Checked"
            End If
        ElseIf grdContact.MouseCol = PCAFFEMAILINDEX Then
            If grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol) = "4" Then
                'grdContact.ToolTipText = "Checked"
                slStr = "Checked"
            End If
        Else
            'grdContact.ToolTipText = grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol)
            slStr = grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol)
        End If
    End If
    grdContact.ToolTipText = slStr
End Sub

Private Sub grdContact_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
'    If y < grdContact.RowHeight(0) Then
'        mSortCol grdContact.Col
'        Exit Sub
'    End If
    'Determine row and col mouse up onto
    On Error GoTo grdContactErr
    pbcArrow.Visible = False
    ilCol = grdContact.MouseCol
    ilRow = grdContact.MouseRow
    If ilCol < grdContact.FixedCols Then
        grdContact.Redraw = True
        Exit Sub
    End If
    If ilRow < grdContact.FixedRows Then
        grdContact.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdContact.TopRow
    DoEvents
    If ilCol = PCDELETEINDEX Then
        llCode = Val(grdContact.TextMatrix(ilRow, PCARTTCODEINDEX))
        If llCode > 0 Then
            If grdContact.TextMatrix(ilRow, PCNAMEINDEX) = "" Then
                ilRet = MsgBox("This will permanently remove " & grdContact.TextMatrix(ilRow, PCEMAILINDEX) & " from the contact list, are you sure", vbYesNo + vbQuestion, "Remove")
            Else
                ilRet = MsgBox("This will permanently remove " & grdContact.TextMatrix(ilRow, PCNAMEINDEX) & " from the contact list, are you sure", vbYesNo + vbQuestion, "Remove")
            End If
            If ilRet = vbYes Then
                'update the artt before deleting the record so we can make the call to delete it from the web
                SQLQuery = "UPDATE artt"
                SQLQuery = SQLQuery & " SET arttEmailToWeb = " & "'" & "D" & "',"
                SQLQuery = SQLQuery & " arttWebEmail = " & "'" & "N" & "'"
                SQLQuery = SQLQuery & " WHERE arttCode = " & llCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "AffContactGrid-grdContact_MouseUp"
                    Exit Sub
                End If
                
                ilRet = gWebTestEmailChange
                
                SQLQuery = "DELETE FROM artt WHERE arttCode = " & llCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "AffContactGrid-grdContact_MouseUp"
                    Exit Sub
                End If
                grdContact.RemoveItem ilRow
            End If
        Else
            grdContact.RemoveItem ilRow
        End If
        grdContact.Redraw = True
        Exit Sub
    End If
    If (grdContact.TextMatrix(ilRow, PCNAMEINDEX) = "") And (grdContact.TextMatrix(ilRow, PCEMAILINDEX) = "") Then
        grdContact.Redraw = False
        Do
            ilRow = ilRow - 1
            If ilRow < grdContact.FixedRows Then
                Exit Do
            End If
        Loop While (Trim(grdContact.TextMatrix(ilRow, PCNAMEINDEX)) = "") And (Trim(grdContact.TextMatrix(ilRow, PCEMAILINDEX)) = "")
        ilRow = ilRow + 1
        ilCol = PCNAMEINDEX
    End If
    grdContact.Col = ilCol
    grdContact.Row = ilRow
    If Not mColOk() Then
        grdContact.Redraw = True
        Exit Sub
    End If
    grdContact.Redraw = True
    mEnableBox
    On Error GoTo 0
    Exit Sub
grdContactErr:
    On Error GoTo 0
    If (lmEnableRow >= grdContact.FixedRows) And (lmEnableRow < grdContact.Rows) Then
        grdContact.Row = lmEnableRow
        grdContact.Col = lmEnableCol
        mSetFocus
    End If
    grdContact.Redraw = False
    grdContact.Redraw = True
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffContactGrid-grdContact"
End Sub

Private Sub grdContact_Scroll()
    mSetShow
    pbcArrow.Visible = False
    If grdContact.RowIsVisible(grdContact.Row) Then
        pbcArrow.Move grdContact.Left - pbcArrow.Width, grdContact.Top + grdContact.RowPos(grdContact.Row) + (grdContact.RowHeight(grdContact.Row) - pbcArrow.Height) / 2
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

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdContact, grdContact, vbHourglass
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
    imCtrlVisible = False
    imLastColSorted = -1
    imLastSort = -1
    imFromArrow = False
    lmEnableRow = -1
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate + 1
    If gGetEMailDistribution Then
        smEMailDistribution = "Y"
    Else
        smEMailDistribution = "N"
    End If
    mInitBox
    mPopTitles
    'ttp 5352
    Set myEmail = New CEmail
    Screen.MousePointer = vbDefault
    gSetMousePointer grdContact, grdContact, vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdContact, grdContact, vbDefault
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

    'grdContact.Move 180, 120, Width - pbcArrow.Width - 120
    'grdContact.Height = Height - grdContact.Top - 120
    'grdContact.Redraw = False
    pbcSTab.Move -100, -100
    pbcTab.Move -100, -100
    pbcClickFocus.Move -100, -100
    mSetGridColumns
    mSetGridTitles
    mClearGrid grdContact
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
    gSetMousePointer grdContact, grdContact, vbDefault
End Sub

Private Sub lbcEMailRights_Click()
    edcDropdown.Text = lbcEMailRights.List(lbcEMailRights.ListIndex)
End Sub

Private Sub lbcTitle_Click()
    edcDropdown.Text = lbcTitle.List(lbcTitle.ListIndex)
End Sub

Private Sub lbcTitle_DblClick()
    Dim ilIndex As Integer
    
    sgTitle = lbcTitle.List(lbcTitle.ListIndex)
    frmTitle.Show vbModal
    mPopTitles
    If bgFrmTitleCanceled Then
        edcDropdown.SetFocus
    Else
        ilIndex = SendMessageByString(lbcTitle.hwnd, LB_FINDSTRING, -1, sgTitle)
        If ilIndex >= 0 Then
            edcDropdown.Text = sgTitle
            pbcTab.SetFocus
        Else
            edcDropdown.SetFocus
        End If
    End If
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
        grdContact.Col = PCNAMEINDEX
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        If lmEnableCol = PCTITLEINDEX Then
            If edcDropdown.Text = "[New]" Then
                sgTitle = ""
                frmTitle.Show vbModal
                If bgFrmTitleCanceled Then
                    edcDropdown.SetFocus
                    Exit Sub
                Else
                    mPopTitles
                    ilIndex = SendMessageByString(lbcTitle.hwnd, LB_FINDSTRING, -1, sgTitle)
                    If ilIndex >= 0 Then
                        edcDropdown.Text = sgTitle
                    Else
                        edcDropdown.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        Do
            ilNext = False
            Select Case grdContact.Col
                Case PCNAMEINDEX
                    If grdContact.Row = grdContact.FixedRows Then
                        mSetShow
                        pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdContact.Row = grdContact.Row - 1
                    If Not grdContact.RowIsVisible(grdContact.Row) Then
                        grdContact.TopRow = grdContact.TopRow - 1
                    End If
                    'grdContact.Col = PCISCIINDEX
                    grdContact.Col = PCAFFEMAILINDEX
                'ttp5352
                Case PCEMAILINDEX
                    'grdcontact.TextMatrix(grdcontact.Row,PCEMAILINDEX) shows previous value when going backwards.
                    If Len(edcDropdown.Text) > 0 Then
                        If Not myEmail.TestAddress(edcDropdown.Text) Then
                            MsgBox "The email is not valid.", vbInformation, "Invalid Email"
                        Else
                            grdContact.Col = grdContact.Col - 1
                        End If
                    Else
                        grdContact.Col = grdContact.Col - 1
                    End If
                Case Else
                    grdContact.Col = grdContact.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdContact.Row = grdContact.FixedRows
        grdContact.Col = grdContact.FixedCols
        Do
            If mColOk() Then
                Exit Do
            Else
                grdContact.Col = grdContact.Col + 1
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

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        If lmEnableCol = PCTITLEINDEX Then
            If edcDropdown.Text = "[New]" Then
                sgTitle = ""
                frmTitle.Show vbModal
                If bgFrmTitleCanceled Then
                    edcDropdown.SetFocus
                    Exit Sub
                Else
                    mPopTitles
                    ilIndex = SendMessageByString(lbcTitle.hwnd, LB_FINDSTRING, -1, sgTitle)
                    If ilIndex >= 0 Then
                        edcDropdown.Text = sgTitle
                    Else
                        edcDropdown.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        llEnableRow = lmEnableRow
        llEnableCol = lmEnableCol
        mSetShow
        grdContact.Row = llEnableRow
        grdContact.Col = llEnableCol
        Do
            ilNext = False
            Select Case grdContact.Col
                'Case PCISCIINDEX
                Case PCAFFEMAILINDEX
                    If (grdContact.Row + 1 >= grdContact.Rows) Then
                        grdContact.Rows = grdContact.Rows + 1
                        grdContact.Row = grdContact.Row + 1
                        grdContact.TextMatrix(grdContact.Row, PCARTTCODEINDEX) = 0
                        grdContact.TextMatrix(grdContact.Row, PCCHGDINDEX) = "0"
                        If Not grdContact.RowIsVisible(grdContact.Row) Then
                            grdContact.TopRow = grdContact.TopRow + 1
                        End If
                        imFromArrow = True
                        pbcArrow.Move grdContact.Left - pbcArrow.Width, grdContact.Top + grdContact.RowPos(grdContact.Row) + (grdContact.RowHeight(grdContact.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                        Exit Sub
                    End If
                    If (grdContact.Row + 1 < grdContact.Rows) Then
                        If (Trim$(grdContact.TextMatrix(grdContact.Row + 1, PCNAMEINDEX)) = "") And (Trim$(grdContact.TextMatrix(grdContact.Row + 1, PCEMAILINDEX)) = "") Then
                            grdContact.Row = grdContact.Row + 1
                            grdContact.Col = PCNAMEINDEX
                            If Not grdContact.RowIsVisible(grdContact.Row) Then
                                grdContact.TopRow = grdContact.TopRow + 1
                            End If
                            imFromArrow = True
                            pbcArrow.Move grdContact.Left - pbcArrow.Width, grdContact.Top + grdContact.RowPos(grdContact.Row) + (grdContact.RowHeight(grdContact.Row) - pbcArrow.Height) / 2
                            pbcArrow.Visible = True
                            pbcArrow.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdContact.Row = grdContact.Row + 1
                    grdContact.Col = PCNAMEINDEX
                    If Not grdContact.RowIsVisible(grdContact.Row) Then
                        grdContact.TopRow = grdContact.TopRow + 1
                    End If
                'ttp5352
                Case PCEMAILINDEX
                    If Len(grdContact.TextMatrix(grdContact.Row, PCEMAILINDEX)) > 0 Then
                        If Not myEmail.TestAddress(Trim(grdContact.TextMatrix(grdContact.Row, PCEMAILINDEX))) Then
                            MsgBox "The email is not valid.", vbInformation, "Invalid Email"
                        Else
                            grdContact.Col = grdContact.Col + 1
                        End If
                    Else
                        grdContact.Col = grdContact.Col + 1
                    End If
                Case Else
                    grdContact.Col = grdContact.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        grdContact.Row = grdContact.FixedRows
        grdContact.Col = grdContact.FixedCols
        Do
            If mColOk() Then
                Exit Do
            Else
                grdContact.Col = grdContact.Col + 1
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

    If (grdContact.Row < grdContact.FixedRows) Or (grdContact.Row >= grdContact.Rows) Or (grdContact.Col < grdContact.FixedCols) Or (grdContact.Col >= grdContact.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdContact.Row
    lmEnableCol = grdContact.Col
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdContact.Left - pbcArrow.Width, grdContact.Top + grdContact.RowPos(grdContact.Row) + (grdContact.RowHeight(grdContact.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True

    Select Case grdContact.Col
        Case PCNAMEINDEX
            mSetEdcGridControl
        Case PCTITLEINDEX
            mSetLbcGridControl lbcTitle
        Case PCPHONEINDEX
            mSetEdcGridControl
        Case PCFAXINDEX
            mSetEdcGridControl
        Case PCEMAILINDEX
            mSetEdcGridControl
        Case PCEMAILRIGHTSINDEX
            mSetLbcGridControl lbcEMailRights
        Case PCAFFLABELINDEX
            ckcAffLabel.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30
            grdContact.Col = PCAFFLABELINDEX
            grdContact.CellFontName = "Monotype Sorts"
            'If ckcAffLabel.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
                ckcAffLabel.FontName = "Arial"
                ckcAffLabel.Height = grdContact.RowHeight(grdContact.Row) - 15
            'End If
            If grdContact.TextMatrix(grdContact.Row, PCAFFLABELINDEX) = "4" Then
                ckcAffLabel.Value = vbChecked
            Else
                ckcAffLabel.Value = vbUnchecked
            End If
            
            ckcAffLabel.Visible = True
            ckcAffLabel.SetFocus
        Case PCISCIINDEX
            ckcISCI2Contact.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30
            grdContact.Col = PCISCIINDEX
            grdContact.CellFontName = "Monotype Sorts"
            'If ckcISCI2Contact.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
                ckcISCI2Contact.FontName = "Arial"
                ckcISCI2Contact.Height = grdContact.RowHeight(grdContact.Row) - 15
            'End If
            If grdContact.TextMatrix(grdContact.Row, PCISCIINDEX) = "4" Then
                ckcISCI2Contact.Value = vbChecked
            Else
                ckcISCI2Contact.Value = vbUnchecked
            End If
            
            ckcISCI2Contact.Visible = True
            ckcISCI2Contact.SetFocus
        Case PCAFFEMAILINDEX
            ckcAffEMail.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30
            grdContact.Col = PCAFFEMAILINDEX
            grdContact.CellFontName = "Monotype Sorts"
            'If ckcAffEMail.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
                ckcAffEMail.FontName = "Arial"
                ckcAffEMail.Height = grdContact.RowHeight(grdContact.Row) - 15
            'End If
            If grdContact.TextMatrix(grdContact.Row, PCAFFEMAILINDEX) = "4" Then
                ckcAffEMail.Value = vbChecked
            Else
                ckcAffEMail.Value = vbUnchecked
            End If
            
            ckcAffEMail.Visible = True
            ckcAffEMail.SetFocus
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
    Dim ilRet As Integer
    Dim llCode As Long
    Dim llSvCol As Long
    Dim llSvRow As Long
    Dim ilVef As Integer
    Dim slCallLetters As String
    
    On Error GoTo ErrHand
    llSvCol = grdContact.Col
    llSvRow = grdContact.Row
    If (lmEnableRow >= grdContact.FixedRows) And (lmEnableRow < grdContact.Rows) Then
        Select Case lmEnableCol
            Case PCNAMEINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdContact.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdContact.TextMatrix(lmEnableRow, PCCHGDINDEX) = "1"
                End If
                grdContact.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PCTITLEINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdContact.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdContact.TextMatrix(lmEnableRow, PCCHGDINDEX) = "1"
                End If
                grdContact.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PCPHONEINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdContact.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdContact.TextMatrix(lmEnableRow, PCCHGDINDEX) = "1"
                    If lmEnableRow = grdContact.FixedRows Then
                        RaiseEvent PhoneChanged(slStr)
                    End If
                End If
                grdContact.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PCFAXINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdContact.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdContact.TextMatrix(lmEnableRow, PCCHGDINDEX) = "1"
                    RaiseEvent FaxChanged(slStr)
                End If
                grdContact.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PCEMAILINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdContact.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdContact.TextMatrix(lmEnableRow, PCCHGDINDEX) = "1"
                End If
                grdContact.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                grdContact.Col = PCEMAILRIGHTSINDEX
                grdContact.Row = lmEnableRow
                If imShttCode <= 0 Then
                    bmStationIsVehicle = False
                    slCallLetters = Trim$(grdContact.TextMatrix(grdContact.FixedRows, PCNAMEINDEX))
                    For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If Trim$(UCase(tgVehicleInfo(ilVef).sVehicle)) = slCallLetters Then
                            bmStationIsVehicle = True
                        End If
                    Next ilVef
                End If
                If (Trim$(slStr) = "") Or (Not bmStationIsVehicle) Then
                    grdContact.CellBackColor = LIGHTYELLOW
                    grdContact.TextMatrix(lmEnableRow, PCEMAILRIGHTSINDEX) = ""
                Else
                    grdContact.CellBackColor = vbWhite
                End If
            Case PCEMAILRIGHTSINDEX
                slStr = edcDropdown.Text
                If StrComp(UCase$(Trim$(grdContact.TextMatrix(lmEnableRow, lmEnableCol))), UCase$(Trim$(slStr)), vbBinaryCompare) <> 0 Then
                    grdContact.TextMatrix(lmEnableRow, PCCHGDINDEX) = "1"
                End If
                grdContact.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PCAFFLABELINDEX
            Case PCISCIINDEX
            Case PCAFFEMAILINDEX
        End Select
    End If
    grdContact.Col = llSvCol
    grdContact.Row = llSvRow
    If smSource = "M" Then
        ilRet = mSaveRec()
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    pbcArrow.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcTitle.Visible = False
    lbcEMailRights.Visible = False
    ckcAffLabel.Visible = False
    ckcISCI2Contact.Visible = False
    ckcAffEMail.Visible = False
    mSetCommands
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffContactGrid-mSetShow"
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

    If (grdContact.Row < grdContact.FixedRows) Or (grdContact.Row >= grdContact.Rows) Or (grdContact.Col < grdContact.FixedCols) Or (grdContact.Col >= grdContact.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdContact.Col - 1 Step 1
        llColPos = llColPos + grdContact.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdContact.ColWidth(grdContact.Col)
    ilCol = grdContact.Col
    Do While ilCol < grdContact.Cols - 1
        If (Trim$(grdContact.TextMatrix(grdContact.Row - 1, grdContact.Col)) <> "") And (Trim$(grdContact.TextMatrix(grdContact.Row - 1, grdContact.Col)) = Trim$(grdContact.TextMatrix(grdContact.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdContact.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdContact.Col
        Case PCNAMEINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case PCTITLEINDEX
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcTitle.Visible = True
            edcDropdown.SetFocus
        Case PCPHONEINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case PCFAXINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case PCEMAILINDEX
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case PCEMAILRIGHTSINDEX
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcEMailRights.Visible = True
            edcDropdown.SetFocus
        Case PCAFFLABELINDEX
            ckcAffLabel.Visible = True
            ckcAffLabel.SetFocus
        Case PCISCIINDEX
            ckcISCI2Contact.Visible = True
            ckcISCI2Contact.SetFocus
        Case PCAFFEMAILINDEX
            ckcAffEMail.Visible = True
            ckcAffEMail.SetFocus
    End Select
End Sub






Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    'gGetEMailDistribution
    grdContact.Width = Width - pbcArrow.Width  'grdStations.Width
    grdContact.Height = Height
    gGrid_IntegralHeight grdContact
    grdContact.Height = grdContact.Height + 30
    'grdContact.Move grdStations.Left, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1)
    grdContact.Move pbcArrow.Width, 0
    grdContact.ColWidth(PCCHGDINDEX) = 0
    grdContact.ColWidth(PCSORTINDEX) = 0
    grdContact.ColWidth(PCARTTCODEINDEX) = 0
    grdContact.ColWidth(PCNAMEINDEX) = grdContact.Width * 0.14
    grdContact.ColWidth(PCTITLEINDEX) = grdContact.Width * 0.08
    grdContact.ColWidth(PCPHONEINDEX) = grdContact.Width * 0.12
    grdContact.ColWidth(PCFAXINDEX) = grdContact.Width * 0.1
    grdContact.ColWidth(PCAFFLABELINDEX) = grdContact.Width * 0.08
    grdContact.ColWidth(PCISCIINDEX) = grdContact.Width * 0.08
    grdContact.ColWidth(PCAFFEMAILINDEX) = grdContact.Width * 0.08
    grdContact.ColWidth(PCDELETEINDEX) = grdContact.Width * 0.05
    If smEMailDistribution = "Y" Then
        grdContact.ColWidth(PCEMAILRIGHTSINDEX) = grdContact.Width * 0.1
    Else
        grdContact.ColWidth(PCEMAILRIGHTSINDEX) = 0
    End If
           
    grdContact.ColWidth(PCEMAILINDEX) = grdContact.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To PCDELETEINDEX Step 1
        If ilCol <> PCEMAILINDEX Then
            grdContact.ColWidth(PCEMAILINDEX) = grdContact.ColWidth(PCEMAILINDEX) - grdContact.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdContact

End Sub

Private Sub mSetGridTitles()
    grdContact.TextMatrix(0, PCNAMEINDEX) = "Name"
    grdContact.TextMatrix(0, PCTITLEINDEX) = "Title"
    grdContact.TextMatrix(0, PCPHONEINDEX) = "Direct #"
    grdContact.TextMatrix(0, PCFAXINDEX) = "Fax"
    grdContact.TextMatrix(0, PCEMAILINDEX) = "E-Mail"
    grdContact.TextMatrix(0, PCEMAILRIGHTSINDEX) = "E-Mail Rights"
    grdContact.TextMatrix(0, PCAFFLABELINDEX) = "Aff-Label"
    grdContact.TextMatrix(0, PCISCIINDEX) = "ISCI Export"
    grdContact.TextMatrix(0, PCAFFEMAILINDEX) = "Aff-Email"
    grdContact.TextMatrix(0, PCDELETEINDEX) = "Delete"
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
    If grdContact.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    If grdContact.CellForeColor = vbRed Then
        mColOk = False
        Exit Function
    End If
    If grdContact.Col = PCAFFLABELINDEX Then
        If Trim$(grdContact.TextMatrix(grdContact.Row, PCNAMEINDEX)) = "" Then
            mColOk = False
            Exit Function
        End If
    End If
    If grdContact.Col = PCISCIINDEX Then
        If Trim$(grdContact.TextMatrix(grdContact.Row, PCEMAILINDEX)) = "" Then
            mColOk = False
            Exit Function
        End If
    End If
    If grdContact.Col = PCAFFEMAILINDEX Then
        If Trim$(grdContact.TextMatrix(grdContact.Row, PCEMAILINDEX)) = "" Then
            mColOk = False
            Exit Function
        End If
    End If
    If grdContact.Col = PCEMAILRIGHTSINDEX Then
        If Trim$(grdContact.TextMatrix(grdContact.Row, PCEMAILINDEX)) = "" Then
            mColOk = False
            Exit Function
        End If
        If smEMailDistribution <> "Y" Then
            mColOk = False
            Exit Function
        End If
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
            'mInit
        Case 3  'Populate
            '6/11/15: Moved mSetGridColumn call here instead of UserControl_ReSize
            'because that event is executed during the compile and will blow-up with the Pervasive call.
            mSetGridColumns
            mPopContactGrid
        Case 4  'Clear
            mSetShow
            pbcArrow.Visible = False
            mClearGrid grdContact
            Screen.MousePointer = vbDefault
            gSetMousePointer grdContact, grdContact, vbDefault
        Case 5  'Save
            mSetShow
            pbcArrow.Visible = False
            ilRet = mSaveRec()
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
    RaiseEvent ContactFocus
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
    Dim slName As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim slEMail As String
    Dim slAffLabel As String
    Dim slISCI2Contact As String
    Dim slAffEmail As String
    Dim slChgd As String
    Dim llCode As Long
    Dim ilTnt As Integer
    Dim iltntCode As Integer
    Dim ilRet As Integer
    Dim ilWebRefID As Integer
    Dim blUpdateWeb As Boolean
    Dim slStr As String
    Dim slRights As String
    '11/26/16
    Dim slCallLetters As String
    Dim ilIndex As Integer
    Dim blRepopRequired As Boolean

    On Error GoTo ErrHand
    If imShttCode <= 0 Then
        mSaveRec = True
        Exit Function
    End If
    blUpdateWeb = False
    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
        slName = Trim(grdContact.TextMatrix(llRow, PCNAMEINDEX))
        slEMail = Trim(grdContact.TextMatrix(llRow, PCEMAILINDEX))
        If (slName <> "") Or (slEMail <> "") Then
            slFirstName = ""
            slLastName = ""
            ilLen = Len(slName)
            If ilLen > 0 Then
                ilPos = InStrRev(slName, " ")
                If ilPos > 0 Then
                    'slFirstName = gFixQuote(Left(slName, ilPos - 1))
                    slFirstName = Left(slName, ilPos - 1)
                    'slLastName = gFixQuote(Trim(right(slName, ilLen - ilPos)))
                    slLastName = Trim(right(slName, ilLen - ilPos))
                Else
                    slLastName = gFixQuote(Trim(slName))
                End If
            End If
            slStr = UCase(grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX))
            If InStr(1, slStr, "PRIMARY", vbTextCompare) > 0 Then
                slRights = "M"
            ElseIf InStr(1, slStr, "BACKUP", vbTextCompare) > 0 Then
                slRights = "A"
            ElseIf InStr(1, slStr, "VIEW", vbTextCompare) > 0 Then
                slRights = "V"
            Else
                slRights = "N"
            End If
            slAffLabel = " "
            If grdContact.TextMatrix(llRow, PCAFFLABELINDEX) = "4" Then
                slAffLabel = "1"
            End If

            slISCI2Contact = " "
            If grdContact.TextMatrix(llRow, PCISCIINDEX) = "4" Then
                slISCI2Contact = "1"
            End If
            
            slAffEmail = "N"  'default, don't send email to the web
            If grdContact.TextMatrix(llRow, PCAFFEMAILINDEX) = "4" Then
                'send email to the web
                slAffEmail = "Y"
            End If

            ilTnt = SendMessageByString(lbcTitle.hwnd, LB_FINDSTRING, -1, grdContact.TextMatrix(llRow, PCTITLEINDEX))
            If ilTnt >= 0 Then
                iltntCode = Val(lbcTitle.ItemData(ilTnt))
            Else
                iltntCode = 0
            End If
            llCode = Val(grdContact.TextMatrix(llRow, PCARTTCODEINDEX))
            slChgd = grdContact.TextMatrix(llRow, PCCHGDINDEX)
            If slChgd = "1" Then
                blUpdateWeb = True
                If llRow = grdContact.FixedRows Then
                    SQLQuery = "Update shtt Set "
                    SQLQuery = SQLQuery & "shttPhone = '" & Trim(grdContact.TextMatrix(llRow, PCPHONEINDEX)) & "', "
                    '7/25/11: Disallow Station E-Mail
                    'SQLQuery = SQLQuery & "shttFax = '" & Trim(grdContact.TextMatrix(llRow, PCFAXINDEX)) & "', "
                    'SQLQuery = SQLQuery & "shttEMail = '" & Trim(grdContact.TextMatrix(llRow, PCEMAILINDEX)) & "' "
                    SQLQuery = SQLQuery & "shttFax = '" & Trim(grdContact.TextMatrix(llRow, PCFAXINDEX)) & "' "
                    SQLQuery = SQLQuery & " Where shttCode = " & imShttCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "AffContactGrid-mSaveRec"
                        mSaveRec = False
                        Exit Function
                    End If
                    '11/26/17
                    blRepopRequired = False
                    ilIndex = gBinarySearchStationInfoByCode(imShttCode)
                    If ilIndex <> -1 Then
                        tgStationInfoByCode(ilIndex).sPhone = Trim(grdContact.TextMatrix(llRow, PCPHONEINDEX))
                        tgStationInfoByCode(ilIndex).sFax = Trim(grdContact.TextMatrix(llRow, PCFAXINDEX))
                        slCallLetters = Trim$(tgStationInfoByCode(ilIndex).sCallLetters)
                        ilIndex = gBinarySearchStation(slCallLetters)
                        If ilIndex <> -1 Then
                            tgStationInfo(ilIndex).sPhone = Trim(grdContact.TextMatrix(llRow, PCPHONEINDEX))
                            tgStationInfo(ilIndex).sFax = Trim(grdContact.TextMatrix(llRow, PCFAXINDEX))
                        Else
                            blRepopRequired = True
                        End If
                    Else
                        blRepopRequired = True
                    End If
                    gFileChgdUpdate "shtt.mkd", blRepopRequired
                Else
                    If llCode > 0 Then
                        'Test if WebEMailRefID needs to be set
                        SQLQuery = "SELECT arttWebEMailRefID FROM artt WHERE arttCode = " & llCode
                        Set rst_artt = gSQLSelectCall(SQLQuery)
                        ilWebRefID = rst_artt!arttWebEMailRefID
                        If (slAffEmail = "Y") And (ilWebRefID = 0) Then
                            SQLQuery = "SELECT MAX(arttWebEMailRefID) FROM artt WHERE arttShttCode = " & imShttCode
                            Set rst_artt = gSQLSelectCall(SQLQuery)
                            If IsNull(rst_artt(0).Value) Then
                                ilWebRefID = 1
                            Else
                                If Not rst_artt.EOF Then
                                    ilWebRefID = rst_artt(0).Value + 1
                                Else
                                    ilWebRefID = 1
                                End If
                            End If
                        End If
                        SQLQuery = "Update artt Set "
                        SQLQuery = SQLQuery & "arttFirstName = '" & gFixQuote(Trim(slFirstName)) & "',"
                        SQLQuery = SQLQuery & " arttLastName = '" & gFixQuote(Trim(slLastName)) & "',"
                        SQLQuery = SQLQuery & " arttPhone = '" & Trim(grdContact.TextMatrix(llRow, PCPHONEINDEX)) & "',"
                        SQLQuery = SQLQuery & " arttFax = '" & Trim(grdContact.TextMatrix(llRow, PCFAXINDEX)) & "',"
                        SQLQuery = SQLQuery & " arttEmail = '" & gFixQuote(Trim(slEMail)) & "',"
                        SQLQuery = SQLQuery & " arttEmailRights = '" & gFixQuote(Trim(slRights)) & "',"
                        If slAffEmail = "Y" Then
                            SQLQuery = SQLQuery & " arttEMailToWeb = '" & "U" & "',"
                        Else
                            SQLQuery = SQLQuery & " arttEMailToWeb = '" & "D" & "',"
                        End If
                        SQLQuery = SQLQuery & " arttTntCode = " & iltntCode & ", "
                        'SQLQuery = SQLQuery & " arttAffContact = '" & slAffLabel & "' "
                        SQLQuery = SQLQuery & " arttAffContact = '" & slAffLabel & "' " & ","
                        SQLQuery = SQLQuery & " arttISCI2Contact = '" & slISCI2Contact & "', "
                        If slAffEmail = "Y" Then
                            SQLQuery = SQLQuery & " arttWebEMail = '" & "Y" & "', "
                        Else
                            SQLQuery = SQLQuery & " arttWebEMail = '" & "N" & "', "
                        End If
                        SQLQuery = SQLQuery & " arttWebEMailRefID = " & ilWebRefID
                        SQLQuery = SQLQuery & " Where arttCode = " & llCode
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/13/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "AffContactGrid-mSaveRec"
                            mSaveRec = False
                            Exit Function
                        End If
                        
                    Else
                        SQLQuery = "SELECT MAX(arttWebEMailRefID) from artt WHERE arttShttCode = " & imShttCode
                        Set rst_artt = gSQLSelectCall(SQLQuery)
                        If IsNull(rst_artt(0).Value) Then
                            ilWebRefID = 1
                        Else
                            If Not rst_artt.EOF Then
                                ilWebRefID = rst_artt(0).Value + 1
                            Else
                                ilWebRefID = 1
                            End If
                        End If
                        Do
                            SQLQuery = "SELECT MAX(arttCode) FROM artt"
                            Set rst_artt = gSQLSelectCall(SQLQuery)
                            If IsNull(rst_artt(0).Value) Then
                                llCode = 1
                            Else
                                If Not rst_artt.EOF Then
                                    llCode = rst_artt(0).Value + 1
                                Else
                                    llCode = 1
                                End If
                            End If
                            ilRet = 0
                            'SQLQuery = "INSERT INTO artt (arttFirstName, arttLastName, arttPhone, arttFax, arttEmail, arttState, arttUsfCode, arttAddress1, arttAddress2, arttCity, arttAddressState, arttZip, ArttCountry, arttType, arttTntCode, arttAffContact, arttShttCode)"
                            SQLQuery = "INSERT INTO artt (arttCode, arttFirstName, arttLastName, arttPhone, arttFax, arttEmail, arttEMailRights, arttType, arttTntCode, arttAffContact, arttISCI2Contact, arttWebEMail, arttEMailToWeb, arttWebEMailRefID, arttShttCode)"
                            SQLQuery = SQLQuery & " VALUES ( "
                            SQLQuery = SQLQuery & llCode & ", "
                            SQLQuery = SQLQuery & " '" & gFixQuote(Trim(slFirstName)) & "', "
                            SQLQuery = SQLQuery & " '" & gFixQuote(Trim(slLastName)) & "', "
                            SQLQuery = SQLQuery & " '" & Trim(grdContact.TextMatrix(llRow, PCPHONEINDEX)) & "', "
                            SQLQuery = SQLQuery & " '" & Trim(grdContact.TextMatrix(llRow, PCFAXINDEX)) & "', "
                            SQLQuery = SQLQuery & " '" & gFixQuote(Trim(slEMail)) & "', "
                            SQLQuery = SQLQuery & " '" & gFixQuote(Trim(slRights)) & "', "
        '                    SQLQuery = SQLQuery & " '' , "  ' State
        '                    SQLQuery = SQLQuery & " '' , "  ' UsfCode
        '                    SQLQuery = SQLQuery & " '' , "  ' Address1
        '                    SQLQuery = SQLQuery & " '' , "  ' Address2
        '                    SQLQuery = SQLQuery & " '' , "  ' City
        '                    SQLQuery = SQLQuery & " '' , "  ' AddressState
        '                    SQLQuery = SQLQuery & " '' , "  ' Zip
        '                    SQLQuery = SQLQuery & " '' , "  ' Country
                            SQLQuery = SQLQuery & " 'P' , "  ' Type
                            SQLQuery = SQLQuery & iltntCode & ", "
                            SQLQuery = SQLQuery & "'" & slAffLabel & "', "
                            SQLQuery = SQLQuery & "'" & slISCI2Contact & "', "
                            SQLQuery = SQLQuery & "'" & slAffEmail & "', "
                            SQLQuery = SQLQuery & "'" & "I" & "', "
                            SQLQuery = SQLQuery & ilWebRefID & ", "
                            SQLQuery = SQLQuery & imShttCode & ")"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand1:
                                Screen.MousePointer = vbDefault
                                If Not gHandleError4994("AffErrorLog.txt", "AffContactGrid-mSaveRec") Then
                                    mSaveRec = False
                                    Exit Function
                                End If
                                ilRet = 1
                            End If
                        Loop While ilRet <> 0
                        grdContact.TextMatrix(llRow, PCARTTCODEINDEX) = llCode
                    End If
                    '07-13-15 Add update to pesonnel EDS
                    'If gGetEMailDistribution Then
                    If smEMailDistribution = "Y" Then
                       ilRet = mAddOrUpdateSingleStationUser(llCode, imShttCode)
                    End If
                End If
                grdContact.TextMatrix(llRow, PCCHGDINDEX) = "0"
            End If
        End If
    Next llRow
    If blUpdateWeb Then
        ilRet = gWebTestEmailChange()
        If ilRet Then
            mSaveRec = True
        Else
            mSaveRec = False
        End If
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffContactGrid-mSaveRec"
    mSaveRec = False
    Exit Function
ErrHand1:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffContactGrid-mSaveRec2"
    mSaveRec = False
End Function
'ttp 5352
Public Property Get InValidEmails() As String
    Dim c As Integer
    Dim slRet As String
    Dim slEMail As String
    
    For c = 1 To grdContact.Rows - 1
        slEMail = Trim$(grdContact.TextMatrix(c, PCEMAILINDEX))
        If Len(slEMail) > 0 Then
            If Not myEmail.TestAddress(slEMail) Then
                slRet = slRet & ";" & slEMail
            End If
        End If
    Next c
    slRet = Mid(slRet, 2)
    InValidEmails = slRet
End Property
Public Property Get Verify() As Integer 'VBC NR
    pbcArrow.Visible = False 'VBC NR
    If imUpdateAllowed Then 'VBC NR
        'Add call to mTestFields
        Verify = True 'VBC NR
    Else 'VBC NR
        Verify = True 'VBC NR
    End If 'VBC NR
End Property 'VBC NR

Public Property Get AnyFieldChanged() As Boolean
    'D.S. TTP 9746 - 2/25/20 - Called in Stations when the Done button is pressed. Checks for any changes in Personnel
    Dim llRow As Integer
    Dim slChgd As String
    Dim slName As String
    Dim slEMail As String
    
    AnyFieldChanged = False
    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
        slName = Trim(grdContact.TextMatrix(llRow, PCNAMEINDEX))
        slEMail = Trim(grdContact.TextMatrix(llRow, PCEMAILINDEX))
        slChgd = grdContact.TextMatrix(llRow, PCCHGDINDEX)
        'I suppose that if there is no name or no email no then sense in saving the row
        If (slName <> "") Or (slEMail <> "") Then
            If slChgd = "1" Then
                AnyFieldChanged = True
                Exit Function
            End If
        End If
    Next llRow
End Property

Public Property Get VerifyRights(slSource As String) As Integer
    Dim llRow As Long
    Dim ilCount As Integer
    Dim slName As String
    Dim slEMail As String
    Dim ilRet As Integer
    Dim slStr As String
    
    VerifyRights = True
    If smEMailDistribution <> "Y" Then
        Exit Property
    End If
    slName = Trim(grdContact.TextMatrix(grdContact.FixedRows, PCNAMEINDEX))
    SQLQuery = "Select * From VEF_Vehicles where vefName = " & "'" & slName & "'"
    Set rst_vef = gSQLSelectCall(SQLQuery)
    If rst_vef.EOF Then
        Exit Property
    End If
        
    
    mSetShow
    pbcArrow.Visible = False
    ilCount = -1
    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
        grdContact.Row = llRow
        grdContact.Col = PCEMAILINDEX
        slName = Trim(grdContact.TextMatrix(llRow, PCNAMEINDEX))
        slEMail = Trim(grdContact.TextMatrix(llRow, PCEMAILINDEX))
        If ((slName <> "") Or (slEMail <> "")) And (grdContact.CellBackColor <> LIGHTYELLOW) Then
            If ilCount = -1 Then
                ilCount = 0
            End If
            slStr = UCase(grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX))
            If InStr(1, slStr, "PRIMARY", vbTextCompare) > 0 Then
                ilCount = ilCount + 1
            End If
        End If
    Next llRow
    If ilCount = 0 Then
        'If slSource = "M" Then  'Affiliate Management
        '    ilRet = MsgBox("One Master must be defined with Insertion Order E-Mail Rights, Personnel Not Saved", vbOKOnly + vbExclamation, "Save")
        'Else
            ilRet = MsgBox("One Primary must be defined with Insertion Order E-Mail Rights, Continue with Save?", vbYesNo + vbExclamation, "Question")
        'End If
        If ilRet = vbNo Then
            VerifyRights = False
        Else
            VerifyRights = True
        End If
    ElseIf ilCount > 1 Then
        'If slSource = "M" Then  'Affiliate Management
        '    ilRet = MsgBox("Only One Personnel Insertion Order E-Mails Rights allowed as Master, Personnel Not Saved", vbOKOnly + vbExclamation, "Save")
        'Else
            ilRet = MsgBox("Only One Personnel Insertion Order E-Mails Rights allowed as Primary, Save Disallowed", vbOKOnly + vbExclamation, "Save")
        'End If
        VerifyRights = False
    End If
End Property
Private Function mPopContactGrid() As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim llSetRow As Long
    Dim ilVef As Integer
    Dim slCallLetters As String
    
    mPopContactGrid = False
    On Error GoTo ErrHand:
    grdContact.Rows = 2
    mClearGrid grdContact
    gGrid_FillWithRows grdContact
    grdContact.Redraw = False
    llRow = grdContact.FixedRows
    'imShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
    bmStationIsVehicle = False
    SQLQuery = "SELECT * FROM shtt"
    SQLQuery = SQLQuery + " WHERE ("
    SQLQuery = SQLQuery & " ShttCode = " & imShttCode & ")"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If imShttCode > 0 Then
        If Not rst_Shtt.EOF Then
            slCallLetters = UCase(Trim$(rst_Shtt!shttCallLetters))
            For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If Trim$(UCase(tgVehicleInfo(ilVef).sVehicle)) = slCallLetters Then
                    bmStationIsVehicle = True
                End If
            Next ilVef
        End If
        SQLQuery = "SELECT * FROM artt"
        SQLQuery = SQLQuery + " WHERE ("
        SQLQuery = SQLQuery & " arttType = 'P'"
        SQLQuery = SQLQuery & " AND arttShttCode = " & imShttCode & ")"
        SQLQuery = SQLQuery & " ORDER BY arttFirstName, arttLastName"
        Set rst_artt = gSQLSelectCall(SQLQuery)
        Do While Not rst_artt.EOF
            If llRow >= grdContact.Rows Then
                grdContact.AddItem ""
            End If
            grdContact.Row = llRow
            grdContact.TextMatrix(llRow, PCNAMEINDEX) = Trim$(Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName))
            grdContact.TextMatrix(llRow, PCTITLEINDEX) = mGetTitle(rst_artt!arttTntCode)
            grdContact.TextMatrix(llRow, PCPHONEINDEX) = Trim$(rst_artt!arttPhone)
            grdContact.TextMatrix(llRow, PCFAXINDEX) = Trim$(rst_artt!arttFax)
            grdContact.TextMatrix(llRow, PCEMAILINDEX) = Trim$(rst_artt!arttEmail)
            grdContact.Col = PCEMAILRIGHTSINDEX
            If (Trim$(grdContact.TextMatrix(llRow, PCEMAILINDEX)) = "") Or (Not bmStationIsVehicle) Then
                grdContact.CellBackColor = LIGHTYELLOW
            Else
                grdContact.CellBackColor = vbWhite
                Select Case rst_artt!arttEmailRights
                    Case "M"
                        grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX) = lbcEMailRights.List(0)
                    Case "A"
                        grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX) = lbcEMailRights.List(1)
                    Case "V"
                        grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX) = lbcEMailRights.List(2)
                    Case Else
                        grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX) = lbcEMailRights.List(3)
                End Select
            End If
            grdContact.Col = PCAFFLABELINDEX
            grdContact.CellFontName = "Monotype Sorts"
            If rst_artt!arttAffContact = "1" Then
                grdContact.TextMatrix(llRow, PCAFFLABELINDEX) = "4"
            Else
                grdContact.TextMatrix(llRow, PCAFFLABELINDEX) = " "
            End If
            grdContact.Col = PCISCIINDEX
            grdContact.CellFontName = "Monotype Sorts"
            If rst_artt!arttISCI2Contact = "1" Then
                grdContact.TextMatrix(llRow, PCISCIINDEX) = "4"
            Else
                grdContact.TextMatrix(llRow, PCISCIINDEX) = " "
            End If
            grdContact.Col = PCAFFEMAILINDEX
            grdContact.CellFontName = "Monotype Sorts"
            If rst_artt!arttWebEMail = "Y" Then
                grdContact.TextMatrix(llRow, PCAFFEMAILINDEX) = "4"
            Else
                grdContact.TextMatrix(llRow, PCAFFEMAILINDEX) = " "
            End If
            grdContact.TextMatrix(llRow, PCARTTCODEINDEX) = rst_artt!arttCode
            grdContact.TextMatrix(llRow, PCDELETEINDEX) = "Delete"
            grdContact.Col = PCDELETEINDEX
            grdContact.CellBackColor = GRAY
            grdContact.TextMatrix(llRow, PCCHGDINDEX) = "0"
            llRow = llRow + 1
            rst_artt.MoveNext
        Loop
    End If
    grdContact.Rows = grdContact.Rows + (grdContact.Height \ grdContact.RowHeight(1))
    For llSetRow = llRow To grdContact.Rows - 1 Step 1
        grdContact.TextMatrix(llSetRow, PCARTTCODEINDEX) = 0
        grdContact.TextMatrix(llSetRow, PCCHGDINDEX) = "0"
    Next llSetRow
    imLastSort = -1
    imLastColSorted = -1
    mContactSortCol PCNAMEINDEX
    'SQLQuery = "SELECT * FROM shtt"
    'SQLQuery = SQLQuery + " WHERE ("
    'SQLQuery = SQLQuery & " ShttCode = " & imShttCode & ")"
    'Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If Not rst_Shtt.EOF Then
        grdContact.AddItem "", grdContact.FixedRows
        llRow = grdContact.FixedRows
        grdContact.Row = llRow
        For ilCol = PCNAMEINDEX To PCDELETEINDEX Step 1
            '7/25/11:  Disallow station E-Mail
            'If (ilCol < PCPHONEINDEX) Or (ilCol > PCEMAILINDEX) Then
            If (ilCol < PCPHONEINDEX) Or (ilCol >= PCEMAILINDEX) Then
                grdContact.Col = ilCol
                grdContact.CellBackColor = LIGHTYELLOW
            End If
        Next ilCol
        grdContact.TextMatrix(llRow, PCNAMEINDEX) = Trim$(rst_Shtt!shttCallLetters)
        grdContact.TextMatrix(llRow, PCTITLEINDEX) = "Main #"
        grdContact.TextMatrix(llRow, PCPHONEINDEX) = Trim$(rst_Shtt!shttPhone)
        grdContact.TextMatrix(llRow, PCFAXINDEX) = Trim$(rst_Shtt!shttFax)
        '7/25/11:  Disallow station E-Mail
        'grdContact.TextMatrix(llRow, PCEMAILINDEX) = Trim$(rst_Shtt!shttEMail)
        grdContact.TextMatrix(llRow, PCEMAILINDEX) = ""
        grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX) = ""
        grdContact.TextMatrix(llRow, PCAFFLABELINDEX) = ""   'Trim$(rst_shtt!shttWebAddress)
        grdContact.TextMatrix(llRow, PCISCIINDEX) = ""
        grdContact.TextMatrix(llRow, PCAFFEMAILINDEX) = ""   'Trim$(rst_shtt!shttWebAddress)
        grdContact.TextMatrix(llRow, PCCHGDINDEX) = "0"
        grdContact.TextMatrix(llRow, PCARTTCODEINDEX) = 0
    Else
        grdContact.AddItem "", grdContact.FixedRows
        llRow = grdContact.FixedRows
        grdContact.Row = llRow
        For ilCol = PCNAMEINDEX To PCDELETEINDEX Step 1
            '7/25/11:  Disallow station E-Mail
            'If (ilCol < PCPHONEINDEX) Or (ilCol > PCEMAILINDEX) Then
            If (ilCol < PCPHONEINDEX) Or (ilCol >= PCEMAILINDEX) Then
                grdContact.Col = ilCol
                grdContact.CellBackColor = LIGHTYELLOW
            End If
        Next ilCol
        grdContact.TextMatrix(llRow, PCNAMEINDEX) = ""
        grdContact.TextMatrix(llRow, PCTITLEINDEX) = "Main #"
        grdContact.TextMatrix(llRow, PCPHONEINDEX) = ""
        grdContact.TextMatrix(llRow, PCFAXINDEX) = ""
        grdContact.TextMatrix(llRow, PCEMAILINDEX) = ""
        grdContact.TextMatrix(llRow, PCEMAILRIGHTSINDEX) = ""
        grdContact.TextMatrix(llRow, PCAFFLABELINDEX) = ""   'Trim$(rst_shtt!shttWebAddress)
        grdContact.TextMatrix(llRow, PCISCIINDEX) = ""
        grdContact.TextMatrix(llRow, PCAFFEMAILINDEX) = ""   'Trim$(rst_shtt!shttWebAddress)
        grdContact.TextMatrix(llRow, PCCHGDINDEX) = "0"
        grdContact.TextMatrix(llRow, PCARTTCODEINDEX) = 0
    End If
    grdContact.Row = 0
    grdContact.Col = PCARTTCODEINDEX
    grdContact.Redraw = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffContactGrid-mPopContactGrid"
    grdContact.Redraw = True
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

Private Function mGetTitle(ilCode As Integer) As String
    Dim ilLoop As Integer
    mGetTitle = ""
    For ilLoop = 1 To lbcTitle.ListCount - 1 Step 1
        If ilCode = lbcTitle.ItemData(ilLoop) Then
            mGetTitle = lbcTitle.List(ilLoop)
        End If
    Next ilLoop
End Function

Public Property Let StationCode(ilShttCode As Long)
    'UserControl.Enabled = ilState
    imShttCode = ilShttCode
    PropertyChanged "StationCode"
End Property
Public Property Let Source(slSource As String)
    'UserControl.Enabled = ilState
    smSource = slSource
    PropertyChanged "Source"
End Property
Public Property Let PhoneNumber(slPhone As String)
    grdContact.TextMatrix(grdContact.FixedRows, PCPHONEINDEX) = slPhone
    PropertyChanged "PhoneNumber"
End Property
Public Property Let FaxNumber(slFax As String)
    grdContact.TextMatrix(grdContact.FixedRows, PCFAXINDEX) = slFax
    PropertyChanged "FaxNumber"
End Property
Public Property Let CALLLETTERS(slCallLetters As String)
    Dim ilCol As Integer
    
    grdContact.TextMatrix(grdContact.FixedRows, PCNAMEINDEX) = slCallLetters
    PropertyChanged "CallLetters"
    grdContact.Row = grdContact.FixedRows
    For ilCol = PCNAMEINDEX To PCDELETEINDEX Step 1
        '7/25/11:  Disallow station E-Mail
        'If (ilCol < PCPHONEINDEX) Or (ilCol > PCEMAILINDEX) Then
        If (ilCol < PCPHONEINDEX) Or (ilCol >= PCEMAILINDEX) Then
            grdContact.Col = ilCol
            grdContact.CellBackColor = LIGHTYELLOW
        End If
    Next ilCol
End Property

Private Sub UserControl_Resize()
    pbcArrow.Width = 90
    grdContact.Width = Width - pbcArrow.Width
    grdContact.Height = Height
    grdContact.Move pbcArrow.Width, 0
    gGrid_IntegralHeight grdContact
    grdContact.Height = grdContact.Height + 30
    UserControl.Height = grdContact.Height
    '6/11/15: Moved mSetGridColumn call to action 4 (Populate) instead of UserControl_ReSize
    'because this event is executed during the compile and will blow-up with the Pervasive call.
    'mSetGridColumns
    
End Sub

Private Sub mContactSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
        slStr = Trim$(grdContact.TextMatrix(llRow, PCNAMEINDEX))
        If (slStr <> "") Or (Trim$(grdContact.TextMatrix(llRow, PCEMAILINDEX)) <> "") Then
            slSort = UCase$(Trim$(grdContact.TextMatrix(llRow, ilCol)))
            If slSort = "" Then
                slSort = Chr(32)
            End If
            slStr = grdContact.TextMatrix(llRow, PCSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdContact.TextMatrix(llRow, PCSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdContact.TextMatrix(llRow, PCSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = PCSORTINDEX
    Else
        imLastColSorted = -1
        imLastColSorted = -1
    End If
    gGrid_SortByCol grdContact, PCNAMEINDEX, PCSORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Sub mSetLbcGridControl(lbcCtrl As ListBox)
    Dim slStr As String
    Dim ilIndex As Integer
    
    If grdContact.Col = PCTITLEINDEX Then
        edcDropdown.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) + grdContact.ColWidth(grdContact.Col + 1) + grdContact.ColWidth(grdContact.Col + 2) / 2, grdContact.RowHeight(grdContact.Row) - 15
    ElseIf grdContact.Col = PCEMAILRIGHTSINDEX Then
        edcDropdown.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) + grdContact.ColWidth(grdContact.Col + 1) + grdContact.ColWidth(grdContact.Col + 2) / 2, grdContact.RowHeight(grdContact.Row) - 15
    Else
        edcDropdown.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - cmcDropDown.Width - 30, grdContact.RowHeight(grdContact.Row) - 15
    End If
    cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
    lbcCtrl.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
    gSetListBoxHeight lbcCtrl, 6
    If lbcCtrl.Top + lbcCtrl.Height > grdContact.Height Then
        lbcCtrl.Top = edcDropdown.Top - lbcCtrl.Height
        If lbcCtrl.Top <= 0 Then
            lbcCtrl.Move cmcDropDown.Left + cmcDropDown.Width, (grdContact.Height - lbcCtrl.Height) / 2, edcDropdown.Width + cmcDropDown.Width
        End If
    Else
        If lbcCtrl.Top + lbcCtrl.Height > grdContact.Height Then
            lbcCtrl.Move cmcDropDown.Left + cmcDropDown.Width, (grdContact.Height - lbcCtrl.Height) / 2, edcDropdown.Width + cmcDropDown.Width
        End If
    End If
    slStr = grdContact.Text
    ilIndex = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRING, -1, slStr)
    If ilIndex >= 0 Then
        lbcCtrl.ListIndex = ilIndex
        edcDropdown.Text = lbcCtrl.List(lbcCtrl.ListIndex)
    Else
        lbcCtrl.ListIndex = -1
        edcDropdown.Text = ""
    End If
    If edcDropdown.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdContact.RowHeight(grdContact.Row) - 15
    End If
    edcDropdown.Visible = True
    cmcDropDown.Visible = True
    lbcCtrl.Visible = True
    edcDropdown.SetFocus
End Sub


Private Sub mSetEdcGridControl()
    edcDropdown.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30, grdContact.RowHeight(grdContact.Row) - 15
    edcDropdown.Text = grdContact.Text
    If edcDropdown.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdContact.RowHeight(grdContact.Row) - 15
    End If
    edcDropdown.Visible = True
    edcDropdown.SetFocus
End Sub

Private Sub mPopTitles()
    Dim slSave As String
    Dim rstTitles As ADODB.Recordset
    

    lbcTitle.Clear
    SQLQuery = "Select tntCode, tntTitle From Tnt"
    Set rstTitles = gSQLSelectCall(SQLQuery)
    While Not rstTitles.EOF
        lbcTitle.AddItem (Trim(rstTitles!tntTitle))
        lbcTitle.ItemData(lbcTitle.NewIndex) = rstTitles!tntCode
        rstTitles.MoveNext
    Wend
    lbcTitle.AddItem "[New]", 0
    lbcTitle.ItemData(lbcTitle.NewIndex) = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AffContactGrid-mPopTitles"
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
            If (InStr(1, slStr, "Arrow", vbTextCompare) > 0) Or ((InStr(1, slStr, "Dropdown", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Then
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
    rst_Shtt.Close
    rst_artt.Close
End Sub


