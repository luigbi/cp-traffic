VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrAudio 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrAudio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11670
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.ListBox lbcANE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   1
      ItemData        =   "EngrAudio.frx":030A
      Left            =   7620
      List            =   "EngrAudio.frx":0311
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   10755
      Top             =   3435
   End
   Begin VB.ListBox lbcANE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   0
      ItemData        =   "EngrAudio.frx":031D
      Left            =   6015
      List            =   "EngrAudio.frx":0324
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1410
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
      Left            =   5490
      Picture         =   "EngrAudio.frx":0330
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2745
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   4545
      TabIndex        =   9
      Top             =   2775
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcCCE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrAudio.frx":042A
      Left            =   4200
      List            =   "EngrAudio.frx":042C
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1935
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4350
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox edcGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2310
      TabIndex        =   5
      Top             =   1710
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   6555
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   195
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   450
      Width           =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   480
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   60
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   105
      Picture         =   "EngrAudio.frx":042E
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6945
      TabIndex        =   15
      Top             =   6705
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10770
      Top             =   6330
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7290
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5175
      TabIndex        =   14
      Top             =   6705
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3390
      TabIndex        =   13
      Top             =   6705
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAudio 
      Height          =   5925
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   10451
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   2
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmcSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   17
      Top             =   75
      Width           =   795
   End
   Begin VB.TextBox edcSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8370
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1725
      Picture         =   "EngrAudio.frx":0738
      Top             =   6690
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Audio Source"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   840
      Picture         =   "EngrAudio.frx":0A42
      Top             =   6690
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   9840
      Picture         =   "EngrAudio.frx":130C
      Top             =   6705
      Width           =   480
   End
End
Attribute VB_Name = "EngrAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrAudio - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private smState As String
Private imInChg As Integer
Private imBSMode As Integer
Private imASECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private imDoubleClickName As Integer
Private imLbcMouseDown As Integer

Private tmASE As ASE


Private imDeleteCodes() As Integer

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer

Const PRIAUDIOINDEX = 0
Const PRIAUDIOTYPEINDEX = 1
Const PRICTRLINDEX = 2
Const DESCRIPTIONINDEX = 3
Const BKUPAUDIOINDEX = 6
Const BKUPCTRLINDEX = 7
Const PROTAUDIOINDEX = 4
Const PROTCTRLINDEX = 5
Const STATEINDEX = 8
Const CODEINDEX = 9
Const USEDFLAGINDEX = 10

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_Search(grdAudio, PRIAUDIOINDEX, slStr)
    If llRow >= 0 Then
        mEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mSetShow
End Sub


Private Function mNameOk() As Integer
    Dim ilError As Integer
    Dim llRow As Long
    Dim llTestRow As Long
    Dim slStr As String
    Dim slTestStr As String
    
    grdAudio.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
        slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdAudio.Rows - 1 Step 1
                slTestStr = Trim$(grdAudio.TextMatrix(llTestRow, PRIAUDIOINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdAudio.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdAudio.Row = llRow
                        grdAudio.Col = PRIAUDIOINDEX
                        grdAudio.CellForeColor = vbRed
                    Else
                        grdAudio.Row = llTestRow
                        grdAudio.Col = PRIAUDIOINDEX
                        grdAudio.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdAudio.Redraw = True
    If ilError Then
        MsgBox "Duplicate Names Found, Save Stopped", vbOKOnly + vbExclamation
        mNameOk = False
        Exit Function
    Else
        mNameOk = True
        Exit Function
    End If
End Function

Private Sub mSortCol(ilCol As Integer)
    Dim llEndRow As Long
    mSetShow
    gGrid_SortByCol grdAudio, PRIAUDIOINDEX, ilCol, imLastColSorted, imLastSort
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
    End If
End Sub

Private Sub mEnableBox()
    Dim slStr As String
    Dim ilIndex As Integer
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdAudio.Row >= grdAudio.FixedRows) And (grdAudio.Row < grdAudio.Rows) And (grdAudio.Col >= 0) And (grdAudio.Col < grdAudio.Cols - 1) Then
        lmEnableRow = grdAudio.Row
        lmEnableCol = grdAudio.Col
        sgReturnCallName = grdAudio.TextMatrix(lmEnableRow, PRIAUDIOINDEX)
        imShowGridBox = True
        pbcArrow.Move grdAudio.Left - pbcArrow.Width - 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + (grdAudio.RowHeight(grdAudio.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdAudio.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdAudio.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdAudio.TextMatrix(lmEnableRow, PRIAUDIOINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdAudio.Col
            Case PRIAUDIOINDEX  'Call Letters
                mPopUnusedANE lmEnableRow
'                edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
'                lbcANE(0).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
'                gSetListBoxHeight lbcANE(0), CLng(grdAudio.Height / 2)
'                If lbcANE(0).Top + lbcANE(0).Height > cmcCancel.Top Then
'                    lbcANE(0).Top = edcDropdown.Top - lbcANE(0).Height
'                End If
                slStr = grdAudio.text
                'ilIndex = SendMessageByString(lbcANE(0).hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcANE(0), slStr)
                If ilIndex >= 0 Then
                    lbcANE(0).ListIndex = ilIndex
                    edcDropdown.text = lbcANE(0).List(lbcANE(0).ListIndex)
                Else
                    edcDropdown.text = ""
                    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
                        lbcANE(0).ListIndex = 0
                        edcDropdown.text = lbcANE(0).List(lbcANE(0).ListIndex)
                    Else
                        lbcANE(0).ListIndex = 0
                        edcDropdown.text = lbcANE(0).List(lbcANE(0).ListIndex)
                    End If
                End If
'                edcDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcANE(0).Visible = True
'                edcDropdown.SetFocus
            Case PRICTRLINDEX
                'edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col - 1) + grdAudio.ColWidth(grdAudio.Col - 1) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
'                lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
'                gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
'                If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
'                    lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
'                End If
'                slStr = grdAudio.Text
                'ilIndex = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcCCE, slStr)
                If ilIndex >= 0 Then
                    lbcCCE.ListIndex = ilIndex
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcCCE.ListIndex = 1
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                End If
'                edcDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcCCE.Visible = True
'                edcDropdown.SetFocus
            Case DESCRIPTIONINDEX  'Date
'                edcGrid.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
                edcGrid.MaxLength = Len(tmASE.sDescription)
                edcGrid.text = grdAudio.text
'                edcGrid.Visible = True
'                edcGrid.SetFocus
            Case BKUPAUDIOINDEX
'                edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
'                lbcANE(1).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
'                gSetListBoxHeight lbcANE(1), CLng(grdAudio.Height / 2)
'                If lbcANE(1).Top + lbcANE(1).Height > cmcCancel.Top Then
'                    lbcANE(1).Top = edcDropdown.Top - lbcANE(1).Height
'                End If
                slStr = grdAudio.text
                'ilIndex = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcANE(1), slStr)
                If ilIndex >= 0 Then
                    lbcANE(1).ListIndex = ilIndex
                    edcDropdown.text = lbcANE(1).List(lbcANE(1).ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcANE(1).ListIndex = 1
                    edcDropdown.text = lbcANE(1).List(lbcANE(1).ListIndex)
                End If
'                edcDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcANE(1).Visible = True
'                edcDropdown.SetFocus
            Case BKUPCTRLINDEX
                'edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col - 1) + grdAudio.ColWidth(grdAudio.Col - 1) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
'                lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
'                gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
'                If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
'                    lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
'                End If
                slStr = grdAudio.text
                'ilIndex = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcCCE, slStr)
                If ilIndex >= 0 Then
                    lbcCCE.ListIndex = ilIndex
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcCCE.ListIndex = 1
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                End If
'                edcDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcCCE.Visible = True
'                edcDropdown.SetFocus
            Case PROTAUDIOINDEX
'                edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
'                lbcANE(1).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
'                gSetListBoxHeight lbcANE(1), CLng(grdAudio.Height / 2)
'                If lbcANE(1).Top + lbcANE(1).Height > cmcCancel.Top Then
'                    lbcANE(1).Top = edcDropdown.Top - lbcANE(1).Height
'                End If
                slStr = grdAudio.text
                'ilIndex = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcANE(1), slStr)
                If ilIndex >= 0 Then
                    lbcANE(1).ListIndex = ilIndex
                    edcDropdown.text = lbcANE(1).List(lbcANE(1).ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcANE(1).ListIndex = 1
                    edcDropdown.text = lbcANE(1).List(lbcANE(1).ListIndex)
                End If
                'Set value as change event might not have happened if same value in two adjacent columns
                grdAudio.text = edcDropdown.text
'                edcDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcANE(1).Visible = True
'                edcDropdown.SetFocus
            Case PROTCTRLINDEX
                'edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col - 1) + grdAudio.ColWidth(grdAudio.Col - 1) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
'                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
'                lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
'                gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
'                If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
'                    lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
'                End If
                slStr = grdAudio.text
                'ilIndex = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcCCE, slStr)
                If ilIndex >= 0 Then
                    lbcCCE.ListIndex = ilIndex
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcCCE.ListIndex = 1
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                End If
                'Set value as change event might not have happened if same value in two adjacent columns
                grdAudio.text = edcDropdown.text
'                edcDropdown.Visible = True
'                cmcDropDown.Visible = True
'                lbcCCE.Visible = True
'                edcDropdown.SetFocus
            Case STATEINDEX
'                pbcState.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
                smState = grdAudio.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
'                pbcState.Visible = True
'                pbcState.SetFocus
        End Select
        smESCValue = grdAudio.text
        mSetFocus
    End If
End Sub
Private Sub mSetShow()
    Dim llRow As Long
    Dim slStr As String
    
    tmcClick.Enabled = False
    If (lmEnableRow >= grdAudio.FixedRows) And (lmEnableRow < grdAudio.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = grdAudio.TextMatrix(lmEnableRow, lmEnableCol)
        If StrComp(slStr, "[None]", vbTextCompare) = 0 Then
            slStr = ""
            grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End If
        Select Case lmEnableCol
            Case PRIAUDIOINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcANE(0).hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcANE(0), slStr)
                If (llRow <= 0) Then
                    grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case PRICTRLINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcCCE, slStr)
                If (llRow <= 0) Then
                    grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case DESCRIPTIONINDEX
            Case BKUPAUDIOINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcANE(1), slStr)
                If (llRow <= 0) Then
                    grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case BKUPCTRLINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcCCE, slStr)
                If (llRow <= 0) Then
                    grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case PROTAUDIOINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcANE(1), slStr)
                If (llRow <= 0) Then
                    grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case PROTCTRLINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcCCE, slStr)
                If (llRow <= 0) Then
                    grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
                If (Trim$(grdAudio.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdAudio.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdAudio.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdAudio.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdAudio.TextMatrix(lmEnableRow, PRIAUDIOINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    edcGrid.Visible = False
    lbcANE(0).Visible = False
    lbcANE(0).Visible = False
    lbcANE(1).Visible = False
    lbcANE(1).Visible = False
    lbcCCE.Visible = False
    cmcDropDown.Visible = False
    edcDropdown.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    Dim llLbc As Long
    Dim slStr1 As String
    
    grdAudio.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
        slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdAudio.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdAudio.TextMatrix(llRow, PRIAUDIOINDEX) = "Missing"
                grdAudio.Row = llRow
                grdAudio.Col = PRIAUDIOINDEX
                grdAudio.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
                'Use lbcANE(1) instead of lbcANE(0) since lbcANE(1) has all names
                'llLbc = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                llLbc = gListBoxFind(lbcANE(1), slStr)
                If (llLbc <= 0) Then
                    ilError = True
                    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                        grdAudio.TextMatrix(llRow, PRIAUDIOINDEX) = "Missing"
                    End If
                    grdAudio.Row = llRow
                    grdAudio.Col = PRIAUDIOINDEX
                    grdAudio.CellForeColor = vbRed
                Else
                    slStr1 = Trim$(grdAudio.TextMatrix(llRow, BKUPAUDIOINDEX))
                    If slStr1 <> "" Then
                        If StrComp(slStr1, "[None]", vbTextCompare) <> 0 Then
                            If StrComp(slStr, slStr1, vbTextCompare) = 0 Then
                                grdAudio.Row = llRow
                                grdAudio.Col = BKUPAUDIOINDEX
                                grdAudio.CellForeColor = vbRed
                            End If
                        End If
                    End If
                    slStr1 = Trim$(grdAudio.TextMatrix(llRow, PROTAUDIOINDEX))
                    If slStr1 <> "" Then
                        If StrComp(slStr1, "[None]", vbTextCompare) <> 0 Then
                            If StrComp(slStr, slStr1, vbTextCompare) = 0 Then
                                grdAudio.Row = llRow
                                grdAudio.Col = PROTAUDIOINDEX
                                grdAudio.CellForeColor = vbRed
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdAudio.TextMatrix(llRow, BKUPAUDIOINDEX))
                If slStr <> "" Then
                    If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                        slStr1 = Trim$(grdAudio.TextMatrix(llRow, PROTAUDIOINDEX))
                        If slStr1 <> "" Then
                            If StrComp(slStr1, "[None]", vbTextCompare) <> 0 Then
                                If StrComp(slStr, slStr1, vbTextCompare) = 0 Then
                                    grdAudio.Row = llRow
                                    grdAudio.Col = PROTAUDIOINDEX
                                    grdAudio.CellForeColor = vbRed
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = grdAudio.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdAudio.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdAudio.Row = llRow
                    grdAudio.Col = STATEINDEX
                    grdAudio.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdAudio.Redraw = True
    If ilError Then
        mCheckFields = False
        Exit Function
    Else
        mCheckFields = True
        Exit Function
    End If
End Function


Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    gGrid_AlignAllColsLeft grdAudio
    mGridColumnWidth
    'Set Titles
    grdAudio.TextMatrix(0, PRIAUDIOINDEX) = "Primary Audio"
    grdAudio.TextMatrix(0, PRIAUDIOTYPEINDEX) = "Primary Audio"
    grdAudio.TextMatrix(0, PRICTRLINDEX) = "Primary Audio"
    grdAudio.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdAudio.TextMatrix(1, PRIAUDIOINDEX) = "Name"
    grdAudio.TextMatrix(1, PRIAUDIOTYPEINDEX) = "Type"
    grdAudio.TextMatrix(1, PRICTRLINDEX) = "Control"
    grdAudio.TextMatrix(0, BKUPAUDIOINDEX) = "Backup Audio"
    grdAudio.TextMatrix(0, BKUPCTRLINDEX) = "Backup Audio"
    grdAudio.TextMatrix(1, BKUPAUDIOINDEX) = "Name"
    grdAudio.TextMatrix(1, BKUPCTRLINDEX) = "Control"
    grdAudio.TextMatrix(0, PROTAUDIOINDEX) = "Protection"
    grdAudio.TextMatrix(0, PROTCTRLINDEX) = "Protection"
    grdAudio.TextMatrix(1, PROTAUDIOINDEX) = "Name"
    grdAudio.TextMatrix(1, PROTCTRLINDEX) = "Control"
    grdAudio.TextMatrix(0, STATEINDEX) = "State"
    grdAudio.TextMatrix(0, CODEINDEX) = "Code"
    grdAudio.TextMatrix(0, USEDFLAGINDEX) = "Used"
    grdAudio.Row = 1
    For ilCol = 0 To grdAudio.Cols - 1 Step 1
        grdAudio.Col = ilCol
        grdAudio.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdAudio.Row = 0
    grdAudio.MergeCells = flexMergeRestrictRows
    grdAudio.MergeRow(0) = True
    grdAudio.Row = 0
    grdAudio.Col = PRIAUDIOINDEX
    grdAudio.CellAlignment = flexAlignCenterCenter
    grdAudio.Row = 0
    grdAudio.Col = BKUPAUDIOINDEX
    grdAudio.CellAlignment = flexAlignCenterCenter
    grdAudio.Row = 0
    grdAudio.Col = PROTAUDIOINDEX
    grdAudio.CellAlignment = flexAlignCenterCenter
    grdAudio.Height = cmcCancel.Top - grdAudio.Top - 120    '8 * grdAudio.RowHeight(0) + 30
    gGrid_IntegralHeight grdAudio
    gGrid_Clear grdAudio, True
    grdAudio.Row = grdAudio.FixedRows

End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdAudio.Width = EngrAudio.Width - 2 * grdAudio.Left
    grdAudio.ColWidth(CODEINDEX) = 0
    grdAudio.ColWidth(USEDFLAGINDEX) = 0
    grdAudio.ColWidth(PRIAUDIOINDEX) = grdAudio.Width / 10
    grdAudio.ColWidth(PRIAUDIOTYPEINDEX) = grdAudio.Width / 6
    If tgUsedSumEPE.sAudioControl <> "Y" Then
        grdAudio.ColWidth(PRICTRLINDEX) = 0
    Else
        grdAudio.ColWidth(PRICTRLINDEX) = grdAudio.Width / 18
    End If
    If tgUsedSumEPE.sBkupAudioName <> "Y" Then
        grdAudio.ColWidth(BKUPAUDIOINDEX) = 0
    Else
        grdAudio.ColWidth(BKUPAUDIOINDEX) = grdAudio.Width / 10
    End If
    If tgUsedSumEPE.sBkupAudioControl <> "Y" Then
        grdAudio.ColWidth(BKUPCTRLINDEX) = 0
    Else
        grdAudio.ColWidth(BKUPCTRLINDEX) = grdAudio.Width / 18
    End If
    If tgUsedSumEPE.sProtAudioName <> "Y" Then
        grdAudio.ColWidth(PROTAUDIOINDEX) = 0
    Else
        grdAudio.ColWidth(PROTAUDIOINDEX) = grdAudio.Width / 10
    End If
    If tgUsedSumEPE.sProtAudioControl <> "Y" Then
        grdAudio.ColWidth(PROTCTRLINDEX) = 0
    Else
        grdAudio.ColWidth(PROTCTRLINDEX) = grdAudio.Width / 18
    End If
    grdAudio.ColWidth(STATEINDEX) = grdAudio.Width / 19
    grdAudio.ColWidth(DESCRIPTIONINDEX) = grdAudio.Width - GRIDSCROLLWIDTH
    For ilCol = PRIAUDIOINDEX To USEDFLAGINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdAudio.ColWidth(DESCRIPTIONINDEX) > grdAudio.ColWidth(ilCol) Then
                grdAudio.ColWidth(DESCRIPTIONINDEX) = grdAudio.ColWidth(DESCRIPTIONINDEX) - grdAudio.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdAudio, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim ilCCE As Integer
    Dim ilANE As Integer
    Dim ilASE As Integer
    Dim slStr As String
        
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdAudio.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdAudio.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmASE.iCode = Val(grdAudio.TextMatrix(llRow, CODEINDEX))
    tmASE.iPriAneCode = 0
    slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
            tmASE.iPriAneCode = tgCurrANE(ilANE).iCode
            Exit For
        End If
    Next ilANE
    tmASE.iPriCceCode = 0
    slStr = Trim$(grdAudio.TextMatrix(llRow, PRICTRLINDEX))
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
            tmASE.iPriCceCode = tgCurrAudioCCE(ilCCE).iCode
            Exit For
        End If
    Next ilCCE
    tmASE.sDescription = grdAudio.TextMatrix(llRow, DESCRIPTIONINDEX)
    tmASE.iBkupAneCode = 0
    slStr = Trim$(grdAudio.TextMatrix(llRow, BKUPAUDIOINDEX))
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
            tmASE.iBkupAneCode = tgCurrANE(ilANE).iCode
            Exit For
        End If
    Next ilANE
    tmASE.iBkupCceCode = 0
    slStr = Trim$(grdAudio.TextMatrix(llRow, BKUPCTRLINDEX))
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
            tmASE.iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode
            Exit For
        End If
    Next ilCCE
    tmASE.iProtAneCode = 0
    slStr = Trim$(grdAudio.TextMatrix(llRow, PROTAUDIOINDEX))
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
            tmASE.iProtAneCode = tgCurrANE(ilANE).iCode
            Exit For
        End If
    Next ilANE
    tmASE.iProtCceCode = 0
    slStr = Trim$(grdAudio.TextMatrix(llRow, PROTCTRLINDEX))
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
            tmASE.iProtCceCode = tgCurrAudioCCE(ilCCE).iCode
            Exit For
        End If
    Next ilCCE
    If grdAudio.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmASE.sState = "D"
    Else
        tmASE.sState = "A"
    End If
    If tmASE.iCode <= 0 Then
        tmASE.sUsedFlag = "N"
    Else
        tmASE.sUsedFlag = grdAudio.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmASE.iVersion = 0
    tmASE.iOrigAseCode = tmASE.iCode
    tmASE.sCurrent = "Y"
    'tmASE.sEnteredDate = smNowDate
    'tmASE.sEnteredTime = smNowTime
    tmASE.sEnteredDate = Format(Now, sgShowDateForm)
    tmASE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmASE.iUieCode = tgUIE.iCode
    tmASE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilCCE As Integer
    Dim ilANE As Integer
    Dim ilASE As Integer
    Dim ilATE As Integer
    
    'gGrid_Clear grdAudio, True
    llRow = grdAudio.FixedRows
    For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
        If llRow + 1 > grdAudio.Rows Then
            grdAudio.AddItem ""
        End If
        grdAudio.Row = llRow
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            If tgCurrASE(ilLoop).iPriAneCode = tgCurrANE(ilANE).iCode Then
                grdAudio.TextMatrix(llRow, PRIAUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                    If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                        grdAudio.TextMatrix(llRow, PRIAUDIOTYPEINDEX) = Trim$(tgCurrATE(ilATE).sName)
                        Exit For
                    End If
                Next ilATE
                Exit For
            End If
        Next ilANE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrASE(ilLoop).iPriCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdAudio.TextMatrix(llRow, PRICTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                Exit For
            End If
        Next ilCCE
        grdAudio.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrASE(ilLoop).sDescription)
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            If tgCurrASE(ilLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
                grdAudio.TextMatrix(llRow, BKUPAUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                Exit For
            End If
        Next ilANE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrASE(ilLoop).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdAudio.TextMatrix(llRow, BKUPCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                Exit For
            End If
        Next ilCCE
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            If tgCurrASE(ilLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
                grdAudio.TextMatrix(llRow, PROTAUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                Exit For
            End If
        Next ilANE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrASE(ilLoop).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdAudio.TextMatrix(llRow, PROTCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                Exit For
            End If
        Next ilCCE
        If tgCurrASE(ilLoop).sState = "A" Then
            grdAudio.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdAudio.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdAudio.TextMatrix(llRow, CODEINDEX) = tgCurrASE(ilLoop).iCode
        grdAudio.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrASE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdAudio.Rows Then
        grdAudio.AddItem ""
    End If
    For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
        grdAudio.Row = llRow
        grdAudio.Col = PRIAUDIOTYPEINDEX
        grdAudio.CellBackColor = LIGHTYELLOW
    Next llRow
    grdAudio.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrAudio-mPopulate Audio", tgCurrASE())

End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlASE As ASE
    
    gSetMousePointer grdAudio, grdAudio, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdAudio, grdAudio, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdAudio, grdAudio, vbDefault
        mSave = False
        Exit Function
    End If
    grdAudio.Redraw = False
    For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If tmASE.iPriAneCode > 0 Then
            imASECode = tmASE.iCode
            If tmASE.iCode > 0 Then
                ilRet = gGetRec_ASE_AudioSource(imASECode, "Audio Source-mSave: Get ASE", tlASE)
                If ilRet Then
                    If mCompare(tmASE, tlASE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmASE.iVersion = tlASE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmASE.iCode <= 0 Then
                    ilRet = gPutInsert_ASE_AudioSource(0, tmASE, "Audio Source-mSave: Insert ASE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_ASE_AudioSource(1, tmASE, "Audio Source-mSave: Update ASE")
                    ilRet = gPutDelete_ASE_AudioSource(tmASE.iCode, "Audio Source-mSave: Delete ASE")
                    ilRet = gPutInsert_ASE_AudioSource(1, tmASE, "Audio Source-mSave: Insert ASE")
                End If
                'Set used flag if required
                ilRet = gPutUpdate_ANE_UsedFlag(tmASE.iPriAneCode, tgCurrANE())
                ilRet = gPutUpdate_ANE_UsedFlag(tmASE.iBkupAneCode, tgCurrANE())
                ilRet = gPutUpdate_ANE_UsedFlag(tmASE.iProtAneCode, tgCurrANE())
                ilRet = gPutUpdate_CCE_UsedFlag(tmASE.iPriCceCode, tgCurrAudioCCE())
                ilRet = gPutUpdate_CCE_UsedFlag(tmASE.iBkupCceCode, tgCurrAudioCCE())
                ilRet = gPutUpdate_CCE_UsedFlag(tmASE.iProtCceCode, tgCurrAudioCCE())
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_ASE_AudioSource(imDeleteCodes(ilLoop), "EngrAudio- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    
    grdAudio.Redraw = True
    sgCurrASEStamp = ""
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrAudio-mSave: Populate", tgCurrASE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrAudio
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrAudio
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdAudio, grdAudio, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdAudio, grdAudio, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdAudio, grdAudio, vbDefault
    Unload EngrAudio
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdAudio.Col
        Case PRIAUDIOINDEX
            lbcANE(0).Visible = Not lbcANE(0).Visible
        Case PRICTRLINDEX, BKUPCTRLINDEX, PROTCTRLINDEX
            lbcCCE.Visible = Not lbcCCE.Visible
        Case BKUPAUDIOINDEX, PROTAUDIOINDEX
            lbcANE(1).Visible = Not lbcANE(1).Visible
    End Select
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim llTopRow As Long
    
    If imFieldChgd = True Then
        gSetMousePointer grdAudio, grdAudio, vbHourglass
        llTopRow = grdAudio.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdAudio, grdAudio, vbDefault
            Exit Sub
        End If
        grdAudio.Redraw = False
        mClearControls
        mMoveRecToCtrls
        If imLastColSorted >= 0 Then
            If imLastSort = flexSortStringNoCaseDescending Then
                imLastSort = flexSortStringNoCaseAscending
            Else
                imLastSort = flexSortStringNoCaseDescending
            End If
            mSortCol imLastColSorted
        Else
            imLastSort = -1
            mSortCol 0
        End If
        grdAudio.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdAudio, grdAudio, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    Dim ilANE As Integer
    Dim ilCCE As Integer
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdAudio.Col
        Case PRIAUDIOINDEX
            'llRow = SendMessageByString(lbcANE(0).hwnd, LB_FINDSTRING, 1, slStr)
            llRow = gListBoxFind(lbcANE(0), slStr)
            If llRow >= 0 Then
                lbcANE(0).ListIndex = llRow
                edcDropdown.text = lbcANE(0).List(lbcANE(0).ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case PRICTRLINDEX, BKUPCTRLINDEX, PROTCTRLINDEX
            'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcCCE, slStr)
            If llRow >= 0 Then
                lbcCCE.ListIndex = llRow
                edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case BKUPAUDIOINDEX, PROTAUDIOINDEX
            'llRow = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcANE(1), slStr)
            If llRow >= 0 Then
                lbcANE(1).ListIndex = llRow
                edcDropdown.text = lbcANE(1).List(lbcANE(1).ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If StrComp(grdAudio.text, edcDropdown.text, vbTextCompare) <> 0 Then
        imFieldChgd = True
        Select Case grdAudio.Col
            Case PRIAUDIOINDEX, BKUPAUDIOINDEX, PROTAUDIOINDEX
                slStr = Trim$(edcDropdown.text)
                For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                        If grdAudio.Col = PRIAUDIOINDEX Then
                            grdAudio.TextMatrix(grdAudio.Row, DESCRIPTIONINDEX) = Trim$(tgCurrANE(ilANE).sDescription)
                        End If
                        If tgCurrANE(ilANE).iCceCode > 0 Then
                            For ilCCE = 0 To lbcCCE.ListCount - 1 Step 1
                                If tgCurrANE(ilANE).iCceCode = lbcCCE.ItemData(ilCCE) Then
                                    If grdAudio.Col = PRIAUDIOINDEX Then
                                        grdAudio.TextMatrix(grdAudio.Row, PRICTRLINDEX) = Trim$(lbcCCE.List(ilCCE))
                                    ElseIf grdAudio.Col = BKUPAUDIOINDEX Then
                                        grdAudio.TextMatrix(grdAudio.Row, BKUPCTRLINDEX) = Trim$(lbcCCE.List(ilCCE))
                                    ElseIf grdAudio.Col = PROTAUDIOINDEX Then
                                        grdAudio.TextMatrix(grdAudio.Row, PROTCTRLINDEX) = Trim$(lbcCCE.List(ilCCE))
                                    End If
                                    Exit For
                                End If
                            Next ilCCE
                        End If
                        Exit For
                    End If
                Next ilANE
        End Select
    End If
    If StrComp(Trim$(edcDropdown.text), "[None]", vbTextCompare) <> 0 Then
        grdAudio.text = edcDropdown.text
    Else
        grdAudio.text = ""
    End If
    grdAudio.CellForeColor = vbBlack
    mSetCommands
End Sub

Private Sub edcDropdown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
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

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdAudio.Col
            Case PRIAUDIOINDEX
                gProcessArrowKey Shift, KeyCode, lbcANE(0), True
            Case PRICTRLINDEX, BKUPCTRLINDEX, PROTCTRLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE, True
            Case BKUPAUDIOINDEX, PROTAUDIOINDEX
                gProcessArrowKey Shift, KeyCode, lbcANE(1), True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcDropdown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilRet As Integer
    
    If imDoubleClickName Then
        ilRet = mBranch()
    End If

End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdAudio.Col
        Case DESCRIPTIONINDEX
            If grdAudio.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudio.text = edcGrid.text
            grdAudio.CellForeColor = vbBlack
        Case STATEINDEX
    End Select
    mSetCommands
End Sub

Private Sub edcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSearch_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
        mFindMatch True
    End If
    imFirstActivate = False
    Me.KeyPreview = True
End Sub

Private Sub Form_Click()
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrAudio
    gCenterFormModal EngrAudio
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdAudio.FixedRows) And (lmEnableRow < grdAudio.Rows) Then
            If (lmEnableCol >= grdAudio.FixedCols) And (lmEnableCol < grdAudio.Cols) Then
                If lmEnableCol <> STATEINDEX Then
                    grdAudio.text = smESCValue
                Else
                    smState = smESCValue
                End If
                mSetShow
                mEnableBox
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
End Sub

Private Sub Form_Resize()
    Dim llRow As Long
    
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdAudio.Height = cmcCancel.Top - grdAudio.Top - 120    '8 * grdAudio.RowHeight(0) + 30
    gGrid_IntegralHeight grdAudio
    gGrid_FillWithRows grdAudio
    For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
        grdAudio.Row = llRow
        grdAudio.Col = PRIAUDIOTYPEINDEX
        grdAudio.CellBackColor = LIGHTYELLOW
    Next llRow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrAudio = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdAudio, grdAudio, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim imDeleteCodes(0 To 0) As Integer
    cmcSearch.Top = 30
    edcSearch.Top = cmcSearch.Top
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imFirstActivate = True
    imInChg = True
    mPopulate
    mPopATE
    mPopANE
    mPopCCE
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    imLbcMouseDown = False
    imDoubleClickName = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdAudio, grdAudio, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdAudio, grdAudio, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Relay Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Relay Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub grdAudio_DblClick()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
        Select Case grdAudio.Col
            Case PRIAUDIOINDEX, BKUPAUDIOINDEX, PROTAUDIOINDEX
                igInitCallInfo = 1
                sgInitCallName = grdAudio.TextMatrix(grdAudio.Row, grdAudio.Col)
                EngrAudioName.Show vbModal
                cmcCancel.SetFocus
            Case PRICTRLINDEX, BKUPCTRLINDEX, PROTCTRLINDEX
                igInitCallInfo = 1
                sgInitCallName = grdAudio.TextMatrix(grdAudio.Row, grdAudio.Col)
                EngrControlChar.Show vbModal
                cmcCancel.SetFocus
        End Select
    End If

End Sub

Private Sub imcInsert_Click()
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = AUDIOSOURCE_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub lbcANE_Click(Index As Integer)
    tmcClick.Enabled = False
    edcDropdown.text = lbcANE(Index).List(lbcANE(Index).ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        'lbcANE(Index).Visible = False
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcANE_DblClick(Index As Integer)
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcANE(Index).Visible = False
End Sub

Private Sub lbcANE_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcANE(Index), y)
    If (llRow < lbcANE(Index).ListCount) And (lbcANE(Index).ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcANE(Index).ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
            If ilCode = tgCurrANE(ilLoop).iCode Then
                lbcANE(Index).ToolTipText = Trim$(tgCurrANE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub



Private Sub lbcCCE_Click()
    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        'lbcCCE.Visible = False
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcCCE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcCCE.Visible = False
End Sub

Private Sub lbcCCE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcCCE, y)
    If (llRow < lbcCCE.ListCount) And (lbcCCE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcCCE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If ilCode = tgCurrAudioCCE(ilLoop).iCode Then
                lbcCCE.ToolTipText = Trim$(tgCurrAudioCCE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub grdAudio_Click()
    If grdAudio.Col >= grdAudio.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAudio_EnterCell()
    mSetShow
End Sub

Private Sub grdAudio_GotFocus()
    If grdAudio.Col >= grdAudio.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAudio_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdAudio.TopRow
    grdAudio.Redraw = False
End Sub

Private Sub grdAudio_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdAudio.RowHeight(0) Then
        mSortCol grdAudio.Col
        Exit Sub
    End If
    If (y > grdAudio.RowHeight(0)) And (y < grdAudio.RowHeight(0) + grdAudio.RowHeight(1)) Then
        mSortCol grdAudio.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdAudio, x, y)
    If Not ilFound Then
        grdAudio.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAudio.Col >= grdAudio.Cols - 1 Then
        grdAudio.Redraw = True
        Exit Sub
    End If
    If grdAudio.Col = PRIAUDIOTYPEINDEX Then
        grdAudio.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    lmTopRow = grdAudio.TopRow
    DoEvents
    llRow = grdAudio.Row
    If grdAudio.TextMatrix(llRow, PRIAUDIOINDEX) = "" Then
        grdAudio.Redraw = False
        Do
            llRow = llRow - 1
        Loop While (grdAudio.TextMatrix(llRow, PRIAUDIOINDEX) = "") And (llRow > grdAudio.FixedRows - 1)
        grdAudio.Row = llRow + 1
        grdAudio.Col = PRIAUDIOINDEX
        grdAudio.Redraw = True
    End If
    grdAudio.Redraw = True
    If gColOk(grdAudio, grdAudio.Row, grdAudio.Col) Then
        mEnableBox
    Else
        Beep
        pbcClickFocus.SetFocus
    End If
    mEnableBox
End Sub

Private Sub grdAudio_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdAudio.Redraw = False Then
        grdAudio.Redraw = True
        If lmTopRow < grdAudio.FixedRows Then
            grdAudio.TopRow = grdAudio.FixedRows
        Else
            grdAudio.TopRow = lmTopRow
        End If
        grdAudio.Refresh
        grdAudio.Redraw = False
    End If
    If (imShowGridBox) And (grdAudio.Row >= grdAudio.FixedRows) And (grdAudio.Col >= 0) And (grdAudio.Col < grdAudio.Cols - 1) Then
        If grdAudio.RowIsVisible(grdAudio.Row) Then
            'edcGrid.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 30, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 30
            pbcArrow.Move grdAudio.Left - pbcArrow.Width - 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + (grdAudio.RowHeight(grdAudio.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            pbcArrow.Visible = False
            edcGrid.Visible = False
            lbcANE(0).Visible = False
            lbcANE(0).Visible = False
            lbcANE(1).Visible = False
            lbcANE(1).Visible = False
            lbcCCE.Visible = False
            cmcDropDown.Visible = False
            edcDropdown.Visible = False
            pbcState.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer
    
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Or pbcState.Visible Then
        If Not mBranch() Then
            mEnableBox
            Exit Sub
        End If
        mSetShow
        Do
            ilPrev = False
            If grdAudio.Col = PRIAUDIOINDEX Then
                If grdAudio.Row > grdAudio.FixedRows Then
                    lmTopRow = -1
                    grdAudio.Row = grdAudio.Row - 1
                    If Not grdAudio.RowIsVisible(grdAudio.Row) Then
                        grdAudio.TopRow = grdAudio.TopRow - 1
                    End If
                    grdAudio.Col = STATEINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                If grdAudio.Col = PRICTRLINDEX Then
                    grdAudio.Col = grdAudio.Col - 2
                Else
                    grdAudio.Col = grdAudio.Col - 1
                End If
                If gColOk(grdAudio, grdAudio.Row, grdAudio.Col) Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdAudio.TopRow = grdAudio.FixedRows
        grdAudio.Col = PRIAUDIOINDEX
        grdAudio.Row = grdAudio.FixedRows
        If gColOk(grdAudio, grdAudio.Row, grdAudio.Col) Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If smState <> "Active" Then
            imFieldChgd = True
        End If
        smState = "Active"
        pbcState_Paint
        grdAudio.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdAudio.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdAudio.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdAudio.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdAudio.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdAudio.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = 30  'fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    pbcState.Print smState
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If Not mBranch() Then
        mEnableBox
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Or pbcState.Visible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdAudio.Col = STATEINDEX Then
                llRow = grdAudio.Rows
                Do
                    llRow = llRow - 1
                Loop While grdAudio.TextMatrix(llRow, PRIAUDIOINDEX) = ""
                llRow = llRow + 1
                If (grdAudio.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdAudio.Row = grdAudio.Row + 1
                    If Not grdAudio.RowIsVisible(grdAudio.Row) Then
                        imIgnoreScroll = True
                        grdAudio.TopRow = grdAudio.TopRow + 1
                    End If
                    grdAudio.Col = PRIAUDIOINDEX
                    'grdAudio.TextMatrix(grdAudio.Row, CODEINDEX) = 0
                    If Trim$(grdAudio.TextMatrix(grdAudio.Row, PRIAUDIOINDEX)) <> "" Then
                        If gColOk(grdAudio, grdAudio.Row, grdAudio.Col) Then
                            mEnableBox
                        Else
                            cmcCancel.SetFocus
                        End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdAudio.Left - pbcArrow.Width - 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + (grdAudio.RowHeight(grdAudio.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdAudio.TextMatrix(llEnableRow, PRIAUDIOINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdAudio.Row + 1 >= grdAudio.Rows Then
                            grdAudio.AddItem ""
                        End If
                        grdAudio.Row = grdAudio.Row + 1
                        grdAudio.Col = PRIAUDIOTYPEINDEX
                        grdAudio.CellBackColor = LIGHTYELLOW
                        If Not grdAudio.RowIsVisible(grdAudio.Row) Then
                            imIgnoreScroll = True
                            grdAudio.TopRow = grdAudio.TopRow + 1
                        End If
                        grdAudio.Col = PRIAUDIOINDEX
                        grdAudio.TextMatrix(grdAudio.Row, CODEINDEX) = 0
                        grdAudio.TextMatrix(grdAudio.Row, USEDFLAGINDEX) = "N"
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdAudio.Left - pbcArrow.Width - 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + (grdAudio.RowHeight(grdAudio.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                If grdAudio.Col = PRIAUDIOINDEX Then
                    grdAudio.Col = grdAudio.Col + 2
                Else
                    grdAudio.Col = grdAudio.Col + 1
                End If
                If gColOk(grdAudio, grdAudio.Row, grdAudio.Col) Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdAudio.TopRow = grdAudio.FixedRows
        grdAudio.Col = PRIAUDIOINDEX
        grdAudio.Row = grdAudio.FixedRows
        If gColOk(grdAudio, grdAudio.Row, grdAudio.Col) Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdAudio.TopRow
    llRow = grdAudio.Row
    slMsg = "Insert above " & Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdAudio.Redraw = False
    grdAudio.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdAudio.Row = llRow
    grdAudio.Col = PRIAUDIOTYPEINDEX
    grdAudio.CellBackColor = LIGHTYELLOW
    grdAudio.Redraw = False
    grdAudio.TopRow = llTRow
    grdAudio.Redraw = True
    DoEvents
    grdAudio.Col = PRIAUDIOINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdAudio.TopRow
    llRow = grdAudio.Row
    If (Val(grdAudio.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdAudio.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdAudio.Redraw = False
    If (Val(grdAudio.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdAudio.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdAudio.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdAudio.AddItem ""
    grdAudio.Redraw = False
    grdAudio.TopRow = llTRow
    grdAudio.Redraw = True
    DoEvents
    grdAudio.Col = PRIAUDIOINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function


Private Sub mPopANE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrAudio-mPopulate Audio Names", tgCurrANE())
'    lbcANE(0).Clear
    lbcANE(1).Clear
    For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
'        lbcANE(0).AddItem Trim$(tgCurrANE(ilLoop).sName)
'        lbcANE(0).ItemData(lbcANE(0).NewIndex) = tgCurrANE(ilLoop).iCode
        lbcANE(1).AddItem Trim$(tgCurrANE(ilLoop).sName)
        lbcANE(1).ItemData(lbcANE(1).NewIndex) = tgCurrANE(ilLoop).iCode
    Next ilLoop
    lbcANE(1).AddItem "[None]", 0
    lbcANE(1).ItemData(lbcANE(1).NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
'        lbcANE(0).AddItem "[New]", 0
'        lbcANE(0).ItemData(lbcANE(0).NewIndex) = 0
        lbcANE(1).AddItem "[New]", 0
        lbcANE(1).ItemData(lbcANE(1).NewIndex) = 0
    Else
'        lbcANE(0).AddItem "[View]", 0
'        lbcANE(0).ItemData(lbcANE(0).NewIndex) = 0
        lbcANE(1).AddItem "[View]", 0
        lbcANE(1).ItemData(lbcANE(1).NewIndex) = 0
    End If
End Sub

Private Sub mPopCCE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrControlChar-mPopulate Audio Controls", tgCurrAudioCCE())
    lbcCCE.Clear
    For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        lbcCCE.AddItem Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
        lbcCCE.ItemData(lbcCCE.NewIndex) = tgCurrAudioCCE(ilLoop).iCode
    Next ilLoop
    lbcCCE.AddItem "[None]", 0
    lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcCCE.AddItem "[New]", 0
        lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    Else
        lbcCCE.AddItem "[View]", 0
        lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case grdAudio.Col
        Case PRIAUDIOINDEX
            lbcANE(0).Visible = False
        Case PRICTRLINDEX, BKUPCTRLINDEX, PROTCTRLINDEX
            lbcCCE.Visible = False
        Case BKUPAUDIOINDEX, PROTAUDIOINDEX
            lbcANE(1).Visible = False
    End Select
End Sub

Private Function mBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    
    mBranch = True
    If (lmEnableRow >= grdAudio.FixedRows) And (lmEnableRow < grdAudio.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdAudio.TextMatrix(lmEnableRow, lmEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEnableCol
                Case PRIAUDIOINDEX
                    'llRow = SendMessageByString(lbcANE(0).hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcANE(0), slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrAudioName.Show vbModal
                        sgCurrANEStamp = ""
                        mPopANE
                        mPopUnusedANE lmEnableRow
                        lbcANE(0).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcANE(0), CLng(grdAudio.Height / 2)
                        If lbcANE(0).Top + lbcANE(0).Height > cmcCancel.Top Then
                            lbcANE(0).Top = edcDropdown.Top - lbcANE(0).Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcANE(0).hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcANE(0), slStr)
                            If llRow > 0 Then
                                lbcANE(0).ListIndex = llRow
                                edcDropdown.text = lbcANE(0).List(lbcANE(0).ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case PRICTRLINDEX, BKUPCTRLINDEX, PROTCTRLINDEX
                    'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcCCE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrControlChar.Show vbModal
                        sgCurrAudioCCEStamp = ""
                        mPopCCE
                        lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
                        If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                            lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcCCE, slStr)
                            If llRow > 0 Then
                                lbcCCE.ListIndex = llRow
                                edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case BKUPAUDIOINDEX, PROTAUDIOINDEX
                    'llRow = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcANE(1), slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrAudioName.Show vbModal
                        sgCurrANEStamp = ""
                        mPopANE
                        lbcANE(1).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcANE(1), CLng(grdAudio.Height / 2)
                        If lbcANE(1).Top + lbcANE(1).Height > cmcCancel.Top Then
                            lbcANE(1).Top = edcDropdown.Top - lbcANE(1).Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcANE(1).hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcANE(1), slStr)
                            If llRow > 0 Then
                                lbcANE(1).ListIndex = llRow
                                edcDropdown.text = lbcANE(1).List(lbcANE(1).ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case STATEINDEX
            End Select
        End If
    End If
    imDoubleClickName = False
End Function

Private Function mCompare(tlNew As ASE, tlOld As ASE) As Integer
    If StrComp(tlNew.sDescription, tlOld.sDescription, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iPriAneCode <> tlOld.iPriAneCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iPriCceCode <> tlOld.iPriCceCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iBkupAneCode <> tlOld.iBkupAneCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iBkupCceCode <> tlOld.iBkupCceCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iProtAneCode <> tlOld.iProtAneCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iProtCceCode <> tlOld.iProtCceCode) Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
End Function

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrASE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
            slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdAudio.Row = llRow
                    Do While Not grdAudio.RowIsVisible(grdAudio.Row)
                        imIgnoreScroll = True
                        grdAudio.TopRow = grdAudio.TopRow + 1
                    Loop
                    grdAudio.Col = PRIAUDIOINDEX
                    mEnableBox
                    Exit Sub
                End If
            End If
        Next llRow
    End If
    If (Not ilCreateNew) Or (Not cmcDone.Enabled) Then
        Exit Sub
    End If
    'Find first blank row
    For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
        slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
        If (slStr = "") Then
            grdAudio.Row = llRow
            Do While Not grdAudio.RowIsVisible(grdAudio.Row)
                imIgnoreScroll = True
                grdAudio.TopRow = grdAudio.TopRow + 1
            Loop
            grdAudio.Col = PRIAUDIOINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdAudio.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mPopUnusedANE(llCurrRow As Long)
    Dim ilANE As Integer
    Dim llRow As Long
    Dim ilFound As Integer
    Dim slStr As String
    Dim slANE As String
    
    lbcANE(0).Clear
    grdAudio.Redraw = False
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        slANE = Trim$(tgCurrANE(ilANE).sName)
        ilFound = False
        For llRow = grdAudio.FixedRows To grdAudio.Rows - 1 Step 1
            slStr = Trim$(grdAudio.TextMatrix(llRow, PRIAUDIOINDEX))
            If StrComp(slStr, slANE, vbTextCompare) = 0 Then
                If llRow = llCurrRow Then
                    Exit For
                End If
                ilFound = True
                Exit For
            End If
        Next llRow
        If Not ilFound Then
            lbcANE(0).AddItem Trim$(tgCurrANE(ilANE).sName)
            lbcANE(0).ItemData(lbcANE(0).NewIndex) = tgCurrANE(ilANE).iCode
        End If
    Next ilANE
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcANE(0).AddItem "[New]", 0
        lbcANE(0).ItemData(lbcANE(0).NewIndex) = 0
    Else
        lbcANE(0).AddItem "[View]", 0
        lbcANE(0).ItemData(lbcANE(0).NewIndex) = 0
    End If
    grdAudio.Redraw = True
End Sub


Private Sub mSetDefaults()

End Sub

Private Sub mSetFocus()
    Select Case grdAudio.Col
        Case PRIAUDIOINDEX  'Call Letters
            edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcANE(0).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcANE(0), CLng(grdAudio.Height / 2)
            If lbcANE(0).Top + lbcANE(0).Height > cmcCancel.Top Then
                lbcANE(0).Top = edcDropdown.Top - lbcANE(0).Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcANE(0).Visible = True
            edcDropdown.SetFocus
        Case PRICTRLINDEX
            edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col - 2) + grdAudio.ColWidth(grdAudio.Col - 2) + grdAudio.ColPos(grdAudio.Col - 1) + grdAudio.ColWidth(grdAudio.Col - 1) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
            If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcCCE.Visible = True
            edcDropdown.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case BKUPAUDIOINDEX
            edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcANE(1).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcANE(1), CLng(grdAudio.Height / 2)
            If lbcANE(1).Top + lbcANE(1).Height > cmcCancel.Top Then
                lbcANE(1).Top = edcDropdown.Top - lbcANE(1).Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcANE(1).Visible = True
            edcDropdown.SetFocus
        Case BKUPCTRLINDEX
            'edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col - 1) + grdAudio.ColWidth(grdAudio.Col - 1) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
            If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcCCE.Visible = True
            edcDropdown.SetFocus
        Case PROTAUDIOINDEX
            edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - cmcDropDown.Width - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcANE(1).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcANE(1), CLng(grdAudio.Height / 2)
            If lbcANE(1).Top + lbcANE(1).Height > cmcCancel.Top Then
                lbcANE(1).Top = edcDropdown.Top - lbcANE(1).Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcANE(1).Visible = True
            edcDropdown.SetFocus
        Case PROTCTRLINDEX
            edcDropdown.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col - 1) + grdAudio.ColWidth(grdAudio.Col - 1) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcCCE, CLng(grdAudio.Height / 2)
            If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcCCE.Visible = True
            edcDropdown.SetFocus
        Case STATEINDEX
            pbcState.Move grdAudio.Left + grdAudio.ColPos(grdAudio.Col) + 30, grdAudio.Top + grdAudio.RowPos(grdAudio.Row) + 15, grdAudio.ColWidth(grdAudio.Col) - 30, grdAudio.RowHeight(grdAudio.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub

Private Sub mPopATE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudio-mPopulate Audio Source", tgCurrATE())
End Sub

