VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrBus 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrBus.frx":0000
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
      Left            =   11745
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.CommandButton cmcNone 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "[None]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4875
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox pbcDefine 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   6555
      ScaleHeight     =   165
      ScaleWidth      =   1035
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmcDefine 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "[New]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4875
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox lbcASE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrBus.frx":030A
      Left            =   7470
      List            =   "EngrBus.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCCE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrBus.frx":030E
      Left            =   6090
      List            =   "EngrBus.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   11100
      Top             =   6165
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2685
      TabIndex        =   12
      Top             =   2430
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
      Left            =   3630
      Picture         =   "EngrBus.frx":0312
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcBGE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrBus.frx":040C
      Left            =   4170
      List            =   "EngrBus.frx":040E
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
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
      TabIndex        =   14
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
      Left            =   2475
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   150
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   6870
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
      Picture         =   "EngrBus.frx":0410
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
      Left            =   6960
      TabIndex        =   18
      Top             =   6675
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10530
      Top             =   6465
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
      Left            =   5190
      TabIndex        =   17
      Top             =   6675
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3405
      TabIndex        =   16
      Top             =   6675
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBus 
      Height          =   5970
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   10530
      _Version        =   393216
      Rows            =   3
      Cols            =   9
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
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
      TabIndex        =   20
      Top             =   75
      Width           =   795
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2235
      Picture         =   "EngrBus.frx":071A
      Top             =   6585
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Bus Definition"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1350
      Picture         =   "EngrBus.frx":0A24
      Top             =   6585
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8970
      Picture         =   "EngrBus.frx":12EE
      Top             =   6585
      Width           =   480
   End
End
Attribute VB_Name = "EngrBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrBus - enters affiliate representative information
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
Private ilBdeCode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private smCurrBSEStamp As String
Private tmCurrBSE() As BSE
Private smBusGroups() As String

Private imDoubleClickName As Integer
Private imLbcMouseDown As Integer


Private tmBDE As BDE
Private tmBSE As BSE

Private imDeleteCodes() As Integer


'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer

Const NAMEINDEX = 0
Const DESCRIPTIONINDEX = 1
Const BUSCTRLINDEX = 2
Const CHANNELINDEX = 3
Const BUSGROUPINDEX = 4
Const AUDIOINDEX = 5
Const STATEINDEX = 6
Const CODEINDEX = 7
Const USEDFLAGINDEX = 8


Private Sub mPopASE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilANE As Integer

    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrBus-mPop Audio Source", tgCurrASE())
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrBus-mPop Audio Source", tgCurrANE())
    lbcASE.Clear
    For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tgCurrASE(ilLoop).iPriAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tgCurrASE(ilLoop).iPriAneCode, tgCurrANE())
            If ilANE <> -1 Then
                lbcASE.AddItem Trim$(tgCurrANE(ilANE).sName)
                lbcASE.ItemData(lbcASE.NewIndex) = tgCurrASE(ilLoop).iCode
        '        Exit For
            End If
        'Next ilANE
    Next ilLoop
    lbcASE.AddItem "[None]", 0
    lbcASE.ItemData(lbcASE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcASE.AddItem "[New]", 0
        lbcASE.ItemData(lbcASE.NewIndex) = 0
    Else
        lbcASE.AddItem "[View]", 0
        lbcASE.ItemData(lbcASE.NewIndex) = 0
    End If
End Sub
Private Sub mPopBGE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrBus-mPopulate Bus Groups", tgCurrBGE())
    lbcBGE.Clear
    For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
        lbcBGE.AddItem Trim$(tgCurrBGE(ilLoop).sName)
        lbcBGE.ItemData(lbcBGE.NewIndex) = tgCurrBGE(ilLoop).iCode
    Next ilLoop
'    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
'        lbcBGE.AddItem "[New]", 0
'        lbcBGE.ItemData(lbcBGE.NewIndex) = 0
'    Else
'        lbcBGE.AddItem "[View]", 0
'        lbcBGE.ItemData(lbcBGE.NewIndex) = 0
'    End If
End Sub
Private Function mNameOk() As Integer
    Dim ilError As Integer
    Dim llRow As Long
    Dim llTestRow As Long
    Dim slStr As String
    Dim slTestStr As String
    
    grdBus.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdBus.FixedRows To grdBus.Rows - 1 Step 1
        slStr = Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdBus.Rows - 1 Step 1
                slTestStr = Trim$(grdBus.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdBus.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdBus.Row = llRow
                        grdBus.Col = NAMEINDEX
                        grdBus.CellForeColor = vbRed
                    Else
                        grdBus.Row = llTestRow
                        grdBus.Col = NAMEINDEX
                        grdBus.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdBus.Redraw = True
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
    gGrid_SortByCol grdBus, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilFieldChgd As Integer
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(BUSLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdBus.Row >= grdBus.FixedRows) And (grdBus.Row < grdBus.Rows) And (grdBus.Col >= 0) And (grdBus.Col < grdBus.Cols - 1) Then
        lmEnableRow = grdBus.Row
        lmEnableCol = grdBus.Col
        sgReturnCallName = grdBus.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdBus.Left - pbcArrow.Width - 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + (grdBus.RowHeight(grdBus.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdBus.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdBus.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdBus.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdBus.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
                'edcGrid.MaxLength = Len(tmBDE.sName)
                edcGrid.MaxLength = gGetAllowedChars("BUSNAME", Len(tmBDE.sName))
                edcGrid.text = grdBus.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
                edcGrid.MaxLength = Len(tmBDE.sDescription)
                edcGrid.text = grdBus.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case BUSCTRLINDEX
                edcDropdown.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - cmcDropDown.Width - 30, grdBus.RowHeight(grdBus.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcCCE, CLng(grdBus.Height / 2)
                If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                    lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
                End If
                slStr = grdBus.text
                'ilIndex = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcCCE, slStr)
                If ilIndex >= 0 Then
                    lbcCCE.ListIndex = ilIndex
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcCCE.ListCount <= 0 Then
                        lbcCCE.ListIndex = -1
                        edcDropdown.text = ""
                    Else
                        lbcCCE.ListIndex = 1
                        edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                    End If
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcCCE.Visible = True
                edcDropdown.SetFocus
            Case CHANNELINDEX  'Date
                edcGrid.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
                edcGrid.MaxLength = Len(tmBDE.sChannel)
                edcGrid.text = grdBus.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case BUSGROUPINDEX
                'cmcDefine.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
                pbcDefine.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
                cmcDefine.Move pbcDefine.Left, pbcDefine.Top + pbcDefine.Height, pbcDefine.Width, pbcDefine.Height
                cmcNone.Move pbcDefine.Left, cmcDefine.Top + cmcDefine.Height, pbcDefine.Width, pbcDefine.Height
                'lbcBGE.Move cmcDefine.Left, cmcDefine.Top + cmcDefine.Height, cmcDefine.Width
                lbcBGE.Move pbcDefine.Left, cmcNone.Top + cmcNone.Height, pbcDefine.Width
                gSetListBoxHeight lbcBGE, CLng(grdBus.Height / 2)
                If lbcBGE.Top + lbcBGE.Height > cmcCancel.Top Then
                    lbcBGE.Top = pbcDefine.Top - lbcBGE.Height
                End If
                If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSGROUPLIST) = 2) Then
                    cmcDefine.Caption = "[New]"
                Else
                    cmcDefine.Caption = "[View]"
                End If
                slStr = grdBus.TextMatrix(grdBus.Row, BUSGROUPINDEX)
                gParseCDFields slStr, False, smBusGroups()
                ilFieldChgd = imFieldChgd
                lbcBGE.ListIndex = -1
                For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
                    lbcBGE.Selected(ilLoop) = False
                Next ilLoop
                For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
                    slStr = Trim$(smBusGroups(ilLoop))
                    If slStr <> "" Then
                        'llRow = SendMessageByString(lbcBGE.hwnd, LB_FINDSTRING, -1, slStr)
                        llRow = gListBoxFind(lbcBGE, slStr)
                        If llRow >= 0 Then
                            lbcBGE.Selected(llRow) = True
                        End If
                    End If
                Next ilLoop
                imFieldChgd = ilFieldChgd
                mSetCommands
                cmcDefine.Visible = True
                cmcNone.Visible = True
                pbcDefine.Visible = True
                lbcBGE.Visible = True
                lbcBGE.SetFocus
            Case AUDIOINDEX
                edcDropdown.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - cmcDropDown.Width - 30, grdBus.RowHeight(grdBus.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcASE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcASE, CLng(grdBus.Height / 2)
                If lbcASE.Top + lbcASE.Height > cmcCancel.Top Then
                    lbcASE.Top = edcDropdown.Top - lbcASE.Height
                End If
                slStr = grdBus.text
                'ilIndex = SendMessageByString(lbcASE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcASE, slStr)
                If ilIndex >= 0 Then
                    lbcASE.ListIndex = ilIndex
                    edcDropdown.text = lbcASE.List(lbcASE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcASE.ListCount <= 0 Then
                        lbcASE.ListIndex = -1
                        edcDropdown.text = ""
                    Else
                        lbcASE.ListIndex = 1
                        edcDropdown.text = lbcASE.List(lbcASE.ListIndex)
                    End If
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcASE.Visible = True
                edcDropdown.SetFocus
            Case STATEINDEX
                pbcState.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
                smState = grdBus.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdBus.text
    End If
End Sub
Private Sub mSetShow()
    Dim llRow As Long
    Dim slStr As String
    
    tmcClick.Enabled = False
    If (lmEnableRow >= grdBus.FixedRows) And (lmEnableRow < grdBus.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = grdBus.TextMatrix(lmEnableRow, lmEnableCol)
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
            Case BUSCTRLINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcCCE, slStr)
                If (llRow <= 0) Then
                    grdBus.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case CHANNELINDEX
            Case BUSGROUPINDEX
            Case AUDIOINDEX
                If (Trim$(grdBus.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdBus.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdBus.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdBus.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdBus.TextMatrix(lmEnableRow, NAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    cmcNone.Visible = False
    cmcDefine.Visible = False
    pbcDefine.Visible = False
    lbcBGE.Visible = False
    lbcCCE.Visible = False
    lbcASE.Visible = False
    cmcDropDown.Visible = False
    edcDropdown.Visible = False
    pbcArrow.Visible = False
    edcGrid.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llLbc As Long
    Dim llRow As Long
    
    grdBus.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdBus.FixedRows To grdBus.Rows - 1 Step 1
        slStr = Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdBus.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdBus.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdBus.Row = llRow
                grdBus.Col = NAMEINDEX
                grdBus.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdBus.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdBus.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdBus.Row = llRow
                    grdBus.Col = STATEINDEX
                    grdBus.CellForeColor = vbRed
                End If
'                slStr = Trim$(grdBus.TextMatrix(llRow, BUSGROUPINDEX))
'                If slStr = "" Then
'                    ilError = True
'                    If slStr = "" Then
'                        grdBus.TextMatrix(llRow, BUSGROUPINDEX) = "Missing"
'                    End If
'                    grdBus.Row = llRow
'                    grdBus.Col = BUSGROUPINDEX
'                    grdBus.CellForeColor = vbRed
'                End If
            End If
        End If
    Next llRow
    grdBus.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdBus
    mGridColumnWidth
    'Set Titles
    grdBus.TextMatrix(0, NAMEINDEX) = "Bus Name"
    grdBus.TextMatrix(0, BUSCTRLINDEX) = "Control"
    grdBus.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdBus.TextMatrix(0, CHANNELINDEX) = "Channel Name"
    grdBus.TextMatrix(0, BUSGROUPINDEX) = "Bus Group"
    grdBus.TextMatrix(0, AUDIOINDEX) = "Cmml Audio"
    grdBus.TextMatrix(0, STATEINDEX) = "State"
    grdBus.Row = 1
    For ilCol = 0 To grdBus.Cols - 1 Step 1
        grdBus.Col = ilCol
        grdBus.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdBus.Height = cmcCancel.Top - grdBus.Top - 120    '8 * grdBus.RowHeight(0) + 30
    gGrid_IntegralHeight grdBus
    gGrid_Clear grdBus, True
    grdBus.Row = grdBus.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdBus.Width = EngrBus.Width - 2 * grdBus.Left
    grdBus.ColWidth(CODEINDEX) = 0
    grdBus.ColWidth(USEDFLAGINDEX) = 0
    grdBus.ColWidth(NAMEINDEX) = grdBus.Width / 12
    If tgUsedSumEPE.sBusControl <> "Y" Then
        grdBus.ColWidth(BUSCTRLINDEX) = 0
    Else
        grdBus.ColWidth(BUSCTRLINDEX) = grdBus.Width / 16
    End If
    grdBus.ColWidth(CHANNELINDEX) = grdBus.Width / 6
    grdBus.ColWidth(BUSGROUPINDEX) = grdBus.Width / 4
    grdBus.ColWidth(AUDIOINDEX) = grdBus.Width / 11
    grdBus.ColWidth(STATEINDEX) = grdBus.Width / 18
    grdBus.ColWidth(DESCRIPTIONINDEX) = grdBus.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdBus.ColWidth(DESCRIPTIONINDEX) > grdBus.ColWidth(ilCol) Then
                grdBus.ColWidth(DESCRIPTIONINDEX) = grdBus.ColWidth(DESCRIPTIONINDEX) - grdBus.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdBus, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim ilASE As Integer
    Dim ilCCE As Integer
    Dim ilANE As Integer
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdBus.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdBus.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmBDE.iCode = Val(grdBus.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmBDE.sName = ""
    Else
        tmBDE.sName = slStr
    End If
    tmBDE.sDescription = grdBus.TextMatrix(llRow, DESCRIPTIONINDEX)
    tmBDE.iCceCode = 0
    slStr = Trim$(grdBus.TextMatrix(llRow, BUSCTRLINDEX))
    For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
        If StrComp(Trim$(tgCurrBusCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
            tmBDE.iCceCode = tgCurrBusCCE(ilCCE).iCode
            Exit For
        End If
    Next ilCCE
    tmBDE.sChannel = grdBus.TextMatrix(llRow, CHANNELINDEX)
    slStr = grdBus.TextMatrix(llRow, BUSGROUPINDEX)
    gParseCDFields slStr, False, smBusGroups()
    tmBDE.iAseCode = 0
    slStr = Trim$(grdBus.TextMatrix(llRow, AUDIOINDEX))
    For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
            If ilANE <> -1 Then
                If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                    tmBDE.iAseCode = tgCurrASE(ilASE).iCode
        '            Exit For
                End If
            End If
        'Next ilANE
        If tmBDE.iAseCode <> 0 Then
            Exit For
        End If
    Next ilASE
    If grdBus.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmBDE.sState = "D"
    Else
        tmBDE.sState = "A"
    End If
    If tmBDE.iCode <= 0 Then
        tmBDE.sUsedFlag = "N"
    Else
        tmBDE.sUsedFlag = grdBus.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmBDE.iVersion = 0
    tmBDE.iOrigBdeCode = tmBDE.iCode
    tmBDE.sCurrent = "Y"
    'tmBDE.sEnteredDate = smNowDate
    'tmBDE.sEnteredTime = smNowTime
    tmBDE.sEnteredDate = Format(Now, sgShowDateForm)
    tmBDE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmBDE.iUieCode = tgUIE.iCode
    tmBDE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilASE As Integer
    Dim ilCCE As Integer
    Dim slStr As String
    Dim ilBSE As Integer
    Dim ilRet As Integer
    Dim ilBGE As Integer
    Dim ilANE As Integer
    
    'gGrid_Clear grdBus, True
    llRow = grdBus.FixedRows
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        If llRow + 1 > grdBus.Rows Then
            grdBus.AddItem ""
        End If
        grdBus.Row = llRow
        grdBus.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrBDE(ilLoop).sName)
        grdBus.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrBDE(ilLoop).sDescription)
        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            If tgCurrBDE(ilLoop).iCceCode = tgCurrBusCCE(ilCCE).iCode Then
                grdBus.TextMatrix(llRow, BUSCTRLINDEX) = Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
                Exit For
            End If
        Next ilCCE
        grdBus.TextMatrix(llRow, CHANNELINDEX) = Trim$(tgCurrBDE(ilLoop).sChannel)
        slStr = ""
        Erase tmCurrBSE
        ilRet = gGetRecs_BSE_BusSelGroup("B", smCurrBSEStamp, tgCurrBDE(ilLoop).iCode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
        For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
            For ilBGE = 0 To UBound(tgCurrBGE) - 1 Step 1
                If tmCurrBSE(ilBSE).iBgeCode = tgCurrBGE(ilBGE).iCode Then
                    slStr = slStr & Trim$(tgCurrBGE(ilBGE).sName) & ","
                    Exit For
                End If
            Next ilBGE
        Next ilBSE
        If slStr <> "" Then
            slStr = Left$(slStr, Len(slStr) - 1)
        End If
        grdBus.TextMatrix(llRow, BUSGROUPINDEX) = slStr
        'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        '    If tgCurrBDE(ilLoop).iAseCode = tgCurrASE(ilASE).iCode Then
            ilASE = gBinarySearchASE(tgCurrBDE(ilLoop).iAseCode, tgCurrASE())
            If ilASE <> -1 Then
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        grdBus.TextMatrix(llRow, AUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                '        Exit For
                    End If
                'Next ilANE
        '        Exit For
            End If
        'Next ilASE
        If tgCurrBDE(ilLoop).sState = "A" Then
            grdBus.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdBus.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdBus.TextMatrix(llRow, CODEINDEX) = tgCurrBDE(ilLoop).iCode
        grdBus.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrBDE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdBus.Rows Then
        grdBus.AddItem ""
    End If
    grdBus.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrBus-mPopulate", tgCurrBDE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilBGE As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim slStr As String
    Dim tlBDE As BDE
    
    gSetMousePointer grdBus, grdBus, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdBus, grdBus, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdBus, grdBus, vbDefault
        MsgBox "Duplicated names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdBus.Redraw = False
    For llRow = grdBus.FixedRows To grdBus.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmBDE.sName) <> "" Then
            ilBdeCode = tmBDE.iCode
            If tmBDE.iCode > 0 Then
                ilRet = gGetRec_BDE_BusDefinition(ilBdeCode, "Bus Definition-mSave: Get BDE", tlBDE)
                If ilRet Then
                    slStr = grdBus.TextMatrix(llRow, BUSGROUPINDEX)
                    gParseCDFields slStr, False, smBusGroups()
                    If mCompare(tmBDE, tlBDE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmBDE.iVersion = tlBDE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmBDE.iCode <= 0 Then
                    ilRet = gPutInsert_BDE_BusDefinition(0, tmBDE, "Bus Definition-mSave: Insert BDE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_BDE_BusDefinition(1, tmBDE, "Bus Definition-mSave: Update BDE")
                    ilRet = gPutDelete_BDE_BusDefinition(tmBDE.iCode, "Bus Definition-mSave: Delete BDE")
                    ilRet = gPutInsert_BDE_BusDefinition(1, tmBDE, "Bus Definition-mSave: Insert BDE")
                End If
                ilRet = gPutUpdate_ASE_UsedFlag(tmBDE.iAseCode, tgCurrASE())
                ilRet = gPutUpdate_CCE_UsedFlag(tmBDE.iCceCode, tgCurrBusCCE())
                For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
                    If Trim$(smBusGroups(ilLoop)) <> "" Then
                        tmBSE.iCode = 0
                        tmBSE.iBdeCode = tmBDE.iCode
                        tmBSE.iBgeCode = -1
                        For ilBGE = 0 To UBound(tgCurrBGE) Step 1
                            If StrComp(Trim$(tgCurrBGE(ilBGE).sName), Trim$(smBusGroups(ilLoop)), vbTextCompare) = 0 Then
                                tmBSE.iBgeCode = tgCurrBGE(ilBGE).iCode
                                Exit For
                            End If
                        Next ilBGE
                        tmBSE.sUnused = ""
                        If tmBSE.iBgeCode > 0 Then
                            ilRet = gPutInsert_BSE_BusSelGroup(tmBSE, "Bus Definition-mSave: Insert BSE")
                        End If
                        ilRet = gPutUpdate_BGE_UsedFlag(tmBSE.iBgeCode, tgCurrBGE())
                    End If
                Next ilLoop
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_BDE_BusDefinition(imDeleteCodes(ilLoop), "EngrBus- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdBus.Redraw = True
    sgCurrBDEStamp = ""
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrBus-mPopulate", tgCurrBDE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrBus
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcDefine_Click()
    Dim ilRet As Integer
    ilRet = mBranch()
    cmcDefine.SetFocus
End Sub

Private Sub cmcDone_Click()
    
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrBus
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdBus, grdBus, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdBus, grdBus, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdBus, grdBus, vbDefault
    Unload EngrBus
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdBus.Col
        Case BUSCTRLINDEX
            lbcCCE.Visible = Not lbcCCE.Visible
        Case AUDIOINDEX
            lbcASE.Visible = Not lbcASE.Visible
    End Select
End Sub

Private Sub cmcNone_Click()
    Dim llRg As Long
    Dim llRet As Long
    Dim ilValue As Integer
    
    ilValue = False
    If lbcBGE.ListCount > 0 Then         'at least 1 entries exists in check box
        llRg = CLng(lbcBGE.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcBGE.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    If Trim$(grdBus.TextMatrix(grdBus.Row, BUSGROUPINDEX)) <> "" Then
        grdBus.CellForeColor = vbBlack
        grdBus.TextMatrix(grdBus.Row, BUSGROUPINDEX) = ""
        imFieldChgd = True
        mSetCommands
    End If
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim llTopRow As Long
    
    If imFieldChgd = True Then
        gSetMousePointer grdBus, grdBus, vbHourglass
        llTopRow = grdBus.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdBus, grdBus, vbDefault
            Exit Sub
        End If
        grdBus.Redraw = False
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
        grdBus.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdBus, grdBus, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_Search(grdBus, NAMEINDEX, slStr)
    If llRow >= 0 Then
        mEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mSetShow
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdBus.Col
        Case BUSCTRLINDEX
            'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcCCE, slStr)
            If llRow >= 0 Then
                lbcCCE.ListIndex = llRow
                edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case AUDIOINDEX
            'llRow = SendMessageByString(lbcASE.hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcASE, slStr)
            If llRow >= 0 Then
                lbcASE.ListIndex = llRow
                edcDropdown.text = lbcASE.List(lbcASE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If StrComp(grdBus.text, edcDropdown.text, vbTextCompare) <> 0 Then
        imFieldChgd = True
    End If
    If StrComp(Trim$(edcDropdown.text), "[None]", vbTextCompare) <> 0 Then
        grdBus.text = edcDropdown.text
    Else
        grdBus.text = ""
    End If
    grdBus.CellForeColor = vbBlack
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
        Select Case grdBus.Col
            Case BUSCTRLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE, True
            Case AUDIOINDEX
                gProcessArrowKey Shift, KeyCode, lbcASE, True
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
    
    Select Case grdBus.Col
        Case NAMEINDEX
            If grdBus.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdBus.text = edcGrid.text
            grdBus.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdBus.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdBus.text = edcGrid.text
            grdBus.CellForeColor = vbBlack
        Case CHANNELINDEX
            If grdBus.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdBus.text = edcGrid.text
            grdBus.CellForeColor = vbBlack
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
    gSetFonts EngrBus
    gCenterFormModal EngrBus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdBus.FixedRows) And (lmEnableRow < grdBus.Rows) Then
            If (lmEnableCol >= grdBus.FixedCols) And (lmEnableCol < grdBus.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdBus.text = smESCValue
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
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdBus.Height = cmcCancel.Top - grdBus.Top - 120    '8 * grdBus.RowHeight(0) + 30
    gGrid_IntegralHeight grdBus
    gGrid_FillWithRows grdBus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Erase smBusGroups
    Erase tmCurrBSE
    Set EngrBus = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdBus, grdBus, vbHourglass
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
    mPopASE
    mPopBGE
    mPopCCE
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    imLbcMouseDown = False
    imDoubleClickName = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdBus, grdBus, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdBus, grdBus, vbDefault
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

Private Sub grdBus_DblClick()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(BUSLIST) <> 2) Then
        Select Case grdBus.Col
            Case BUSGROUPINDEX
                igInitCallInfo = 1
                sgInitCallName = grdBus.TextMatrix(grdBus.Row, grdBus.Col)
                EngrAudioType.Show vbModal
                cmcCancel.SetFocus
        End Select
    End If
End Sub

Private Sub imcInsert_Click()
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = BUS_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub


Private Sub lbcASE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcASE.List(lbcASE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        'lbcBGE.Visible = False
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcASE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcASE.Visible = False
End Sub

Private Sub lbcASE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcASE, y)
    If (llRow < lbcASE.ListCount) And (lbcASE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcASE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
        '    If ilCode = tgCurrASE(ilLoop).iCode Then
            ilLoop = gBinarySearchASE(ilCode, tgCurrASE())
            If ilLoop <> -1 Then
                lbcASE.ToolTipText = Trim$(tgCurrASE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub lbcBGE_Click()
    Dim slStr As String
    Dim ilLoop As Integer
'    tmcClick.Enabled = False
'    edcDropdown.Text = lbcBGE.List(lbcBGE.ListIndex)
'    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
'        edcDropdown.SetFocus
'        'lbcBGE.Visible = False
'        tmcClick.Enabled = True
'    End If
    slStr = ""
    For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
        If lbcBGE.Selected(ilLoop) Then
            slStr = slStr & lbcBGE.List(ilLoop) & ","
        End If
    Next ilLoop
    If slStr <> "" Then
        slStr = Left$(slStr, Len(slStr) - 1)
    End If
    grdBus.text = slStr
    grdBus.CellForeColor = vbBlack
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub lbcBGE_DblClick()
'    tmcClick.Enabled = False
'    Sleep 300
'    DoEvents
'    edcDropdown.SetFocus
'    imDoubleClickName = True    'Double click event is followed by a mouse up event
'                                'Process the double click event in the mouse up event
'                                'to avoid the mouse up event being in next form
'    edcDropdown_MouseUp 0, 0, 0, 0
'    lbcBGE.Visible = False
End Sub

Private Sub lbcBGE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcBGE, y)
    If (llRow < lbcBGE.ListCount) And (lbcBGE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcBGE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
            If ilCode = tgCurrBGE(ilLoop).iCode Then
                lbcBGE.ToolTipText = Trim$(tgCurrBGE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcCCE_Click()
    tmcClick.Enabled = False
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
        For ilLoop = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            If ilCode = tgCurrBusCCE(ilLoop).iCode Then
                lbcCCE.ToolTipText = Trim$(tgCurrBusCCE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub grdBus_Click()
    If grdBus.Col >= grdBus.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdBus_EnterCell()
    mSetShow
End Sub

Private Sub grdBus_GotFocus()
    If grdBus.Col >= grdBus.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdBus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdBus.TopRow
    grdBus.Redraw = False
End Sub

Private Sub grdBus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdBus.RowHeight(0) Then
        mSortCol grdBus.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdBus, x, y)
    If Not ilFound Then
        grdBus.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdBus.Col >= grdBus.Cols - 1 Then
        grdBus.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdBus.TopRow
    DoEvents
    llRow = grdBus.Row
    If grdBus.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdBus.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdBus.TextMatrix(llRow, NAMEINDEX) = ""
        grdBus.Row = llRow + 1
        grdBus.Col = NAMEINDEX
        grdBus.Redraw = True
    End If
    grdBus.Redraw = True
    If gColOk(grdBus, grdBus.Row, grdBus.Col) Then
        mEnableBox
    Else
        Beep
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub grdBus_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdBus.Redraw = False Then
        grdBus.Redraw = True
        If lmTopRow < grdBus.FixedRows Then
            grdBus.TopRow = grdBus.FixedRows
        Else
            grdBus.TopRow = lmTopRow
        End If
        grdBus.Refresh
        grdBus.Redraw = False
    End If
    If (imShowGridBox) And (grdBus.Row >= grdBus.FixedRows) And (grdBus.Col >= 0) And (grdBus.Col < grdBus.Cols - 1) Then
        If grdBus.RowIsVisible(grdBus.Row) Then
            pbcArrow.Move grdBus.Left - pbcArrow.Width - 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + (grdBus.RowHeight(grdBus.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            cmcNone.Visible = False
            cmcDefine.Visible = False
            pbcDefine.Visible = False
            lbcBGE.Visible = False
            lbcCCE.Visible = False
            lbcASE.Visible = False
            cmcDropDown.Visible = False
            edcDropdown.Visible = False
            edcGrid.Visible = False
            pbcState.Visible = False
            pbcArrow.Visible = False
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

Private Sub pbcDefine_Paint()
    pbcDefine.CurrentX = 30
    pbcDefine.CurrentY = 0
    pbcDefine.Print "Multi-Select"
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
    If edcGrid.Visible Or edcDropdown.Visible Or pbcState.Visible Or lbcBGE.Visible Then
        If Not lbcBGE.Visible Then
            If Not mBranch() Then
                mEnableBox
                Exit Sub
            End If
        End If
        mSetShow
        Do
            ilPrev = False
            If grdBus.Col = NAMEINDEX Then
                If grdBus.Row > grdBus.FixedRows Then
                    lmTopRow = -1
                    grdBus.Row = grdBus.Row - 1
                    If Not grdBus.RowIsVisible(grdBus.Row) Then
                        grdBus.TopRow = grdBus.TopRow - 1
                    End If
                    grdBus.Col = STATEINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdBus.Col = grdBus.Col - 1
                If gColOk(grdBus, grdBus.Row, grdBus.Col) Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdBus.TopRow = grdBus.FixedRows
        grdBus.Col = NAMEINDEX
        grdBus.Row = grdBus.FixedRows
        If gColOk(grdBus, grdBus.Row, grdBus.Col) Then
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
        grdBus.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdBus.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdBus.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdBus.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdBus.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdBus.CellForeColor = vbBlack
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
    If edcGrid.Visible Or edcDropdown.Visible Or pbcState.Visible Or lbcBGE.Visible Then
        If Not lbcBGE.Visible Then
            If Not mBranch() Then
                mEnableBox
                Exit Sub
            End If
        End If
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdBus.Col = STATEINDEX Then
                llRow = grdBus.Rows
                Do
                    llRow = llRow - 1
                Loop While grdBus.TextMatrix(llRow, NAMEINDEX) = ""
                llRow = llRow + 1
                If (grdBus.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdBus.Row = grdBus.Row + 1
                    If Not grdBus.RowIsVisible(grdBus.Row) Then
                        imIgnoreScroll = True
                        grdBus.TopRow = grdBus.TopRow + 1
                    End If
                    grdBus.Col = NAMEINDEX
                    'grdBus.TextMatrix(grdBus.Row, CODEINDEX) = 0
                    If Trim$(grdBus.TextMatrix(grdBus.Row, NAMEINDEX)) <> "" Then
                        If gColOk(grdBus, grdBus.Row, grdBus.Col) Then
                            mEnableBox
                        Else
                            cmcCancel.SetFocus
                        End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdBus.Left - pbcArrow.Width - 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + (grdBus.RowHeight(grdBus.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdBus.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdBus.Row + 1 >= grdBus.Rows Then
                            grdBus.AddItem ""
                        End If
                        grdBus.Row = grdBus.Row + 1
                        If Not grdBus.RowIsVisible(grdBus.Row) Then
                            imIgnoreScroll = True
                            grdBus.TopRow = grdBus.TopRow + 1
                        End If
                        grdBus.Col = NAMEINDEX
                        grdBus.TextMatrix(grdBus.Row, CODEINDEX) = 0
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdBus.Left - pbcArrow.Width - 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + (grdBus.RowHeight(grdBus.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdBus.Col = grdBus.Col + 1
                If gColOk(grdBus, grdBus.Row, grdBus.Col) Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdBus.TopRow = grdBus.FixedRows
        grdBus.Col = NAMEINDEX
        grdBus.Row = grdBus.FixedRows
        If gColOk(grdBus, grdBus.Row, grdBus.Col) Then
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
    
    llTRow = grdBus.TopRow
    llRow = grdBus.Row
    slMsg = "Insert above " & Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdBus.Redraw = False
    grdBus.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdBus.Row = llRow
    grdBus.Redraw = False
    grdBus.TopRow = llTRow
    grdBus.Redraw = True
    DoEvents
    grdBus.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdBus.TopRow
    llRow = grdBus.Row
    If (Val(grdBus.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdBus.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdBus.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdBus.Redraw = False
    If (Val(grdBus.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdBus.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdBus.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdBus.AddItem ""
    grdBus.Redraw = False
    grdBus.TopRow = llTRow
    grdBus.Redraw = True
    DoEvents
    grdBus.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilBGE As Integer
    
    mBranch = True
    If (lmEnableRow >= grdBus.FixedRows) And (lmEnableRow < grdBus.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdBus.TextMatrix(lmEnableRow, lmEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEnableCol
                Case NAMEINDEX
                Case DESCRIPTIONINDEX
                Case BUSCTRLINDEX
                    'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcCCE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 2
                        sgInitCallName = slStr
                        EngrControlChar.Show vbModal
                        sgCurrBusCCEStamp = ""
                        mPopCCE
                        lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcCCE, CLng(grdBus.Height / 2)
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
                Case CHANNELINDEX
                Case BUSGROUPINDEX
                    ReDim ilBusGroupSel(0 To 0) As Integer
                    For ilLoop = 0 To lbcBGE.ListCount - 1 Step 1
                        If lbcBGE.Selected(ilLoop) Then
                            ilBusGroupSel(UBound(ilBusGroupSel)) = lbcBGE.ItemData(ilLoop)
                            ReDim Preserve ilBusGroupSel(0 To UBound(ilBusGroupSel) + 1) As Integer
                        End If
                    Next ilLoop
                    igInitCallInfo = 1
                    sgInitCallName = ""
                    EngrBusGroup.Show vbModal
                    sgCurrBGEStamp = ""
                    mPopBGE
                    For ilLoop = 0 To UBound(ilBusGroupSel) - 1 Step 1
                        For ilBGE = 0 To lbcBGE.ListCount - 1 Step 1
                            If ilBusGroupSel(ilLoop) = lbcBGE.ItemData(ilBGE) Then
                                lbcBGE.Selected(ilBGE) = True
                                Exit For
                            End If
                        Next ilBGE
                    Next ilLoop
                    lbcBGE.Move pbcDefine.Left, cmcNone.Top + cmcNone.Height, pbcDefine.Width
                    gSetListBoxHeight lbcBGE, CLng(grdBus.Height / 2)
                    If lbcBGE.Top + lbcBGE.Height > cmcCancel.Top Then
                        lbcBGE.Top = pbcDefine.Top - lbcBGE.Height
                    End If
                    If igReturnCallStatus = CALLDONE Then
                        mBranch = True
                    ElseIf igReturnCallStatus = CALLCANCELLED Then
                        mBranch = False
                    ElseIf igReturnCallStatus = CALLTERMINATED Then
                        mBranch = False
                    End If
                Case AUDIOINDEX
                    'llRow = SendMessageByString(lbcASE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcASE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrAudio.Show vbModal
                        sgCurrASEStamp = ""
                        mPopASE
                        lbcASE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcASE, CLng(grdBus.Height / 2)
                        If lbcASE.Top + lbcASE.Height > cmcCancel.Top Then
                            lbcASE.Top = edcDropdown.Top - lbcASE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcASE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcASE, slStr)
                            If llRow > 0 Then
                                lbcASE.ListIndex = llRow
                                edcDropdown.text = lbcASE.List(lbcASE.ListIndex)
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

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrBDE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdBus.FixedRows To grdBus.Rows - 1 Step 1
            slStr = Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdBus.Row = llRow
                    Do While Not grdBus.RowIsVisible(grdBus.Row)
                        imIgnoreScroll = True
                        grdBus.TopRow = grdBus.TopRow + 1
                    Loop
                    grdBus.Col = NAMEINDEX
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
    For llRow = grdBus.FixedRows To grdBus.Rows - 1 Step 1
        slStr = Trim$(grdBus.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdBus.Row = llRow
            Do While Not grdBus.RowIsVisible(grdBus.Row)
                imIgnoreScroll = True
                grdBus.TopRow = grdBus.TopRow + 1
            Loop
            grdBus.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdBus.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Function mCompare(tlNew As BDE, tlOld As BDE) As Integer
    Dim ilLoop As Integer
    Dim ilBGE As Integer
    Dim ilBSE As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim slStr As String
    
    If StrComp(tlNew.sName, tlOld.sName, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sDescription, tlOld.sDescription, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sChannel, tlOld.sChannel, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iCceCode <> tlOld.iCceCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iAseCode <> tlOld.iAseCode) Then
        mCompare = False
        Exit Function
    End If
    Erase tmCurrBSE
    ilRet = gGetRecs_BSE_BusSelGroup("B", smCurrBSEStamp, tlOld.iCode, "Bus Definition-mMoveRecToCtrls", tmCurrBSE())
    For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
        slStr = Trim$(smBusGroups(ilLoop))
        If slStr <> "" Then
            For ilBGE = 0 To UBound(tgCurrBGE) Step 1
                If StrComp(Trim$(tgCurrBGE(ilBGE).sName), slStr, vbTextCompare) = 0 Then
                    ilFound = False
                    For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
                        If tgCurrBGE(ilBGE).iCode = tmCurrBSE(ilBSE).iBgeCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilBSE
                    If Not ilFound Then
                        mCompare = False
                        Exit Function
                    End If
                End If
            Next ilBGE
        End If
    Next ilLoop
    For ilBSE = 0 To UBound(tmCurrBSE) - 1 Step 1
        For ilBGE = 0 To UBound(tgCurrBGE) Step 1
            If tgCurrBGE(ilBGE).iCode = tmCurrBSE(ilBSE).iBgeCode Then
                ilFound = False
                For ilLoop = LBound(smBusGroups) To UBound(smBusGroups) Step 1
                    slStr = Trim$(smBusGroups(ilLoop))
                    If slStr <> "" Then
                        If StrComp(Trim$(tgCurrBGE(ilBGE).sName), slStr, vbTextCompare) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilLoop
                If Not ilFound Then
                    mCompare = False
                    Exit Function
                End If
            End If
        Next ilBGE
    Next ilBSE
    mCompare = True
End Function

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case grdBus.Col
        Case BUSCTRLINDEX
            lbcCCE.Visible = False
        Case AUDIOINDEX
            lbcASE.Visible = False
    End Select
End Sub

Private Sub mPopCCE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrBus-mPop Bus Controls", tgCurrBusCCE())
    lbcCCE.Clear
    For ilLoop = 0 To UBound(tgCurrBusCCE) - 1 Step 1
        lbcCCE.AddItem Trim$(tgCurrBusCCE(ilLoop).sAutoChar)
        lbcCCE.ItemData(lbcCCE.NewIndex) = tgCurrBusCCE(ilLoop).iCode
    Next ilLoop
    lbcCCE.AddItem "[None]", 0
    lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
        lbcCCE.AddItem "[New]", 0
        lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    Else
        lbcCCE.AddItem "[View]", 0
        lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    End If
End Sub

Private Sub mSetFocus()
    Select Case grdBus.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case BUSCTRLINDEX
            edcDropdown.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - cmcDropDown.Width - 30, grdBus.RowHeight(grdBus.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcCCE, CLng(grdBus.Height / 2)
            If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcCCE.Visible = True
            edcDropdown.SetFocus
        Case CHANNELINDEX  'Date
            edcGrid.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case BUSGROUPINDEX
            pbcDefine.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
            cmcDefine.Move pbcDefine.Left, pbcDefine.Top + pbcDefine.Height, pbcDefine.Width, pbcDefine.Height
            cmcNone.Move pbcDefine.Left, cmcDefine.Top + cmcDefine.Height, pbcDefine.Width, pbcDefine.Height
            'lbcBGE.Move cmcDefine.Left, cmcDefine.Top + cmcDefine.Height, cmcDefine.Width
            lbcBGE.Move pbcDefine.Left, cmcNone.Top + cmcNone.Height, pbcDefine.Width
            gSetListBoxHeight lbcBGE, CLng(grdBus.Height / 2)
            If lbcBGE.Top + lbcBGE.Height > cmcCancel.Top Then
                lbcBGE.Top = pbcDefine.Top - lbcBGE.Height
            End If
            If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
                cmcDefine.Caption = "[New]"
            Else
                cmcDefine.Caption = "[View]"
            End If
            cmcDefine.Visible = True
            cmcNone.Visible = True
            pbcDefine.Visible = True
            lbcBGE.Visible = True
            lbcBGE.SetFocus
        Case AUDIOINDEX
            edcDropdown.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - cmcDropDown.Width - 30, grdBus.RowHeight(grdBus.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcASE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcASE, CLng(grdBus.Height / 2)
            If lbcASE.Top + lbcASE.Height > cmcCancel.Top Then
                lbcASE.Top = edcDropdown.Top - lbcASE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcASE.Visible = True
            edcDropdown.SetFocus
        Case STATEINDEX
            pbcState.Move grdBus.Left + grdBus.ColPos(grdBus.Col) + 30, grdBus.Top + grdBus.RowPos(grdBus.Row) + 15, grdBus.ColWidth(grdBus.Col) - 30, grdBus.RowHeight(grdBus.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub


