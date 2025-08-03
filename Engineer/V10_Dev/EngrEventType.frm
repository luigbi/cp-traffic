VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrEventType 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrEventType.frx":0000
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
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   45
      Width           =   45
   End
   Begin VB.PictureBox pbcPTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   120
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   6570
      Width           =   60
   End
   Begin VB.PictureBox pbcPSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   165
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   5205
      Width           =   60
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6165
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcCategory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2970
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4335
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   6
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
      TabIndex        =   8
      Top             =   5055
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
      Picture         =   "EngrEventType.frx":030A
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
      Top             =   6615
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11265
      Top             =   1065
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
      Top             =   6615
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3390
      TabIndex        =   13
      Top             =   6615
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEventType 
      Height          =   4380
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   7726
      _Version        =   393216
      Rows            =   3
      Cols            =   10
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdEventProperties 
      Height          =   1335
      Left            =   1050
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5130
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   5
      Cols            =   35
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
      _Band(0).Cols   =   35
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
      Left            =   10095
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
      Left            =   8400
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Label lacExport 
      Caption         =   "Export:"
      Height          =   255
      Left            =   135
      TabIndex        =   21
      Top             =   6165
      Width           =   1650
   End
   Begin VB.Label lacUsed 
      Caption         =   "Used:"
      Height          =   180
      Left            =   150
      TabIndex        =   19
      Top             =   5655
      Width           =   1650
   End
   Begin VB.Label lacMandatory 
      Caption         =   "Mandatory:"
      Height          =   255
      Left            =   150
      TabIndex        =   18
      Top             =   5910
      Width           =   1650
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1335
      Picture         =   "EngrEventType.frx":0614
      Top             =   6540
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Event Type"
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   450
      Picture         =   "EngrEventType.frx":091E
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   10530
      Picture         =   "EngrEventType.frx":11E8
      Top             =   6465
      Width           =   480
   End
End
Attribute VB_Name = "EngrEventType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrEventType - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private smState As String
Private smCategory As String
Private smYN As String
Private imInChg As Integer
Private imBSMode As Integer
Private imETECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer
Private imMaxCols As Integer

Private smESCValue As String    'Value used if ESC pressed
Private smPESCValue As String    'Value used if ESC pressed

Private tmETE As ETE
Private smCurrEPEStamp
Private tmCurrEPE() As EPE

Private imDeleteCodes() As Integer



'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer
Private lmPEnableRow As Long         'Current or last row focus was on
Private lmPEnableCol As Long         'Current or last column focus was on

Const CATEGORYINDEX = 0
Const NAMEINDEX = 1
Const DESCRIPTIONINDEX = 2
Const AUTOCODEINDEX = 3
Const STATEINDEX = 4
Const CODEINDEX = 5
Const USEDFLAGINDEX = 6
Const USEDINDEX = 7
Const MANDATORYINDEX = 8
Const EXPORTINDEX = 9

Const EXPORTROW = 4
Const BUSNAMEINDEX = 0
Const BUSCTRLINDEX = 1
Const TIMEINDEX = 2
Const STARTTYPEINDEX = 3
Const FIXEDINDEX = 4
Const ENDTYPEINDEX = 5
Const DURATIONINDEX = 6
Const MATERIALINDEX = 7
Const AUDIONAMEINDEX = 8
Const AUDIOITEMIDINDEX = 9
Const AUDIOISCIINDEX = 10
Const AUDIOCTRLINDEX = 11
Const BACKUPNAMEINDEX = 12  '14
Const BACKUPCTRLINDEX = 13  '15
Const PROTNAMEINDEX = 14    '11
Const PROTITEMIDINDEX = 15  '12
Const PROTISCIINDEX = 16
Const PROTCTRLINDEX = 17    '13
Const RELAY1INDEX = 18
Const RELAY2INDEX = 19
Const FOLLOWINDEX = 20
Const SILENCETIMEINDEX = 21
Const SILENCE1INDEX = 22
Const SILENCE2INDEX = 23
Const SILENCE3INDEX = 24
Const SILENCE4INDEX = 25
Const NETCUE1INDEX = 26
Const NETCUE2INDEX = 27
Const TITLE1INDEX = 28
Const TITLE2INDEX = 29
Const ABCFORMATINDEX = 30
Const ABCPGMCODEINDEX = 31
Const ABCXDSMODEINDEX = 32
Const ABCRECORDITEMINDEX = 33
Const PCODEINDEX = 34

Private Sub cmcCancel_GotFocus()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_Search(grdEventType, NAMEINDEX, slStr)
    If llRow >= 0 Then
        mEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub


Private Function mNameOk() As Integer
    Dim ilError As Integer
    Dim llRow As Long
    Dim llTestRow As Long
    Dim slStr As String
    Dim slTestStr As String
    
    grdEventType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdEventType.FixedRows To grdEventType.Rows - 1 Step 1
        slStr = Trim$(grdEventType.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdEventType.Rows - 1 Step 1
                slTestStr = Trim$(grdEventType.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdEventType.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdEventType.Row = llRow
                        grdEventType.Col = NAMEINDEX
                        grdEventType.CellForeColor = vbRed
                    Else
                        grdEventType.Row = llTestRow
                        grdEventType.Col = NAMEINDEX
                        grdEventType.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdEventType.Redraw = True
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
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    gGrid_SortByCol grdEventType, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(EVENTTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdEventType.Row >= grdEventType.FixedRows) And (grdEventType.Row < grdEventType.Rows) And (grdEventType.Col >= 0) And (grdEventType.Col < grdEventType.Cols - 1) Then
        If lmEnableRow <> grdEventType.Row Then
            mMoveEPECtrlsToRec lmEnableRow
            mPSetShow
            mMoveEPERecToCtrls grdEventType.Row
        End If
        lmEnableRow = grdEventType.Row
        lmEnableCol = grdEventType.Col
        sgReturnCallName = grdEventType.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
        If ((Val(grdEventType.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdEventType.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdEventType.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdEventType.Col
            Case CATEGORYINDEX
                pbcCategory.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
                smCategory = grdEventType.text
                If (Trim$(smCategory) = "") Or (smCategory = "Missing") Then
                    If lmEnableRow = grdEventType.FixedRows Then
                        smCategory = "Program"
                    ElseIf lmEnableRow = grdEventType.FixedRows + 1 Then
                        If grdEventType.TextMatrix(grdEventType.FixedRows, CATEGORYINDEX) = "Program" Then
                            smCategory = "Avail"
                        Else
                            smCategory = "Program"
                        End If
                    Else
                        smCategory = "Program"
                    End If
                End If
                pbcCategory.Visible = True
                pbcCategory.SetFocus
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
                edcGrid.MaxLength = Len(tmETE.sName)
                slStr = grdEventType.text
                If (slStr = "Missing") Then
                    slStr = ""
                End If
                edcGrid.text = slStr
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
                edcGrid.MaxLength = Len(tmETE.sDescription)
                slStr = grdEventType.text
                If (slStr = "Missing") Then
                    slStr = ""
                End If
                edcGrid.text = slStr
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case AUTOCODEINDEX  'Call Letters
                edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
                edcGrid.MaxLength = Len(tmETE.sAutoCodeChar)
                slStr = grdEventType.text
                If (slStr = "Missing") Then
                    slStr = ""
                End If
                edcGrid.text = slStr
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
                smState = grdEventType.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdEventType.text
    End If
End Sub

Private Sub mPEnableBox()
    Dim ilCol As Integer
    Dim llColPos As Integer
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(EVENTTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (lmEnableRow >= grdEventType.FixedRows) And (lmEnableRow < grdEventType.Rows) And (grdEventType.Col >= 0) Then
        If (grdEventProperties.Row >= grdEventProperties.FixedRows) And (grdEventProperties.Row < grdEventProperties.Rows) And (grdEventProperties.Col >= 0) And (grdEventProperties.Col < grdEventProperties.Cols - 1) Then
            mSetShow
            pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            lmPEnableRow = grdEventProperties.Row
            'lmPEnableCol = grdEventProperties.Col
            'llColPos = 0
            'For ilCol = 0 To grdEventProperties.Col - 1 Step 1
            '    llColPos = llColPos + grdEventProperties.ColWidth(ilCol)
            'Next ilCol
            ilCol = grdEventProperties.Col
            If grdEventProperties.Col >= TITLE1INDEX Then
                grdEventProperties.LeftCol = grdEventProperties.LeftCol + 6
                DoEvents
            End If
            lmPEnableRow = grdEventProperties.Row
            grdEventProperties.Col = ilCol
            lmPEnableCol = grdEventProperties.Col
            llColPos = 0
            For ilCol = 0 To grdEventProperties.Col - 1 Step 1
                If grdEventProperties.ColIsVisible(ilCol) Then
                    llColPos = llColPos + grdEventProperties.ColWidth(ilCol)
                End If
            Next ilCol
            pbcYN.Move grdEventProperties.Left + llColPos + 30, grdEventProperties.Top + grdEventProperties.RowPos(grdEventProperties.Row) + 15, grdEventProperties.ColWidth(grdEventProperties.Col) - 30, grdEventProperties.RowHeight(grdEventProperties.Row) - 15
            smYN = grdEventProperties.text
            If (Trim$(smYN) = "") Or (smYN = "Missing") Then
                smYN = "N"
            End If
            pbcYN.Visible = True
            pbcYN.SetFocus
        End If
        smPESCValue = grdEventProperties.text
    End If
End Sub
Private Sub mSetShow()
    Dim ilIndex As Integer
    
    If (lmEnableRow >= grdEventType.FixedRows) And (lmEnableRow < grdEventType.Rows) Then
        If lmEnableRow <> grdEventType.Row Then
            mMoveEPECtrlsToRec lmEnableRow
        End If
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case CATEGORYINDEX
                grdEventType.TextMatrix(lmEnableRow, lmEnableCol) = smCategory
                ilIndex = Val(grdEventType.TextMatrix(lmEnableRow, EXPORTINDEX))
                If Trim$(tmCurrEPE(ilIndex).sBus) = "" Then
                    If smCategory = "Avail" Then
                        mInitEPE tmCurrEPE(ilIndex), "E", "N"
                    Else
                        mInitEPE tmCurrEPE(ilIndex), "E", "Y"
                    End If
                    mMoveEPERecToCtrls lmEnableRow
                    mMoveEPERecToCtrls lmEnableRow
                End If
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
            Case AUTOCODEINDEX
                If (Trim$(grdEventType.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdEventType.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdEventType.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdEventType.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdEventType.TextMatrix(lmEnableRow, NAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    edcGrid.Visible = False
    pbcCategory.Visible = False
    pbcState.Visible = False
End Sub

Private Sub mPSetShow()
    Dim llCurrRow As Long
    Dim llCurrCol As Long
    
    If (lmEnableRow >= grdEventType.FixedRows) And (lmEnableRow < grdEventType.Rows) Then
        If (lmPEnableRow >= grdEventProperties.FixedRows) And (lmPEnableRow < grdEventProperties.Rows) Then
            If (lmPEnableCol >= BUSNAMEINDEX) And (lmPEnableCol <= imMaxCols) Then
                grdEventProperties.TextMatrix(lmPEnableRow, lmPEnableCol) = smYN
                If lmPEnableRow = EXPORTROW Then
                    llCurrRow = grdEventProperties.Row
                    llCurrCol = grdEventProperties.Col
                    grdEventProperties.Row = lmPEnableRow
                    grdEventProperties.Col = lmPEnableCol
                    If smYN = "N" Then
                        grdEventProperties.CellForeColor = vbBlue
                    Else
                        grdEventProperties.CellForeColor = vbBlack
                    End If
                    grdEventProperties.Row = llCurrRow
                    grdEventProperties.Col = llCurrCol
                End If
            End If
        End If
    End If
    pbcArrow.Visible = False
    pbcYN.Visible = False
    lmPEnableRow = -1
    lmPEnableCol = -1
End Sub

Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    
    grdEventType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdEventType.FixedRows To grdEventType.Rows - 1 Step 1
        slStr = Trim$(grdEventType.TextMatrix(llRow, CATEGORYINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdEventType.TextMatrix(llRow, NAMEINDEX)
            If slStr <> "" Then
                ilError = True
                grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Missing"
                grdEventType.Row = llRow
                grdEventType.Col = CATEGORYINDEX
                grdEventType.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdEventType.TextMatrix(llRow, NAMEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdEventType.TextMatrix(llRow, NAMEINDEX) = "Missing"
                    grdEventType.Row = llRow
                    grdEventType.Col = NAMEINDEX
                    grdEventType.CellForeColor = vbRed
                End If
                slStr = grdEventType.TextMatrix(llRow, AUTOCODEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdEventType.TextMatrix(llRow, AUTOCODEINDEX) = "Missing"
                    grdEventType.Row = llRow
                    grdEventType.Col = AUTOCODEINDEX
                    grdEventType.CellForeColor = vbRed
                End If
                slStr = grdEventType.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdEventType.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdEventType.Row = llRow
                    grdEventType.Col = STATEINDEX
                    grdEventType.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdEventType.Redraw = True
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
    
    
    gGrid_AlignAllColsLeft grdEventType
    mGridColumnWidth
    'Set Titles
    grdEventType.TextMatrix(0, CATEGORYINDEX) = "Category"
    grdEventType.TextMatrix(0, NAMEINDEX) = "Name"
    grdEventType.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdEventType.TextMatrix(0, AUTOCODEINDEX) = "Code"
    grdEventType.TextMatrix(0, STATEINDEX) = "State"
    grdEventType.Row = 1
    For ilCol = 0 To grdEventType.Cols - 1 Step 1
        grdEventType.Col = ilCol
        grdEventType.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdEventType.Height = grdEventProperties.Top - grdEventType.Top - 120    '8 * grdEventType.RowHeight(0) + 30
    gGrid_IntegralHeight grdEventType
    gGrid_Clear grdEventType, True
    grdEventType.Row = grdEventType.FixedRows
    
    gGrid_AlignAllColsLeft grdEventProperties
    mGridColumnWidth
    'Set Titles
    'Set Titles
    For ilCol = BUSNAMEINDEX To BUSCTRLINDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Bus"
    Next ilCol
    For ilCol = AUDIONAMEINDEX To AUDIOCTRLINDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Audio"
    Next ilCol
    For ilCol = BACKUPNAMEINDEX To BACKUPCTRLINDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Backup"
    Next ilCol
    For ilCol = PROTNAMEINDEX To PROTCTRLINDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Protection"
    Next ilCol
    For ilCol = RELAY1INDEX To RELAY2INDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Relay"
    Next ilCol
    For ilCol = SILENCETIMEINDEX To SILENCE4INDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Silence"
    Next ilCol
    For ilCol = NETCUE1INDEX To NETCUE2INDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Netcue"
    Next ilCol
    For ilCol = TITLE1INDEX To TITLE2INDEX Step 1
        grdEventProperties.TextMatrix(0, ilCol) = "Title"
    Next ilCol
    grdEventProperties.TextMatrix(1, BUSNAMEINDEX) = "Name"
    grdEventProperties.TextMatrix(1, BUSCTRLINDEX) = "C"
    grdEventProperties.TextMatrix(0, TIMEINDEX) = "Offset"
    grdEventProperties.TextMatrix(1, TIMEINDEX) = "Time"
    grdEventProperties.TextMatrix(0, STARTTYPEINDEX) = "Start "
    grdEventProperties.TextMatrix(1, STARTTYPEINDEX) = "Type"
    grdEventProperties.TextMatrix(0, FIXEDINDEX) = "Fix"
    grdEventProperties.TextMatrix(0, ENDTYPEINDEX) = "End"
    grdEventProperties.TextMatrix(1, ENDTYPEINDEX) = "Type"
    grdEventProperties.TextMatrix(0, DURATIONINDEX) = "Dur"
    grdEventProperties.TextMatrix(0, MATERIALINDEX) = "Mat"
    grdEventProperties.TextMatrix(1, MATERIALINDEX) = "Type"
    grdEventProperties.TextMatrix(1, AUDIONAMEINDEX) = "Name"
    grdEventProperties.TextMatrix(1, AUDIOITEMIDINDEX) = "Item"
    grdEventProperties.TextMatrix(1, AUDIOISCIINDEX) = "ISCI"
    grdEventProperties.TextMatrix(1, AUDIOCTRLINDEX) = "C"
    grdEventProperties.TextMatrix(1, BACKUPNAMEINDEX) = "Name"
    grdEventProperties.TextMatrix(1, BACKUPCTRLINDEX) = "C"
    grdEventProperties.TextMatrix(1, PROTNAMEINDEX) = "Name"
    grdEventProperties.TextMatrix(1, PROTITEMIDINDEX) = "Item"
    grdEventProperties.TextMatrix(1, PROTISCIINDEX) = "ISCI"
    grdEventProperties.TextMatrix(1, PROTCTRLINDEX) = "C"
    grdEventProperties.TextMatrix(1, RELAY1INDEX) = "1"
    grdEventProperties.TextMatrix(1, RELAY2INDEX) = "2"
    grdEventProperties.TextMatrix(0, FOLLOWINDEX) = "Fol-"
    grdEventProperties.TextMatrix(1, FOLLOWINDEX) = "low"
    grdEventProperties.TextMatrix(1, SILENCETIMEINDEX) = "Time"
    grdEventProperties.TextMatrix(1, SILENCE1INDEX) = "1"
    grdEventProperties.TextMatrix(1, SILENCE2INDEX) = "2"
    grdEventProperties.TextMatrix(1, SILENCE3INDEX) = "3"
    grdEventProperties.TextMatrix(1, SILENCE4INDEX) = "4"
    grdEventProperties.TextMatrix(1, NETCUE1INDEX) = "Start"
    grdEventProperties.TextMatrix(1, NETCUE2INDEX) = "Stop"
    grdEventProperties.TextMatrix(1, TITLE1INDEX) = "1"
    grdEventProperties.TextMatrix(1, TITLE2INDEX) = "2"
    grdEventProperties.TextMatrix(0, ABCFORMATINDEX) = "For-"
    grdEventProperties.TextMatrix(1, ABCFORMATINDEX) = "mat"
    grdEventProperties.TextMatrix(0, ABCPGMCODEINDEX) = "Pgm"
    grdEventProperties.TextMatrix(1, ABCPGMCODEINDEX) = "Code"
    grdEventProperties.TextMatrix(0, ABCXDSMODEINDEX) = "XDS"
    grdEventProperties.TextMatrix(1, ABCXDSMODEINDEX) = "Mode"
    grdEventProperties.TextMatrix(0, ABCRECORDITEMINDEX) = "Rec'd"
    grdEventProperties.TextMatrix(1, ABCRECORDITEMINDEX) = "Item"
    
    grdEventProperties.Row = 1
    For ilCol = 0 To grdEventProperties.Cols - 1 Step 1
        grdEventProperties.Col = ilCol
        grdEventProperties.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdEventProperties.Row = 0
    grdEventProperties.MergeCells = flexMergeRestrictRows
    grdEventProperties.MergeRow(0) = True
    grdEventProperties.Row = 0
    grdEventProperties.Col = BUSNAMEINDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = AUDIONAMEINDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = BACKUPNAMEINDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = PROTNAMEINDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = RELAY1INDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = SILENCETIMEINDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = NETCUE1INDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Row = 0
    grdEventProperties.Col = TITLE1INDEX
    grdEventProperties.CellAlignment = flexAlignCenterCenter
    grdEventProperties.Height = 5 * grdEventProperties.RowHeight(0) + 15
    'grdEventProperties.Height = grdEventProperties.RowHeight(0) + grdEventProperties.RowHeight(1) + grdEventProperties.RowHeight(2) + grdEventProperties.RowHeight(3) '- 15
    gGrid_IntegralHeight grdEventProperties
    gGrid_Clear grdEventProperties, True
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdEventType.Width = EngrEventType.Width - 2 * grdEventType.Left
    grdEventType.ColWidth(CODEINDEX) = 0
    grdEventType.ColWidth(USEDFLAGINDEX) = 0
    grdEventType.ColWidth(USEDINDEX) = 0
    grdEventType.ColWidth(MANDATORYINDEX) = 0
    grdEventType.ColWidth(EXPORTINDEX) = 0
    grdEventType.ColWidth(CATEGORYINDEX) = grdEventType.Width / 15
    grdEventType.ColWidth(NAMEINDEX) = grdEventType.Width / 7
    grdEventType.ColWidth(AUTOCODEINDEX) = grdEventType.Width / 13
    grdEventType.ColWidth(STATEINDEX) = grdEventType.Width / 15
    grdEventType.ColWidth(DESCRIPTIONINDEX) = grdEventType.Width - GRIDSCROLLWIDTH
    For ilCol = CATEGORYINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdEventType.ColWidth(DESCRIPTIONINDEX) > grdEventType.ColWidth(ilCol) Then
                grdEventType.ColWidth(DESCRIPTIONINDEX) = grdEventType.ColWidth(DESCRIPTIONINDEX) - grdEventType.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    
    grdEventProperties.ColWidth(PCODEINDEX) = 0
    grdEventProperties.ColWidth(BUSNAMEINDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(BUSCTRLINDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(STARTTYPEINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(FIXEDINDEX) = grdEventProperties.Width / 35
    grdEventProperties.ColWidth(ENDTYPEINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(DURATIONINDEX) = grdEventProperties.Width / 25
    grdEventProperties.ColWidth(MATERIALINDEX) = grdEventProperties.Width / 25
    grdEventProperties.ColWidth(AUDIONAMEINDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(AUDIOITEMIDINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(AUDIOISCIINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(AUDIOCTRLINDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(BACKUPNAMEINDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(BACKUPCTRLINDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(PROTNAMEINDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(PROTITEMIDINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(PROTISCIINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(PROTCTRLINDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(RELAY1INDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(RELAY2INDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(FOLLOWINDEX) = grdEventProperties.Width / 30
    grdEventProperties.ColWidth(SILENCETIMEINDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(SILENCE1INDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(SILENCE2INDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(SILENCE3INDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(SILENCE4INDEX) = grdEventProperties.Width / 57
    grdEventProperties.ColWidth(NETCUE1INDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(NETCUE2INDEX) = grdEventProperties.Width / 21
    grdEventProperties.ColWidth(TITLE1INDEX) = grdEventProperties.Width / 52
    grdEventProperties.ColWidth(TITLE2INDEX) = grdEventProperties.Width / 52
    
    If sgClientFields = "A" Then
        grdEventProperties.ColWidth(ABCFORMATINDEX) = grdEventProperties.Width / 28
        grdEventProperties.ColWidth(ABCPGMCODEINDEX) = grdEventProperties.Width / 28
        grdEventProperties.ColWidth(ABCXDSMODEINDEX) = grdEventProperties.Width / 28
        grdEventProperties.ColWidth(ABCRECORDITEMINDEX) = grdEventProperties.Width / 28
    Else
        grdEventProperties.ColWidth(ABCFORMATINDEX) = 0
        grdEventProperties.ColWidth(ABCPGMCODEINDEX) = 0
        grdEventProperties.ColWidth(ABCXDSMODEINDEX) = 0
        grdEventProperties.ColWidth(ABCRECORDITEMINDEX) = 0
    End If
    
    grdEventProperties.ColWidth(TIMEINDEX) = grdEventProperties.Width '- GRIDSCROLLWIDTH
    For ilCol = BUSNAMEINDEX To TITLE2INDEX Step 1
        If ilCol <> TIMEINDEX Then
            If grdEventProperties.ColWidth(TIMEINDEX) > grdEventProperties.ColWidth(ilCol) Then
                grdEventProperties.ColWidth(TIMEINDEX) = grdEventProperties.ColWidth(TIMEINDEX) - grdEventProperties.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol


End Sub


Private Sub mClearControls()
    gGrid_Clear grdEventType, True
    'Can't be 0 to 0 because index stored into grid
    ReDim tmCurrEPE(1 To 1) As EPE
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdEventType.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdEventType.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmETE.iCode = Val(grdEventType.TextMatrix(llRow, CODEINDEX))
    If grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Avail" Then
        tmETE.sCategory = "A"
    ElseIf grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Spot" Then
        tmETE.sCategory = "S"
    Else
        tmETE.sCategory = "P"
    End If
    slStr = Trim$(grdEventType.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmETE.sName = ""
    Else
        tmETE.sName = slStr
    End If
    tmETE.sDescription = grdEventType.TextMatrix(llRow, DESCRIPTIONINDEX)
    tmETE.sAutoCodeChar = grdEventType.TextMatrix(llRow, AUTOCODEINDEX)
    If grdEventType.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmETE.sState = "D"
    Else
        tmETE.sState = "A"
    End If
    If tmETE.iCode <= 0 Then
        tmETE.sUsedFlag = "N"
    Else
        tmETE.sUsedFlag = grdEventType.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmETE.iVersion = 0
    tmETE.iOrigEteCode = tmETE.iCode
    tmETE.sCurrent = "Y"
    'tmETE.sEnteredDate = smNowDate
    'tmETE.sEnteredTime = smNowTime
    tmETE.sEnteredDate = Format(Now, sgShowDateForm)
    tmETE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmETE.iUieCode = tgUIE.iCode
    tmETE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilEPE As Integer
    Dim ilFound As Integer
    
    'gGrid_Clear grdEventType, True
    llRow = grdEventType.FixedRows
    For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
        If llRow + 1 > grdEventType.Rows Then
            grdEventType.AddItem ""
        End If
        grdEventType.Row = llRow
        If tgCurrETE(ilLoop).sCategory = "A" Then
            grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Avail"
        ElseIf tgCurrETE(ilLoop).sCategory = "S" Then
            grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Spot"
        Else
            grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Program"
        End If
        grdEventType.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrETE(ilLoop).sName)
        grdEventType.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrETE(ilLoop).sDescription)
        grdEventType.TextMatrix(llRow, AUTOCODEINDEX) = Trim$(tgCurrETE(ilLoop).sAutoCodeChar)
        If tgCurrETE(ilLoop).sState = "A" Then
            grdEventType.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdEventType.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdEventType.TextMatrix(llRow, CODEINDEX) = tgCurrETE(ilLoop).iCode
        grdEventType.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrETE(ilLoop).sUsedFlag
        grdEventType.TextMatrix(llRow, USEDINDEX) = "0"
        For ilEPE = LBound(tgCurrEPE) To UBound(tgCurrEPE) - 1 Step 1
            If (tgCurrEPE(ilEPE).iEteCode = tgCurrETE(ilLoop).iCode) And (tgCurrEPE(ilEPE).sType = "U") Then
                LSet tmCurrEPE(UBound(tmCurrEPE)) = tgCurrEPE(ilEPE)
                grdEventType.TextMatrix(llRow, USEDINDEX) = UBound(tmCurrEPE)
                ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                Exit For
            End If
        Next ilEPE
        grdEventType.TextMatrix(llRow, MANDATORYINDEX) = "0"
        For ilEPE = LBound(tgCurrEPE) To UBound(tgCurrEPE) - 1 Step 1
            If (tgCurrEPE(ilEPE).iEteCode = tgCurrETE(ilLoop).iCode) And (tgCurrEPE(ilEPE).sType = "M") Then
                LSet tmCurrEPE(UBound(tmCurrEPE)) = tgCurrEPE(ilEPE)
                grdEventType.TextMatrix(llRow, MANDATORYINDEX) = UBound(tmCurrEPE)
                ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                Exit For
            End If
        Next ilEPE
        grdEventType.TextMatrix(llRow, EXPORTINDEX) = "0"
        For ilEPE = LBound(tgCurrEPE) To UBound(tgCurrEPE) - 1 Step 1
            If (tgCurrEPE(ilEPE).iEteCode = tgCurrETE(ilLoop).iCode) And (tgCurrEPE(ilEPE).sType = "E") Then
                LSet tmCurrEPE(UBound(tmCurrEPE)) = tgCurrEPE(ilEPE)
                grdEventType.TextMatrix(llRow, EXPORTINDEX) = UBound(tmCurrEPE)
                ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                Exit For
            End If
        Next ilEPE
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdEventType.Rows Then
        grdEventType.AddItem ""
    End If
    grdEventType.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrEventType-mPopulate", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrEventType-mPopulate", tgCurrEPE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim ilOldUsedEPECode As Integer
    Dim ilOldManEPECode As Integer
    Dim ilOldExpEPECode As Integer
    Dim ilETECompare As Integer
    Dim ilUsedCompare As Integer
    Dim ilManCompare As Integer
    Dim ilExpCompare As Integer
    Dim ilCount As Integer
    Dim tlETE As ETE
    
    gSetMousePointer grdEventType, grdEventType, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdEventType, grdEventType, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdEventType, grdEventType, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    ilCount = mCheckForOneSpot()
    If ilCount <= 0 Then
        gSetMousePointer grdEventType, grdEventType, vbDefault
        ilRet = MsgBox("No Spot Category defined, Press Ok to continue with Save anyway", vbOKCancel + vbQuestion)
        If ilRet = vbCancel Then
            mSave = False
            Exit Function
        End If
        gSetMousePointer grdEventType, grdEventType, vbHourglass
    ElseIf ilCount > 1 Then
        gSetMousePointer grdEventType, grdEventType, vbDefault
        MsgBox "Only One Spot type allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdEventType.Redraw = False
    For llRow = grdEventType.FixedRows To grdEventType.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmETE.sName) <> "" Then
            ilUsedCompare = True
            ilManCompare = True
            ilETECompare = True
            imETECode = tmETE.iCode
            If tmETE.iCode > 0 Then
                ilRet = gGetRec_ETE_EventType(imETECode, "Event Type-mSave: Get ETE", tlETE)
                If ilRet Then
                    ilOldUsedEPECode = tmCurrEPE(Val(grdEventType.TextMatrix(llRow, USEDINDEX))).iCode
                    ilOldManEPECode = tmCurrEPE(Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX))).iCode
                    ilOldExpEPECode = tmCurrEPE(Val(grdEventType.TextMatrix(llRow, EXPORTINDEX))).iCode
                    ilUsedCompare = mCompareEPE(ilOldUsedEPECode)
                    ilManCompare = mCompareEPE(ilOldManEPECode)
                    ilExpCompare = mCompareEPE(ilOldExpEPECode)
                    ilETECompare = mCompare(tmETE, tlETE)
                    If (ilETECompare) And (ilUsedCompare) And (ilManCompare) And (ilExpCompare) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmETE.iVersion = tlETE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmETE.iCode <= 0 Then
                    ilRet = gPutInsert_ETE_EventType(0, tmETE, "Event Type-mSave: Insert ETE")
                    ilIndex = Val(grdEventType.TextMatrix(llRow, USEDINDEX))
                    tmCurrEPE(ilIndex).iCode = 0
                    tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    ilIndex = Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX))
                    tmCurrEPE(ilIndex).iCode = 0
                    tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    ilIndex = Val(grdEventType.TextMatrix(llRow, EXPORTINDEX))
                    tmCurrEPE(ilIndex).iCode = 0
                    tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                Else
                    '7/12/11: History no longer retained
                    'If ilETECompare Then
                    '    ilRet = gPutUpdate_ETE_EventType(0, tmETE, "Event Type-mSave: Update ETE")
                    'Else
                    '    ilRet = gPutUpdate_ETE_EventType(1, tmETE, "Event Type-mSave: Update ETE")
                    'End If
                    'ilIndex = Val(grdEventType.TextMatrix(llRow, USEDINDEX))
                    'tmCurrEPE(ilIndex).iCode = 0
                    'tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    'ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    'If (Not ilUsedCompare) Then
                    '    ilRet = gUpdateAIE(1, tmETE.iVersion, "EPE", CLng(ilOldUsedEPECode), CLng(tmCurrEPE(ilIndex).iCode), CLng(tmETE.iOrigEteCode), "Event Type- mSave: Insert EPE:AIE")
                    'End If
                    'ilIndex = Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX))
                    'tmCurrEPE(ilIndex).iCode = 0
                    'tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    'ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    'If (Not ilManCompare) Then
                    '    ilRet = gUpdateAIE(1, tmETE.iVersion, "EPE", CLng(ilOldManEPECode), CLng(tmCurrEPE(ilIndex).iCode), CLng(tmETE.iOrigEteCode), "Event Type- mSave: Insert EPE:AIE")
                    'End If
                    'ilIndex = Val(grdEventType.TextMatrix(llRow, EXPORTINDEX))
                    'tmCurrEPE(ilIndex).iCode = 0
                    'tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    'ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    'If (Not ilExpCompare) Then
                    '    ilRet = gUpdateAIE(1, tmETE.iVersion, "EPE", CLng(ilOldExpEPECode), CLng(tmCurrEPE(ilIndex).iCode), CLng(tmETE.iOrigEteCode), "Event Type- mSave: Insert EPE:AIE")
                    'End If
                    ilRet = gPutDelete_ETE_EventType(tmETE.iCode, "Event Type-mSave: Delete ETE")
                    ilRet = gPutInsert_ETE_EventType(1, tmETE, "Event Type-mSave: Insert ETE")
                    ilIndex = Val(grdEventType.TextMatrix(llRow, USEDINDEX))
                    tmCurrEPE(ilIndex).iCode = 0
                    tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    ilIndex = Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX))
                    tmCurrEPE(ilIndex).iCode = 0
                    tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                    ilIndex = Val(grdEventType.TextMatrix(llRow, EXPORTINDEX))
                    tmCurrEPE(ilIndex).iCode = 0
                    tmCurrEPE(ilIndex).iEteCode = tmETE.iCode
                    ilRet = gPutInsert_EPE_EventProperties(tmCurrEPE(ilIndex), "Event Type- mSave: Insert EPE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_ETE_EventType(imDeleteCodes(ilLoop), "EngrEventType- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdEventType.Redraw = True
    sgCurrETEStamp = ""
    sgCurrEPEStamp = ""
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrEventType-mSave", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrEventType-mSave", tgCurrEPE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrEventType
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrEventType
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdEventType, grdEventType, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdEventType, grdEventType, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdEventType, grdEventType, vbDefault
    Unload EngrEventType
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim llTopRow As Long
    
    If imFieldChgd = True Then
        gSetMousePointer grdEventType, grdEventType, vbHourglass
        llTopRow = grdEventType.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdEventType, grdEventType, vbDefault
            Exit Sub
        End If
        grdEventType.Redraw = False
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
        grdEventType.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdEventType, grdEventType, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdEventType.Col
        Case CATEGORYINDEX
        Case NAMEINDEX
            If grdEventType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdEventType.text = edcGrid.text
            grdEventType.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdEventType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdEventType.text = edcGrid.text
            grdEventType.CellForeColor = vbBlack
        Case AUTOCODEINDEX
            If grdEventType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdEventType.text = edcGrid.text
            grdEventType.CellForeColor = vbBlack
        Case STATEINDEX
    End Select
    mSetCommands
End Sub

Private Sub edcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSearch_GotFocus()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
        mFindMatch True
        lacUsed.Top = grdEventProperties.Top + 2 * grdEventProperties.RowHeight(0)
        lacMandatory.Top = grdEventProperties.Top + 3 * grdEventProperties.RowHeight(0)
        lacExport.Top = grdEventProperties.Top + 4 * grdEventProperties.RowHeight(0)
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
    gSetFonts EngrEventType
    gCenterFormModal EngrEventType
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If Not pbcYN.Visible Then
            If (lmEnableRow >= grdEventType.FixedRows) And (lmEnableRow < grdEventType.Rows) Then
                If (lmEnableCol >= grdEventType.FixedCols) And (lmEnableCol < grdEventType.Cols) Then
                If lmEnableCol = CATEGORYINDEX Then
                    smCategory = smESCValue
                ElseIf lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdEventType.text = smESCValue
                End If
                    mSetShow
                    mEnableBox
                End If
            End If
        Else
            If (lmPEnableRow >= grdEventProperties.FixedRows) And (lmPEnableRow < grdEventProperties.Rows) Then
                If (lmPEnableCol >= grdEventProperties.FixedCols) And (lmPEnableCol < grdEventProperties.Cols) Then
                    smYN = smPESCValue
                    mPSetShow
                    mPEnableBox
                End If
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
    grdEventProperties.Height = 5 * grdEventProperties.RowHeight(0) + 15
    gGrid_IntegralHeight grdEventProperties
    grdEventType.Height = grdEventProperties.Top - grdEventType.Top - 120    '8 * grdEventType.RowHeight(0) + 30
    gGrid_IntegralHeight grdEventType
    gGrid_FillWithRows grdEventType
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Erase tmCurrEPE
    Set EngrEventType = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdEventType, grdEventType, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim imDeleteCodes(0 To 0) As Integer
    'Can't be 0 to 0 because of index in grid
    ReDim tmCurrEPE(1 To 1) As EPE
    cmcSearch.Top = 30
    edcSearch.Top = cmcSearch.Top
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imFirstActivate = True
    imInChg = True
    mPopulate
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    If sgClientFields = "A" Then
        grdEventProperties.ScrollBars = flexScrollBarHorizontal
        imMaxCols = ABCRECORDITEMINDEX
    Else
        imMaxCols = TITLE2INDEX
    End If
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(EVENTTYPELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdEventType, grdEventType, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdEventType, grdEventType, vbDefault
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

Private Sub grdEventProperties_Click()
    If grdEventProperties.Col >= grdEventType.Cols - 1 Then
        Exit Sub
    End If

End Sub

Private Sub grdEventProperties_EnterCell()
    mPSetShow
    mSetShow
End Sub

Private Sub grdEventProperties_GotFocus()
    If grdEventProperties.Col >= grdEventProperties.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdEventProperties_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    
    'If same cell entered after clicking some other place, a enter cell event does not happen
    mSetShow
    If (lmEnableRow < grdEventType.FixedRows) And (lmEnableRow >= grdEventType.Rows) Then
        Exit Sub
    End If
    'Determine if in header
    If y < grdEventProperties.RowHeight(0) Then
        Exit Sub
    End If
    'ilFound = gGrid_DetermineRowCol(grdEventProperties, X, Y)
    'If Not ilFound Then
    '    pbcClickFocus.SetFocus
    '    Exit Sub
    'End If
    If grdEventProperties.Col >= grdEventProperties.Cols - 1 Then
        Exit Sub
    End If
    If mPColOk(grdEventProperties.Row, grdEventProperties.Col) Then
        mPEnableBox
    Else
        Beep
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub imcInsert_Click()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = EVENT_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
    mDeleteRow
End Sub

Private Sub pbcCategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If smCategory <> "Avail" Then
            imFieldChgd = True
        End If
        smCategory = "Avail"
        pbcCategory_Paint
        grdEventType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("S") Or (KeyAscii = Asc("s")) Then
        If smCategory <> "Spot" Then
            imFieldChgd = True
        End If
        smCategory = "Spot"
        pbcCategory_Paint
        grdEventType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
        If smCategory <> "Program" Then
            imFieldChgd = True
        End If
        smCategory = "Program"
        pbcCategory_Paint
        grdEventType.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smCategory = "Avail" Then
            imFieldChgd = True
            smCategory = "Spot"
            pbcCategory_Paint
            grdEventType.CellForeColor = vbBlack
        ElseIf smCategory = "Spot" Then
            imFieldChgd = True
            smCategory = "Program"
            pbcCategory_Paint
            grdEventType.CellForeColor = vbBlack
        ElseIf smCategory = "Program" Then
            imFieldChgd = True
            smCategory = "Avail"
            pbcCategory_Paint
            grdEventType.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcCategory_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smCategory = "Avail" Then
        imFieldChgd = True
        smCategory = "Spot"
        pbcCategory_Paint
        grdEventType.CellForeColor = vbBlack
    ElseIf smCategory = "Spot" Then
        imFieldChgd = True
        smCategory = "Program"
        pbcCategory_Paint
        grdEventType.CellForeColor = vbBlack
    ElseIf smCategory = "Program" Then
        imFieldChgd = True
        smCategory = "Avail"
        pbcCategory_Paint
        grdEventType.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcCategory_Paint()
    pbcCategory.Cls
    pbcCategory.CurrentX = 30  'fgBoxInsetX
    pbcCategory.CurrentY = 0 'fgBoxInsetY
    pbcCategory.Print smCategory
End Sub

Private Sub grdEventType_Click()
    If grdEventType.Col >= grdEventType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdEventType_EnterCell()
    mPSetShow
    mSetShow
End Sub

Private Sub grdEventType_GotFocus()
    If grdEventType.Col >= grdEventType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdEventType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdEventType.TopRow
    grdEventType.Redraw = False
End Sub

Private Sub grdEventType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'If same cell entered after clicking some other place, a enter cell event does not happen
    mPSetShow
    'Determine if in header
    If y < grdEventType.RowHeight(0) Then
        mSortCol grdEventType.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdEventType, x, y)
    If Not ilFound Then
        grdEventType.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdEventType.Col >= grdEventType.Cols - 1 Then
        grdEventType.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdEventType.TopRow
    DoEvents
    llRow = grdEventType.Row
    If grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "" Then
        grdEventType.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdEventType.TextMatrix(llRow, CATEGORYINDEX) = ""
        grdEventType.Row = llRow + 1
        grdEventType.Col = CATEGORYINDEX
        grdEventType.Redraw = True
    End If
    grdEventType.Redraw = True
    mEnableBox
End Sub

Private Sub grdEventType_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdEventType.Redraw = False Then
        grdEventType.Redraw = True
        If lmTopRow < grdEventType.FixedRows Then
            grdEventType.TopRow = grdEventType.FixedRows
        Else
            grdEventType.TopRow = lmTopRow
        End If
        grdEventType.Refresh
        grdEventType.Redraw = False
    End If
    If (imShowGridBox) And (grdEventType.Row >= grdEventType.FixedRows) And (grdEventType.Col >= 0) And (grdEventType.Col < grdEventType.Cols - 1) Then
        If grdEventType.RowIsVisible(grdEventType.Row) Then
            'edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 30, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 30
            pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            pbcArrow.Visible = False
            edcGrid.Visible = False
            pbcCategory.Visible = False
            pbcState.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mPSetShow
    mSetShow
    mMoveEPECtrlsToRec lmEnableRow
    lmPEnableRow = -1
    lmPEnableCol = -1
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcPSTab_GotFocus()
    Dim ilNext As Integer
    
    If GetFocus() <> pbcPSTab.hwnd Then
        Exit Sub
    End If
    If pbcYN.Visible Then
        mPSetShow
        Do
            ilNext = False
            If grdEventProperties.Col = BUSNAMEINDEX Then
                If grdEventProperties.Row > grdEventProperties.FixedRows Then
                    grdEventProperties.Row = grdEventProperties.Row - 1
                    grdEventProperties.Col = imMaxCols  'TITLE2INDEX
                    'mPEnableBox
                Else
                    cmcCancel.SetFocus
                    Exit Sub
                End If
            Else
                grdEventProperties.Col = grdEventProperties.Col - 1
                'mPEnableBox
            End If
            If mPColOk(grdEventProperties.Row, grdEventProperties.Col) Then
                mPEnableBox
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        grdEventProperties.LeftCol = BUSNAMEINDEX
        grdEventProperties.Col = BUSNAMEINDEX
        grdEventProperties.Row = grdEventProperties.FixedRows
        mPEnableBox
    End If
End Sub

Private Sub pbcPTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llPEnableRow As Long
    
    If GetFocus() <> pbcPTab.hwnd Then
        Exit Sub
    End If
    If pbcYN.Visible Then
        llPEnableRow = lmPEnableRow
        mPSetShow
        Do
            ilNext = False
            If grdEventProperties.Col = imMaxCols Then
                If grdEventProperties.Row >= grdEventProperties.Rows - 1 Then
                    If grdEventType.Col = STATEINDEX Then
                        llRow = grdEventType.Rows
                        Do
                            llRow = llRow - 1
                        Loop While grdEventType.TextMatrix(llRow, CATEGORYINDEX) = ""
                        llRow = llRow + 1
                        If (grdEventType.Row + 1 < llRow) Then
                            lmTopRow = -1
                            grdEventType.Row = grdEventType.Row + 1
                            If Not grdEventType.RowIsVisible(grdEventType.Row) Then
                                imIgnoreScroll = True
                                grdEventType.TopRow = grdEventType.TopRow + 1
                            End If
                            grdEventType.Col = CATEGORYINDEX
                            'grdEventType.TextMatrix(grdEventType.Row, CODEINDEX) = 0
                            If Trim$(grdEventType.TextMatrix(grdEventType.Row, CATEGORYINDEX)) <> "" Then
                                mEnableBox
                            Else
                                imFromArrow = True
                                pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
                                pbcArrow.Visible = True
                                pbcArrow.SetFocus
                            End If
                        Else
                            If Trim$(grdEventType.TextMatrix(lmEnableRow, CATEGORYINDEX)) <> "" Then
                                lmTopRow = -1
                                If grdEventType.Row + 1 >= grdEventType.Rows Then
                                    grdEventType.AddItem ""
                                End If
                                grdEventType.Row = grdEventType.Row + 1
                                If Not grdEventType.RowIsVisible(grdEventType.Row) Then
                                    imIgnoreScroll = True
                                    grdEventType.TopRow = grdEventType.TopRow + 1
                                End If
                                grdEventType.Col = CATEGORYINDEX
                                grdEventType.TextMatrix(grdEventType.Row, CODEINDEX) = 0
                                'mEnableBox
                                imFromArrow = True
                                pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
                                pbcArrow.Visible = True
                                pbcArrow.SetFocus
                            Else
                                pbcClickFocus.SetFocus
                            End If
                        End If
                        Exit Sub
                    Else
                        mEnableBox
                        Exit Sub
                    End If
                Else
                    grdEventProperties.Row = grdEventProperties.Row + 1
                    grdEventProperties.LeftCol = BUSNAMEINDEX
                    grdEventProperties.Col = BUSNAMEINDEX
                    'mPEnableBox
                End If
            Else
                grdEventProperties.Col = grdEventProperties.Col + 1
                'mPEnableBox
            End If
            If mPColOk(grdEventProperties.Row, grdEventProperties.Col) Then
                mPEnableBox
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        grdEventProperties.LeftCol = BUSNAMEINDEX
        grdEventProperties.Col = BUSNAMEINDEX
        grdEventProperties.Row = grdEventProperties.FixedRows
        mPEnableBox
    End If
End Sub

Private Sub pbcSTab_GotFocus()
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If edcGrid.Visible Or pbcCategory.Visible Or pbcState.Visible Then
        mSetShow
        If grdEventType.Col = CATEGORYINDEX Then
            If grdEventType.Row > grdEventType.FixedRows Then
                lmTopRow = -1
                grdEventType.Row = grdEventType.Row - 1
                If Not grdEventType.RowIsVisible(grdEventType.Row) Then
                    grdEventType.TopRow = grdEventType.TopRow - 1
                End If
                grdEventType.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdEventType.Col = grdEventType.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdEventType.TopRow = grdEventType.FixedRows
        grdEventType.Col = CATEGORYINDEX
        grdEventType.Row = grdEventType.FixedRows
        mEnableBox
    End If
End Sub

Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If smState <> "Active" Then
            imFieldChgd = True
        End If
        smState = "Active"
        pbcState_Paint
        grdEventType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdEventType.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdEventType.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdEventType.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdEventType.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdEventType.CellForeColor = vbBlack
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
    Dim llEnableRow As Long
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If edcGrid.Visible Or pbcCategory.Visible Or pbcState.Visible Then
        llEnableRow = lmEnableRow
        mSetShow
        If grdEventType.Col = STATEINDEX Then
            
            llRow = grdEventType.Rows
            Do
                llRow = llRow - 1
            Loop While grdEventType.TextMatrix(llRow, CATEGORYINDEX) = ""
            llRow = llRow + 1
            If (grdEventType.Row + 1 < llRow) Then
                lmTopRow = -1
                grdEventType.Row = grdEventType.Row + 1
                If Not grdEventType.RowIsVisible(grdEventType.Row) Then
                    imIgnoreScroll = True
                    grdEventType.TopRow = grdEventType.TopRow + 1
                End If
                grdEventType.Col = CATEGORYINDEX
                'grdEventType.TextMatrix(grdEventType.Row, CODEINDEX) = 0
                If Trim$(grdEventType.TextMatrix(grdEventType.Row, CATEGORYINDEX)) <> "" Then
                    mEnableBox
                Else
'                    imFromArrow = True
'                    pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
'                    pbcArrow.Visible = True
'                    pbcArrow.SetFocus
                    grdEventProperties.LeftCol = BUSNAMEINDEX
                    grdEventProperties.Col = BUSNAMEINDEX
                    grdEventProperties.Row = grdEventProperties.FixedRows
                    mPEnableBox
                End If
            Else
                If Trim$(grdEventType.TextMatrix(llEnableRow, CATEGORYINDEX)) <> "" Then
'                    lmTopRow = -1
'                    If grdEventType.Row + 1 >= grdEventType.Rows Then
'                        grdEventType.AddItem ""
'                    End If
'                    grdEventType.Row = grdEventType.Row + 1
'                    If Not grdEventType.RowIsVisible(grdEventType.Row) Then
'                        grdEventType.TopRow = grdEventType.TopRow + 1
'                    End If
'                    grdEventType.Col = CATEGORYINDEX
'                    grdEventType.TextMatrix(grdEventType.Row, CODEINDEX) = 0
'                    'mEnableBox
'                    imFromArrow = True
'                    pbcArrow.Move grdEventType.Left - pbcArrow.Width - 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + (grdEventType.RowHeight(grdEventType.Row) - pbcArrow.Height) / 2
'                    pbcArrow.Visible = True
'                    pbcArrow.SetFocus
                    grdEventProperties.LeftCol = BUSNAMEINDEX
                    grdEventProperties.Col = BUSNAMEINDEX
                    grdEventProperties.Row = grdEventProperties.FixedRows
                    mPEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdEventType.Col = grdEventType.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdEventType.TopRow = grdEventType.FixedRows
        grdEventType.Col = CATEGORYINDEX
        grdEventType.Row = grdEventType.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdEventType.TopRow
    llRow = grdEventType.Row
    slMsg = "Insert above " & Trim$(grdEventType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdEventType.Redraw = False
    grdEventType.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdEventType.Row = llRow
    grdEventType.Redraw = False
    grdEventType.TopRow = llTRow
    grdEventType.Redraw = True
    DoEvents
    grdEventType.Col = CATEGORYINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdEventType.TopRow
    llRow = grdEventType.Row
    If (Val(grdEventType.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdEventType.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdEventType.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdEventType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdEventType.Redraw = False
    If (Val(grdEventType.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdEventType.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
'    If (Val(grdEventType.TextMatrix(llRow, USEDINDEX)) <> 0) Then
'        If Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX)) <> 0 Then
'            If (Val(grdEventType.TextMatrix(llRow, USEDINDEX)) <> 0) < (Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX)) <> 0) Then
'            Else
'            End If
'        Else
'        End If
'    Else
'        If Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX)) <> 0 Then
'
'        End If
'    End If
    
    grdEventType.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdEventType.AddItem ""
    grdEventType.Redraw = False
    grdEventType.TopRow = llTRow
    grdEventType.Redraw = True
    DoEvents
    'grdEventType.Col = CATEGORYINDEX
    'mEnableBox
    cmcCancel.SetFocus
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As ETE, tlOld As ETE) As Integer
    If StrComp(tlNew.sCategory, tlOld.sCategory, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sName, tlOld.sName, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sDescription, tlOld.sDescription, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sAutoCodeChar, tlOld.sAutoCodeChar, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
End Function

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrETE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdEventType.FixedRows To grdEventType.Rows - 1 Step 1
            slStr = Trim$(grdEventType.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdEventType.Row = llRow
                    Do While Not grdEventType.RowIsVisible(grdEventType.Row)
                        imIgnoreScroll = True
                        grdEventType.TopRow = grdEventType.TopRow + 1
                    Loop
                    grdEventType.Col = NAMEINDEX
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
    For llRow = grdEventType.FixedRows To grdEventType.Rows - 1 Step 1
        slStr = Trim$(grdEventType.TextMatrix(llRow, CATEGORYINDEX))
        If (slStr = "") Then
            grdEventType.Row = llRow
            Do While Not grdEventType.RowIsVisible(grdEventType.Row)
                imIgnoreScroll = True
                grdEventType.TopRow = grdEventType.TopRow + 1
            Loop
            grdEventType.Col = CATEGORYINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdEventType.TextMatrix(llRow, NAMEINDEX) = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub


Private Sub mMoveEPERecToCtrls(llRow As Long)
    Dim llIndex As Long
    Dim ilPass As Integer
    Dim llPRow As Long
    Dim llCol As Long
    Dim slStr As String
    
    If (llRow >= grdEventType.FixedRows) And (llRow < grdEventType.Rows) Then
        For ilPass = 0 To 2 Step 1
            If ilPass = 0 Then
                If Val(grdEventType.TextMatrix(llRow, USEDINDEX)) = 0 Then
                    grdEventType.TextMatrix(llRow, USEDINDEX) = UBound(tmCurrEPE)
                    mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "U", ""
                    ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                End If
                llIndex = Val(grdEventType.TextMatrix(llRow, USEDINDEX))
            ElseIf ilPass = 1 Then
                If Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX)) = 0 Then
                    grdEventType.TextMatrix(llRow, MANDATORYINDEX) = UBound(tmCurrEPE)
                    mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "M", ""
                    ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                End If
                llIndex = Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX))
            ElseIf ilPass = 2 Then
                If Val(grdEventType.TextMatrix(llRow, EXPORTINDEX)) = 0 Then
                    grdEventType.TextMatrix(llRow, EXPORTINDEX) = UBound(tmCurrEPE)
                    If Trim$(grdEventType.TextMatrix(llRow, CATEGORYINDEX)) = "" Then
                        mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "E", ""
                    Else
                        If grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Avail" Then
                            mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "E", "N"
                        Else
                            mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "E", "Y"
                        End If
                    End If
                    ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                End If
                llIndex = Val(grdEventType.TextMatrix(llRow, EXPORTINDEX))
            End If
            llPRow = ilPass + 2
            grdEventProperties.TextMatrix(llPRow, BUSNAMEINDEX) = tmCurrEPE(llIndex).sBus
            grdEventProperties.TextMatrix(llPRow, BUSCTRLINDEX) = tmCurrEPE(llIndex).sBusControl
            grdEventProperties.TextMatrix(llPRow, TIMEINDEX) = tmCurrEPE(llIndex).sTime
            grdEventProperties.TextMatrix(llPRow, STARTTYPEINDEX) = tmCurrEPE(llIndex).sStartType
            grdEventProperties.TextMatrix(llPRow, FIXEDINDEX) = tmCurrEPE(llIndex).sFixedTime
            grdEventProperties.TextMatrix(llPRow, ENDTYPEINDEX) = tmCurrEPE(llIndex).sEndType
            grdEventProperties.TextMatrix(llPRow, DURATIONINDEX) = tmCurrEPE(llIndex).sDuration
            grdEventProperties.TextMatrix(llPRow, MATERIALINDEX) = tmCurrEPE(llIndex).sMaterialType
            grdEventProperties.TextMatrix(llPRow, AUDIONAMEINDEX) = tmCurrEPE(llIndex).sAudioName
            grdEventProperties.TextMatrix(llPRow, AUDIOITEMIDINDEX) = tmCurrEPE(llIndex).sAudioItemID
            grdEventProperties.TextMatrix(llPRow, AUDIOISCIINDEX) = tmCurrEPE(llIndex).sAudioISCI
            grdEventProperties.TextMatrix(llPRow, AUDIOCTRLINDEX) = tmCurrEPE(llIndex).sAudioControl
            grdEventProperties.TextMatrix(llPRow, BACKUPNAMEINDEX) = tmCurrEPE(llIndex).sBkupAudioName
            grdEventProperties.TextMatrix(llPRow, BACKUPCTRLINDEX) = tmCurrEPE(llIndex).sBkupAudioControl
            grdEventProperties.TextMatrix(llPRow, PROTNAMEINDEX) = tmCurrEPE(llIndex).sProtAudioName
            grdEventProperties.TextMatrix(llPRow, PROTITEMIDINDEX) = tmCurrEPE(llIndex).sProtAudioItemID
            grdEventProperties.TextMatrix(llPRow, PROTISCIINDEX) = tmCurrEPE(llIndex).sProtAudioISCI
            grdEventProperties.TextMatrix(llPRow, PROTCTRLINDEX) = tmCurrEPE(llIndex).sProtAudioControl
            grdEventProperties.TextMatrix(llPRow, RELAY1INDEX) = tmCurrEPE(llIndex).sRelay1
            grdEventProperties.TextMatrix(llPRow, RELAY2INDEX) = tmCurrEPE(llIndex).sRelay2
            grdEventProperties.TextMatrix(llPRow, FOLLOWINDEX) = tmCurrEPE(llIndex).sFollow
            grdEventProperties.TextMatrix(llPRow, SILENCETIMEINDEX) = tmCurrEPE(llIndex).sSilenceTime
            grdEventProperties.TextMatrix(llPRow, SILENCE1INDEX) = tmCurrEPE(llIndex).sSilence1
            grdEventProperties.TextMatrix(llPRow, SILENCE2INDEX) = tmCurrEPE(llIndex).sSilence2
            grdEventProperties.TextMatrix(llPRow, SILENCE3INDEX) = tmCurrEPE(llIndex).sSilence3
            grdEventProperties.TextMatrix(llPRow, SILENCE4INDEX) = tmCurrEPE(llIndex).sSilence4
            grdEventProperties.TextMatrix(llPRow, NETCUE1INDEX) = tmCurrEPE(llIndex).sStartNetcue
            grdEventProperties.TextMatrix(llPRow, NETCUE2INDEX) = tmCurrEPE(llIndex).sStopNetcue
            grdEventProperties.TextMatrix(llPRow, TITLE1INDEX) = tmCurrEPE(llIndex).sTitle1
            grdEventProperties.TextMatrix(llPRow, TITLE2INDEX) = tmCurrEPE(llIndex).sTitle2
            If sgClientFields = "A" Then
                grdEventProperties.TextMatrix(llPRow, ABCFORMATINDEX) = tmCurrEPE(llIndex).sABCFormat
                grdEventProperties.TextMatrix(llPRow, ABCPGMCODEINDEX) = tmCurrEPE(llIndex).sABCPgmCode
                grdEventProperties.TextMatrix(llPRow, ABCXDSMODEINDEX) = tmCurrEPE(llIndex).sABCXDSMode
                grdEventProperties.TextMatrix(llPRow, ABCRECORDITEMINDEX) = tmCurrEPE(llIndex).sABCRecordItem
            Else
                grdEventProperties.TextMatrix(llPRow, ABCFORMATINDEX) = ""
                grdEventProperties.TextMatrix(llPRow, ABCPGMCODEINDEX) = ""
                grdEventProperties.TextMatrix(llPRow, ABCXDSMODEINDEX) = ""
                grdEventProperties.TextMatrix(llPRow, ABCRECORDITEMINDEX) = ""
            End If
            grdEventProperties.TextMatrix(llPRow, PCODEINDEX) = tmCurrEPE(llIndex).iCode
            If llPRow = EXPORTROW Then
                For llCol = BUSNAMEINDEX To imMaxCols Step 1
                    If grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Avail" Then
                        If grdEventProperties.TextMatrix(llPRow, llCol) <> "N" Then
                            grdEventProperties.TextMatrix(llPRow, llCol) = "N"
                            imFieldChgd = True
                        End If
                        grdEventProperties.Row = llPRow
                        grdEventProperties.Col = llCol
                        grdEventProperties.CellForeColor = vbBlue
                        grdEventProperties.CellBackColor = LIGHTYELLOW
                    Else
                        slStr = grdEventProperties.TextMatrix(llPRow, llCol)
                        grdEventProperties.Row = llPRow
                        grdEventProperties.Col = llCol
                        If slStr = "N" Then
                            grdEventProperties.CellForeColor = vbBlue
                        ElseIf slStr = "Y" Then
                            grdEventProperties.CellForeColor = vbBlack
                        End If
                        grdEventProperties.CellBackColor = vbWhite
                    End If
                Next llCol
            End If
        Next ilPass
    End If
    mSetCommands
End Sub

Private Sub mMoveEPECtrlsToRec(llRow As Long)
    Dim llIndex As Long
    Dim ilPass As Integer
    Dim llPRow As Long
    
    If (llRow >= grdEventType.FixedRows) And (llRow < grdEventType.Rows) Then
        For ilPass = 0 To 2 Step 1
            If ilPass = 0 Then
                If Val(grdEventType.TextMatrix(llRow, USEDINDEX)) = 0 Then
                    grdEventType.TextMatrix(llRow, USEDINDEX) = UBound(tmCurrEPE)
                    mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "U", ""
                    ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                End If
                llIndex = Val(grdEventType.TextMatrix(llRow, USEDINDEX))
                tmCurrEPE(llIndex).sType = "U"
            ElseIf ilPass = 1 Then
                If Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX)) = 0 Then
                    grdEventType.TextMatrix(llRow, MANDATORYINDEX) = UBound(tmCurrEPE)
                    mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "M", ""
                    ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                End If
                llIndex = Val(grdEventType.TextMatrix(llRow, MANDATORYINDEX))
                tmCurrEPE(llIndex).sType = "M"
            ElseIf ilPass = 2 Then
                If Val(grdEventType.TextMatrix(llRow, EXPORTINDEX)) = 0 Then
                    grdEventType.TextMatrix(llRow, EXPORTINDEX) = UBound(tmCurrEPE)
                    If Trim$(grdEventType.TextMatrix(llRow, CATEGORYINDEX)) = "" Then
                        mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "E", ""
                    Else
                        If grdEventType.TextMatrix(llRow, CATEGORYINDEX) = "Avail" Then
                            mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "E", "N"
                        Else
                            mInitEPE tmCurrEPE(UBound(tmCurrEPE)), "E", "Y"
                        End If
                    End If
                    ReDim Preserve tmCurrEPE(1 To UBound(tmCurrEPE) + 1) As EPE
                End If
                llIndex = Val(grdEventType.TextMatrix(llRow, EXPORTINDEX))
                tmCurrEPE(llIndex).sType = "E"
            End If
            llPRow = ilPass + 2
            tmCurrEPE(llIndex).sBus = grdEventProperties.TextMatrix(llPRow, BUSNAMEINDEX)
            tmCurrEPE(llIndex).sBusControl = grdEventProperties.TextMatrix(llPRow, BUSCTRLINDEX)
            tmCurrEPE(llIndex).sTime = grdEventProperties.TextMatrix(llPRow, TIMEINDEX)
            tmCurrEPE(llIndex).sStartType = grdEventProperties.TextMatrix(llPRow, STARTTYPEINDEX)
            tmCurrEPE(llIndex).sFixedTime = grdEventProperties.TextMatrix(llPRow, FIXEDINDEX)
            tmCurrEPE(llIndex).sEndType = grdEventProperties.TextMatrix(llPRow, ENDTYPEINDEX)
            tmCurrEPE(llIndex).sDuration = grdEventProperties.TextMatrix(llPRow, DURATIONINDEX)
            tmCurrEPE(llIndex).sMaterialType = grdEventProperties.TextMatrix(llPRow, MATERIALINDEX)
            tmCurrEPE(llIndex).sAudioName = grdEventProperties.TextMatrix(llPRow, AUDIONAMEINDEX)
            tmCurrEPE(llIndex).sAudioItemID = grdEventProperties.TextMatrix(llPRow, AUDIOITEMIDINDEX)
            tmCurrEPE(llIndex).sAudioISCI = grdEventProperties.TextMatrix(llPRow, AUDIOISCIINDEX)
            tmCurrEPE(llIndex).sAudioControl = grdEventProperties.TextMatrix(llPRow, AUDIOCTRLINDEX)
            tmCurrEPE(llIndex).sBkupAudioName = grdEventProperties.TextMatrix(llPRow, BACKUPNAMEINDEX)
            tmCurrEPE(llIndex).sBkupAudioControl = grdEventProperties.TextMatrix(llPRow, BACKUPCTRLINDEX)
            tmCurrEPE(llIndex).sProtAudioName = grdEventProperties.TextMatrix(llPRow, PROTNAMEINDEX)
            tmCurrEPE(llIndex).sProtAudioItemID = grdEventProperties.TextMatrix(llPRow, PROTITEMIDINDEX)
            tmCurrEPE(llIndex).sProtAudioISCI = grdEventProperties.TextMatrix(llPRow, PROTISCIINDEX)
            tmCurrEPE(llIndex).sProtAudioControl = grdEventProperties.TextMatrix(llPRow, PROTCTRLINDEX)
            tmCurrEPE(llIndex).sRelay1 = grdEventProperties.TextMatrix(llPRow, RELAY1INDEX)
            tmCurrEPE(llIndex).sRelay2 = grdEventProperties.TextMatrix(llPRow, RELAY2INDEX)
            tmCurrEPE(llIndex).sFollow = grdEventProperties.TextMatrix(llPRow, FOLLOWINDEX)
            tmCurrEPE(llIndex).sSilenceTime = grdEventProperties.TextMatrix(llPRow, SILENCETIMEINDEX)
            tmCurrEPE(llIndex).sSilence1 = grdEventProperties.TextMatrix(llPRow, SILENCE1INDEX)
            tmCurrEPE(llIndex).sSilence2 = grdEventProperties.TextMatrix(llPRow, SILENCE2INDEX)
            tmCurrEPE(llIndex).sSilence3 = grdEventProperties.TextMatrix(llPRow, SILENCE3INDEX)
            tmCurrEPE(llIndex).sSilence4 = grdEventProperties.TextMatrix(llPRow, SILENCE4INDEX)
            tmCurrEPE(llIndex).sStartNetcue = grdEventProperties.TextMatrix(llPRow, NETCUE1INDEX)
            tmCurrEPE(llIndex).sStopNetcue = grdEventProperties.TextMatrix(llPRow, NETCUE2INDEX)
            tmCurrEPE(llIndex).sTitle1 = grdEventProperties.TextMatrix(llPRow, TITLE1INDEX)
            tmCurrEPE(llIndex).sTitle2 = grdEventProperties.TextMatrix(llPRow, TITLE2INDEX)
            If sgClientFields = "A" Then
                tmCurrEPE(llIndex).sABCFormat = grdEventProperties.TextMatrix(llPRow, ABCFORMATINDEX)
                tmCurrEPE(llIndex).sABCPgmCode = grdEventProperties.TextMatrix(llPRow, ABCPGMCODEINDEX)
                tmCurrEPE(llIndex).sABCXDSMode = grdEventProperties.TextMatrix(llPRow, ABCXDSMODEINDEX)
                tmCurrEPE(llIndex).sABCRecordItem = grdEventProperties.TextMatrix(llPRow, ABCRECORDITEMINDEX)
            Else
                tmCurrEPE(llIndex).sABCFormat = ""
                tmCurrEPE(llIndex).sABCPgmCode = ""
                tmCurrEPE(llIndex).sABCXDSMode = ""
                tmCurrEPE(llIndex).sABCRecordItem = ""
            End If
            tmCurrEPE(llIndex).sUnused = ""
            tmCurrEPE(llIndex).iCode = Val(grdEventProperties.TextMatrix(llPRow, PCODEINDEX))
        Next ilPass
    End If
    mSetCommands
End Sub

Private Sub mInitEPE(tlEPE As EPE, slType As String, slSetting As String)
    If slType = "U" Then
        tlEPE.sBus = "Y"
        tlEPE.sBusControl = "N"
        tlEPE.sType = slType
        tlEPE.sTime = "Y"
        tlEPE.sStartType = "Y"
        tlEPE.sFixedTime = "Y"
        tlEPE.sEndType = "Y"
        tlEPE.sDuration = "Y"
        tlEPE.sMaterialType = "N"
        tlEPE.sAudioName = "Y"
        tlEPE.sAudioItemID = "Y"
        If smCategory = "Spot" Then
            tlEPE.sAudioISCI = "Y"
        Else
            tlEPE.sAudioISCI = "N"
        End If
        tlEPE.sAudioControl = "N"
        tlEPE.sBkupAudioName = "Y"
        tlEPE.sBkupAudioControl = "N"
        tlEPE.sProtAudioName = "Y"
        tlEPE.sProtAudioItemID = "Y"
        If smCategory = "Spot" Then
            tlEPE.sProtAudioISCI = "Y"
        Else
            tlEPE.sProtAudioISCI = "N"
        End If
        tlEPE.sProtAudioControl = "N"
        tlEPE.sRelay1 = "Y"
        tlEPE.sRelay2 = "Y"
        tlEPE.sFollow = "N"
        tlEPE.sSilenceTime = "Y"
        tlEPE.sSilence1 = "N"
        tlEPE.sSilence2 = "N"
        tlEPE.sSilence3 = "N"
        tlEPE.sSilence4 = "N"
        tlEPE.sStartNetcue = "Y"
        tlEPE.sStopNetcue = "Y"
        tlEPE.sTitle1 = "Y"
        tlEPE.sTitle2 = "Y"
        If sgClientFields = "A" Then
            tlEPE.sABCFormat = "Y"
            tlEPE.sABCPgmCode = "Y"
            tlEPE.sABCXDSMode = "Y"
            tlEPE.sABCRecordItem = "Y"
        Else
            tlEPE.sABCFormat = "N"
            tlEPE.sABCPgmCode = "N"
            tlEPE.sABCXDSMode = "N"
            tlEPE.sABCRecordItem = "N"
        End If
        tlEPE.iCode = 0
    ElseIf slType = "M" Then
        tlEPE.sBus = "Y"
        tlEPE.sBusControl = "N"
        tlEPE.sTime = "Y"
        tlEPE.sStartType = "N"
        tlEPE.sFixedTime = "N"
        tlEPE.sEndType = "N"
        tlEPE.sDuration = "N"
        tlEPE.sMaterialType = "N"
        tlEPE.sAudioName = "Y"
        tlEPE.sAudioItemID = "N"
        tlEPE.sAudioISCI = "N"
        tlEPE.sAudioControl = "N"
        tlEPE.sBkupAudioName = "N"
        tlEPE.sBkupAudioControl = "N"
        tlEPE.sProtAudioName = "N"
        tlEPE.sProtAudioItemID = "N"
        tlEPE.sProtAudioISCI = "N"
        tlEPE.sProtAudioControl = "N"
        tlEPE.sRelay1 = "N"
        tlEPE.sRelay2 = "N"
        tlEPE.sFollow = "N"
        tlEPE.sSilenceTime = "N"
        tlEPE.sSilence1 = "N"
        tlEPE.sSilence2 = "N"
        tlEPE.sSilence3 = "N"
        tlEPE.sSilence4 = "N"
        tlEPE.sStartNetcue = "N"
        tlEPE.sStopNetcue = "N"
        tlEPE.sTitle1 = "N"
        tlEPE.sTitle2 = "N"
        If sgClientFields = "A" Then
            tlEPE.sABCFormat = "N"
            tlEPE.sABCPgmCode = "N"
            tlEPE.sABCXDSMode = "N"
            tlEPE.sABCRecordItem = "N"
        Else
            tlEPE.sABCFormat = "N"
            tlEPE.sABCPgmCode = "N"
            tlEPE.sABCXDSMode = "N"
            tlEPE.sABCRecordItem = "N"
        End If
        tlEPE.iCode = 0
    ElseIf slType = "E" Then
        tlEPE.sBus = slSetting
        tlEPE.sBusControl = slSetting
        tlEPE.sType = slType
        tlEPE.sTime = slSetting
        tlEPE.sStartType = slSetting
        tlEPE.sFixedTime = slSetting
        tlEPE.sEndType = slSetting
        tlEPE.sDuration = slSetting
        tlEPE.sMaterialType = slSetting
        tlEPE.sAudioName = slSetting
        tlEPE.sAudioItemID = slSetting
        tlEPE.sAudioISCI = slSetting
        tlEPE.sAudioControl = slSetting
        tlEPE.sBkupAudioName = slSetting
        tlEPE.sBkupAudioControl = slSetting
        tlEPE.sProtAudioName = slSetting
        tlEPE.sProtAudioItemID = slSetting
        tlEPE.sProtAudioISCI = slSetting
        tlEPE.sProtAudioControl = slSetting
        tlEPE.sRelay1 = slSetting
        tlEPE.sRelay2 = slSetting
        tlEPE.sFollow = slSetting
        tlEPE.sSilenceTime = slSetting
        tlEPE.sSilence1 = slSetting
        tlEPE.sSilence2 = slSetting
        tlEPE.sSilence3 = slSetting
        tlEPE.sSilence4 = slSetting
        tlEPE.sStartNetcue = slSetting
        tlEPE.sStopNetcue = slSetting
        tlEPE.sTitle1 = slSetting
        tlEPE.sTitle2 = slSetting
        If sgClientFields = "A" Then
            tlEPE.sABCFormat = "Y"
            tlEPE.sABCPgmCode = "Y"
            tlEPE.sABCXDSMode = "Y"
            tlEPE.sABCRecordItem = "Y"
        Else
            tlEPE.sABCFormat = "N"
            tlEPE.sABCPgmCode = "N"
            tlEPE.sABCXDSMode = "N"
            tlEPE.sABCRecordItem = "N"
        End If
        tlEPE.iCode = 0
    End If
    imFieldChgd = True
End Sub

Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If smYN <> "Y" Then
            imFieldChgd = True
        End If
        smYN = "Y"
        pbcYN_Paint
        grdEventProperties.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If smYN <> "N" Then
            imFieldChgd = True
        End If
        smYN = "N"
        pbcYN_Paint
        grdEventProperties.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smYN = "Y" Then
            imFieldChgd = True
            smYN = "N"
            pbcYN_Paint
            grdEventProperties.CellForeColor = vbBlack
        ElseIf smYN = "N" Then
            imFieldChgd = True
            smYN = "Y"
            pbcYN_Paint
            grdEventProperties.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smYN = "Y" Then
        imFieldChgd = True
        smYN = "N"
        pbcYN_Paint
        grdEventProperties.CellForeColor = vbBlack
    ElseIf smYN = "N" Then
        imFieldChgd = True
        smYN = "Y"
        pbcYN_Paint
        grdEventProperties.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = 30  'fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    pbcYN.Print smYN
End Sub

Private Function mCompareEPE(ilCode As Integer) As Integer
    Dim ilEPENew As Integer
    Dim ilEPEOld As Integer
    
    If ilCode > 0 Then
        For ilEPENew = LBound(tmCurrEPE) To UBound(tmCurrEPE) - 1 Step 1
            If ilCode = tmCurrEPE(ilEPENew).iCode Then
                For ilEPEOld = LBound(tgCurrEPE) To UBound(tgCurrEPE) - 1 Step 1
                    If ilCode = tgCurrEPE(ilEPEOld).iCode Then
                        'Compare fields
                        If tmCurrEPE(ilEPENew).sBus <> tgCurrEPE(ilEPEOld).sBus Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sBusControl <> tgCurrEPE(ilEPEOld).sBusControl Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sTime <> tgCurrEPE(ilEPEOld).sTime Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sStartType <> tgCurrEPE(ilEPEOld).sStartType Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sFixedTime <> tgCurrEPE(ilEPEOld).sFixedTime Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sEndType <> tgCurrEPE(ilEPEOld).sEndType Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sDuration <> tgCurrEPE(ilEPEOld).sDuration Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sMaterialType <> tgCurrEPE(ilEPEOld).sMaterialType Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sAudioName <> tgCurrEPE(ilEPEOld).sAudioName Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sAudioItemID <> tgCurrEPE(ilEPEOld).sAudioItemID Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sAudioISCI <> tgCurrEPE(ilEPEOld).sAudioISCI Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sAudioControl <> tgCurrEPE(ilEPEOld).sAudioControl Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sBkupAudioName <> tgCurrEPE(ilEPEOld).sBkupAudioName Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sBkupAudioControl <> tgCurrEPE(ilEPEOld).sBkupAudioControl Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sProtAudioName <> tgCurrEPE(ilEPEOld).sProtAudioName Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sProtAudioItemID <> tgCurrEPE(ilEPEOld).sProtAudioItemID Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sProtAudioISCI <> tgCurrEPE(ilEPEOld).sProtAudioISCI Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sProtAudioControl <> tgCurrEPE(ilEPEOld).sProtAudioControl Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sRelay1 <> tgCurrEPE(ilEPEOld).sRelay1 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sRelay2 <> tgCurrEPE(ilEPEOld).sRelay2 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sFollow <> tgCurrEPE(ilEPEOld).sFollow Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sSilenceTime <> tgCurrEPE(ilEPEOld).sSilenceTime Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sSilence1 <> tgCurrEPE(ilEPEOld).sSilence1 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sSilence2 <> tgCurrEPE(ilEPEOld).sSilence2 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sSilence3 <> tgCurrEPE(ilEPEOld).sSilence3 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sSilence4 <> tgCurrEPE(ilEPEOld).sSilence4 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sStartNetcue <> tgCurrEPE(ilEPEOld).sStartNetcue Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sStopNetcue <> tgCurrEPE(ilEPEOld).sStopNetcue Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sTitle1 <> tgCurrEPE(ilEPEOld).sTitle1 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sTitle2 <> tgCurrEPE(ilEPEOld).sTitle2 Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sABCFormat <> tgCurrEPE(ilEPEOld).sABCFormat Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sABCPgmCode <> tgCurrEPE(ilEPEOld).sABCPgmCode Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sABCXDSMode <> tgCurrEPE(ilEPEOld).sABCXDSMode Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        If tmCurrEPE(ilEPENew).sABCRecordItem <> tgCurrEPE(ilEPEOld).sABCRecordItem Then
                            mCompareEPE = False
                            Exit Function
                        End If
                        mCompareEPE = True
                        Exit Function
                    End If
                Next ilEPEOld
                mCompareEPE = True
                Exit Function
            End If
        Next ilEPENew
    Else
        mCompareEPE = False
    End If
    
    
    
End Function

Private Function mCheckForOneSpot() As Integer
    Dim llRow As Long
    Dim slCategory As String
    Dim slState As String
    Dim ilCount As Integer
    ilCount = 0
    For llRow = grdEventType.FixedRows To grdEventType.Rows - 1 Step 1
        slCategory = Trim$(grdEventType.TextMatrix(llRow, CATEGORYINDEX))
        If slCategory <> "" Then
            slState = Trim$(grdEventType.TextMatrix(llRow, STATEINDEX))
            If (StrComp(slCategory, "Spot", vbTextCompare) = 0) And (StrComp(slState, "Dormant", vbTextCompare) <> 0) Then
                ilCount = ilCount + 1
            End If
        End If
    Next llRow
    mCheckForOneSpot = ilCount
    Exit Function
End Function

Private Sub mSetFocus()
    Select Case grdEventType.Col
        Case CATEGORYINDEX
            pbcCategory.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
            pbcCategory.Visible = True
            pbcCategory.SetFocus
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case AUTOCODEINDEX  'Call Letters
            edcGrid.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdEventType.Left + grdEventType.ColPos(grdEventType.Col) + 30, grdEventType.Top + grdEventType.RowPos(grdEventType.Row) + 15, grdEventType.ColWidth(grdEventType.Col) - 30, grdEventType.RowHeight(grdEventType.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub

Private Function mPColOk(llRow As Long, llCol As Long) As Integer
    mPColOk = True
    If grdEventProperties.ColWidth(llCol) <= 0 Then
        mPColOk = False
        Exit Function
    End If
    grdEventProperties.Row = llRow
    grdEventProperties.Col = llCol
    If grdEventProperties.CellBackColor = LIGHTYELLOW Then
        mPColOk = False
        Exit Function
    End If
End Function
