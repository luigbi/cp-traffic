VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGetGame 
   Caption         =   "Get Event"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffGetGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   Begin VB.PictureBox plcKey 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   225
      ScaleHeight     =   240
      ScaleWidth      =   6120
      TabIndex        =   6
      Top             =   4230
      Width           =   6120
      Begin VB.Label Label4 
         Caption         =   "Black = Not Posted"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   15
         TabIndex        =   9
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Green = Posting Completed"
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   1
         Left            =   1590
         TabIndex        =   8
         Top             =   0
         Width           =   1980
      End
      Begin VB.Label Label4 
         Caption         =   "Magenta = Partially Posted"
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Index           =   3
         Left            =   3675
         TabIndex        =   7
         Top             =   0
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5145
      TabIndex        =   5
      Top             =   4515
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AffGetGame.frx":08CA
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcGameFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   0
      Width           =   60
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   450
      Top             =   4620
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5025
      FormDesignWidth =   9105
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2685
      TabIndex        =   4
      Top             =   4515
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGame 
      Height          =   3810
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmGetGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmGetGame - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private lmRowSelected As Long

Private smPledgeByEvent As String

Private imLastGameColSorted As Integer
Private imLastGameSort As Integer

Private smEventTitle1 As String
Private smEventTitle2 As String

Private rst_Gsf As ADODB.Recordset
Private rst_Ast As ADODB.Recordset
Private rst_att As ADODB.Recordset
Private rst_Pet As ADODB.Recordset
Private rst_Lst As ADODB.Recordset

'Grid Controls

Const GAMENOINDEX = 0
Const FEEDSOURCEINDEX = 1
Const LANGUAGEINDEX = 2
Const VISITTEAMINDEX = 3
Const HOMETEAMINDEX = 4
Const AIRDATEINDEX = 5
Const AIRTIMEINDEX = 6
Const DECLAREDINDEX = 7
Const CLEAREDINDEX = 8
Const GSFCODEINDEX = 9
Const SORTINDEX = 10




Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    gGrid_Clear grdGame, True
    'Set color within cells
    For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
        'For llCol = 0 To SORTINDEX Step 1
        '    grdGame.Row = llRow
        '    grdGame.Col = llCol
        '    grdGame.CellBackColor = LIGHTYELLOW
        'Next llCol
    Next llRow
End Sub


Private Sub cmdCancel_Click()
    lgSelGameGsfCode = -1
    igSelGameNo = 0
    sgSelGameDate = ""
    Unload frmGetGame
End Sub

Private Sub cmdDone_Click()
    Dim iLoop As Integer
    
    Screen.MousePointer = vbHourglass
    If lmRowSelected > 0 Then
        lgSelGameGsfCode = Val(grdGame.TextMatrix(lmRowSelected, GSFCODEINDEX))
        igSelGameNo = Val(grdGame.TextMatrix(lmRowSelected, GAMENOINDEX))
        sgSelGameDate = grdGame.TextMatrix(lmRowSelected, AIRDATEINDEX)
    Else
        lgSelGameGsfCode = -1
        igSelGameNo = 0
        sgSelGameDate = ""
    End If
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Unload frmGetGame
    Exit Sub
   
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        gGetEventTitles igGameVefCode, smEventTitle1, smEventTitle2
        mSetGridColumns
        mSetGridTitles
        gGrid_IntegralHeight grdGame
        gGrid_FillWithRows grdGame
        mPopulate
        imFirstTime = False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.15
    Me.Height = Screen.Height / 1.55
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmGetGame
    gCenterForm frmGetGame
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    
    mInit
    Screen.MousePointer = vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Gsf.Close
    rst_Ast.Close
    rst_att.Close
    rst_Pet.Close
    rst_Lst.Close
    Set frmGetGame = Nothing
End Sub





Private Sub grdGame_Click()
    Dim llRow As Long
    
    If grdGame.Row >= grdGame.FixedRows Then
        If grdGame.TextMatrix(grdGame.Row, GAMENOINDEX) <> "" Then
            If (lmRowSelected = grdGame.Row) Then
                If imCtrlKey Then
                    lmRowSelected = -1
                    grdGame.Row = 0
                    grdGame.Col = GSFCODEINDEX
                End If
            Else
                lmRowSelected = grdGame.Row
            End If
        Else
            lmRowSelected = -1
            grdGame.Row = 0
            grdGame.Col = GSFCODEINDEX
        End If
    End If

End Sub

Private Sub grdGame_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdGame_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    
    llVeh = gBinarySearchVef(CLng(igGameVefCode))
    If llVeh <> -1 Then
        frmGetGame.Caption = "Event Selection-" & Trim$(tgVehicleInfo(llVeh).sVehicle)
    Else
        frmGetGame.Caption = "Event Selection"
    End If
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    
    ilRet = gPopVff()
    ilRet = gPopTeams()
    ilRet = gPopLangs()
    
    mGetPledgeBy
    
    mClearGrid

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slFWkDate As String
    Dim slLWkDate As String
    Dim llCellColor As Long
    Dim ilShttCode As Integer
    Dim ilTimeAdj As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slFed As String
    Dim blShowEvent As Boolean
    
    On Error GoTo ErrHand:
    grdGame.Redraw = False
    grdGame.Row = 0
    For llCol = GAMENOINDEX To CLEAREDINDEX Step 1
        grdGame.Col = llCol
        grdGame.CellBackColor = LIGHTBLUE
    Next llCol
    llRow = grdGame.FixedRows
    ilShttCode = -1
    ilTimeAdj = 0
    If lgGameAttCode > 0 Then
        SQLQuery = "SELECT attShfCode FROM att WHERE attCode = " & lgGameAttCode
        Set rst_att = gSQLSelectCall(SQLQuery)
        If Not rst_att.EOF Then
            ilShttCode = rst_att!attshfCode
            ilTimeAdj = gGetTimeAdj(ilShttCode, igGameVefCode, slFed)
        End If
    End If
    slFWkDate = Format$(gObtainPrevMonday(sgGameStartDate), sgShowDateForm)
    slLWkDate = Format$(gObtainNextSunday(sgGameStartDate), sgShowDateForm)
    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & igGameVefCode & " AND gsfAirDate >= '" & Format$(sgGameStartDate, sgSQLDateForm) & "'" & " AND gsfAirDate <= '" & Format$(sgGameEndDate, sgSQLDateForm) & "'" & ")"
    Set rst_Gsf = gSQLSelectCall(SQLQuery)
    Do While Not rst_Gsf.EOF
        If llRow >= grdGame.Rows Then
            grdGame.AddItem ""
        End If
        llCellColor = vbBlack
        If lgGameAttCode > 0 Then
            SQLQuery = "Select Count(astCode) FROM ast, lst WHERE "
            SQLQuery = SQLQuery + " astAtfCode = " & lgGameAttCode
            SQLQuery = SQLQuery + " AND lstCode = astLsfCode"
            SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_Gsf!gsfCode
            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')"
            Set rst_Ast = gSQLSelectCall(SQLQuery)
            If Not rst_Ast.EOF Then
                If Val(rst_Ast(0).Value) > 0 Then
                    plcKey.Visible = True
                    SQLQuery = "Select astCode FROM ast, lst WHERE astCPStatus = 0"
                    SQLQuery = SQLQuery + " AND astAtfCode = " & lgGameAttCode
                    SQLQuery = SQLQuery + " AND lstCode = astLsfCode"
                    SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_Gsf!gsfCode
                    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')"
                    Set rst_Ast = gSQLSelectCall(SQLQuery)
                    If rst_Ast.EOF Then
                        '3/19/13: Add VefCode to speed up call
                        'SQLQuery = "Select Count(lstCode) FROM lst WHERE lstGsfCode = " & rst_Gsf!gsfCode
                        SQLQuery = "Select Count(lstCode) FROM lst WHERE lstLogVefCode = " & igGameVefCode & " AND lstGsfCode = " & rst_Gsf!gsfCode
                        Set rst_Ast = gSQLSelectCall(SQLQuery)
                        If Not rst_Ast.EOF Then
                            llCellColor = DARKGREEN
                        End If
                    Else
                        SQLQuery = "Select astCode FROM ast, lst WHERE astCPStatus = 1"
                        SQLQuery = SQLQuery + " AND astAtfCode = " & lgGameAttCode
                        SQLQuery = SQLQuery + " AND lstCode = astLsfCode"
                        '3/19/13: Add vefCode
                        SQLQuery = SQLQuery + " AND lstLogVefCode = " & igGameVefCode
                        SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_Gsf!gsfCode
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')"
                        Set rst_Ast = gSQLSelectCall(SQLQuery)
                        If Not rst_Ast.EOF Then
                            llCellColor = vbMagenta
                        End If
                    End If
                Else
                    plcKey.Visible = False
                End If
            Else
                plcKey.Visible = False
            End If
        Else
            plcKey.Visible = False
        End If
        '3/11/17: Test if LST exist
        blShowEvent = plcKey.Visible
        If plcKey.Visible = False Then
            SQLQuery = "Select Count(*) FROM lst WHERE "
            SQLQuery = SQLQuery + " lstLogVefCode = " & igGameVefCode
            SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_Gsf!gsfCode
            Set rst_Lst = gSQLSelectCall(SQLQuery)
            If Not rst_Lst.EOF Then
                If Val(rst_Lst(0).Value) > 0 Then
                    blShowEvent = True
                End If
            End If
        End If
        If blShowEvent Then
            For llCol = GAMENOINDEX To CLEAREDINDEX Step 1
                grdGame.Row = llRow
                grdGame.Col = llCol
                grdGame.CellForeColor = llCellColor
            Next llCol
            'Game Number
            grdGame.TextMatrix(llRow, GAMENOINDEX) = rst_Gsf!gsfGameNo
            'Feed Source
            If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
                If rst_Gsf!gsfFeedSource = "V" Then
                    grdGame.TextMatrix(llRow, FEEDSOURCEINDEX) = smEventTitle1  '"Visiting"
                ElseIf rst_Gsf!gsfFeedSource = "N" Then
                    grdGame.TextMatrix(llRow, FEEDSOURCEINDEX) = "National"
                Else
                    grdGame.TextMatrix(llRow, FEEDSOURCEINDEX) = smEventTitle2  '"Home"
                End If
            End If
            'Language
            If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
                For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
                    If tgLangInfo(ilLang).iCode = rst_Gsf!gsfLangMnfCode Then
                        grdGame.TextMatrix(llRow, LANGUAGEINDEX) = Trim$(tgLangInfo(ilLang).sName)
                        Exit For
                    End If
                Next ilLang
            End If
            'Visiting Team
            For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfVisitMnfCode Then
                    grdGame.TextMatrix(llRow, VISITTEAMINDEX) = Trim$(tgTeamInfo(ilTeam).sName)
                    Exit For
                End If
            Next ilTeam
            'Home Team
            For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfHomeMnfCode Then
                    grdGame.TextMatrix(llRow, HOMETEAMINDEX) = Trim$(tgTeamInfo(ilTeam).sName)
                    Exit For
                End If
            Next ilTeam
            'Air Date
            slDate = Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
            'Start Time
            slTime = Format$(rst_Gsf!gsfAirTime, sgShowTimeWSecForm)
            gAdjustEventTime ilTimeAdj, slDate, slTime
            grdGame.TextMatrix(llRow, AIRDATEINDEX) = Format$(slDate, sgShowDateForm)
            grdGame.TextMatrix(llRow, AIRTIMEINDEX) = Format$(slTime, sgShowTimeWSecForm)
            'Declared and Cleared status
            If smPledgeByEvent = "Y" Then
                SQLQuery = "SELECT petCode, petGsfCode, petDeclaredStatus, petClearStatus"
                SQLQuery = SQLQuery + " FROM pet"
                SQLQuery = SQLQuery & " WHERE (petAttCode = " & lgGameAttCode & " AND petGsfCode = " & rst_Gsf!gsfCode & ")"
                SQLQuery = SQLQuery + " ORDER BY petGsfCode"
                Set rst_Pet = gSQLSelectCall(SQLQuery)
                If Not rst_Pet.EOF Then
                    If rst_Pet!petDeclaredStatus = "Y" Then
                        grdGame.TextMatrix(llRow, DECLAREDINDEX) = "Air"
                    ElseIf rst_Pet!petDeclaredStatus = "N" Then
                        grdGame.TextMatrix(llRow, DECLAREDINDEX) = "Not Airing"
                    Else
                        grdGame.TextMatrix(llRow, DECLAREDINDEX) = "Unknown"
                    End If
                    'D.S. 11/20/13 changed from declared to cleared status
                    If rst_Pet!petClearStatus = "Y" Then
                        grdGame.TextMatrix(llRow, CLEAREDINDEX) = "Aired"
                    'D.S. 11/20/13 changed from declared to cleared status
                    ElseIf rst_Pet!petClearStatus = "N" Then
                        grdGame.TextMatrix(llRow, CLEAREDINDEX) = "Not Aired"
                    Else
                        grdGame.TextMatrix(llRow, CLEAREDINDEX) = "Unknown"
                    End If
                Else
                    grdGame.TextMatrix(llRow, DECLAREDINDEX) = "Unknown"
                    grdGame.TextMatrix(llRow, CLEAREDINDEX) = "Unknown"
                End If
            End If
            grdGame.TextMatrix(llRow, GSFCODEINDEX) = rst_Gsf!gsfCode
            llRow = llRow + 1
        End If
        rst_Gsf.MoveNext
    Loop
    mGameSortCol AIRTIMEINDEX
    mGameSortCol AIRDATEINDEX
    If llRow = grdGame.FixedRows + 1 Then
        grdGame.Row = grdGame.FixedRows
        grdGame.RowSel = grdGame.FixedRows
        grdGame.Col = GAMENOINDEX
        grdGame.ColSel = AIRTIMEINDEX
        lmRowSelected = grdGame.FixedRows
    Else
        grdGame.Row = 0
        grdGame.Col = GSFCODEINDEX
        lmRowSelected = -1
    End If
    grdGame.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmGetGame-mPopulate"
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdGame.ColWidth(GSFCODEINDEX) = 0
    grdGame.ColWidth(SORTINDEX) = 0
    grdGame.ColWidth(GAMENOINDEX) = grdGame.Width * 0.06
    If smPledgeByEvent <> "Y" Then
        If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
            grdGame.ColWidth(FEEDSOURCEINDEX) = grdGame.Width * 0.1
        Else
            grdGame.ColWidth(FEEDSOURCEINDEX) = 0
        End If
        If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
            grdGame.ColWidth(LANGUAGEINDEX) = grdGame.Width * 0.1
        Else
            grdGame.ColWidth(LANGUAGEINDEX) = 0
        End If
        grdGame.ColWidth(DECLAREDINDEX) = 0
        grdGame.ColWidth(CLEAREDINDEX) = 0
    Else
        If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
            grdGame.ColWidth(FEEDSOURCEINDEX) = grdGame.Width * 0.1
        Else
            grdGame.ColWidth(FEEDSOURCEINDEX) = 0
        End If
        If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
            grdGame.ColWidth(LANGUAGEINDEX) = grdGame.Width * 0.09
        Else
            grdGame.ColWidth(LANGUAGEINDEX) = 0
        End If
        grdGame.ColWidth(DECLAREDINDEX) = grdGame.Width * 0.08
        grdGame.ColWidth(CLEAREDINDEX) = grdGame.Width * 0.08
    End If
    grdGame.ColWidth(VISITTEAMINDEX) = grdGame.Width * 0.16
    grdGame.ColWidth(HOMETEAMINDEX) = grdGame.Width * 0.16
    grdGame.ColWidth(AIRDATEINDEX) = grdGame.Width * 0.08
    grdGame.ColWidth(AIRTIMEINDEX) = grdGame.Width * 0.09
    
    grdGame.ColWidth(VISITTEAMINDEX) = grdGame.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To CLEAREDINDEX Step 1
        If ilCol <> VISITTEAMINDEX Then
            grdGame.ColWidth(VISITTEAMINDEX) = grdGame.ColWidth(VISITTEAMINDEX) - grdGame.ColWidth(ilCol)
        End If
    Next ilCol
    grdGame.ColWidth(VISITTEAMINDEX) = (grdGame.ColWidth(HOMETEAMINDEX) + grdGame.ColWidth(VISITTEAMINDEX)) \ 2
    grdGame.ColWidth(HOMETEAMINDEX) = grdGame.ColWidth(VISITTEAMINDEX)
    'Align columns to left
    gGrid_AlignAllColsLeft grdGame
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdGame.TextMatrix(0, GAMENOINDEX) = "Game #"
    grdGame.TextMatrix(0, FEEDSOURCEINDEX) = "Feed"
    grdGame.TextMatrix(0, LANGUAGEINDEX) = "Language"
    grdGame.TextMatrix(0, VISITTEAMINDEX) = smEventTitle1   '"Visiting Team"
    grdGame.TextMatrix(0, HOMETEAMINDEX) = smEventTitle2    '"Home Team"
    grdGame.TextMatrix(0, AIRDATEINDEX) = "Air Date"
    grdGame.TextMatrix(0, AIRTIMEINDEX) = "Start Time"
    grdGame.TextMatrix(0, DECLAREDINDEX) = "Declared"
    grdGame.TextMatrix(0, CLEAREDINDEX) = "Cleared"
    'Set height of grid

End Sub

Private Sub mGameSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
        slStr = Trim$(grdGame.TextMatrix(llRow, GAMENOINDEX))
        If slStr <> "" Then
            If ilCol = AIRDATEINDEX Then
                slSort = Trim$(Str$(gDateValue(grdGame.TextMatrix(llRow, AIRDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = AIRTIMEINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdGame.TextMatrix(llRow, AIRTIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = GAMENOINDEX) Then
                slSort = Trim$(grdGame.TextMatrix(llRow, GAMENOINDEX))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdGame.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdGame.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastGameColSorted) Or ((ilCol = imLastGameColSorted) And (imLastGameSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdGame.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdGame.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastGameColSorted Then
        imLastGameColSorted = SORTINDEX
    Else
        imLastGameColSorted = -1
        imLastGameSort = -1
    End If
    gGrid_SortByCol grdGame, GAMENOINDEX, SORTINDEX, imLastGameColSorted, imLastGameSort
    imLastGameColSorted = ilCol
End Sub

Private Sub grdGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llGsfCode As Long
    
    If Y < grdGame.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        llGsfCode = -1
        If lmRowSelected >= grdGame.FixedRows Then
            If Trim$(grdGame.TextMatrix(lmRowSelected, GAMENOINDEX)) <> "" Then
                llGsfCode = grdGame.TextMatrix(lmRowSelected, GSFCODEINDEX)
            End If
        End If
        grdGame.Col = grdGame.MouseCol
        mGameSortCol grdGame.Col
        grdGame.Row = 0
        grdGame.Col = GSFCODEINDEX
        lmRowSelected = -1
        If llGsfCode <> -1 Then
            For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
                If llGsfCode = grdGame.TextMatrix(llRow, GSFCODEINDEX) Then
                    grdGame.Row = llRow
                    grdGame.RowSel = llRow
                    grdGame.Col = GAMENOINDEX
                    grdGame.ColSel = AIRTIMEINDEX
                    lmRowSelected = llRow
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next llRow
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

End Sub

Private Sub mGetPledgeBy()
    Dim ilVff As Integer
    smPledgeByEvent = "N"
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        Exit Sub
    End If
    If igGameVefCode <= 0 Then
        Exit Sub
    End If
    ilVff = gBinarySearchVff(igGameVefCode)
    If ilVff <> -1 Then
        smPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
        If smPledgeByEvent = "" Then
            smPledgeByEvent = "N"
        End If
    End If
End Sub
