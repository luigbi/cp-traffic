VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMerge 
   Caption         =   "Merge"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffMerge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   Begin VB.CommandButton cmcMerge 
      Caption         =   "&Merge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   15
      Top             =   4560
      Width           =   1335
   End
   Begin VB.PictureBox plcMerge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      ScaleHeight     =   1155
      ScaleWidth      =   8760
      TabIndex        =   13
      Top             =   2895
      Width           =   8820
      Begin VB.ListBox lbcMerge 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "AffMerge.frx":08CA
         Left            =   30
         List            =   "AffMerge.frx":08CC
         TabIndex        =   14
         Top             =   75
         Width           =   8685
      End
   End
   Begin VB.PictureBox pbcDnMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   4155
      ScaleHeight     =   90
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   2520
      Width           =   105
   End
   Begin VB.CommandButton cmcDnMove 
      Appearance      =   0  'Flat
      Caption         =   "M&ove   "
      Enabled         =   0   'False
      Height          =   285
      Left            =   3435
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2430
      Width           =   945
   End
   Begin VB.PictureBox pbcUpMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   5115
      ScaleHeight     =   90
      ScaleWidth      =   105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2520
      Width           =   105
   End
   Begin VB.PictureBox plcTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   4590
      ScaleHeight     =   1785
      ScaleWidth      =   4140
      TabIndex        =   6
      Top             =   405
      Width           =   4200
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTo 
         Height          =   1665
         Left            =   30
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   30
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   2937
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton cmcCancel 
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
      Height          =   375
      Left            =   4950
      TabIndex        =   2
      Top             =   4560
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.PictureBox pbcUsersLogFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   0
      Width           =   60
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   90
      Top             =   4635
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
   Begin VB.PictureBox plcFrom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   135
      ScaleHeight     =   1785
      ScaleWidth      =   4140
      TabIndex        =   3
      Top             =   405
      Width           =   4200
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFrom 
         Height          =   1665
         Left            =   30
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   2937
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton cmcUpMove 
      Appearance      =   0  'Flat
      Caption         =   "Mo&ve   "
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      Caption         =   "Replacement Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5760
      TabIndex        =   7
      Top             =   195
      Width           =   1815
   End
   Begin VB.Label lacFrom 
      Appearance      =   0  'Flat
      Caption         =   "Name to Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1440
      TabIndex        =   4
      Top             =   195
      Width           =   1560
   End
End
Attribute VB_Name = "frmMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmMerge - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private lmScrollTop As Long
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private lmLastClickedRow As Long
Private imAllowOkMsg As Integer
Private imShowGroup As Integer

Private hmMsg As Integer

Private smSelectedCodes() As String

Private imLastFromColSort As Integer
Private imLastFromSort As Integer
Private imPrevLastFromColSort As Integer

Private imLastToColSort As Integer
Private imLastToSort As Integer
Private imPrevLastToColSort As Integer

Private rst_Mkt As ADODB.Recordset
Private rst_Fmt As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset

'Grid Controls

Const NAMEINDEX = 0
Const NOREFINDEX = 1
Const GROUPINDEX = 2
Const SORTINDEX = 3
Const SELECTEDINDEX = 4
Const CODEINDEX = 5





Private Sub mClearGrid(grdCtrl As MSHFlexGrid)
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    gGrid_Clear grdCtrl, True
    'Set color within cells
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        For llCol = 0 To CODEINDEX Step 1
            grdCtrl.TextMatrix(llRow, llCol) = ""
        Next llCol
    Next llRow
End Sub

Private Sub cmcCancel_Click()
    Unload frmMerge
End Sub


Private Sub cmcDnMove_Click()
    Dim ilRet As Integer
    Dim llFromRow As Long
    Dim llToRow As Long
    Dim ilMove As Integer
    Dim slNames As String
    Dim slCodes As String
    Dim slGroup1 As String
    Dim slGroup2 As String
    Dim ilGroupsMatch As Integer
    
    ilMove = False
    ilGroupsMatch = True
    For llToRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
        If grdTo.TextMatrix(llToRow, SELECTEDINDEX) = "1" Then
            slGroup1 = Trim$(grdTo.TextMatrix(llToRow, GROUPINDEX))
            For llFromRow = grdFrom.FixedRows To grdFrom.Rows - 1 Step 1
                If grdFrom.TextMatrix(llFromRow, SELECTEDINDEX) = "1" Then
                    If llToRow = llFromRow Then
                        Beep
                        Exit Sub
                    End If
                    slGroup2 = Trim$(grdFrom.TextMatrix(llFromRow, GROUPINDEX))
                    If StrComp(slGroup1, slGroup2, vbTextCompare) <> 0 Then
                        ilGroupsMatch = False
                    End If
                    ilMove = True
                End If
            Next llFromRow
            Exit For
        End If
    Next llToRow
    If (ilMove) And (Not ilGroupsMatch) Then
        ilRet = MsgBox("Group Names Not Matching, Continue Anyway", vbQuestion + vbYesNo, "Merge")
        If ilRet = vbNo Then
            Exit Sub
        End If
    End If
    If ilMove Then
        slNames = "Replace:"
        llFromRow = grdFrom.FixedRows
        Do
            If grdFrom.TextMatrix(llFromRow, SELECTEDINDEX) = "1" Then
                slNames = slNames & Trim$(grdFrom.TextMatrix(llFromRow, NAMEINDEX)) & "(" & Trim$(grdFrom.TextMatrix(llFromRow, NOREFINDEX)) & ") "
                slCodes = slCodes & Trim$(grdFrom.TextMatrix(llFromRow, CODEINDEX)) & "|"
                grdFrom.RemoveItem llFromRow
                grdTo.RemoveItem llFromRow
            Else
                llFromRow = llFromRow + 1
            End If
        Loop While llFromRow < grdFrom.Rows
        slNames = slNames & " With:"
        For llToRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
            If grdTo.TextMatrix(llToRow, SELECTEDINDEX) = "1" Then
                slNames = slNames & Trim$(grdTo.TextMatrix(llToRow, NAMEINDEX)) & "(" & Trim$(grdTo.TextMatrix(llToRow, NOREFINDEX)) & ") "
                slCodes = slCodes & Trim$(grdTo.TextMatrix(llToRow, CODEINDEX))
                grdTo.RemoveItem llToRow
                grdFrom.RemoveItem llToRow
                Exit For
            End If
        Next llToRow
        lbcMerge.AddItem slNames
        smSelectedCodes(UBound(smSelectedCodes)) = slCodes
        lbcMerge.ItemData(lbcMerge.NewIndex) = UBound(smSelectedCodes)
        ReDim Preserve smSelectedCodes(0 To UBound(smSelectedCodes) + 1) As String
    End If
    mSetCommands
End Sub

Private Sub cmcMerge_Click()
    Dim ilRet As Integer
    Dim ilMerge As Integer
    Dim ilLoop As Integer
    Dim ilNoItems As Integer
    Dim llNoRowsRemoved As Long
    Dim llReplaceMktCode As Long
    
    If igMergeType = 0 Then
        ilRet = MsgBox("Continue with Market Merge", vbQuestion + vbYesNo, "Merge Markets")
        If ilRet = vbNo Then
            Exit Sub
        End If
    ElseIf igMergeType = 1 Then
        ilRet = MsgBox("Continue with Format Merge", vbQuestion + vbYesNo, "Merge Formats")
        If ilRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    On Error GoTo ErrHand:
    llNoRowsRemoved = 0
    If igMergeType = 0 Then
        gLogMsgWODT "O", hmMsg, sgMsgDirectory & "MergeMarket.Txt"
        gLogMsgWODT "W", hmMsg, "Merge Markets Started on " & Format(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    If igMergeType = 1 Then
        gLogMsgWODT "O", hmMsg, sgMsgDirectory & "MergeFormat.Txt"
        gLogMsgWODT "W", hmMsg, "Merge Formats Started on " & Format(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    For ilMerge = 0 To lbcMerge.ListCount - 1 Step 1
        ReDim slCodes(0 To 0) As String
        ilNoItems = 1
        Do
            ilRet = gParseItem(smSelectedCodes(ilMerge), ilNoItems, "|", slCodes(UBound(slCodes)))
            If ilRet <> CSI_MSG_NONE Then
                Exit Do
            End If
            ReDim Preserve slCodes(0 To UBound(slCodes) + 1) As String
            ilNoItems = ilNoItems + 1
        Loop While ilRet = CSI_MSG_NONE
        gLogMsgWODT "W", hmMsg, Trim$(smSelectedCodes(ilMerge))
        ilNoItems = ilNoItems - 1
        llNoRowsRemoved = llNoRowsRemoved - ilNoItems - 1
        llReplaceMktCode = Val(slCodes(ilNoItems - 1))
        For ilLoop = 0 To ilNoItems - 2 Step 1
            If igMergeType = 0 Then
                ilRet = mMergeMarket(Val(slCodes(ilLoop)), llReplaceMktCode)
            End If
            If igMergeType = 1 Then
                ilRet = mMergeFormat(Val(slCodes(ilLoop)), llReplaceMktCode)
            End If
        Next ilLoop
    Next ilMerge
    If igMergeType = 0 Then
        gLogMsgWODT "W", hmMsg, "Merge Markets Completed on " & Format(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    If igMergeType = 1 Then
        gLogMsgWODT "W", hmMsg, "Merge Formats Completed on " & Format(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    gLogMsgWODT "C", hmMsg, ""
    grdFrom.Rows = grdFrom.Rows - llNoRowsRemoved
    mClearGrid grdFrom
    grdTo.Rows = grdTo.Rows - llNoRowsRemoved
    lbcMerge.Clear
    mClearGrid grdTo
    mPopulate
    mSetCommands
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMerge-cmcMerge"
End Sub

Private Sub cmcUpMove_Click()
    ReDim slCodes(0 To 0) As String
    Dim ilMerge As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilItemNo As Integer
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFirstColSort As Integer
    Dim ilSecondColSort As Integer
    
    On Error GoTo ErrHand:
    ilFirstColSort = imPrevLastFromColSort
    ilSecondColSort = imLastFromColSort
    If lbcMerge.ListIndex >= 0 Then
        ilIndex = lbcMerge.ListIndex
        ilItemNo = 1
        Do
            ilRet = gParseItem(smSelectedCodes(ilIndex), ilItemNo, "|", slCodes(UBound(slCodes)))
            If ilRet <> CSI_MSG_NONE Then
                Exit Do
            End If
            ReDim Preserve slCodes(0 To UBound(slCodes) + 1) As String
            ilItemNo = ilItemNo + 1
        Loop While ilRet = CSI_MSG_NONE
        ilItemNo = ilItemNo - 1
        For ilMerge = lbcMerge.ListIndex + 1 To lbcMerge.ListCount - 1 Step 1
            lbcMerge.ItemData(ilMerge) = lbcMerge.ItemData(ilMerge) - 1
        Next ilMerge
        lbcMerge.RemoveItem lbcMerge.ListIndex
        For ilLoop = ilIndex To UBound(smSelectedCodes) - 1 Step 1
            smSelectedCodes(ilLoop) = smSelectedCodes(ilLoop + 1)
        Next ilLoop
        ReDim Preserve smSelectedCodes(0 To UBound(smSelectedCodes) - 1) As String
        'Find first blank row
        llRow = grdFrom.Rows - 1
        Do
            If Trim$(grdFrom.TextMatrix(llRow, NAMEINDEX)) <> "" Then
                Exit Do
            End If
            llRow = llRow - 1
        Loop While llRow >= grdFrom.FixedRows
        llRow = llRow + 1
        'Restore items to List
        grdFrom.Redraw = False
        grdTo.Redraw = False
        For ilLoop = 0 To UBound(slCodes) - 1 Step 1
            If llRow >= grdFrom.Rows Then
                grdFrom.AddItem ""
            End If
            If llRow >= grdTo.Rows Then
                grdTo.AddItem ""
            End If
            For llCol = NAMEINDEX To GROUPINDEX Step 1
                grdFrom.Row = llRow
                grdFrom.Col = llCol
                grdFrom.CellBackColor = vbWhite 'LIGHTBLUE
                grdFrom.CellForeColor = vbBlue
                grdTo.Row = llRow
                grdTo.Col = llCol
                grdTo.CellBackColor = vbWhite 'LIGHTBLUE
                grdTo.CellForeColor = vbBlue
            Next llCol
            If igMergeType = 0 Then
                SQLQuery = "SELECT * FROM MKT WHERE mktCode = " & slCodes(ilLoop)
                Set rst_Mkt = gSQLSelectCall(SQLQuery)
                If Not rst_Mkt.EOF Then
                    grdFrom.TextMatrix(llRow, NAMEINDEX) = Trim$(rst_Mkt!mktName)
                    SQLQuery = "SELECT Count(shttMktCode) FROM SHTT Where shttMktCode = " & rst_Mkt!mktCode
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    If Not rst_Shtt.EOF Then
                        grdFrom.TextMatrix(llRow, NOREFINDEX) = rst_Shtt(0).Value
                    Else
                        grdFrom.TextMatrix(llRow, NOREFINDEX) = ""
                    End If
                    grdFrom.TextMatrix(llRow, GROUPINDEX) = Trim$(rst_Mkt!mktGroupName)
                    grdFrom.TextMatrix(llRow, CODEINDEX) = rst_Mkt!mktCode
                    
                    grdTo.TextMatrix(llRow, NAMEINDEX) = grdFrom.TextMatrix(llRow, NAMEINDEX)
                    grdTo.TextMatrix(llRow, NOREFINDEX) = grdFrom.TextMatrix(llRow, NOREFINDEX)
                    grdTo.TextMatrix(llRow, GROUPINDEX) = grdFrom.TextMatrix(llRow, GROUPINDEX)
                    grdTo.TextMatrix(llRow, CODEINDEX) = grdFrom.TextMatrix(llRow, CODEINDEX)
                End If
            End If
            If igMergeType = 1 Then
                SQLQuery = "SELECT * FROM FMT_Station_Format WHERE fmtCode = " & slCodes(ilLoop)
                Set rst_Fmt = gSQLSelectCall(SQLQuery)
                If Not rst_Fmt.EOF Then
                    'Name
                    grdFrom.TextMatrix(llRow, NAMEINDEX) = Trim$(rst_Fmt!FmtName)
                    SQLQuery = "SELECT Count(shttFmtCode) FROM SHTT Where shttFmtCode = " & rst_Fmt!FmtCode
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    If Not rst_Shtt.EOF Then
                        grdFrom.TextMatrix(llRow, NOREFINDEX) = rst_Shtt(0).Value
                    Else
                        grdFrom.TextMatrix(llRow, NOREFINDEX) = ""
                    End If
                    grdFrom.TextMatrix(llRow, GROUPINDEX) = Trim$(rst_Fmt!fmtGroupName)
                    grdFrom.TextMatrix(llRow, CODEINDEX) = rst_Fmt!FmtCode
                    
                    grdTo.TextMatrix(llRow, NAMEINDEX) = grdFrom.TextMatrix(llRow, NAMEINDEX)
                    grdTo.TextMatrix(llRow, NOREFINDEX) = grdFrom.TextMatrix(llRow, NOREFINDEX)
                    grdTo.TextMatrix(llRow, GROUPINDEX) = grdFrom.TextMatrix(llRow, GROUPINDEX)
                    grdTo.TextMatrix(llRow, CODEINDEX) = grdFrom.TextMatrix(llRow, CODEINDEX)
                End If
            End If
            llRow = llRow + 1
        Next ilLoop
        If igMergeType = 0 Then
            On Error Resume Next
            rst_Mkt.Close
        End If
        If igMergeType = 1 Then
            On Error Resume Next
            rst_Fmt.Close
        End If
        imLastFromColSort = -1
        imLastFromSort = -1
        imPrevLastFromColSort = -1
        imLastToColSort = -1
        imLastToSort = -1
        imPrevLastToColSort = -1
        mSortCol ilFirstColSort, grdFrom, imLastFromColSort, imLastFromSort, imPrevLastFromColSort
        mSortCol ilSecondColSort, grdFrom, imLastFromColSort, imLastFromSort, imPrevLastFromColSort
        grdFrom.Row = 0
        grdFrom.Col = CODEINDEX
        grdFrom.Redraw = True
        mSortCol ilFirstColSort, grdTo, imLastToColSort, imLastToSort, imPrevLastToColSort
        mSortCol ilSecondColSort, grdTo, imLastToColSort, imLastToSort, imPrevLastToColSort
        grdTo.Row = 0
        grdTo.Col = CODEINDEX
        grdTo.Redraw = True
    End If
    On Error Resume Next
    rst_Shtt.Close
    mSetCommands
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMerge-cmcUpMove"
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        mPopulate
        mSetGridColumns grdFrom
        mSetGridTitles grdFrom
        gGrid_IntegralHeight grdFrom
        gGrid_FillWithRows grdFrom
        mSetGridColumns grdTo
        mSetGridTitles grdTo
        gGrid_IntegralHeight grdTo
        gGrid_FillWithRows grdTo
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
    gSetFonts frmMerge
    pbcDnMove.FontName = "Monotype Sorts"
    pbcDnMove.FontBold = False
    pbcDnMove.FontSize = 8
    pbcUpMove.FontName = "Monotype Sorts"
    pbcUpMove.FontBold = False
    pbcUpMove.FontSize = 8
    lbcMerge.Height = plcMerge.Height - 120
    gCenterForm frmMerge
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    mInit
    Screen.MousePointer = vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase smSelectedCodes
    rst_Mkt.Close
    rst_Fmt.Close
    rst_Shtt.Close
    Set frmMerge = Nothing
End Sub

Private Sub grdFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    'If (Shift And CTRLMASK) > 0 Then
    '    imCtrlKey = True
    'Else
    '    imCtrlKey = False
    'End If
    'If (Shift And SHIFTMASK) > 0 Then
    '    imShiftKey = True
    'Else
    '    imShiftKey = False
    'End If
End Sub

Private Sub grdFrom_KeyUp(KeyCode As Integer, Shift As Integer)
    'imCtrlKey = False
    'imShiftKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    lmLastClickedRow = -1
    lmScrollTop = grdFrom.FixedRows
    imLastFromColSort = -1
    imLastFromSort = -1
    imPrevLastFromColSort = -1
    imLastToColSort = -1
    imLastToSort = -1
    imPrevLastToColSort = -1
    ReDim smSelectedCodes(0 To 0) As String
    mClearGrid grdFrom
    mClearGrid grdTo
    mSetCommands
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMerge-mInit"
End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    
    On Error GoTo ErrHand:
    imShowGroup = False
    grdFrom.Redraw = False
    grdFrom.Row = 0
    For llCol = NAMEINDEX To GROUPINDEX Step 1
        grdFrom.Col = llCol
        grdFrom.CellBackColor = LIGHTBLUE
    Next llCol
    llRow = grdFrom.FixedRows
    grdTo.Redraw = False
    grdTo.Row = 0
    For llCol = NAMEINDEX To GROUPINDEX Step 1
        grdTo.Col = llCol
        grdTo.CellBackColor = LIGHTBLUE
    Next llCol
    
    If igMergeType = 0 Then
        SQLQuery = "SELECT * FROM MKT"
        Set rst_Mkt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Mkt.EOF
            If llRow >= grdFrom.Rows Then
                grdFrom.AddItem ""
            End If
            If llRow >= grdTo.Rows Then
                grdTo.AddItem ""
            End If
            For llCol = NAMEINDEX To GROUPINDEX Step 1
                grdFrom.Row = llRow
                grdFrom.Col = llCol
                grdFrom.CellBackColor = vbWhite 'LIGHTBLUE
                grdFrom.CellForeColor = vbBlue
                grdTo.Row = llRow
                grdTo.Col = llCol
                grdTo.CellBackColor = vbWhite 'LIGHTBLUE
                grdTo.CellForeColor = vbBlue
            Next llCol
            'Name
            grdFrom.TextMatrix(llRow, NAMEINDEX) = Trim$(rst_Mkt!mktName)
            SQLQuery = "SELECT Count(shttMktCode) FROM SHTT Where shttMktCode = " & rst_Mkt!mktCode
            Set rst_Shtt = gSQLSelectCall(SQLQuery)
            If Not rst_Shtt.EOF Then
                grdFrom.TextMatrix(llRow, NOREFINDEX) = rst_Shtt(0).Value
            Else
                grdFrom.TextMatrix(llRow, NOREFINDEX) = ""
            End If
            If Trim$(rst_Mkt!mktGroupName) <> "" Then
                imShowGroup = True
            End If
            grdFrom.TextMatrix(llRow, GROUPINDEX) = Trim$(rst_Mkt!mktGroupName)
            grdFrom.TextMatrix(llRow, CODEINDEX) = rst_Mkt!mktCode
            
            grdTo.TextMatrix(llRow, NAMEINDEX) = grdFrom.TextMatrix(llRow, NAMEINDEX)
            grdTo.TextMatrix(llRow, NOREFINDEX) = grdFrom.TextMatrix(llRow, NOREFINDEX)
            grdTo.TextMatrix(llRow, GROUPINDEX) = grdFrom.TextMatrix(llRow, GROUPINDEX)
            grdTo.TextMatrix(llRow, CODEINDEX) = grdFrom.TextMatrix(llRow, CODEINDEX)
            llRow = llRow + 1
            rst_Mkt.MoveNext
        Loop
        On Error Resume Next
        rst_Mkt.Close
    End If
    
    If igMergeType = 1 Then
        SQLQuery = "SELECT * FROM FMT_Station_Format"
        Set rst_Fmt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Fmt.EOF
            If llRow >= grdFrom.Rows Then
                grdFrom.AddItem ""
            End If
            If llRow >= grdTo.Rows Then
                grdTo.AddItem ""
            End If
            For llCol = NAMEINDEX To GROUPINDEX Step 1
                grdFrom.Row = llRow
                grdFrom.Col = llCol
                grdFrom.CellBackColor = vbWhite 'LIGHTBLUE
                grdFrom.CellForeColor = vbBlue
                grdTo.Row = llRow
                grdTo.Col = llCol
                grdTo.CellBackColor = vbWhite 'LIGHTBLUE
                grdTo.CellForeColor = vbBlue
            Next llCol
            'Name
            grdFrom.TextMatrix(llRow, NAMEINDEX) = Trim$(rst_Fmt!FmtName)
            SQLQuery = "SELECT Count(shttFmtCode) FROM SHTT Where shttFmtCode = " & rst_Fmt!FmtCode
            Set rst_Shtt = gSQLSelectCall(SQLQuery)
            If Not rst_Shtt.EOF Then
                grdFrom.TextMatrix(llRow, NOREFINDEX) = rst_Shtt(0).Value
            Else
                grdFrom.TextMatrix(llRow, NOREFINDEX) = ""
            End If
            If Trim$(rst_Fmt!fmtGroupName) <> "" Then
                imShowGroup = True
            End If
            grdFrom.TextMatrix(llRow, GROUPINDEX) = Trim$(rst_Fmt!fmtGroupName)
            grdFrom.TextMatrix(llRow, CODEINDEX) = rst_Fmt!FmtCode
            
            grdTo.TextMatrix(llRow, NAMEINDEX) = grdFrom.TextMatrix(llRow, NAMEINDEX)
            grdTo.TextMatrix(llRow, NOREFINDEX) = grdFrom.TextMatrix(llRow, NOREFINDEX)
            grdTo.TextMatrix(llRow, GROUPINDEX) = grdFrom.TextMatrix(llRow, GROUPINDEX)
            grdTo.TextMatrix(llRow, CODEINDEX) = grdFrom.TextMatrix(llRow, CODEINDEX)
            llRow = llRow + 1
            rst_Fmt.MoveNext
        Loop
        On Error Resume Next
        rst_Fmt.Close
    End If
    On Error Resume Next
    rst_Shtt.Close
    mSortCol NAMEINDEX, grdFrom, imLastFromColSort, imLastFromSort, imPrevLastFromColSort
    mSortCol GROUPINDEX, grdFrom, imLastFromColSort, imLastFromSort, imPrevLastFromColSort
    grdFrom.Row = 0
    grdFrom.Col = CODEINDEX
    grdFrom.Redraw = True
    mSortCol NAMEINDEX, grdTo, imLastToColSort, imLastToSort, imPrevLastToColSort
    mSortCol GROUPINDEX, grdTo, imLastToColSort, imLastToSort, imPrevLastToColSort
    grdTo.Row = 0
    grdTo.Col = CODEINDEX
    grdTo.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMerge-mPopulate"
End Sub

Private Sub mSetGridColumns(grdCtrl As MSHFlexGrid)
    Dim ilCol As Integer
    
    grdCtrl.ColWidth(CODEINDEX) = 0
    grdCtrl.ColWidth(SELECTEDINDEX) = 0
    grdCtrl.ColWidth(SORTINDEX) = 0
    grdCtrl.ColWidth(NAMEINDEX) = grdCtrl.Width * 0.5
    grdCtrl.ColWidth(NOREFINDEX) = grdCtrl.Width * 0.1
    If imShowGroup Then
        grdCtrl.ColWidth(GROUPINDEX) = grdCtrl.Width * 0.3
    Else
        grdCtrl.ColWidth(GROUPINDEX) = 0
    End If
    grdCtrl.ColWidth(NAMEINDEX) = grdCtrl.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To GROUPINDEX Step 1
        If ilCol <> NAMEINDEX Then
            grdCtrl.ColWidth(NAMEINDEX) = grdCtrl.ColWidth(NAMEINDEX) - grdCtrl.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdCtrl
End Sub

Private Sub mSetGridTitles(grdCtrl As MSHFlexGrid)
    'Set column titles
    grdCtrl.TextMatrix(0, NAMEINDEX) = "Name"
    grdCtrl.TextMatrix(0, NOREFINDEX) = "# Ref."
    grdCtrl.TextMatrix(0, GROUPINDEX) = "Group Name"
    'Set height of grid

End Sub

Private Sub mSortCol(ilCol As Integer, grdCtrl As MSHFlexGrid, ilLastColSort As Integer, ilLastSort As Integer, ilPrevLastColSort As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        slStr = Trim$(grdCtrl.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            If ilCol = NOREFINDEX Then
                slStr = grdCtrl.TextMatrix(llRow, NOREFINDEX)
                Do While Len(slStr) < 4
                    slStr = "0" & slStr
                Loop
                slSort = slStr
            Else
                slSort = UCase$(Trim$(grdCtrl.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdCtrl.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> ilLastColSort) Or ((ilCol = ilLastColSort) And (ilLastSort = flexSortStringNoCaseDescending)) Then
                'Ascending
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCtrl.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                'Descending
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCtrl.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    ilPrevLastColSort = ilLastColSort
    If ilCol = ilLastColSort Then
        ilLastColSort = SORTINDEX
    Else
        ilLastColSort = -1
        ilLastSort = -1
    End If
    gGrid_SortByCol grdCtrl, NAMEINDEX, SORTINDEX, ilLastColSort, ilLastSort
    ilLastColSort = ilCol
End Sub

Private Sub grdFrom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long
    
    For llRow = 0 To grdFrom.FixedRows - 1 Step 1
        llHeight = grdFrom.RowHeight(llRow)
    Next llRow
    grdFrom.ToolTipText = ""
    If Y <= llHeight Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdFrom, X, Y, llRow, llCol)
    'If (llCol = EMAILFLAGINDEX) Then
    '    grdFrom.ToolTipText = Trim$(grdFrom.TextMatrix(llRow, EMAILVALUEINDEX))
    'End If
    'If (llCol = USERNAMEINDEX) Then
    '    grdFrom.ToolTipText = Trim$(grdFrom.TextMatrix(llRow, PCMACINDEX))
    'End If
End Sub

Private Sub grdFrom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdFrom.RowHeight(0) Then
        grdFrom.Col = grdFrom.MouseCol
        grdTo.Col = grdFrom.Col
        mSortCol grdFrom.Col, grdFrom, imLastFromColSort, imLastFromSort, imPrevLastFromColSort
        grdFrom.Row = 0
        grdFrom.Col = CODEINDEX
        mSortCol grdTo.Col, grdTo, imLastToColSort, imLastToSort, imPrevLastToColSort
        grdTo.Row = 0
        grdTo.Col = CODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdFrom, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdFrom.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdFrom.FixedRows Then
        If grdFrom.TextMatrix(llCurrentRow, NAMEINDEX) <> "" Then
            'grdFrom.TopRow = lmScrollTop
            'llTopRow = grdFrom.TopRow
            'If (Shift And CTRLMASK) > 0 Then
            '    If grdFrom.TextMatrix(grdFrom.Row, CODEINDEX) <> "0" Then
            '        If grdFrom.TextMatrix(grdFrom.Row, SELECTEDINDEX) <> 1 Then
            '            grdFrom.TextMatrix(grdFrom.Row, SELECTEDINDEX) = 1
            '        Else
            '            grdFrom.TextMatrix(grdFrom.Row, SELECTEDINDEX) = 0
            '        End If
            '        mPaintRowColor grdFrom, grdFrom.Row
            '    End If
            'Else
            '    For llRow = grdFrom.FixedRows To grdFrom.Rows - 1 Step 1
            '        If grdFrom.TextMatrix(llRow, NAMEINDEX) <> "" Then
            '            grdFrom.TextMatrix(llRow, SELECTEDINDEX) = "0"
            '            If grdFrom.TextMatrix(llRow, CODEINDEX) <> "0" Then
            '                If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
            '                    If llRow = llCurrentRow Then
            '                        grdFrom.TextMatrix(llRow, SELECTEDINDEX) = "1"
            '                    Else
            '                        grdFrom.TextMatrix(llRow, SELECTEDINDEX) = "0"
            '                    End If
            '                ElseIf lmLastClickedRow < llCurrentRow Then
            '                    If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
            '                        grdFrom.TextMatrix(llRow, SELECTEDINDEX) = "1"
            '                    End If
            '                Else
            '                    If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
            '                        grdFrom.TextMatrix(llRow, SELECTEDINDEX) = "1"
            '                    End If
            '                End If
            '                mPaintRowColor grdFrom, llRow
            '            End If
            '        End If
            '    Next llRow
            '    grdFrom.TopRow = llTopRow
            '    grdFrom.Row = llCurrentRow
            'End If
            'lmLastClickedRow = llCurrentRow
            If grdFrom.TextMatrix(llCurrentRow, SELECTEDINDEX) <> "1" Then
                grdFrom.TextMatrix(llCurrentRow, SELECTEDINDEX) = "1"
            Else
                grdFrom.TextMatrix(llCurrentRow, SELECTEDINDEX) = "0"
            End If
            mPaintRowColor grdFrom, llCurrentRow
        End If
    End If
    mSetCommands
End Sub

Private Sub grdFrom_Scroll()
    'lmScrollTop = grdFrom.TopRow
End Sub

Private Sub mPaintRowColor(grdCtrl As MSHFlexGrid, llRow As Long)
    Dim llCol As Long
    
    grdCtrl.Row = llRow
    For llCol = NAMEINDEX To GROUPINDEX Step 1
        grdCtrl.Col = llCol
        If grdCtrl.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdCtrl.CellBackColor = vbWhite
        Else
            grdCtrl.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub



Private Sub mSetCommands()
    Dim llFromRow As Long
    Dim llToRow As Long
    
    If lbcMerge.ListCount > 0 Then
        cmcMerge.Enabled = True
        If lbcMerge.ListIndex >= 0 Then
            cmcUpMove.Enabled = True
        Else
            cmcUpMove.Enabled = False
        End If
    Else
        cmcMerge.Enabled = False
        cmcUpMove.Enabled = False
    End If
    cmcDnMove.Enabled = False
    For llToRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
        If grdTo.TextMatrix(llToRow, SELECTEDINDEX) = "1" Then
            For llFromRow = grdFrom.FixedRows To grdFrom.Rows - 1 Step 1
                If grdFrom.TextMatrix(llFromRow, SELECTEDINDEX) = "1" Then
                    cmcDnMove.Enabled = True
                    Exit Sub
                End If
            Next llFromRow
            Exit Sub
        End If
    Next llToRow
    
End Sub


Private Sub grdTo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdTo.RowHeight(0) Then
        grdFrom.Col = grdTo.MouseCol
        grdTo.Col = grdFrom.Col
        mSortCol grdFrom.Col, grdFrom, imLastFromColSort, imLastFromSort, imPrevLastFromColSort
        grdFrom.Row = 0
        grdFrom.Col = CODEINDEX
        mSortCol grdTo.Col, grdTo, imLastToColSort, imLastToSort, imPrevLastToColSort
        grdTo.Row = 0
        grdTo.Col = CODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdTo, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdTo.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdTo.FixedRows Then
        If grdTo.TextMatrix(llCurrentRow, NAMEINDEX) <> "" Then
            For llRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
                If grdTo.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                    If llRow <> llCurrentRow Then
                        grdTo.TextMatrix(llRow, SELECTEDINDEX) = "0"
                        mPaintRowColor grdTo, llRow
                    End If
                    Exit For
                End If
            Next llRow
            If grdTo.TextMatrix(llCurrentRow, SELECTEDINDEX) <> "1" Then
                grdTo.TextMatrix(llCurrentRow, SELECTEDINDEX) = "1"
            Else
                grdTo.TextMatrix(llCurrentRow, SELECTEDINDEX) = "0"
            End If
            mPaintRowColor grdTo, llCurrentRow
        End If
    End If
    mSetCommands
End Sub

Private Sub lbcMerge_Click()
    mSetCommands
End Sub

Private Sub pbcDnMove_Paint()
    pbcDnMove.CurrentX = 0
    pbcDnMove.CurrentY = 0
    pbcDnMove.Print "t"
End Sub

Private Sub pbcUpMove_Paint()
    pbcUpMove.CurrentX = 0
    pbcUpMove.CurrentY = 0
    pbcUpMove.Print "s"
End Sub

Private Function mMergeMarket(llRemoveMktCode As Long, llReplaceMktCode As Long) As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    SQLQuery = "Update shtt Set shttMktCode = " & llReplaceMktCode & " Where shttMktCode = " & llRemoveMktCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeMarket"
        mMergeMarket = False
        Exit Function
    End If
    '11/26/17
    gFileChgdUpdate "shtt.mkd", True
    SQLQuery = "Update mat Set matMktCode = " & llReplaceMktCode & " Where matMktCode = " & llRemoveMktCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeMarket"
        mMergeMarket = False
        Exit Function
    End If
    SQLQuery = "Update mgt Set mgtMktCode = " & llReplaceMktCode & " Where mgtMktCode = " & llRemoveMktCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeMarket"
        mMergeMarket = False
        Exit Function
    End If
    ilRet = gUpdateRegions("M", llRemoveMktCode, llReplaceMktCode, "MergeMarket.Txt")
    If Not ilRet Then
        mMergeMarket = False
        Exit Function
    End If
    SQLQuery = "DELETE FROM mkt WHERE mktCode = " & llRemoveMktCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeMarket"
        mMergeMarket = False
        Exit Function
    End If
    mMergeMarket = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMerge-mMergeMarket"
    mMergeMarket = False
End Function

Private Function mMergeFormat(llRemoveFmtCode As Long, llReplaceFmtCode As Long) As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    SQLQuery = "Update shtt Set shttFmtCode = " & llReplaceFmtCode & " Where shttFmtCode = " & llRemoveFmtCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeFormat"
        mMergeFormat = False
        Exit Function
    End If
    '11/26/17
    gFileChgdUpdate "shtt.mkd", True
    SQLQuery = "Update flt Set fltIntFmtCode = " & llReplaceFmtCode & " Where fltIntFmtCode = " & llRemoveFmtCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeFormat"
        mMergeFormat = False
        Exit Function
    End If
    ilRet = gUpdateRegions("F", llRemoveFmtCode, llReplaceFmtCode, "MergeFormat.Txt")
    If Not ilRet Then
        mMergeFormat = False
        Exit Function
    End If
    SQLQuery = "DELETE FROM FMT_Station_Format WHERE fmtCode = " & llRemoveFmtCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Merge-mMergeFormat"
        mMergeFormat = False
        Exit Function
    End If
    mMergeFormat = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmMerge-mMergeMarket"
    mMergeFormat = False
End Function
