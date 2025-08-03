VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCP 
   Caption         =   "Certificate of Performance"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "AffCPLog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9240
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   15
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6270
      Width           =   15
   End
   Begin VB.OptionButton rbcVeh 
      Caption         =   "All Veh"
      Height          =   285
      Index           =   1
      Left            =   8145
      TabIndex        =   6
      Top             =   765
      Width           =   1050
   End
   Begin VB.OptionButton rbcVeh 
      Caption         =   "Active Veh"
      Height          =   240
      Index           =   0
      Left            =   6705
      TabIndex        =   5
      Top             =   765
      Value           =   -1  'True
      Width           =   1545
   End
   Begin VB.Frame frcPrintCPNotCarried 
      Caption         =   "Spots Not Carried"
      Height          =   525
      Left            =   225
      TabIndex        =   43
      Top             =   5955
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton rbcShowAll 
         Caption         =   "Suppress"
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   45
         Top             =   225
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton rbcShowAll 
         Caption         =   "Show"
         Height          =   255
         Index           =   2
         Left            =   1425
         TabIndex        =   44
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Reprint"
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   1
      Left            =   7440
      TabIndex        =   20
      Top             =   2685
      Visible         =   0   'False
      Width           =   8460
      Begin VB.Frame frcRePrintCPNotCarried 
         Caption         =   "Spots Not Carried"
         Height          =   960
         Left            =   75
         TabIndex        =   40
         Top             =   2895
         Width           =   2640
         Begin VB.OptionButton rbcShowAll 
            Caption         =   "Show"
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   42
            Top             =   570
            Width           =   2085
         End
         Begin VB.OptionButton rbcShowAll 
            Caption         =   "Suppress"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   41
            Top             =   255
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.TextBox txtDays 
         Height          =   360
         Left            =   1470
         TabIndex        =   24
         Top             =   960
         Width           =   1260
      End
      Begin VB.CheckBox ckcCover 
         Caption         =   "Cover page only"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   2625
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame frcZone 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Reprint Zone"
         ForeColor       =   &H80000008&
         Height          =   1365
         Left            =   105
         TabIndex        =   28
         Top             =   1395
         Visible         =   0   'False
         Width           =   2400
         Begin VB.CheckBox chkZone 
            Caption         =   "EST"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   29
            Top             =   225
            Value           =   1  'Checked
            Width           =   780
         End
         Begin VB.CheckBox chkZone 
            Caption         =   "CST"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   30
            Top             =   495
            Value           =   1  'Checked
            Width           =   795
         End
         Begin VB.CheckBox chkZone 
            Caption         =   "MST"
            Height          =   225
            Index           =   2
            Left            =   105
            TabIndex        =   31
            Top             =   750
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.CheckBox chkZone 
            Caption         =   "PST"
            Height          =   225
            Index           =   3
            Left            =   105
            TabIndex        =   32
            Top             =   1035
            Value           =   1  'Checked
            Width           =   780
         End
      End
      Begin VB.Frame frcOrder 
         Caption         =   "Selection by"
         Height          =   1035
         Left            =   75
         TabIndex        =   25
         Top             =   1425
         Width           =   2625
         Begin VB.OptionButton optSort 
            Caption         =   "Station, then Vehicles"
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   27
            Top             =   600
            Width           =   2340
         End
         Begin VB.OptionButton optSort 
            Caption         =   "Vehicle, then Stations"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   26
            Top             =   285
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   195
         Left            =   5565
         TabIndex        =   38
         Top             =   3915
         Width           =   900
      End
      Begin VB.TextBox txtDate 
         Height          =   360
         Left            =   1470
         TabIndex        =   22
         Top             =   495
         Width           =   1260
      End
      Begin VB.ListBox lbcSingle 
         Height          =   3375
         ItemData        =   "AffCPLog.frx":08CA
         Left            =   2970
         List            =   "AffCPLog.frx":08CC
         TabIndex        =   35
         Top             =   495
         Width           =   2355
      End
      Begin VB.ListBox lbcMulti 
         Height          =   3375
         ItemData        =   "AffCPLog.frx":08CE
         Left            =   5550
         List            =   "AffCPLog.frx":08D0
         MultiSelect     =   2  'Extended
         TabIndex        =   37
         Top             =   510
         Width           =   2625
      End
      Begin VB.Label lacReDays 
         Caption         =   "Number Days"
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   1005
         Width           =   1455
      End
      Begin VB.Label lacReDate 
         Caption         =   "Reprint Date"
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lacTitle1 
         Alignment       =   2  'Center
         Caption         =   "Vehicle"
         Height          =   255
         Left            =   2940
         TabIndex        =   34
         Top             =   180
         Width           =   2385
      End
      Begin VB.Label lacTitle2 
         Alignment       =   2  'Center
         Caption         =   "Stations"
         Height          =   255
         Left            =   5550
         TabIndex        =   36
         Top             =   195
         Width           =   2595
      End
   End
   Begin VB.Frame frcDest 
      Caption         =   "Report Destination"
      Height          =   660
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   8760
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   1230
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   2145
         TabIndex        =   3
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffCPLog.frx":08D2
         Left            =   3075
         List            =   "AffCPLog.frx":08D4
         TabIndex        =   4
         Top             =   210
         Width           =   2595
      End
   End
   Begin VB.Frame frcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Print"
      ForeColor       =   &H80000008&
      Height          =   4560
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1185
      Width           =   8445
      Begin VB.PictureBox pbcDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2040
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1365
         Visible         =   0   'False
         Width           =   210
         Begin VB.CheckBox ckcDay 
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
            Left            =   15
            TabIndex        =   15
            Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
            Top             =   15
            Width           =   180
         End
      End
      Begin VB.TextBox txtDropdown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2355
         TabIndex        =   13
         Top             =   1260
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.PictureBox pbcLogSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   12
         Top             =   210
         Width           =   60
      End
      Begin VB.PictureBox pbcLogTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   90
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   16
         Top             =   3120
         Width           =   60
      End
      Begin VB.PictureBox pbcLogFocus 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   285
         ScaleHeight     =   90
         ScaleWidth      =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   60
      End
      Begin VB.PictureBox pbcArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   0
         Picture         =   "AffCPLog.frx":08D6
         ScaleHeight     =   165
         ScaleWidth      =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   450
         Visible         =   0   'False
         Width           =   90
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLog 
         Height          =   4035
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   6
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
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2415
      Top             =   6180
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6480
      FormDesignWidth =   9240
   End
   Begin VB.CommandButton cmdReprint 
      Caption         =   "Reprint"
      Height          =   375
      Left            =   3900
      TabIndex        =   19
      Top             =   6015
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   5625
      TabIndex        =   17
      Top             =   6030
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7185
      TabIndex        =   18
      Top             =   6030
      Width           =   1335
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4980
      Left            =   105
      TabIndex        =   7
      Top             =   840
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   8784
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Print Logs"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Reprint Logs"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Studio Logs"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lacProcess 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   105
      TabIndex        =   39
      Top             =   5775
      Width           =   8775
   End
End
Attribute VB_Name = "frmCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmCP - shows certificate of performance information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text
Private imIntegralSet As Integer
Private imTabIndex As Integer
Private imVefCode As Integer
Private imShfCode As Integer
Private imAllClick As Integer
Private smCntrNo As String
Private smChfType As String
Private tmCmmlSum() As CMMLSUM
Private imMaxDays As Integer
Private chfrst As ADODB.Recordset
Private smDate As String
Private tmCPInfo() As CPINFO
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private cprst As ADODB.Recordset
Private lstrst As ADODB.Recordset
Private imFirstTime As Integer
Private bFormWasAlreadyResized As Boolean
Private hmAst As Integer

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long
Private lmEnableCol As Long


Private Sub mLogSetShow()
    
    If (lmEnableRow >= grdLog.FixedRows) And (lmEnableRow < grdLog.Rows) Then
        'Set any field that that should only be set after user leaves the cell
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imShowGridBox = False
    pbcArrow.Visible = False
    pbcDay.Visible = False
    txtDropdown.Visible = False
End Sub

Private Sub mLogEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim iLoop As Integer
    Dim iIndex As Integer
    
    If (grdLog.Row >= grdLog.FixedRows) And (grdLog.Row < grdLog.Rows) And (grdLog.Col >= 0) And (grdLog.Col < grdLog.Cols) Then
        lmEnableRow = grdLog.Row
        lmEnableCol = grdLog.Col
        imShowGridBox = True
        pbcArrow.Move grdLog.Left - pbcArrow.Width - 15, grdLog.Top + grdLog.RowPos(grdLog.Row) + (grdLog.RowHeight(grdLog.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdLog.Col
            Case 0
                pbcDay.Move grdLog.Left + grdLog.ColPos(grdLog.Col) + 30, grdLog.Top + grdLog.RowPos(grdLog.Row) + 15, grdLog.ColWidth(grdLog.Col) - 30, grdLog.RowHeight(grdLog.Row) - 15
                If grdLog.Text = "4" Then
                    ckcDay.Value = vbChecked
                Else
                    ckcDay.Value = vbUnchecked
                End If
                pbcDay.Visible = True
                ckcDay.SetFocus
            Case 4  'Log or CP
                txtDropdown.Move grdLog.Left + grdLog.ColPos(grdLog.Col) + 30, grdLog.Top + grdLog.RowPos(grdLog.Row) + 15, grdLog.ColWidth(grdLog.Col) - 30, grdLog.RowHeight(grdLog.Row) - 15
                If grdLog.Text <> "Missing" Then
                    txtDropdown.Text = grdLog.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdLog.RowHeight(grdLog.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdLog.RowHeight(grdLog.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            Case 5  'Other
                txtDropdown.Move grdLog.Left + grdLog.ColPos(grdLog.Col) + 30, grdLog.Top + grdLog.RowPos(grdLog.Row) + 15, grdLog.ColWidth(grdLog.Col) - 30, grdLog.RowHeight(grdLog.Row) - 15
                If grdLog.Text <> "Missing" Then
                    txtDropdown.Text = grdLog.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdLog.RowHeight(grdLog.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdLog.RowHeight(grdLog.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
        End Select
    End If
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    gGrid_Clear grdLog, True
    For llRow = grdLog.FixedRows To grdLog.Rows - 1 Step 1
        For llCol = 1 To 3 Step 1
            grdLog.Row = llRow
            grdLog.Col = llCol
            grdLog.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llRow
End Sub

Private Sub mFill()
    Dim iLoop As Integer
    Dim sMarket As String
    
    On Error GoTo ErrHand
    chkAll.Value = 0
    If igCPOrLog = 0 Then
        lbcSingle.Clear
        lbcMulti.Clear
    Else
        lbcMulti.Clear
    End If
    If ((igCPOrLog = 0) And (optSort(0).Value)) Or (igCPOrLog = 1) Then
        For iLoop = 0 To UBound(tmCPInfo) - 1 Step 1
            If tmCPInfo(iLoop).iStatus = 0 Then
                If igCPOrLog = 0 Then
                    If rbcVeh(0).Value = True Then
                        If tmCPInfo(iLoop).sVefState = "A" Then
                            lbcSingle.AddItem Trim$(tmCPInfo(iLoop).sVefName)
                            lbcSingle.ItemData(lbcSingle.NewIndex) = tmCPInfo(iLoop).iVefCode
                        End If
                    Else
                        lbcSingle.AddItem Trim$(tmCPInfo(iLoop).sVefName)
                        lbcSingle.ItemData(lbcSingle.NewIndex) = tmCPInfo(iLoop).iVefCode
                    End If
                Else
                    If rbcVeh(0).Value = True Then
                        If tmCPInfo(iLoop).sVefState = "A" Then
                            lbcMulti.AddItem Trim$(tmCPInfo(iLoop).sVefName)
                            lbcMulti.ItemData(lbcMulti.NewIndex) = tmCPInfo(iLoop).iVefCode
                        End If
                    Else
                        lbcMulti.AddItem Trim$(tmCPInfo(iLoop).sVefName)
                        lbcMulti.ItemData(lbcMulti.NewIndex) = tmCPInfo(iLoop).iVefCode
                    End If
                End If
            End If
        Next iLoop
    Else
        'SQLQuery = "SELECT DISTINCT shttCallLetters, shttMarket, shttCode"
        'SQLQuery = SQLQuery + " FROM shtt, att"
        SQLQuery = "SELECT DISTINCT shttCallLetters, mktName, shttCode"
        SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, att"
        SQLQuery = SQLQuery + " WHERE (shttCode = attShfCode" & ")"
        'SQLQuery = SQLQuery + " ORDER BY shttCallLetters, shttMarket"
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters, mktName"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            'If IsNull(rst!shttMarket) = True Then
            '    sMarket = ""
            'Else
            '    sMarket = rst!shttMarket  'Trim$(rst!shttMarket)
            'End If
            'lbcSingle.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)  ', " & rst(1).Value & ""
            If IsNull(rst!mktName) = True Then
                sMarket = ""
                lbcSingle.AddItem Trim$(rst!shttCallLetters)
            Else
                sMarket = rst!mktName  'Trim$(rst!shttMarket)
                lbcSingle.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)  ', " & rst(1).Value & ""
            End If
            lbcSingle.ItemData(lbcSingle.NewIndex) = rst!shttCode
            rst.MoveNext
        Wend
    End If
    On Error GoTo 0
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-mFill"
End Sub


Private Sub cboFileType_GotFocus()
    mLogSetShow
End Sub

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcMulti.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcMulti.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcMulti.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If
End Sub

Private Sub ckcDay_Click()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llRow As Long
    
    If ckcDay.Value = vbChecked Then
        grdLog.Text = 4
    Else
        grdLog.Text = ""
    End If
    grdLog.TextMatrix(grdLog.Row, grdLog.Col) = grdLog.Text
End Sub

Private Sub cmdCancel_Click()
    Unload frmCP
End Sub

Private Sub cmdCancel_GotFocus()
    mLogSetShow
End Sub

Private Sub cmdGenerate_Click()

    Dim iRet As Integer
    
    lacProcess.Caption = ""
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            cmdCancel.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(4) <> "I" Then
            cmdCancel.SetFocus
            Exit Sub
        End If
    End If
    
    'igCPOrLog = 0 is a CP
    'igCPOrLog = 1 is a Log
    'imTabIndex = 1 is print CP
    'imTabIndex = 2 is a reprint of a CP

    If igCPOrLog = 0 Then
        If imTabIndex = 2 Then
            If mPrintPrePass(False) Then
                iRet = mCPGen(False)
            End If
        Else
            If mPrintPrePass(True) Then
                iRet = mCPGen(True)
            End If

        End If
    Else
        If imTabIndex = 3 Then      'Studio Log
            If mLogPrePass(False, True) Then
                iRet = mLogGen(False, False)
            End If
        ElseIf imTabIndex = 2 Then  'Reprint
            If mLogPrePass(False, True) Then
                iRet = mLogGen(False, True)
            End If
        Else
            If mLogPrePass(True, True) Then
                iRet = mLogGen(True, True) 'Print Log
            End If
        End If
    End If
    lacProcess.Caption = ""
    cmdCancel.Caption = "&Done"
End Sub

Private Sub cmdGenerate_GotFocus()
    mLogSetShow
End Sub

Private Sub cmdReprint_Click()
    'frmReprintCP.Show
    Dim iRet As Integer
    
    If igCPOrLog = 0 Then
        iRet = mCPGen(False)
    Else
        iRet = mLogGen(False, True)
    End If
End Sub

Private Sub cmdReprint_GotFocus()
    mLogSetShow
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        'Hide column 2
        If igCPOrLog = 0 Then
            grdLog.ColWidth(5) = 0
        Else
            grdLog.ColWidth(5) = grdLog.Width * 0.1
        End If
        grdLog.ColWidth(0) = grdLog.Width * 0.08
        grdLog.ColWidth(2) = grdLog.Width * 0.12
        grdLog.ColWidth(3) = grdLog.Width * 0.1
        grdLog.ColWidth(4) = grdLog.Width * 0.1
        grdLog.ColWidth(1) = grdLog.Width - grdLog.ColWidth(0) - grdLog.ColWidth(2) - grdLog.ColWidth(3) - grdLog.ColWidth(4) - grdLog.ColWidth(5) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdLog
        grdLog.TextMatrix(0, 0) = "Check"
        grdLog.TextMatrix(0, 1) = "Vehicle"
        If igCPOrLog = 0 Then
            grdLog.TextMatrix(0, 2) = "CP Date"
            grdLog.TextMatrix(0, 3) = "Cycle"
            grdLog.TextMatrix(0, 4) = "CP"
        Else
            grdLog.TextMatrix(0, 2) = "Log Date"
            grdLog.TextMatrix(0, 3) = "Cycle"
            grdLog.TextMatrix(0, 4) = "Log"
            grdLog.TextMatrix(0, 5) = "Other"
        End If
        gGrid_IntegralHeight grdLog
        mClearGrid
        mCPMain
        ckcDay.Height = 180
        ckcDay.Width = 180
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Visible = False
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmCP
    gCenterForm frmCP
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
   
   On Error GoTo ErrHand
   
    imAllClick = False
    bFormWasAlreadyResized = False
    
    imIntegralSet = False
    imTabIndex = 1
    If igCPOrLog = 0 Then
        frmCP.Caption = "Certificate of Performance - " & sgClientName
        TabStrip1.Tabs(1).Caption = "&Print Certificates of Performance"
        TabStrip1.Tabs(2).Caption = "&Reprint Certificates of Performance"
        TabStrip1.Tabs.Remove 3
        lbcSingle.Visible = True
        lacTitle2.Visible = True
        frcOrder.Visible = True
        frcZone.Visible = False
        frcPrintCPNotCarried.Visible = True
    Else
        frmCP.Caption = "Log - " & sgClientName
        TabStrip1.Tabs(1).Caption = "&Print Logs"
        TabStrip1.Tabs(2).Caption = "&Reprint Logs"
        TabStrip1.Tabs(3).Caption = "&Studio Logs"
        lbcSingle.Visible = False
        lacTitle2.Visible = False
        frcOrder.Visible = False
        frcZone.Visible = True
        frcPrintCPNotCarried.Visible = False
        frcRePrintCPNotCarried.Visible = False
        lbcMulti.Left = lbcSingle.Left
        chkAll.Left = lbcMulti.Left
        lbcMulti.Width = lbcSingle.Width + lbcMulti.Width
    End If
    imFirstTime = True
    
    imShowGridBox = False
    imFromArrow = False
    lmTopRow = -1
    lmEnableRow = -1
    
    cboFileType.Enabled = False
    gPopExportTypes cboFileType     '3-12-04
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
'    mCPMain
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in CP-Form Load: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub


Private Sub Form_Resize()
    If bFormWasAlreadyResized Then
        Exit Sub
    End If
    bFormWasAlreadyResized = True
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    TabStrip1.Left = frcDest.Left
    TabStrip1.Height = TabStrip1.ClientTop - TabStrip1.Top + (10 * frcTab(1).Height) / 9
    TabStrip1.Width = frcDest.Width
    frcTab(0).Move TabStrip1.ClientLeft, TabStrip1.ClientTop
    frcTab(1).Move TabStrip1.ClientLeft, TabStrip1.ClientTop
    frcTab(0).BorderStyle = 0
    frcTab(1).BorderStyle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    
    Erase tmCPInfo
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmCmmlSum
    
    Set frmCP = Nothing
End Sub



Private Sub grdLog_Click()
    Dim llRow As Long
    
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(4) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
'    If grdLog.Col >= grdLog.Cols Then
'        Exit Sub
'    End If
'    If (grdLog.Col >= 1) And (grdLog.Col <= 3) Then
'        pbcClickFocus.SetFocus
'        Exit Sub
'    End If
'    lmTopRow = grdLog.TopRow
'    llRow = grdLog.Row
'    mLogEnableBox
End Sub

Private Sub grdLog_EnterCell()
    mLogSetShow
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(4) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub grdLog_GotFocus()
    If grdLog.Col >= grdLog.Cols Then
        Exit Sub
    End If
    'grdLog_Click
End Sub

Private Sub grdLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdLog.TopRow
    grdLog.Redraw = False
End Sub

Private Sub grdLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(4) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdLog, X, Y)
    If Not ilFound Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdLog.Col >= grdLog.Cols Then
        Exit Sub
    End If
    If (grdLog.Col >= 1) And (grdLog.Col <= 3) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    lmTopRow = grdLog.TopRow
    llRow = grdLog.Row
    grdLog.Redraw = True
    mLogEnableBox
End Sub

Private Sub grdLog_Scroll()
    If grdLog.Redraw = False Then
        grdLog.Redraw = True
        grdLog.TopRow = lmTopRow
        grdLog.Refresh
        grdLog.Redraw = False
    End If
    If (imShowGridBox) And (grdLog.Row >= grdLog.FixedRows) And (grdLog.Col >= 0) And (grdLog.Col < grdLog.Cols) Then
        If grdLog.RowIsVisible(grdLog.Row) Then
            pbcArrow.Move grdLog.Left - pbcArrow.Width, grdLog.Top + grdLog.RowPos(grdLog.Row) + (grdLog.RowHeight(grdLog.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            If grdLog.Col = 0 Then
                pbcDay.Move grdLog.Left + grdLog.ColPos(grdLog.Col) + 30, grdLog.Top + grdLog.RowPos(grdLog.Row) + 30, grdLog.ColWidth(grdLog.Col) - 30, grdLog.RowHeight(grdLog.Row) - 30
                pbcDay.Visible = True
                ckcDay.SetFocus
            ElseIf grdLog.Col = 4 Then  'Date
                txtDropdown.Move grdLog.Left + grdLog.ColPos(grdLog.Col) + 30, grdLog.Top + grdLog.RowPos(grdLog.Row) + 15, grdLog.ColWidth(grdLog.Col) - 30, grdLog.RowHeight(grdLog.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            ElseIf grdLog.Col = 5 Then  'Time
                txtDropdown.Move grdLog.Left + grdLog.ColPos(grdLog.Col) + 30, grdLog.Top + grdLog.RowPos(grdLog.Row) + 15, grdLog.ColWidth(grdLog.Col) - 30, grdLog.RowHeight(grdLog.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            End If
        Else
            pbcLogFocus.SetFocus
            pbcDay.Visible = False
            txtDropdown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcLogFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub lbcMulti_Click()
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
End Sub

Private Sub lbcSingle_Click()
    Dim sDateRange As String
    Dim sMarket As String
    
    On Error GoTo ErrHand
    'D.S.
    lbcMulti.Clear
    chkAll.Value = 0

    If lbcSingle.ListIndex < 0 Then
        Exit Sub
    End If
    If txtDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If gIsDate(txtDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtDate.SetFocus
    Else
        smDate = Format(txtDate.Text, "m/d/yyyy")
    End If
    smDate = gObtainPrevMonday(smDate)
    Screen.MousePointer = vbHourglass
    If optSort(0).Value Then
        imVefCode = lbcSingle.ItemData(lbcSingle.ListIndex)
        
        'sDateRange = "(att.attOffAir >=" + smDate + ") And (att.attOnAir <=" + smDate + ")"
        sDateRange = "(attOffAir >= '" & Format$(smDate, sgSQLDateForm) & "') and (attDropDate >= '" & Format$(smDate, sgSQLDateForm) & "') And (attOnAir <= '" + Format$(smDate, sgSQLDateForm) & "')"
        
        'SQLQuery = "SELECT DISTINCT shttCallLetters, shttMarket, shttCode"
        'SQLQuery = SQLQuery + " FROM shtt, att" 'cptt"
        SQLQuery = "SELECT DISTINCT shttCallLetters, mktName, shttCode"
        SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, att" 'cptt"
        SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND shttCode = attShfCode"
        SQLQuery = SQLQuery & " AND " & sDateRange & ")"
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            'If IsNull(rst!shttMarket) = True Then
            '    sMarket = ""
            'Else
            '    sMarket = rst!shttMarket  'Trim$(rst!shttMarket)
            'End If
            'lbcMulti.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)  ', " & rst(1).Value & ""
            If IsNull(rst!mktName) = True Then
                sMarket = ""
                lbcMulti.AddItem Trim$(rst!shttCallLetters)
            Else
                sMarket = rst!mktName  'Trim$(rst!shttMarket)
                lbcMulti.AddItem Trim$(rst!shttCallLetters) & ", " & Trim$(sMarket)  ', " & rst(1).Value & ""
            End If
            lbcMulti.ItemData(lbcMulti.NewIndex) = rst!shttCode
            rst.MoveNext
        Wend
    Else
        imShfCode = lbcSingle.ItemData(lbcSingle.ListIndex)
        
        'sDateRange = "(att.attOffAir >=" + smDate + ") And (att.attOnAir <=" + smDate + ")"
        sDateRange = "(attOffAir >= '" & Format$(smDate, sgSQLDateForm) + "') and (attDropDate >= '" & Format$(smDate, sgSQLDateForm) & "') And (attOnAir <= '" & Format$(smDate, sgSQLDateForm) & "')"
        
        SQLQuery = "SELECT DISTINCT vefType, vefName, vefCode"
        'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, att" 'cptt"
        SQLQuery = SQLQuery + " FROM VEF_Vehicles, att"
        SQLQuery = SQLQuery & " WHERE (attShfCode = " & imShfCode
        SQLQuery = SQLQuery & " AND vefCode = attVefCode"
        SQLQuery = SQLQuery & " AND " & sDateRange & ")"
        SQLQuery = SQLQuery + " ORDER BY vefName"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            If sgShowByVehType = "Y" Then
                lbcMulti.AddItem Trim$(rst!vefType) & ":" & Trim$(rst!vefName)
            Else
                lbcMulti.AddItem Trim$(rst!vefName)  ', " & rst(1).Value & ""
            End If
            lbcMulti.ItemData(lbcMulti.NewIndex) = rst!vefCode
            rst.MoveNext
        Wend
    End If
    If lbcMulti.ListCount = 1 Then
        chkAll.Value = 1
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-lbcSingle"
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Sub optSort_Click(Index As Integer)
    If optSort(0).Value Then
        lacTitle1.Caption = "Vehicle"
        lacTitle2.Caption = "Stations"
    Else
        lacTitle1.Caption = "Station"
        lacTitle2.Caption = "Vehicles"
    End If
    Screen.MousePointer = vbHourglass
    mFill
    Screen.MousePointer = vbDefault
End Sub

Private Sub pbcClickFocus_GotFocus()
    mLogSetShow
End Sub

Private Sub pbcDay_Click()
    If ckcDay.Value = vbChecked Then
        ckcDay.Value = vbUnchecked
    Else
        ckcDay.Value = vbChecked
    End If
End Sub

Private Sub pbcLogSTab_GotFocus()
    If GetFocus() <> pbcLogSTab.hwnd Then
        Exit Sub
    End If
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(4) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    If imFromArrow Then
        imFromArrow = False
        mLogEnableBox
        Exit Sub
    End If
    If pbcDay.Visible Or txtDropdown.Visible Then
        mLogSetShow
        If grdLog.Col = 0 Then
            If grdLog.Row > grdLog.FixedRows Then
                lmTopRow = -1
                grdLog.Row = grdLog.Row - 1
                If Not grdLog.RowIsVisible(grdLog.Row) Then
                    grdLog.TopRow = grdLog.TopRow - 1
                End If
                grdLog.Col = 1
                mLogEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        ElseIf grdLog.Col = 4 Then
            grdLog.Col = 0
            mLogEnableBox
        Else
            grdLog.Col = grdLog.Col - 1
            mLogEnableBox
        End If
    Else
        lmTopRow = -1
        grdLog.TopRow = grdLog.FixedRows
        grdLog.Col = 0
        grdLog.Row = grdLog.FixedRows
        mLogEnableBox
    End If
End Sub

Private Sub pbcLogTab_GotFocus()
    If GetFocus() <> pbcLogTab.hwnd Then
        Exit Sub
    End If
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(4) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    If pbcDay.Visible Or txtDropdown.Visible Then
        mLogSetShow
        If ((grdLog.Col = grdLog.Cols - 1) And (igCPOrLog <> 0)) Or ((grdLog.Col = grdLog.Cols - 2) And (igCPOrLog = 0)) Then
            If grdLog.Row + 1 < grdLog.Rows Then
                lmTopRow = -1
                grdLog.Row = grdLog.Row + 1
                If Not grdLog.RowIsVisible(grdLog.Row) Then
                    grdLog.TopRow = grdLog.TopRow + 1
                End If
                grdLog.Col = 0
                mLogEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        ElseIf grdLog.Col = 0 Then
            grdLog.Col = 4
            mLogEnableBox
        Else
            grdLog.Col = grdLog.Col + 1
            mLogEnableBox
        End If
    Else
        lmTopRow = -1
        grdLog.TopRow = grdLog.FixedRows
        grdLog.Col = 0
        grdLog.Row = grdLog.FixedRows
        mLogEnableBox
    End If
End Sub

Private Sub rbcVeh_Click(Index As Integer)
    mCPMain
End Sub

Private Sub rbcVeh_GotFocus(Index As Integer)
    mLogSetShow
End Sub

Private Sub TabStrip1_Click()
    Dim iLoop As Integer
    Dim iZone As Integer
    
    'Log tab index 1 = Print Logs, 2 = Reprint Logs, 3 = Studio Logs
    'CPs 1 = Print CP, 2 = Reprint CP
    'igCPorLog 0 = CP, 1 = Log
    
    If imTabIndex = TabStrip1.SelectedItem.Index Then
        Exit Sub
    End If
    'Logs
    If TabStrip1.SelectedItem.Index = 1 Then
        frcTab(0).Visible = True
        frcTab(1).Visible = False
        ckcCover.Visible = False
        If igCPOrLog = 0 Then
            frcPrintCPNotCarried.Visible = True
            frcRePrintCPNotCarried.Visible = False
        End If
    'Reprint logs
    Else
        If TabStrip1.SelectedItem.Index = 2 Then
            lacReDate.Caption = "Reprint Date"
            If igCPOrLog = 0 Then
                lacReDays.Caption = "Number Weeks"
            Else
                lacReDays.Caption = "Number Days"
            End If
            frcZone.Caption = "Reprint Zone"
            chkZone(0).Enabled = True
            chkZone(1).Enabled = True
            chkZone(2).Enabled = True
            chkZone(3).Enabled = True
            chkZone(0).Value = 1
            chkZone(1).Value = 1
            chkZone(2).Value = 1
            chkZone(3).Value = 1
            If igCPOrLog = 0 Then
                frcPrintCPNotCarried.Visible = False
                frcRePrintCPNotCarried.Visible = True
                ckcCover.Visible = True
            End If
        'Studio logs
        Else
            lacReDate.Caption = "Studio Date"
            lacReDays.Caption = "Number Days"
            frcZone.Caption = "Studio Zone"
            '*********************************
            'Allow all zones to be selected- reset to false to have only
            'Fed = "*" enabled
            'Also test for * removed from mGenLog code
            'chkZone(0).Enabled = False
            'chkZone(1).Enabled = False
            'chkZone(2).Enabled = False
            'chkZone(3).Enabled = False
            chkZone(0).Enabled = True
            chkZone(1).Enabled = True
            chkZone(2).Enabled = True
            chkZone(3).Enabled = True
            '*************************************************
            chkZone(0).Value = 0
            chkZone(1).Value = 0
            chkZone(2).Value = 0
            chkZone(3).Value = 0

            For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                For iZone = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                    Select Case Left$(tgVehicleInfo(iLoop).sZone(iZone), 1)
                        Case "E"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(0).Enabled = True
                                chkZone(0).Value = 1
                            End If
                        Case "C"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(1).Enabled = True
                                chkZone(1).Value = 1
                            End If
                        Case "M"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(2).Enabled = True
                                chkZone(2).Value = 1
                            End If
                        Case "P"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(3).Enabled = True
                                chkZone(3).Value = 1
                            End If
                    End Select
                Next iZone
            Next iLoop
        End If
        frcTab(1).Visible = True
        frcTab(0).Visible = False
    End If
    imTabIndex = TabStrip1.SelectedItem.Index
End Sub

Private Sub TabStrip1_GotFocus()
    mLogSetShow
End Sub

Private Sub txtDate_Change()
    lbcSingle.ListIndex = -1
    txtDays.Text = ""
End Sub

Private Sub txtDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Function mLogGen(iPrint As Integer, iLog As Integer) As Integer
'
'   iPrint(I)- True = From Log selection; False= From Reprint or Studio selection
'   iLog(I)- True = Generate Log, False = Generate Studio
'
    Dim iTRow As Integer
    Dim iRow As Integer
    Dim sDate As String
    Dim sSDate As String
    Dim sEDate As String
    Dim lSDate As Long
    Dim lEDate As Long
    Dim iVefCode As Integer
    Dim iSeqNo As Integer
    Dim iWkNo As Integer
    Dim iDay As Integer
    Dim iCycle As Integer
    Dim iZone As Integer
    Dim iOutput As Integer
    Dim iLoop As Integer
    Dim iRnfCode As Integer
    Dim sRptExe As String
    Dim sLogName As String
    Dim sCommd As String
    Dim sLength As String
    Dim sGenDate As String
    Dim sGenTime As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim lAirTime As Long
    Dim lTime As Long
    Dim iIndex As Integer
    Dim sStr As String
    Dim sZone As String
    Dim sProd As String
    Dim sCart As String
    Dim iSelected As Integer
    Dim iPass As Integer
    Dim iRnf As Integer
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    Dim sFileName As String
    Dim sLetter As String
    Dim iRptType As Integer
    Dim iPrtZone As Integer
    Dim iZoneNo As Integer
    Dim iSPass As Integer
    Dim iEPass As Integer
    Dim lAvailTime As Long
    Dim lRunTime As Long
    Dim sFileZone As String * 1
    Dim sOutput As String
    Dim sNoDays As String
    Dim iNoDays As Integer
    Dim iDayLp As Integer
    Dim iGenL32 As Integer
    Dim sStdDate As String
    Dim rstDat As ADODB.Recordset
    Dim ilExportType As Integer
    Dim slExportName As String
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim llRow As Long
    'Dim NewForm As New frmViewReport
    
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    smCntrNo = "0"
    If Not iPrint Then
        sDate = txtDate.Text
        sNoDays = txtDays.Text
        If Not gIsDate(sDate) Then
            Screen.MousePointer = vbDefault
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            mLogGen = False
            Exit Function
        End If
        If Trim$(sNoDays) = "" Then
            Screen.MousePointer = vbDefault
            Beep
            gMsgBox "Please enter Number Days.", vbCritical
            mLogGen = False
            Exit Function
        End If
        Select Case Weekday(sDate)
            Case vbMonday
                imMaxDays = 7
            Case vbTuesday
                imMaxDays = 6
            Case vbWednesday
                imMaxDays = 5
            Case vbThursday
                imMaxDays = 4
            Case vbFriday
                imMaxDays = 3
            Case vbSaturday
                imMaxDays = 2
            Case vbSunday
                imMaxDays = 1
            Case Else
                txtDays.Text = ""
        End Select
        If Val(sNoDays) > imMaxDays Then
            Screen.MousePointer = vbDefault
            Beep
            gMsgBox "Number Days to large, can't be more then " & Str$(imMaxDays), vbCritical
            mLogGen = False
            Exit Function
        End If
        iNoDays = Val(sNoDays)
    End If
    If optRptDest(0).Value Then
        iOutput = 0
    ElseIf optRptDest(2).Value Then
        iOutput = 2
    Else
        iOutput = 1
    End If
    If iLog Then
        iSPass = 4
        iEPass = 5
    Else
        iSPass = 1
        iEPass = 1
    End If
    grdLog.Redraw = False
    llRow = grdLog.FixedRows
    For iRow = 0 To UBound(tmCPInfo) - 1 Step 1
        If (tmCPInfo(iRow).iStatus = 0) And ((rbcVeh(0).Value = False) Or ((rbcVeh(0).Value = True) And (tmCPInfo(iRow).sVefState = "A"))) Then
            iGenL32 = False
            If iLog Then
                For iPass = iSPass To iEPass Step 1
                    If StrComp(Trim$(grdLog.TextMatrix(llRow, iPass)), "L32", 1) = 0 Then
                        iGenL32 = True
                        Exit For
                    End If
                Next iPass
            End If
            ReDim tmCmmlSum(0 To 0) As CMMLSUM
            iVefCode = tmCPInfo(iRow).iVefCode
            iIndex = -1
            For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If tgVehicleInfo(iLoop).iCode = iVefCode Then
                    iIndex = iLoop
                    Exit For
                End If
            Next iLoop
            If iIndex >= 0 Then
                If iPrint Then 'Print Logs
                    sDate = Trim$(grdLog.TextMatrix(llRow, 2))
                    sNoDays = Val(grdLog.TextMatrix(llRow, 3))
                    If Trim$(grdLog.TextMatrix(llRow, 0)) <> "" Then
                        iSelected = True
                    Else
                        iSelected = False
                    End If
                    If Not gIsDate(sDate) Then
                        iSelected = False
                    End If
                    If Trim$(sNoDays) = "" Then
                        iSelected = False
                    End If
                Else 'Reprint or Studio Logs
                    sDate = txtDate.Text
                    'If imVefCode = iVefCode Then
                    '    iSelected = True
                    'Else
                    '    iSelected = False
                    'End If
                    For iLoop = 0 To lbcMulti.ListCount - 1 Step 1
                        If iVefCode = lbcMulti.ItemData(iLoop) Then
                            If lbcMulti.Selected(iLoop) Then
                                iSelected = True
                                imVefCode = iVefCode
                            Else
                                iSelected = False
                            End If
                            Exit For
                        End If
                    Next iLoop
                End If
                If (iSelected) And (Len(sDate) > 0) Then
                    For iLoop = 0 To UBound(tgRnfInfo) - 1 Step 1
                        If (StrComp(Trim$(tgRnfInfo(iLoop).sName), Trim$(grdLog.TextMatrix(llRow, 4)), 1) = 0) Or (StrComp(Trim$(tgRnfInfo(iLoop).sName), Trim$(grdLog.TextMatrix(llRow, 5)), 1) = 0) Then
                            lacProcess.Caption = "Processing : " & Trim$(tgVehicleInfo(iIndex).sVehicle)
                            DoEvents
                            'SQLQuery = "DELETE FROM ODF_One_Day_Log odf WHERE (odfVefCode = " & iVefCode & ")"
                            'cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            'cnn.CommitTrans
                            If iPrint Then
                                sDate = Trim$(grdLog.TextMatrix(llRow, 2))
                                sNoDays = Val(grdLog.TextMatrix(llRow, 3))
                                iNoDays = Val(sNoDays)
                                iCycle = Val(grdLog.TextMatrix(llRow, 3))
                            Else
                                iCycle = Val(grdLog.TextMatrix(llRow, 3))
                            End If
                            For iDayLp = 1 To iNoDays Step iCycle
                                sGenDate = Format$(gNow(), sgShowDateForm)
                                sGenTime = Format$(gNow(), sgShowTimeWSecForm)
                                
                                lAvailTime = -1
                                lRunTime = 0
                                iZone = 0   'All
                                iSeqNo = 1
                                sSDate = Format$(DateValue(gAdjYear(sDate)) + iDayLp - 1, sgShowDateForm)
                                If iCycle >= iNoDays Then
                                    sEDate = Format$(DateValue(gAdjYear(sSDate)) + iNoDays - 1, sgShowDateForm)
                                Else
                                    sEDate = Format$(DateValue(gAdjYear(sSDate)) + iCycle - 1, sgShowDateForm)
                                End If
                                lSDate = DateValue(gAdjYear(sSDate))
                                lEDate = DateValue(gAdjYear(sEDate))
                                SQLQuery = "SELECT * FROM lst "
                                SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & iVefCode
                                SQLQuery = SQLQuery + " AND lstType = " & 0
                                SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
                                SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(lSDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lEDate + 1, sgSQLDateForm) & "')" & ")"
                                SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
                                Set rst = gSQLSelectCall(SQLQuery)
                                While Not rst.EOF
                                    If tgStatusTypes(rst!lstStatus).iPledged <> 2 Then
                                        If rst!lstLen < 60 Then
                                            'sLength = "12:00:" & Trim$(Str$(rst!lstLen)) & " AM"
                                            sStr = Trim$(Str$(rst!lstLen))
                                            If Len(sStr) = 1 Then
                                                sStr = "0" & sStr
                                            End If
                                            sLength = "00:00:" & sStr
                                        Else
                                            'sLength = "12:01:" & Trim$(Str$(rst!lstLen - 60)) & " AM"
                                            sStr = Trim$(Str$(rst!lstLen - 60))
                                            If Len(sStr) = 1 Then
                                                sStr = "0" & sStr
                                            End If
                                            sLength = "00:01:" & sStr
                                        End If
                                        lAirTime = gTimeToLong(rst!lstLogTime, False)
                                        If lAirTime = lAvailTime Then
                                            lAirTime = lRunTime
                                            lRunTime = lRunTime + rst!lstLen
                                        Else
                                            lAvailTime = lAirTime
                                            lRunTime = lAvailTime + rst!lstLen
                                        End If
                                        For iZone = LBound(tgVehicleInfo(iIndex).sZone) To UBound(tgVehicleInfo(iIndex).sZone) Step 1
                                            If (tgVehicleInfo(iIndex).sZone(iZone) <> "~~~") Then
                                                If IsNull(rst!lstZone) Then
                                                    sZone = "   "
                                                Else
                                                    sZone = rst!lstZone
                                                End If
                                                If (tgVehicleInfo(iIndex).sZone(iZone) = sZone) Or ((tgVehicleInfo(iIndex).sFed(iZone) = Left$(sZone, 1)) And (sZone <> "   ") And (Len(Trim$(tgVehicleInfo(iIndex).sFed(iZone))) <> 0)) Then
                                                    sAirDate = Format$(rst!lstLogDate, sgShowDateForm)
                                                    lTime = lAirTime + 3600 * tgVehicleInfo(iIndex).iLocalAdj(iZone)
                                                    If lTime < 0 Then
                                                        lTime = lTime + 86400
                                                        sAirDate = Format$(DateValue(gAdjYear(sAirDate)) - 1, sgShowDateForm)
                                                    ElseIf lTime > 86400 Then
                                                        lTime = lTime - 86400
                                                        sAirDate = Format$(DateValue(gAdjYear(sAirDate)) + 1, sgShowDateForm)
                                                    End If
                                                    sAirTime = Format$(gLongToTime(lTime), sgShowTimeWSecForm)
                                                    If (DateValue(gAdjYear(sAirDate)) >= lSDate) And (DateValue(gAdjYear(sAirDate)) <= lEDate) Then
                                                        iWkNo = (DateValue(gAdjYear(sAirDate)) - DateValue(gAdjYear(sSDate))) \ 7 + 1
                                                        Select Case Weekday(sAirDate)
                                                            Case vbSaturday
                                                                iDay = 6
                                                            Case vbSunday
                                                                iDay = 7
                                                            Case Else
                                                                iDay = 1
                                                        End Select
                                                        If IsNull(rst!lstProd) Then
                                                            sProd = ""
                                                        Else
                                                            sProd = gFixQuote(rst!lstProd)
                                                        End If
                                                        If IsNull(rst!lstCart) Or Left$(rst!lstCart, 1) = Chr$(0) Then
                                                            sCart = ""
                                                        Else
                                                            sCart = gFixQuote(rst!lstCart)
                                                        End If
                                                        'Build ODF
                                                        'SQLQuery = "INSERT INTO ODF_One_Day_Log (odfUrfCode, odfVefCode, odfAirDate, odfAirTime, "
                                                        SQLQuery = "INSERT INTO " & "ODF_One_Day_Log"
                                                        SQLQuery = SQLQuery & " (odfUrfCode, odfVefCode, odfAirDate, odfAirTime, "
                                                        SQLQuery = SQLQuery & "odfSeqNo, odfLocalTime, odfFeedTime, odfZone, odfEtfCode, "
                                                        SQLQuery = SQLQuery & "odfEnfCode, odfProgCode, odfMnfFeed, odfWkNo, odfAnfCode, "
                                                        SQLQuery = SQLQuery & "odfUnits, odfLength, odfAdfCode, odfCifCode, odfProduct, "
                                                        SQLQuery = SQLQuery & "odfMnfSubFeed, odfCntrNo, odfBreakNo, odfPositionNo, odfType, "
                                                        SQLQuery = SQLQuery & "odfCefCode, odfShortTitle, odfPageEjectFlag, odfSortSeq, "
                                                        SQLQuery = SQLQuery & "odfAvailCefCode, odfRdfSortCode, odfDPDesc, odfChfCxfCode, odfDaySort, "
                                                        SQLQuery = SQLQuery & "odfEvtCefCode, odfEvtCefSort, odfEvtIDCefCode, odfDupeAvailID, odfLogType, "
                                                        SQLQuery = SQLQuery & "odfAvailLen, odfAvailLock, "
                                                        SQLQuery = SQLQuery & "odfGenDate, odfGenTime)" ', "
                                                        'Temporary remove header and footer and vehicle name so release can be sent-  to make work we need to have
                                                        'a blank comment record to point to if ---cefcode values are zero
                                                        'SQLQuery = SQLQuery & "odfHd1CefCode, odfFt1CefCode, odfFt2CefCode, odfVehNmCefCode) "
                                                        SQLQuery = SQLQuery & " VALUES (" & 1 & ", " & iVefCode & ", '" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', "
                                                        SQLQuery = SQLQuery & iSeqNo & ", '" & Format$(sAirTime, sgSQLTimeForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', '" & tgVehicleInfo(iIndex).sZone(iZone) & "', " & 0 & ", "
                                                        SQLQuery = SQLQuery & 0 & ", '" & "" & "', " & 0 & ", " & iWkNo & ", " & rst!lstAnfCode & ", "
                                                        SQLQuery = SQLQuery & 0 & ", '" & sLength & "', " & rst!lstAdfCode & ", " & rst!lstCifCode & ", '" & sProd & "', "
                                                        SQLQuery = SQLQuery & 0 & ", " & rst!lstCntrNo & ", " & rst!lstBreakNo & ", " & rst!lstPositionNo & ", " & 4 & ", "
                                                        SQLQuery = SQLQuery & 0 & ", '" & sCart & "', '" & "N" & "', " & iSeqNo & ", "
                                                        SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & "" & "', " & 0 & ", " & iDay & ", "
                                                        SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "" & "', '" & "F" & "'" & ", "
                                                        SQLQuery = SQLQuery & "0" & ", " & "'N'" & ", "
                                                        'SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Format(sGenTime, sgSQLTimeForm) & "')"   '", "
                                                        SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"    '", "
                                                        ''Temporary remove header and footer and vehicle name so release can be sent-  to make work we need to have
                                                        ''a blank comment record to point to if ---cefcode values are zero
                                                        ''SQLQuery = SQLQuery & ", " & tgVehicleInfo(iIndex).lHd1CefCode & ", " & tgVehicleInfo(iIndex).lLgNmCefCode & ", " & tgVehicleInfo(iIndex).lFt1CefCode & ", " & tgVehicleInfo(iIndex).lFt2CefCode & ")"
                                                        'cnn.BeginTrans
                                                        'cnn.Execute SQLQuery, rdExecDirect
                                                        'cnn.CommitTrans
                                                        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                                            'grdLog.Redraw = True
                                                            'Screen.MousePointer = vbDefault
                                                            'mLogGen = False
                                                            'Exit Function
                                                            '6/10/16: Replaced GoSub
                                                            'GoSub ErrHand:
                                                            Screen.MousePointer = vbDefault
                                                            gHandleError "AffErrorLog.txt", "CP-mLogGen"
                                                            mLogGen = False
                                                            Exit Function
                                                        End If
                                                        If iGenL32 Then
                                                            mCmmlSum "L32", Weekday(sAirDate, vbMonday) - 1, iIndex, rst!lstCntrNo, tgVehicleInfo(iIndex).sZone(iZone), sAirTime, iVefCode, rst!lstAdfCode, sProd, rst!lstLen
                                                        End If
                                                        iSeqNo = iSeqNo + 1
                                                    End If
                                                End If
                                            End If
                                            If tgVehicleInfo(iIndex).iNoZones = 0 Then
                                                Exit For
                                            End If
                                        Next iZone
                                    End If
                                    rst.MoveNext
                                Wend
                                If iGenL32 Then
                                    If Not mGenCmmlSum(sGenDate, sGenTime) Then
                                        grdLog.Redraw = True
                                        Screen.MousePointer = vbDefault
                                        mLogGen = False
                                        Exit Function
                                    End If
                                End If
                                For iPass = iSPass To iEPass Step 1
                                    For iRnf = 0 To UBound(tgRnfInfo) - 1 Step 1
                                        If (StrComp(Trim$(tgRnfInfo(iRnf).sName), Trim$(grdLog.TextMatrix(llRow, iPass)), 1) = 0) Or (Not iLog) Then
                                            sLogName = Trim$(grdLog.TextMatrix(llRow, iPass))
                                            For iZone = LBound(tgVehicleInfo(iIndex).sZone) To UBound(tgVehicleInfo(iIndex).sZone) Step 1
                                                If iPrint Then
                                                    iPrtZone = True
                                                Else
                                                    iPrtZone = False
                                                End If
                                                sFileZone = ""
                                                If tgVehicleInfo(iIndex).iNoZones <> 0 Then
                                                    Select Case UCase$(Left$(tgVehicleInfo(iIndex).sZone(iZone), 1))
                                                        Case "E"
                                                            iZoneNo = 1
                                                            If chkZone(0).Value = 1 Then
                                                                If iLog Then
                                                                    iPrtZone = True
                                                                '***********************************
                                                                'Remove test so all zones will be printed that are select
                                                                'ElseIf tgVehicleInfo(iIndex).sFed(iZone) = "*" Then
                                                                Else
                                                                '**********************************
                                                                    iPrtZone = True
                                                                End If
                                                            End If
                                                            sFileZone = "E"
                                                        Case "C"
                                                            iZoneNo = 2
                                                            If chkZone(1).Value = 1 Then
                                                                If iLog Then
                                                                    iPrtZone = True
                                                                '***********************************
                                                                'Remove test so all zones will be printed that are select
                                                                'ElseIf tgVehicleInfo(iIndex).sFed(iZone) = "*" Then
                                                                Else
                                                                '**********************************
                                                                    iPrtZone = True
                                                                End If
                                                            End If
                                                            sFileZone = "C"
                                                        Case "M"
                                                            iZoneNo = 3
                                                            If chkZone(2).Value = 1 Then
                                                                If iLog Then
                                                                    iPrtZone = True
                                                                '***********************************
                                                                'Remove test so all zones will be printed that are select
                                                                'ElseIf tgVehicleInfo(iIndex).sFed(iZone) = "*" Then
                                                                Else
                                                                '**********************************
                                                                    iPrtZone = True
                                                                End If
                                                            End If
                                                            sFileZone = "M"
                                                        Case "P"
                                                            iZoneNo = 4
                                                            If chkZone(3).Value = 1 Then
                                                                If iLog Then
                                                                    iPrtZone = True
                                                                '***********************************
                                                                'Remove test so all zones will be printed that are select
                                                                'ElseIf tgVehicleInfo(iIndex).sFed(iZone) = "*" Then
                                                                Else
                                                                '**********************************
                                                                    iPrtZone = True
                                                                End If
                                                            End If
                                                            sFileZone = "P"
                                                    End Select
                                                Else
                                                    iPrtZone = True
                                                    sFileName = ""
                                                End If
                                                If iPrtZone Then
                                                    If iLog Then
                                                        If StrComp(sLogName, "L31", 1) = 0 Then
                                                            'Create Report Call
                                                            'CRpt1.Connect = "DSN = " & sgDatabaseName
                                                            If optRptDest(0).Value = True Then
                                                                'CRpt1.Destination = crptToWindow
                                                                ilRptDest = 0
                                                            ElseIf optRptDest(1).Value = True Then
                                                                'CRpt1.Destination = crptToPrinter
                                                                ilRptDest = 1
                                                            ElseIf optRptDest(2).Value = True Then
                                                                gObtainYearMonthDayStr sSDate, True, sYear, sMonth, sDay
                                                                'If Val(sMonth) <= 9 Then
                                                                '    sMonth = Right$(sMonth, 1)
                                                                'ElseIf Val(sMonth) = 10 Then
                                                                '    sMonth = "A"
                                                                'ElseIf Val(sMonth) = 11 Then
                                                                '    sMonth = "B"
                                                                'ElseIf Val(sMonth) = 12 Then
                                                                '    sMonth = "C"
                                                                'End If
                                                                sLetter = Trim$(Left$(tgVehicleInfo(iIndex).sCodeStn, 3))
                                                                If iPass = iSPass Then
                                                                    sFileName = Trim$(sFileZone) & sMonth & sDay & "L_" & sLetter '& ".Txt"
                                                                Else
                                                                    sFileName = Trim$(sFileZone) & sMonth & sDay & "O_" & sLetter '& ".Txt"
                                                                End If
                                                                'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
                                                                ilExportType = cboFileType.ListIndex    '3-12-04
                                                                slExportName = sFileName
                                                                ilRptDest = 2
                                                                'gOutputMethod frmCP, sFileName, sOutput
                                                            Else
                                                                Screen.MousePointer = vbDefault
                                                                Exit Function
                                                            End If
                                                            If igSQLSpec = 0 Then
                                                                SQLQuery = "SELECT *"
                                                                'SQLQuery = SQLQuery + " FROM ODF_One_Day_Log odf, ADF_Advertisers adf, VEF_Vehicles vef"
                                                                SQLQuery = SQLQuery & " FROM ODF_One_Day_Log, "
                                                                SQLQuery = SQLQuery & "ADF_Advertisers, "
                                                                SQLQuery = SQLQuery & "VEF_Vehicles, "
                                                                SQLQuery = SQLQuery & "MNF_Multi_Names"
                                                                If Trim$(sFileZone) <> "" Then
                                                                    SQLQuery = SQLQuery + " WHERE (odfZone = '" & sFileZone & "ST'"
                                                                Else
                                                                    SQLQuery = SQLQuery + " WHERE (odfZone = '" & "   '"
                                                                End If
                                                                SQLQuery = SQLQuery + " AND odfVefCode = " & iVefCode
                                                                SQLQuery = SQLQuery + " AND odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND odfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                                SQLQuery = SQLQuery + " AND odfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
                                                                SQLQuery = SQLQuery + " AND (odfAirDate >= '" & Format$(sSDate, sgSQLDateForm) & "' AND odfAirDate <= '" & Format$(sEDate, sgSQLDateForm) & "')"
                                                                SQLQuery = SQLQuery + " AND adfCode = odfAdfCode"
                                                                SQLQuery = SQLQuery & " AND vefMnfVehGp4Fmt = mnfCode (+)"
                                                                SQLQuery = SQLQuery & " AND vefCode = odfVefCode" & ")"
                                                            Else
                                                                'SQLQuery = "SELECT *"
                                                                'SQLQuery = SQLQuery & " FROM " '& """" & "ODF_One_Day_Log" & """" & " odf, "
                                                                ''SQLQuery = SQLQuery & "([ODF_One_Day_Log] Odf INNER JOIN ([VEF_Vehicles] Vef LEFT OUTER JOIN [MNF_Multi_Names] Mnf on vefMnfVehGp4Fmt = mnfCode) "
                                                                ''SQLQuery = SQLQuery & " ON odfVefCode = vefCode) INNER JOIN [ADF_Advertisers] adf ON odfAdfCode = adfCode "
                                                                'SQLQuery = SQLQuery & "(ODF_One_Day_Log INNER JOIN (VEF_Vehicles LEFT OUTER JOIN MNF_Multi_Names on vefMnfVehGp4Fmt = mnfCode) "
                                                                'SQLQuery = SQLQuery & " ON odfVefCode = vefCode) INNER JOIN ADF_Advertisers ON odfAdfCode = adfCode "
                                                                'SQLQuery = SQLQuery + " WHERE (odfZone = '" & sFileZone & "ST'"
                                                                'SQLQuery = SQLQuery + " AND odfVefCode = " & iVefCode
                                                                'SQLQuery = SQLQuery + " AND odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND odfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND (odfAirDate BETWEEN '" & Format$(sSDate, sgSQLDateForm) & "' AND '" & Format$(sEDate, sgSQLDateForm) & "')"
                                                                'SQLQuery = SQLQuery & ")"
                                                                SQLQuery = "SELECT *"
                                                                SQLQuery = SQLQuery & " FROM ODF_One_Day_Log, "
                                                                SQLQuery = SQLQuery & "ADF_Advertisers, "
                                                                'SQLQuery = SQLQuery & """" & "VEF_Vehicles" & """" & " vef, "
                                                                'SQLQuery = SQLQuery & """" & "MNF_Multi_Names" & """" & " mnf"
                                                                SQLQuery = SQLQuery & "VEF_Vehicles LEFT OUTER JOIN MNF_Multi_Names On vefMnfVehGp4Fmt = mnfCode"
                                                                If Trim$(sFileZone) <> "" Then
                                                                    SQLQuery = SQLQuery + " WHERE (odfZone = '" & sFileZone & "ST'"
                                                                Else
                                                                    SQLQuery = SQLQuery + " WHERE (odfZone = '" & "   '"
                                                                End If
                                                                SQLQuery = SQLQuery + " AND odfVefCode = " & iVefCode
                                                                SQLQuery = SQLQuery + " AND odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND odfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                                SQLQuery = SQLQuery + " AND odfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
                                                                SQLQuery = SQLQuery + " AND (odfAirDate >= '" & Format$(sSDate, sgSQLDateForm) & "' AND odfAirDate <= '" & Format$(sEDate, sgSQLDateForm) & "')"
                                                                SQLQuery = SQLQuery + " AND adfCode = odfAdfCode"
                                                                'SQLQuery = SQLQuery & " AND vefMnfVehGp4Fmt = mnfCode (+)"
                                                                SQLQuery = SQLQuery & " AND vefCode = odfVefCode" & ")"
                                                            End If
                                                            'CRpt1.SQLQuery = SQLQuery
                                                            'CRpt1.ReportFileName = sgReportDirectory & sLogName & "a.Rpt"    '"\Affiliate5\Reports\Stations.rpt"
                                                            slRptName = sLogName & "a.rpt"
                                                            sYear = Format$(gObtainEndStd(sSDate), "yyyy")
                                                            sStdDate = gObtainStartStd("1/15/" & sYear)
                                                            sYear = Format$(gObtainEndStd(sStdDate), "yy")
                                                            'CRpt1.Formulas(0) = "StdYear =  '" & sYear & "'"
                                                            sgCrystlFormula1 = "'" & sYear & "'"
                                                            'CRpt1.Formulas(1) = "Week = " & ((DateValue(sSDate) - DateValue(sStdDate)) \ 7 + 1)
                                                            sgCrystlFormula2 = ((DateValue(gAdjYear(sSDate)) - DateValue(gAdjYear(sStdDate))) \ 7 + 1)
                                                            'CRpt1.Formulas(2) = "InputDate = Date(" + Format$(sSDate, "yyyy") + "," + Format$(sSDate, "mm") + "," + Format$(sSDate, "dd") + ")"
                                                            sgCrystlFormula3 = "Date(" + Format$(sSDate, "yyyy") + "," + Format$(sSDate, "mm") + "," + Format$(sSDate, "dd") + ")"
                                                            'CRpt1.Formulas(3) = "NumberDays = " & (DateValue(sEDate) - DateValue(sSDate) + 1)
                                                            sgCrystlFormula4 = (DateValue(gAdjYear(sEDate)) - DateValue(gAdjYear(sSDate)) + 1)
                                                            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
                                                            'CRpt1.Action = 1
                                                            'CRpt1.Formulas(0) = ""
                                                            'CRpt1.Formulas(1) = ""
                                                            'CRpt1.Formulas(2) = ""
                                                            'CRpt1.Formulas(3) = ""
                                                        ElseIf StrComp(sLogName, "L32", 1) = 0 Then
                                                            'Create Report Call
                                                            'CRpt1.Connect = "DSN = " & sgDatabaseName
                                                            If optRptDest(0).Value = True Then
                                                                'CRpt1.Destination = crptToWindow
                                                                ilRptDest = 0
                                                            ElseIf optRptDest(1).Value = True Then
                                                                'CRpt1.Destination = crptToPrinter
                                                                ilRptDest = 1
                                                            ElseIf optRptDest(2).Value = True Then
                                                                gObtainYearMonthDayStr sSDate, True, sYear, sMonth, sDay
                                                                'If Val(sMonth) <= 9 Then
                                                                '    sMonth = Right$(sMonth, 1)
                                                                'ElseIf Val(sMonth) = 10 Then
                                                                '    sMonth = "A"
                                                                'ElseIf Val(sMonth) = 11 Then
                                                                '    sMonth = "B"
                                                                'ElseIf Val(sMonth) = 12 Then
                                                                '    sMonth = "C"
                                                                'End If
                                                                sLetter = Trim$(Left$(tgVehicleInfo(iIndex).sCodeStn, 3))
                                                                
                                                                'sFileName = sFileZone & sMonth & sDay & "O" & sLetter '& ".Txt"
                                                                If iPass = iSPass Then
                                                                    sFileName = Trim$(sFileZone) & sMonth & sDay & "L_" & sLetter '& ".Txt"
                                                                Else
                                                                    sFileName = Trim$(sFileZone) & sMonth & sDay & "O_" & sLetter '& ".Txt"
                                                                End If
                                                                'gOutputMethod frmCP, sFileName, sOutput
                                                                'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
                                                                ilExportType = cboFileType.ListIndex    '3-12-04
                                                                slExportName = sFileName
                                                                ilRptDest = 2
                                                            Else
                                                                Screen.MousePointer = vbDefault
                                                                Exit Function
                                                            End If
                                                            If igSQLSpec = 0 Then
                                                                SQLQuery = "SELECT *"
                                                                'SQLQuery = SQLQuery + " FROM ODF_One_Day_Log odf, ADF_Advertisers adf, VEF_Vehicles vef"
                                                                SQLQuery = SQLQuery & " FROM GRF_Generic_Report, "
                                                                SQLQuery = SQLQuery & "ADF_Advertisers, "
                                                                SQLQuery = SQLQuery & "VEF_Vehicles, "
                                                                SQLQuery = SQLQuery & "VPF_Vehicle_Options, "
                                                                SQLQuery = SQLQuery & "MNF_Multi_Names"
                                                                SQLQuery = SQLQuery + " WHERE (grfBktType = '" & sFileZone & "'"
                                                                SQLQuery = SQLQuery + " AND grfVefCode = " & iVefCode
                                                                SQLQuery = SQLQuery + " AND grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND grfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                                SQLQuery = SQLQuery + " AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
                                                                
                                                                'SQLQuery = SQLQuery + " AND (grf.grfAirDate BETWEEN '" & Format$(sSDate, "mm/dd/yyyy") & "' AND '" & Format$(sEDate, "mm/dd/yyyy") & "')"
                                                                SQLQuery = SQLQuery + " AND adfCode = grfAdfCode"
                                                                SQLQuery = SQLQuery & " AND vefMnfVehGp4Fmt = mnfCode (+)"
                                                                SQLQuery = SQLQuery & " AND vpfvefKCode = grfVefCode"
                                                                SQLQuery = SQLQuery & " AND vefCode = grfVefCode" & ")"
                                                            Else
                                                                'SQLQuery = "SELECT *"
                                                                'SQLQuery = SQLQuery & " FROM " '& """" & "GRF_Generic_Report" & """" & " grf, "
                                                                'SQLQuery = SQLQuery & "(GRF_Generic_Report INNER JOIN ((VEF_Vehicles INNER JOIN VPF_Vehicle_Options on vefCode = vpfVefKCode) LEFT OUTER JOIN MNF_Multi_Names on vefMnfVehGp4Fmt = mnfCode) "
                                                                'SQLQuery = SQLQuery & " ON grfVefCode = vefCode) INNER JOIN ADF_Advertisers ON grfAdfCode = adfCode "
                                                                'SQLQuery = SQLQuery + " WHERE (grf.grfBktType = '" & sFileZone & "'"
                                                                'SQLQuery = SQLQuery + " AND grfVefCode = " & iVefCode
                                                                'SQLQuery = SQLQuery + " AND grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND grfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                                'SQLQuery = SQLQuery & ")"
                                                                SQLQuery = "SELECT *"
                                                                SQLQuery = SQLQuery & " FROM GRF_Generic_Report, "
                                                                SQLQuery = SQLQuery & "ADF_Advertisers, "
                                                                'SQLQuery = SQLQuery & """" & "VEF_Vehicles" & """" & " vef, "
                                                                SQLQuery = SQLQuery & "VPF_Vehicle_Options, "
                                                                'SQLQuery = SQLQuery & """" & "MNF_Multi_Names" & """" & " mnf"
                                                                SQLQuery = SQLQuery & "VEF_Vehicles LEFT OUTER JOIN MNF_Multi_Names on vefMnfVehGp4Fmt = mnfCode"
                                                                SQLQuery = SQLQuery + " WHERE (grfBktType = '" & sFileZone & "'"
                                                                SQLQuery = SQLQuery + " AND grfVefCode = " & iVefCode
                                                                SQLQuery = SQLQuery + " AND grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                                'SQLQuery = SQLQuery + " AND grfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                                SQLQuery = SQLQuery + " AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
                                                                
                                                                'SQLQuery = SQLQuery + " AND (grf.grfAirDate BETWEEN '" & Format$(sSDate, "mm/dd/yyyy") & "' AND '" & Format$(sEDate, "mm/dd/yyyy") & "')"
                                                                SQLQuery = SQLQuery + " AND adfCode = grfAdfCode"
                                                                'SQLQuery = SQLQuery & " AND vefMnfVehGp4Fmt = mnfCode (+)"
                                                                SQLQuery = SQLQuery & " AND vpfvefKCode = grfVefCode"
                                                                SQLQuery = SQLQuery & " AND vefCode = grfVefCode" & ")"
                                                            End If
                                                            'CRpt1.SQLQuery = SQLQuery
                                                            'CRpt1.ReportFileName = sgReportDirectory & sLogName & "a.Rpt"    '"\Affiliate5\Reports\Stations.rpt"
                                                            slRptName = sLogName & "a.rpt"
                                                            
                                                            sYear = Format$(gObtainEndStd(sSDate), "yyyy")
                                                            sStdDate = gObtainStartStd("1/15/" & sYear)
                                                            sYear = Format$(gObtainEndStd(sStdDate), "yy")
                                                            'CRpt1.Formulas(0) = "StdYear =  '" & sYear & "'"
                                                            sgCrystlFormula1 = "'" & sYear & "'"
                                                            'CRpt1.Formulas(1) = "Week = " & ((DateValue(sSDate) - DateValue(sStdDate)) \ 7 + 1)
                                                            sgCrystlFormula2 = ((DateValue(gAdjYear(sSDate)) - DateValue(gAdjYear(sStdDate))) \ 7 + 1)
                                                            'CRpt1.Formulas(2) = "InputDate = Date(" + Format$(sSDate, "yyyy") + "," + Format$(sSDate, "mm") + "," + Format$(sSDate, "dd") + ")"
                                                            sgCrystlFormula3 = "Date(" + Format$(sSDate, "yyyy") + "," + Format$(sSDate, "mm") + "," + Format$(sSDate, "dd") + ")"
                                                            'CRpt1.Formulas(3) = "NumberDays = " & (DateValue(sEDate) - DateValue(sSDate) + 1)
                                                            sgCrystlFormula4 = (DateValue(gAdjYear(sEDate)) - DateValue(gAdjYear(sSDate)) + 1)
                                                            frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
                                                            'CRpt1.Action = 1
                                                            'CRpt1.Formulas(0) = ""
                                                            'CRpt1.Formulas(1) = ""
                                                            'CRpt1.Formulas(2) = ""
                                                            'CRpt1.Formulas(3) = ""
                                                        Else
                                                            'D.S. 10/8/04
                                                            gMsgBox "Sorry, only Logs L31 and L32 are allowed at this time."
                                                            Screen.MousePointer = vbDefault
                                                            Exit Function
                                                        
                                                            iRnfCode = tgRnfInfo(iRnf).iCode
                                                            sRptExe = tgRnfInfo(iRnf).sRptExe
                                                            sCommd = sgExeDirectory & sRptExe & " "
                                                            If iCycle >= iNoDays Then
                                                                sCommd = sCommd & "Affiliat^Prod\Guide\" & LOGSJOB & "\" & iRnfCode & "\1\" & Format$(sSDate, "mm/dd/yyyy") & "\" & iNoDays & "\12M\12M\" & iVefCode & "\" & iZoneNo & "\" & iOutput
                                                            Else
                                                                sCommd = sCommd & "Affiliat^Prod\Guide\" & LOGSJOB & "\" & iRnfCode & "\1\" & Format$(sSDate, "mm/dd/yyyy") & "\" & iCycle & "\12M\12M\" & iVefCode & "\" & iZoneNo & "\" & iOutput
                                                            End If
                                                            If iOutput = 2 Then
                                                                'Add one to index so that 0 is 1 and will be sent as -1 (1 will be sent as -2,.)
                                                                'iRptType = cboFileType.ItemData(cboFileType.ListIndex) + 1
                                                                iRptType = cboFileType.ListIndex    '3-12-04

                                                                iRptType = -iRptType
                                                                'If cboFileType.ListIndex = 2 Then
                                                                '    iRptType = 0
                                                                'ElseIf cboFileType.ListIndex = 0 Then
                                                                '    iRptType = 1
                                                                'ElseIf cboFileType.ListIndex = 4 Then
                                                                '    iRptType = 2
                                                                'ElseIf cboFileType.ListIndex = 3 Then
                                                                '    iRptType = 5
                                                                'End If
                                                                sCommd = sCommd & "\" & iRptType
                                                                gObtainYearMonthDayStr sSDate, True, sYear, sMonth, sDay
                                                                'If Val(sMonth) <= 9 Then
                                                                '    sMonth = Right$(sMonth, 1)
                                                                'ElseIf Val(sMonth) = 10 Then
                                                                '    sMonth = "A"
                                                                'ElseIf Val(sMonth) = 11 Then
                                                                '    sMonth = "B"
                                                                'ElseIf Val(sMonth) = 12 Then
                                                                '    sMonth = "C"
                                                                'End If
                                                                sLetter = Trim$(Left$(tgVehicleInfo(iIndex).sCodeStn, 3))
                                                                
                                                                If iPass = 4 Then
                                                                    sFileName = Trim$(sFileZone) & sMonth & sDay & "L_" & sLetter & ".Txt"
                                                                Else
                                                                    sFileName = Trim$(sFileZone) & sMonth & sDay & "O_" & sLetter & ".Txt"
                                                                End If
                                                                sCommd = sCommd & "\" & sFileName
                                                            Else
                                                                sCommd = sCommd & "\\"
                                                            End If
                                                            sCommd = sCommd & "\" & sGenDate & "\" & sGenTime
                                                            gShellAndWait sCommd
                                                        End If
                                                    Else
                                                        'Create Report Call
                                                        'CRpt1.Connect = "DSN = " & sgDatabaseName
                                                    
                                                        If optRptDest(0).Value = True Then
                                                            'CRpt1.Destination = crptToWindow
                                                            ilRptDest = 0
                                                        ElseIf optRptDest(1).Value = True Then
                                                            'CRpt1.Destination = crptToPrinter
                                                            ilRptDest = 1
                                                        ElseIf optRptDest(2).Value = True Then
                                                            gObtainYearMonthDayStr sSDate, True, sYear, sMonth, sDay
                                                            'If Val(sMonth) <= 9 Then
                                                            '    sMonth = Right$(sMonth, 1)
                                                            'ElseIf Val(sMonth) = 10 Then
                                                            '    sMonth = "A"
                                                            'ElseIf Val(sMonth) = 11 Then
                                                            '    sMonth = "B"
                                                            'ElseIf Val(sMonth) = 12 Then
                                                            '    sMonth = "C"
                                                            'End If
                                                            sLetter = Trim$(Left$(tgVehicleInfo(iIndex).sCodeStn, 3))
                                                            
                                                            sFileName = Trim$(sFileZone) & sMonth & sDay & "S" & sLetter '& ".Txt"
                                                            'gOutputMethod frmCP, sFileName, sOutput
                                                            'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
                                                            ilExportType = cboFileType.ListIndex    '3-12-04

                                                            slExportName = sFileName
                                                            ilRptDest = 2
                                                        Else
                                                            Screen.MousePointer = vbDefault
                                                            Exit Function
                                                        End If
                                                        SQLQuery = "SELECT *"
                                                        'SQLQuery = SQLQuery + " FROM ODF_One_Day_Log odf, ADF_Advertisers adf, VEF_Vehicles vef"
                                                        SQLQuery = SQLQuery & " FROM ODF_One_Day_Log, "
                                                        SQLQuery = SQLQuery & "ADF_Advertisers, "
                                                        SQLQuery = SQLQuery & "VEF_Vehicles"
                            '                            SQLQuery = SQLQuery + " WHERE (odfZone = '" & sFileZone & "ST'"
                                                        If Trim$(sFileZone) <> "" Then
                                                            SQLQuery = SQLQuery + " WHERE (odfZone = '" & sFileZone & "ST'"
                                                        Else
                                                            SQLQuery = SQLQuery + " WHERE (odfZone = '" & "   '"
                                                        End If
                                                        SQLQuery = SQLQuery + " AND odfVefCode = " & iVefCode
                                                        SQLQuery = SQLQuery + " AND odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "'"
                                                        'SQLQuery = SQLQuery + " AND odfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "'"
                                                        SQLQuery = SQLQuery + " AND odfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
                                                        
                                                        SQLQuery = SQLQuery + " AND (odfAirDate >= '" & Format$(sSDate, sgSQLDateForm) & "' AND odfAirDate <= '" & Format$(sEDate, sgSQLDateForm) & "')"
                                                        SQLQuery = SQLQuery + " AND adfCode = odfAdfCode"
                                                        SQLQuery = SQLQuery & " AND vefCode = odfVefCode" & ")"
                                                        'CRpt1.SQLQuery = SQLQuery
                                                        'CRpt1.ReportFileName = sgReportDirectory + "Af01.Rpt"    '"\Affiliate5\Reports\Stations.rpt"
                                                        'slRptName = sLogName & "af01.rpt"
                                                        'D.S. 05/30/03 changed slRptName
                                                        slRptName = "af01.rpt"
                                                        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
                                                        'CRpt1.Action = 1
                                                        DoEvents
                                                    End If
                                                    Screen.MousePointer = vbHourglass
                                                End If
                                                If tgVehicleInfo(iIndex).iNoZones = 0 Then
                                                    Exit For
                                                End If
                                            Next iZone
                                            Exit For
                                        End If
                                    Next iRnf
                                Next iPass
                                'SQLQuery = "DELETE FROM ODF_One_Day_Log odf WHERE (odfGenDate = '" & sGenDate & "' AND odfGenTime = " & Format(sGenTime, "hh:mm:ss") & ")"
                                SQLQuery = "DELETE FROM " & "ODF_One_Day_Log"
                                'SQLQuery = SQLQuery & " WHERE (odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND odfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "')"
                                SQLQuery = SQLQuery & " WHERE (odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND odfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
                                'cnn.BeginTrans
                                'cnn.Execute SQLQuery, rdExecDirect
                                'cnn.CommitTrans
                                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                    'grdLog.Redraw = True
                                    'Screen.MousePointer = vbDefault
                                    'mLogGen = False
                                    'Exit Function
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "CP-mLogGen"
                                    mLogGen = False
                                    Exit Function
                                End If
                                If iGenL32 Then
                                    SQLQuery = "DELETE FROM " & "GRF_Generic_Report"
                                    'SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "')"
                                    SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
                                    
                                    'cnn.BeginTrans
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    'cnn.CommitTrans
                                    If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                        'grdLog.Redraw = True
                                        'Screen.MousePointer = vbDefault
                                        'mLogGen = False
                                        'Exit Function
                                        '6/10/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "CP-mLogGen"
                                        mLogGen = False
                                        Exit Function
                                    End If
                                End If
                                If iPrint Then
                                    'Update Last CP date
                                    'SQLQuery = "UPDATE VPF_Vehicle_Options vpf SET "
                                    SQLQuery = "UPDATE " & "VPF_Vehicle_Options" & " SET "
                                    SQLQuery = SQLQuery + "vpfLastLog = '" & Format$(sEDate, sgSQLDateForm) & "'"
                                    SQLQuery = SQLQuery + " WHERE vpfvefKCode = " & iVefCode
                                    'cnn.BeginTrans
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    'cnn.CommitTrans
                                    If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                        'grdLog.Redraw = True
                                        'Screen.MousePointer = vbDefault
                                        'mLogGen = False
                                        'Exit Function
                                        '6/10/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "CP-mLogGen"
                                        mLogGen = False
                                        Exit Function
                                    End If
                                    '11/26/17
                                    gFileChgdUpdate "vpf.btr", True

                                    grdLog.TextMatrix(llRow, 2) = Format$(gObtainNextMonday(sEDate), sgShowDateForm)
                                    grdLog.TextMatrix(llRow, 0) = ""
                                End If
                            Next iDayLp
                            Exit For
                        End If
                    Next iLoop
                End If
            End If
            llRow = llRow + 1
        End If
    Next iRow
    grdLog.Redraw = True
    Screen.MousePointer = vbDefault
    mLogGen = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-mLogGen"
    grdLog.Redraw = True
    mLogGen = False
End Function

Private Function mCPGen(iPrint As Integer)

    'iPrint True = Gen CP, False = Reprint CP
    
    Dim iTRow As Integer
    Dim iRow As Integer
    Dim sDate As String
    Dim sSDate As String
    Dim sEDate As String
    Dim iVefCode As Integer
    Dim iRet As Integer
    Dim iAst As Integer
    Dim iSeqNo As Integer
    Dim iWkNo As Integer
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim iDay As Integer
    Dim iCycle As Integer
    Dim iOutput As Integer
    Dim iRnfCode As Integer
    Dim sRptExe As String
    Dim sCommd As String
    Dim sLength As String
    Dim sGenDate As String
    Dim sGenTime As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim sTimeZone As String
    Dim sStr As String
    Dim iSelected As Integer
    Dim iShf As Integer
    Dim sProd As String
    Dim iRnf As Integer
    Dim sCPName As String
    Dim dFWeek As Date
    Dim sOutput As String
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    Dim iIndex As Integer
    Dim sFileName As String
    Dim sLetter As String
    Dim sNoDays As String
    Dim iNoDays As Integer
    Dim iDayLp As Integer
    Dim ilRptDest As Integer
    Dim ilExportType As Integer
    Dim slExportName As String
    Dim slRptName As String
    Dim llRow As Long
    'Dim NewForm As New frmViewReport
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    If iPrint Then
        sNoDays = 1
        llRow = grdLog.FixedRows
        For iRow = 0 To UBound(tmCPInfo) - 1 Step 1
            If (tmCPInfo(iRow).iStatus = 0) And ((rbcVeh(0).Value = False) Or ((rbcVeh(0).Value = True) And (tmCPInfo(iRow).sVefState = "A"))) Then
                sDate = Trim$(grdLog.TextMatrix(llRow, 2))
                If Trim$(grdLog.TextMatrix(llRow, 0)) <> "" Then
                    If Not gIsDate(sDate) Then
                        Screen.MousePointer = vbDefault
                        Beep
                        gMsgBox "Please enter a valid date (m/d/yy) for " & Trim$(grdLog.TextMatrix(llRow, 1)), vbCritical
                        mCPGen = False
                        Exit Function
                    End If
                End If
                llRow = llRow + 1
            End If
        Next iRow
    Else
        sDate = txtDate.Text
        sNoDays = txtDays.Text
        If Not gIsDate(sDate) Then
            Screen.MousePointer = vbDefault
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            mCPGen = False
            Exit Function
        End If
    End If
    If Not iPrint Then
        If Trim$(sNoDays) = "" Then
            Screen.MousePointer = vbDefault
            Beep
            gMsgBox "Please enter Number Days.", vbCritical
            mCPGen = False
            Exit Function
        End If
    End If
    iNoDays = 7 * Val(sNoDays) - 1
    grdLog.Redraw = False
    llRow = grdLog.FixedRows
    'D.S. 11/21/05
'    iRet = gGetMaxAstCode()
'    If Not iRet Then
'        Exit Function
'    End If
    
    For iRow = 0 To UBound(tmCPInfo) - 1 Step 1
        If (tmCPInfo(iRow).iStatus = 0) And ((rbcVeh(0).Value = False) Or ((rbcVeh(0).Value = True) And (tmCPInfo(iRow).sVefState = "A"))) Then
            iVefCode = tmCPInfo(iRow).iVefCode
            If iPrint Then
                sDate = Trim$(grdLog.TextMatrix(llRow, 2))
                If Trim$(grdLog.TextMatrix(llRow, 0)) <> "" Then
                    iSelected = True
                Else
                    iSelected = False
                End If
                If Not gIsDate(sDate) Then
                    iSelected = False
                End If
            Else 'Reprint CPs
                sDate = txtDate.Text
                'If imVefCode = iVefCode Then
                '    iSelected = True
                'Else
                '    iSelected = False
                'End If
                If optSort(0).Value Then
                    'Vehicle, then Stations
                    If lbcSingle.ListIndex >= 0 Then
                        If iVefCode = lbcSingle.ItemData(lbcSingle.ListIndex) Then
                            iSelected = True
                            imVefCode = iVefCode
                        Else
                            iSelected = False
                        End If
                    Else
                        iSelected = False
                    End If
                Else
                    'Stations, Then vehicle
                    For iLoop = 0 To lbcMulti.ListCount - 1 Step 1
                        If iVefCode = lbcMulti.ItemData(iLoop) Then
                            If lbcMulti.Selected(iLoop) Then
                                iSelected = True
                                imVefCode = iVefCode
                            Else
                                iSelected = False
                            End If
                            Exit For
                        End If
                    Next iLoop
                End If
            End If
            If (iSelected) And (Len(sDate) > 0) Then
                'D.S. 10/8/04  Currently we only support C17
                sCPName = Trim$(grdLog.TextMatrix(llRow, 4))
                If StrComp("C17", sCPName, 1) <> 0 Then
                    gMsgBox "Sorry, only CP C17 is allowed at this time."
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
                iIndex = -1
                For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                    If tgVehicleInfo(iLoop).iCode = iVefCode Then
                        iIndex = iLoop
                        Exit For
                    End If
                Next iLoop
                For iLoop = 0 To UBound(tgRnfInfo) - 1 Step 1
                    If (StrComp(Trim$(tgRnfInfo(iLoop).sName), Trim$(grdLog.TextMatrix(llRow, 4)), 1) = 0) Or (StrComp("C17", Trim$(grdLog.TextMatrix(llRow, 4)), 1) = 0) Then
                        
                        If iPrint Then
                            sDate = Trim$(grdLog.TextMatrix(llRow, 2))
                            sNoDays = Val(grdLog.TextMatrix(llRow, 3))
                            iNoDays = Val(sNoDays)
                            iCycle = Val(grdLog.TextMatrix(llRow, 3))
                        Else
                            iCycle = Val(grdLog.TextMatrix(llRow, 3))
                        End If
                        For iDayLp = 1 To iNoDays Step iCycle
                        
                            sSDate = gAdjYear(Format$(DateValue(sDate) + iDayLp - 1, sgShowDateForm))   'sDate
                            sEDate = gAdjYear(Format$(DateValue(sSDate) + iCycle - 1, sgShowDateForm))   ' sgShowDateForm)
                            ''Get CPTT so that Stations requiring CP can be obtained
                            'SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
                            'SQLQuery = SQLQuery + " FROM shtt, cptt, att"
                            SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, mktName"
                            SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"
                            SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
                            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
                            SQLQuery = SQLQuery + " AND cpttVefCode = " & iVefCode
                            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
                            Set cprst = gSQLSelectCall(SQLQuery)
                            While Not cprst.EOF
                                lacProcess.Caption = "Processing : " & Trim$(grdLog.TextMatrix(llRow, 1))
                                DoEvents
                                If iPrint Then
                                    iSelected = True
                                Else
                                    iSelected = False
                                    If optSort(0).Value Then
                                        'Vehicle, then Stations
                                        For iShf = 0 To lbcMulti.ListCount - 1 Step 1
                                            If lbcMulti.Selected(iShf) Then
                                                If lbcMulti.ItemData(iShf) = cprst!shttCode Then
                                                    iSelected = True
                                                    Exit For
                                                End If
                                            Else
                                                If lbcMulti.ItemData(iShf) = cprst!shttCode Then
                                                    iSelected = False
                                                    Exit For
                                                End If
                                            End If
                                        Next iShf
                                    Else
                                        'Station, Then vehicle
                                        If lbcSingle.ListIndex >= 0 Then
                                            If cprst!shttCode = lbcSingle.ItemData(lbcSingle.ListIndex) Then
                                                iSelected = True
                                            Else
                                                iSelected = False
                                            End If
                                        Else
                                            iSelected = False
                                        End If
                                    End If
                                End If
                                'Remove allowing reprint to include PrintCP as No (1)
                                'If (iSelected) And (((cprst!attPrintCP = 0) And (iPrint)) Or (Not iPrint)) Then
                                
                                If (iSelected) And (cprst!attPrintCP = 0) Then
                                    lacProcess.Caption = "Processing : " & Trim$(grdLog.TextMatrix(llRow, 1)) & "/" & Trim$(cprst!shttCallLetters)
                                    DoEvents
                                    ReDim tgCPPosting(0 To 1) As CPPOSTING
                                    tgCPPosting(0).lCpttCode = cprst!cpttCode
                                    tgCPPosting(0).iStatus = cprst!cpttStatus
                                    tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                                    tgCPPosting(0).lAttCode = cprst!cpttatfCode
                                    tgCPPosting(0).iAttTimeType = cprst!attTimeType
                                    tgCPPosting(0).iVefCode = iVefCode
                                    tgCPPosting(0).iShttCode = cprst!shttCode
                                    If tgVehicleInfo(iIndex).iNoZones <> 0 Then
                                        sTimeZone = cprst!shttTimeZone
                                    Else
                                        sTimeZone = "   "
                                    End If
                                    tgCPPosting(0).sZone = sTimeZone
                                    tgCPPosting(0).sDate = Format$(sSDate, sgShowDateForm)
                                    tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                                    'Create AST records
                                    igTimes = 1 'By Week
                                    iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, False, True)
                                    gClearASTInfo False
                                    sCPName = Trim$(grdLog.TextMatrix(llRow, 4))
                                    iRnfCode = tgRnfInfo(iLoop).iCode
                                    sRptExe = tgRnfInfo(iLoop).sRptExe
                                    If IsNull(cprst!attGenCP) <> True Then
                                        For iRnf = 0 To UBound(tgRnfInfo) - 1 Step 1
                                            If (StrComp(Trim$(tgRnfInfo(iRnf).sName), Trim$(cprst!attGenCP), 1) = 0) Or (StrComp("C17", Trim$(cprst!attGenCP), 1) = 0) Then
                                                sCPName = Trim$(cprst!attGenCP)
                                                iRnfCode = tgRnfInfo(iRnf).iCode
                                                sRptExe = tgRnfInfo(iRnf).sRptExe
                                                Exit For
                                            End If
                                        Next iRnf
                                    End If
                                    If StrComp("C17", sCPName, 1) <> 0 Then
                                        'SQLQuery = "DELETE FROM ODF_One_Day_Log odf WHERE (odfVefCode = " & iVefCode & ")"
                                        'cnn.BeginTrans
                                        'cnn.Execute SQLQuery, rdExecDirect
                                        'cnn.CommitTrans
                                        sGenDate = Format$(gNow(), sgShowDateForm)
                                        sGenTime = Format$(gNow(), sgShowTimeWSecForm)
                                        iZone = 0   'All
                                        Select Case UCase$(Left$(cprst!shttTimeZone, 1))
                                            Case "E"
                                                iZone = 1
                                            Case "C"
                                                iZone = 2
                                            Case "M"
                                                iZone = 3
                                            Case "P"
                                                iZone = 4
                                            Case Else
                                                iZone = 0
                                                'Should test vpf and if defined, then use first one
                                        End Select
                                        iSeqNo = 1
                                        For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
                                            If tgStatusTypes(gGetAirStatus(tmAstInfo(iAst).iStatus)).iPledged <> 2 Then
                                                SQLQuery = "SELECT * FROM lst "
                                                SQLQuery = SQLQuery + " WHERE (lstCode = " & tmAstInfo(iAst).lLstCode & ")"
                                                Set lstrst = gSQLSelectCall(SQLQuery)
                                                If Not lstrst.EOF Then
                                                    'iWkNo = (DateValue(lstrst!lstLogDate) - DateValue(sSDate)) \ 7 + 1
                                                    iWkNo = (DateValue(gAdjYear(tmAstInfo(iAst).sFeedDate)) - DateValue(gAdjYear(sSDate))) \ 7 + 1
                                                    'Select Case WeekDay(lstrst!lstLogDate)
                                                    Select Case Weekday(tmAstInfo(iAst).sFeedDate)
                                                        Case vbSaturday
                                                            iDay = 6
                                                        Case vbSunday
                                                            iDay = 7
                                                        Case Else
                                                            iDay = 1
                                                    End Select
                                                    If lstrst!lstLen < 60 Then
                                                        'sLength = "00:00:30"
                                                        sStr = Trim$(Str$(lstrst!lstLen))
                                                        If Len(sStr) = 1 Then
                                                            sStr = "0" & sStr
                                                        End If
                                                        sLength = "00:00:" & sStr
                                                    Else
                                                        'sLength = "00:01:00"
                                                        sStr = Trim$(Str$(lstrst!lstLen - 60))
                                                        If Len(sStr) = 1 Then
                                                            sStr = "0" & sStr
                                                        End If
                                                        sLength = "00:01:" & sStr
                                                    End If
                                                    'sAirDate = Format$(lstrst!lstLogDate, "mm/dd/yyyy")
                                                    'sAirTime = Format(lstrst!lstLogTime, "hh:mm:ss")
                                                    sAirDate = Format$(tmAstInfo(iAst).sFeedDate, sgShowDateForm)
                                                    sAirTime = Format(tmAstInfo(iAst).sFeedTime, sgShowTimeWSecForm)
                                                    If IsNull(lstrst!lstProd) Then
                                                        sProd = ""
                                                    Else
                                                        sProd = gFixQuote(lstrst!lstProd)
                                                    End If
                                                    'Build ODF
                                                    'SQLQuery = "INSERT INTO ODF_One_Day_Log (odfUrfCode, odfVefCode, odfAirDate, odfAirTime, "
                                                    SQLQuery = "INSERT INTO " & "ODF_One_Day_Log"
                                                    SQLQuery = SQLQuery & " (odfUrfCode, odfVefCode, odfAirDate, odfAirTime, "
                                                    SQLQuery = SQLQuery & "odfSeqNo, odfLocalTime, odfFeedTime, odfZone, odfEtfCode, "
                                                    SQLQuery = SQLQuery & "odfEnfCode, odfProgCode, odfMnfFeed, odfWkNo, odfAnfCode, "
                                                    SQLQuery = SQLQuery & "odfUnits, odfLength, odfAdfCode, odfCifCode, odfProduct, "
                                                    SQLQuery = SQLQuery & "odfMnfSubFeed, odfCntrNo, odfBreakNo, odfPositionNo, odfType, "
                                                    SQLQuery = SQLQuery & "odfCefCode, odfShortTitle, odfPageEjectFlag, odfSortSeq, "
                                                    SQLQuery = SQLQuery & "odfAvailCefCode, odfRdfSortCode, odfDPDesc, odfChfCxfCode, odfDaySort, "
                                                    SQLQuery = SQLQuery & "odfEvtCefCode, odfEvtCefSort, odfEvtIDCefCode, odfDupeAvailID, odfLogType, "
                                                    SQLQuery = SQLQuery & "odfAvailLen, odfAvailLock, "
                                                    SQLQuery = SQLQuery & "odfGenDate, odfGenTime)" ', "
                                                    'Temporary remove header and footer and vehicle name so release can be sent-  to make work we need to have
                                                    'a blank comment record to point to if ---cefcode values are zero
                                                    'SQLQuery = SQLQuery & "odfHd1CefCode, odfFt1CefCode, odfFt2CefCode, odfVehNmCefCode) "
                                                    SQLQuery = SQLQuery & " VALUES (" & 1 & ", " & iVefCode & ", '" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', "
                                                    SQLQuery = SQLQuery & iSeqNo & ", '" & Format$(sAirTime, sgSQLTimeForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', '" & sTimeZone & "', " & 0 & ", "
                                                    SQLQuery = SQLQuery & 0 & ", '" & "" & "', " & 0 & ", " & iWkNo & ", " & lstrst!lstAnfCode & ", "
                                                    SQLQuery = SQLQuery & 0 & ", '" & sLength & "', " & lstrst!lstAdfCode & ", " & lstrst!lstCifCode & ", '" & sProd & "', "
                                                    SQLQuery = SQLQuery & 0 & ", " & lstrst!lstCntrNo & ", " & lstrst!lstBreakNo & ", " & lstrst!lstPositionNo & ", " & 4 & ", "
                                                    SQLQuery = SQLQuery & 0 & ", '" & "" & "', '" & "N" & "', " & iSeqNo & ", "
                                                    SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & "" & "', " & 0 & ", " & iDay & ", "
                                                    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "" & "', '" & "F" & "'" & ", "
                                                    SQLQuery = SQLQuery & "0" & ", " & "'N'" & ", "
                                                    'SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Format(sGenTime, sgSQLTimeForm) & "')" '", "
                                                    SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"  '", "
                                                    
                                                    ''Temporary remove header and footer and vehicle name so release can be sent-  to make work we need to have
                                                    ''a blank comment record to point to if ---cefcode values are zero
                                                    ''SQLQuery = SQLQuery & tgVehicleInfo(iIndex).lHd1CefCode & ", " & tgVehicleInfo(iIndex).lLgNmCefCode & ", " & tgVehicleInfo(iIndex).lFt1CefCode & ", " & tgVehicleInfo(iIndex).lFt2CefCode & ")"
                                                    'cnn.BeginTrans
                                                    'cnn.Execute SQLQuery, rdExecDirect
                                                    'cnn.CommitTrans
                                                    If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                                        'grdLog.Redraw = True
                                                        'Screen.MousePointer = vbDefault
                                                        'mCPGen = False
                                                        'Exit Function
                                                        '6/10/16: Replaced GoSub
                                                        'GoSub ErrHand:
                                                        Screen.MousePointer = vbDefault
                                                        gHandleError "AffErrorLog.txt", "CP-mCPGen"
                                                        mCPGen = False
                                                        Exit Function
                                                    End If
                                                    iSeqNo = iSeqNo + 1
                                                End If
                                            End If
                                        Next iAst
                                        sCommd = sgExeDirectory & sRptExe & " "
                                        sCommd = sCommd & "Affiliat^Prod\Guide\" & LOGSJOB & "\" & iRnfCode & "\1\" & sSDate & "\" & iCycle & "\12:00:00 AM\12:00:00 AM\" & iVefCode & "\" & iZone & "\" & iOutput & "\\\" & sGenDate & "\" & sGenTime
                                        gShellAndWait sCommd
                                        'SQLQuery = "DELETE FROM ODF_One_Day_Log odf WHERE (odfGenDate = '" & sGenDate & "' AND odfGenTime = " & Format(sGenTime, "hh:mm:ss") & ")"
                                        SQLQuery = "DELETE FROM " & "ODF_One_Day_Log"
                                        'SQLQuery = SQLQuery & " WHERE (odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND odfGenTime = '" & Format(sGenTime, sgSQLTimeForm) & "')"
                                        SQLQuery = SQLQuery & " WHERE (odfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND odfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
                                        
                                        'cnn.BeginTrans
                                        'cnn.Execute SQLQuery, rdExecDirect
                                        'cnn.CommitTrans
                                        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                            'grdLog.Redraw = True
                                            'Screen.MousePointer = vbDefault
                                            'mCPGen = False
                                            'Exit Function
                                            '6/10/16: Replaced GoSub
                                            'GoSub ErrHand:
                                            Screen.MousePointer = vbDefault
                                            gHandleError "AffErrorLog.txt", "CP-mCPGen"
                                            mCPGen = False
                                            Exit Function
                                        End If
                                        Screen.MousePointer = vbHourglass
                                    Else
                                        'Create Report Call
                                        'CRpt1.Connect = "DSN = " & sgDatabaseName
                                    
                                        If optRptDest(0).Value = True Then
                                            'CRpt1.Destination = crptToWindow
                                            ilRptDest = 0
                                        ElseIf optRptDest(1).Value = True Then
                                            'CRpt1.Destination = crptToPrinter
                                            ilRptDest = 1
                                        ElseIf optRptDest(2).Value = True Then
                                            gObtainYearMonthDayStr sSDate, True, sYear, sMonth, sDay
                                            'If Val(sMonth) <= 9 Then
                                            '    sMonth = Right$(sMonth, 1)
                                            'ElseIf Val(sMonth) = 10 Then
                                            '    sMonth = "A"
                                            'ElseIf Val(sMonth) = 11 Then
                                            '    sMonth = "B"
                                            'ElseIf Val(sMonth) = 12 Then
                                            '    sMonth = "C"
                                            'End If
                                            sLetter = Trim$(Left$(tgVehicleInfo(iIndex).sCodeStn, 3))

                                            'sFileName = sMonth & sDay & "CP" & sLetter '& ".Txt"
                                            sFileName = sMonth & sDay & "C_" & sLetter & "_" & Trim$(cprst!shttCallLetters)
                                            'gOutputMethod frmCP, sFileName, sOutput
                                            'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
                                            ilExportType = cboFileType.ListIndex    '3-12-04

                                            slExportName = sFileName
                                            ilRptDest = 2
                                        Else
                                            Screen.MousePointer = vbDefault
                                            Exit Function
                                        End If
                                        If igSQLSpec = 0 Then
                                            SQLQuery = "SELECT *"
                                            'SQLQuery = SQLQuery + " FROM ast, shtt, lst, ADF_Advertisers adf, VEF_Vehicles vef, MNF_Multi_Names mnf"
                                            SQLQuery = SQLQuery & " FROM ast, shtt, lst, ADF_Advertisers, "
                                            SQLQuery = SQLQuery & "VEF_Vehicles, "
                                            SQLQuery = SQLQuery & "MNF_Multi_Names"
                                            SQLQuery = SQLQuery + " WHERE (astatfCode = " & tgCPPosting(0).lAttCode
                                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(DateValue(sSDate) - 1, sgSQLDateForm) & "' AND astFeedDate <= '" & gAdjYear(Format$(DateValue(sEDate) + 1, sgSQLDateForm)) & "')"
                                            SQLQuery = SQLQuery + " AND (astAirDate >= '" & Format$(sSDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sEDate, sgSQLDateForm) & "')"
                                            
                                            'D.S. 1/3/02
                                            'Suppress all spots not carried by the affiliate - save paper and confusion
                                            If (rbcShowAll(0).Value And (frcRePrintCPNotCarried.Visible = True)) Or (rbcShowAll(3).Value And (frcPrintCPNotCarried.Visible = True)) Then
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 2"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 3"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 4"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 5"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 6"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 8"
                                                'SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 22"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> " & ASTAIR_MISSED_MG_BYPASS
                                            End If
                                            
                                            SQLQuery = SQLQuery & " AND lstCode = astLsfCode"
                                            SQLQuery = SQLQuery & " AND vefCode = astVefCode"
                                            SQLQuery = SQLQuery & " AND shttCode = astShfCode"
                                            SQLQuery = SQLQuery & " AND vefMnfVehGp4Fmt = mnfCode (+)"
                                            SQLQuery = SQLQuery & " AND adfCode = lstAdfCode" & ")"
                                        Else  'Pervasive 2000
                                            SQLQuery = "SELECT *"
                                            SQLQuery = SQLQuery & " FROM ast, shtt, lst, ADF_Advertisers, "
                                            'SQLQuery = SQLQuery & """" & "VEF_Vehicles" & """" & " vef, "
                                            'SQLQuery = SQLQuery & """" & "MNF_Multi_Names" & """" & " mnf"
                                            SQLQuery = SQLQuery & "VEF_Vehicles LEFT OUTER JOIN MNF_Multi_Names On vefMnfVehGp4Fmt = mnfCode"
                                            SQLQuery = SQLQuery + " WHERE (astatfCode = " & tgCPPosting(0).lAttCode
                                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & gAdjYear(Format$(DateValue(sSDate) - 1, sgSQLDateForm)) & "' AND astFeedDate <= '" & gAdjYear(Format$(DateValue(sEDate) + 1, sgSQLDateForm)) & "')"
                                            SQLQuery = SQLQuery + " AND (astAirDate >= '" & Format$(sSDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(sEDate, sgSQLDateForm) & "')"
    
                                            'D.S. 1/3/02
                                            'Suppress all spots not carried by the affiliate - save paper and confusion
                                            If rbcShowAll(0).Value Or rbcShowAll(3).Value Then
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 2"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 3"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 4"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 5"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 6"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 8"
                                                'SQLQuery = SQLQuery & " AND astStatus Mod 100 <> 22"
                                                SQLQuery = SQLQuery & " AND astStatus Mod 100 <> " & ASTAIR_MISSED_MG_BYPASS
                                            End If
                                            
                                            SQLQuery = SQLQuery & " AND lstCode = astLsfCode"
                                            SQLQuery = SQLQuery & " AND vefCode = astVefCode"
                                            SQLQuery = SQLQuery & " AND shttCode = astShfCode"
                                            'SQLQuery = SQLQuery & " AND vef.vefMnfVehGp4Fmt = mnf.mnfCode (+)"
                                            SQLQuery = SQLQuery & " AND adfCode = lstAdfCode" & ")"
                                        End If
                                        'CRpt1.SQLQuery = SQLQuery
                                        'CRpt1.ReportFileName = sgReportDirectory + "C17.Rpt"    '"\Affiliate5\Reports\Stations.rpt"
                                        slRptName = "C17.rpt"
                                        
                                        dFWeek = CDate(sSDate)
                                        'CRpt1.Formulas(0) = "StartDate = Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
                                        sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
                                        dFWeek = CDate(sEDate)
                                        'CRpt1.Formulas(1) = "EndDate = Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
                                        sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
                                     
                                        If TabStrip1.SelectedItem.Index = 1 Then        'final cp
                                            'CRpt1.Formulas(2) = "CoverPageOnly = 'N'"      'always show detail if its not a reprint
                                            sgCrystlFormula3 = "N"
                                        Else                                            'reprint cp
                                            'reprint, cover page only?
                                            If ckcCover.Value = 1 Then
                                                'CRpt1.Formulas(2) = "CoverPageOnly = 'Y'"
                                                sgCrystlFormula3 = "Y"
                                            Else
                                                'CRpt1.Formulas(2) = "CoverPageOnly = 'N'"
                                                sgCrystlFormula3 = "N"
                                            End If
                                        End If
                                        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
                                        'CRpt1.Action = 1
                                        'CRpt1.Formulas(0) = ""
                                        'CRpt1.Formulas(1) = ""
                                        'CRpt1.Formulas(2) = ""
                                        Screen.MousePointer = vbHourglass
                                    End If
                                End If
                                cprst.MoveNext
                            Wend
                            If iPrint Then
                                'Update Last CP date
                                'SQLQuery = "UPDATE VPF_Vehicle_Options vpf SET "
                                SQLQuery = "UPDATE " & "VPF_Vehicle_Options" & " SET "
                                SQLQuery = SQLQuery + "vpfLastCP = '" & Format$(sEDate, sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery + " WHERE vpfvefKCode = " & iVefCode & ""
                                'cnn.BeginTrans
                                'cnn.Execute SQLQuery, rdExecDirect
                                'cnn.CommitTrans
                                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                    'grdLog.Redraw = True
                                    'Screen.MousePointer = vbDefault
                                    'mCPGen = False
                                    'Exit Function
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "CP-mCPGen"
                                    mCPGen = False
                                    Exit Function
                                End If
                                '11/26/17
                                gFileChgdUpdate "vpf.btr", True
                                grdLog.TextMatrix(llRow, 2) = Format$(gObtainNextMonday(sEDate), sgShowDateForm)
                                grdLog.TextMatrix(llRow, 0) = ""
                            End If
                        Next iDayLp
                        Exit For
                    End If
                Next iLoop
            End If
            llRow = llRow + 1
        End If
    Next iRow
    gCloseRegionSQLRst
    grdLog.Redraw = True
    mCPGen = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-mCPGen"
    grdLog.Redraw = True
    mCPGen = False
End Function

Private Sub txtDays_GotFocus()
    Dim sDate As String
    
    If Trim$(txtDays.Text) = "" Then
        If igCPOrLog = 0 Then
            txtDays.Text = "1"
        Else
            sDate = txtDate.Text
            If Not gIsDate(sDate) Then
                txtDays.Text = "1"
            Else
                Select Case Weekday(sDate)
                    Case vbMonday
                        txtDays.Text = "7"
                        imMaxDays = 7
                    Case vbTuesday
                        txtDays.Text = "6"
                        imMaxDays = 6
                    Case vbWednesday
                        txtDays.Text = "5"
                        imMaxDays = 5
                    Case vbThursday
                        txtDays.Text = "4"
                        imMaxDays = 4
                    Case vbFriday
                        txtDays.Text = "3"
                        imMaxDays = 3
                    Case vbSaturday
                        txtDays.Text = "2"
                        imMaxDays = 2
                    Case vbSunday
                        txtDays.Text = "1"
                        imMaxDays = 1
                    Case Else
                        txtDays.Text = ""
                End Select
            End If
        End If
    End If
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
    If igCPOrLog <> 0 Then
        slStr = txtDays.Text
        slStr = Left$(slStr, txtDays.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - txtDays.SelStart - txtDays.SelLength)
        If Val(slStr) > Val(imMaxDays) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Function mGenCmmlSum(sGenDate As String, sGenTime As String) As Integer
    Dim iLoop As Integer
    Dim sZone As String
    
    'Loop thru tmCmmlSum and create one disk record for entry
    For iLoop = LBound(tmCmmlSum) To UBound(tmCmmlSum) - 1 Step 1
        SQLQuery = "INSERT INTO " & "GRF_Generic_Report"
        SQLQuery = SQLQuery & " (grfBktType, grfvefCode, grfadfCode, "
        SQLQuery = SQLQuery & "grfGenDesc, grfCode2, grfPer1, "
        SQLQuery = SQLQuery & "grfPer2, grfPer3, grfPer4, "
        SQLQuery = SQLQuery & "grfPer5, grfPer6, grfPer7, "
        SQLQuery = SQLQuery & "grfPer8, grfPer9, grfPer10, "
        SQLQuery = SQLQuery & "grfPer11, grfPer12, grfPer13, "
        SQLQuery = SQLQuery & "grfPer14, grfPer15, grfPer16, "
        SQLQuery = SQLQuery & "grfPer17, grfPer18, grfPer1Genl, "
        SQLQuery = SQLQuery & "grfPer2Genl, grfPer3Genl, grfPer4Genl, "
        SQLQuery = SQLQuery & "grfPer5Genl, grfPer6Genl, grfPer7Genl, "
        SQLQuery = SQLQuery & "grfGenDate, grfGenTime)"
        
        SQLQuery = SQLQuery & " VALUES ('" & Left$(Trim$(tmCmmlSum(iLoop).sZone), 1) & "', " & tmCmmlSum(iLoop).iVefCode & ", " & tmCmmlSum(iLoop).iAdfCode & ", "
        SQLQuery = SQLQuery & "'" & Trim$(tmCmmlSum(iLoop).sProduct) & "', " & tmCmmlSum(iLoop).iLen & ", " & tmCmmlSum(iLoop).iMFEarly & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iMFAM & ", " & tmCmmlSum(iLoop).iMFMid & ", " & tmCmmlSum(iLoop).iMFPM & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iMFEve & ", " & tmCmmlSum(iLoop).iSaEarly & ", " & tmCmmlSum(iLoop).iSaAM & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iSaMid & ", " & tmCmmlSum(iLoop).iSaPM & ", " & tmCmmlSum(iLoop).iSaEve & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iSuEarly & ", " & tmCmmlSum(iLoop).iSuAM & ", " & tmCmmlSum(iLoop).iSuMid & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iSuPM & ", " & tmCmmlSum(iLoop).iSuEve & ", " & tmCmmlSum(iLoop).iMFEarliest & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iSaEarliest & ", " & tmCmmlSum(iLoop).iSuEarliest & ", " & tmCmmlSum(iLoop).iDay(0) & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iDay(1) & ", " & tmCmmlSum(iLoop).iDay(2) & ", " & tmCmmlSum(iLoop).iDay(3) & ", "
        SQLQuery = SQLQuery & tmCmmlSum(iLoop).iDay(4) & ", " & tmCmmlSum(iLoop).iDay(5) & ", " & tmCmmlSum(iLoop).iDay(6) & ", "
        'SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Format(sGenTime, sgSQLTimeForm) & "')"   '", "
        SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"    '", "
        
        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
            'mGenCmmlSum = False
            'Exit Function
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "CP-mGenCmmlSum"
            mGenCmmlSum = False
            Exit Function
        End If
                
    Next iLoop
    mGenCmmlSum = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-mGenCmmlSum"
    mGenCmmlSum = False
    Exit Function
End Function

Public Sub mCmmlSum(sRptName As String, iDay As Integer, iVehIndex As Integer, sCntrNo As String, sZone As String, sAirTime As String, iVefCode As Integer, iAdfCode As Integer, sProduct As String, iLen As Integer)
    Dim lAirTime As Long
    Dim lGenEndTime As Long
    Dim iTest As Integer
    Dim iFound As Integer
    Dim iOk As Integer
    ReDim lDPStartTime(0 To 4) As Long
    ReDim lDPEndtime(0 To 4) As Long
    
    On Error GoTo ErrHand
    ''Get contract to eliminate PSA and Promos
    'If Val(smCntrNo) <> Val(sCntrNo) Then
    '    SQLQuery = "SELECT chfType "
    '    SQLQuery = SQLQuery & " FROM " & """" & "CHF_Contract_Header" & """" & " chf"
    '    SQLQuery = SQLQuery + " WHERE (chf.chfCntrNo = " & sCntrNo & ")"
    '    SQLQuery = SQLQuery + " ORDER BY chf.chfCntrNo, chf.chfCntRevNo Desc"
    '    Set chfrst = gSQLSelectCall(SQLQuery)
    '    If Not chfrst.EOF Then
    '        smChfType = chfrst!chfType
    '        smCntrNo = sCntrNo
    '        iOk = True
    '    Else
    '        iOk = False
    '    End If
    'Else
    '    iOk = True
    'End If
    'If iOk Then
    '    If (smChfType <> "S") And (smChfType <> "M") Then
            Select Case UCase$(Left$(sZone, 1))
                Case "E"
                    lDPStartTime(0) = 0
                    'lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1)) - 1 '21599    '6am
                    'lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1))
                    'lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) - 1 '35999
                    'lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) '36000  '10Am
                    'lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) - 1 '53999
                    'lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) '54000  '3pm
                    'lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(4)) - 1 '68399
                    'lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(4)) '68400  '7pm
                    'lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(5)) - 1 '86399
                    lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(0)) - 1 '21599    '6am
                    lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(0))
                    lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1)) - 1 '35999
                    lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1)) '36000  '10Am
                    lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) - 1 '53999
                    lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) '54000  '3pm
                    lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) - 1 '68399
                    lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) '68400  '7pm
                    lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(4)) - 1 '86399
                    
                Case "C"
                    lDPStartTime(0) = 0
                    'lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(1)) - 1 '21599    '6am
                    'lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(1))
                    'lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(2)) - 1 '35999
                    'lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(2)) '36000  '10Am
                    'lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(3)) - 1 '53999
                    'lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(3)) '54000  '3pm
                    'lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(4)) - 1 '68399
                    'lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(4)) '68400  '7pm
                    'lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(5)) - 1 '86399
                    lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(0)) - 1 '21599    '6am
                    lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(0))
                    lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(1)) - 1 '35999
                    lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(1)) '36000  '10Am
                    lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(2)) - 1 '53999
                    lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(2)) '54000  '3pm
                    lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(3)) - 1 '68399
                    lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(3)) '68400  '7pm
                    lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iCSTEndTime(4)) - 1 '86399
                    
                Case "M"
                    lDPStartTime(0) = 0
                    'lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(1)) - 1 '21599    '6am
                    'lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(1))
                    'lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(2)) - 1 '35999
                    'lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(2)) '36000  '10Am
                    'lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(3)) - 1 '53999
                    'lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(3)) '54000  '3pm
                    'lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(4)) - 1 '68399
                    'lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(4)) '68400  '7pm
                    'lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(5)) - 1 '86399
                    lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(0)) - 1 '21599    '6am
                    lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(0))
                    lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(1)) - 1 '35999
                    lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(1)) '36000  '10Am
                    lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(2)) - 1 '53999
                    lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(2)) '54000  '3pm
                    lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(3)) - 1 '68399
                    lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(3)) '68400  '7pm
                    lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iMSTEndTime(4)) - 1 '86399
                    
                Case "P"
                    lDPStartTime(0) = 0
                    'lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(1)) - 1 '21599    '6am
                    'lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(1))
                    'lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(2)) - 1 '35999
                    'lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(2)) '36000  '10Am
                    'lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(3)) - 1 '53999
                    'lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(3)) '54000  '3pm
                    'lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(4)) - 1 '68399
                    'lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(4)) '68400  '7pm
                    'lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(5)) - 1 '86399
                    lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(0)) - 1 '21599    '6am
                    lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(0))
                    lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(1)) - 1 '35999
                    lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(1)) '36000  '10Am
                    lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(2)) - 1 '53999
                    lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(2)) '54000  '3pm
                    lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(3)) - 1 '68399
                    lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(3)) '68400  '7pm
                    lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iPSTEndTime(4)) - 1 '86399
                    
                Case Else
                    lDPStartTime(0) = 0
                    'lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1)) - 1 '21599    '6am
                    'lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1))
                    'lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) - 1 '35999
                    'lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) '36000  '10Am
                    'lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) - 1 '53999
                    'lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) '54000  '3pm
                    'lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(4)) - 1 '68399
                    'lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(4)) '68400  '7pm
                    'lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(5)) - 1 '86399
                    lDPEndtime(0) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(0)) - 1 '21599    '6am
                    lDPStartTime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(0))
                    lDPEndtime(1) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1)) - 1 '35999
                    lDPStartTime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(1)) '36000  '10Am
                    lDPEndtime(2) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) - 1 '53999
                    lDPStartTime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(2)) '54000  '3pm
                    lDPEndtime(3) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) - 1 '68399
                    lDPStartTime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(3)) '68400  '7pm
                    lDPEndtime(4) = 60 * CLng(tgVehicleInfo(iVehIndex).iESTEndTime(4)) - 1 '86399
                    
            End Select
            lAirTime = gTimeToLong(sAirTime, True)
            lGenEndTime = lDPEndtime(4)
            If sRptName = "L35" Then
                lGenEndTime = 86400           '12m
            End If
            If (lAirTime >= lDPStartTime(0)) And (lAirTime <= lGenEndTime) Then  'testing for correct time 8/30/99
                iFound = -1
                For iTest = 0 To UBound(tmCmmlSum) - 1 Step 1
                    'use advt/prod names, not short title
                    If (tmCmmlSum(iTest).iVefCode = iVefCode) And (StrComp(Trim$(tmCmmlSum(iTest).sZone), Trim$(sZone), 1) = 0) And (tmCmmlSum(iTest).iAdfCode = iAdfCode) And (Trim$(tmCmmlSum(iTest).sProduct) = Trim$(sProduct)) And (tmCmmlSum(iTest).iLen = iLen) Then
                        iFound = iTest
                        Exit For
                    End If
                Next iTest
                If iFound = -1 Then
                    iFound = UBound(tmCmmlSum)
                    ReDim Preserve tmCmmlSum(0 To UBound(tmCmmlSum) + 1) As CMMLSUM
                    tmCmmlSum(iFound).iVefCode = iVefCode
                    tmCmmlSum(iFound).sZone = sZone
                    tmCmmlSum(iFound).iAdfCode = iAdfCode
                    tmCmmlSum(iFound).sProduct = Trim$(sProduct)
                    tmCmmlSum(iFound).iLen = iLen
                    tmCmmlSum(iFound).iMFEarliest = 0
                    tmCmmlSum(iFound).iSaEarliest = 0
                    tmCmmlSum(iFound).iSuEarliest = 0
                    tmCmmlSum(iFound).iMFEarly = 0
                    tmCmmlSum(iFound).iSaEarly = 0
                    tmCmmlSum(iFound).iSuEarly = 0
                    tmCmmlSum(iFound).iMFAM = 0
                    tmCmmlSum(iFound).iSaAM = 0
                    tmCmmlSum(iFound).iSuAM = 0
                    tmCmmlSum(iFound).iMFMid = 0
                    tmCmmlSum(iFound).iSaMid = 0
                    tmCmmlSum(iFound).iSuMid = 0
                    tmCmmlSum(iFound).iMFPM = 0
                    tmCmmlSum(iFound).iSaPM = 0
                    tmCmmlSum(iFound).iSuPM = 0
                    tmCmmlSum(iFound).iMFEve = 0
                    tmCmmlSum(iFound).iSaEve = 0
                    tmCmmlSum(iFound).iSuEve = 0

                    tmCmmlSum(iFound).iTotal = 0
                End If

                If sRptName = "L35" Then
                    If lAirTime >= lDPStartTime(1) And lAirTime <= lDPEndtime(1) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFEarly = tmCmmlSum(iFound).iMFEarly + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaEarly = tmCmmlSum(iFound).iSaEarly + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuEarly = tmCmmlSum(iFound).iSuEarly + 1
                        End If
                    ElseIf lAirTime >= lDPStartTime(2) And lAirTime <= lDPEndtime(2) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFAM = tmCmmlSum(iFound).iMFAM + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaAM = tmCmmlSum(iFound).iSaAM + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuAM = tmCmmlSum(iFound).iSuAM + 1
                        End If
                    ElseIf lAirTime >= lDPStartTime(3) And lAirTime <= lDPEndtime(3) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFMid = tmCmmlSum(iFound).iMFMid + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaMid = tmCmmlSum(iFound).iSaMid + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuMid = tmCmmlSum(iFound).iSuMid + 1
                        End If
                    ElseIf lAirTime >= lDPStartTime(4) And lAirTime <= lDPEndtime(4) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFPM = tmCmmlSum(iFound).iMFPM + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaPM = tmCmmlSum(iFound).iSaPM + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuPM = tmCmmlSum(iFound).iSuPM + 1
                        End If
                    Else
                        If lAirTime >= lDPStartTime(0) And lAirTime <= lDPEndtime(0) Then   'is log event time within first DP
                            If iDay <= 4 Then  'M-F
                                tmCmmlSum(iFound).iMFEarliest = tmCmmlSum(iFound).iMFEarliest + 1
                            ElseIf iDay = 5 Then   'Sa
                                tmCmmlSum(iFound).iSaEarliest = tmCmmlSum(iFound).iSaEarliest + 1
                            Else    'Sun
                                tmCmmlSum(iFound).iSuEarliest = tmCmmlSum(iFound).iSuEarliest + 1
                            End If
                        Else
                            If iDay <= 4 Then  'M-F
                                tmCmmlSum(iFound).iMFEve = tmCmmlSum(iFound).iMFEve + 1
                            ElseIf iDay = 5 Then   'Sa
                                tmCmmlSum(iFound).iSaEve = tmCmmlSum(iFound).iSaEve + 1
                            Else    'Sun
                                tmCmmlSum(iFound).iSuEve = tmCmmlSum(iFound).iSuEve + 1
                            End If
                        
                        End If
                    End If
                Else
                    If lAirTime < lDPStartTime(1) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFEarly = tmCmmlSum(iFound).iMFEarly + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaEarly = tmCmmlSum(iFound).iSaEarly + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuEarly = tmCmmlSum(iFound).iSuEarly + 1
                        End If
                    ElseIf lAirTime < lDPStartTime(2) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFAM = tmCmmlSum(iFound).iMFAM + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaAM = tmCmmlSum(iFound).iSaAM + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuAM = tmCmmlSum(iFound).iSuAM + 1
                        End If
                    ElseIf lAirTime < lDPStartTime(3) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFMid = tmCmmlSum(iFound).iMFMid + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaMid = tmCmmlSum(iFound).iSaMid + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuMid = tmCmmlSum(iFound).iSuMid + 1
                        End If
                    ElseIf lAirTime < lDPStartTime(4) Then
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFPM = tmCmmlSum(iFound).iMFPM + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaPM = tmCmmlSum(iFound).iSaPM + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuPM = tmCmmlSum(iFound).iSuPM + 1
                        End If
                    Else
                        If iDay <= 4 Then  'M-F
                            tmCmmlSum(iFound).iMFEve = tmCmmlSum(iFound).iMFEve + 1
                        ElseIf iDay = 5 Then   'Sa
                            tmCmmlSum(iFound).iSaEve = tmCmmlSum(iFound).iSaEve + 1
                        Else    'Sun
                            tmCmmlSum(iFound).iSuEve = tmCmmlSum(iFound).iSuEve + 1
                        End If
                    End If
                End If
                tmCmmlSum(iFound).iTotal = tmCmmlSum(iFound).iTotal + 1
                'tmCmmlSum(iFound).iDay(iDay) = 1
                tmCmmlSum(iFound).iDay(iDay) = tmCmmlSum(iFound).iDay(iDay) + 1
            End If
    '    End If
    'End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If Err.Number <> 0 Then             'SQLSetConnectAttr vs. SQLSetOpenConnection
        gMsg = "A SQL error has occured in CP-mCmmlSum: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in CP-mCmmlSum: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    On Error GoTo 0
    Exit Sub
End Sub



Public Sub mCPMain()
    Dim sDate As String
    Dim sCP As String
    Dim iLoop As Integer
    Dim iUpper As Integer
    Dim sLogNum As String
    Dim iAdd As Integer
    Dim rstATT As ADODB.Recordset
    Dim llRow As Long
    Dim llCol As Long
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass

    SQLQuery = "SELECT DISTINCT vefName, vefCode, vefState, vefType, vpfLNoDaysCycle, vpfLastLog, vpfLastCP, vpfRnfPlayCode, rnfName " 'vpf.vpfLLeadTime
    'SQLQuery = SQLQuery + " FROM VEF_Vehicles vef, VPF_Vehicle_Options vpf, RNF_Report_Name rnf"
    SQLQuery = SQLQuery + " FROM VEF_Vehicles, "
    SQLQuery = SQLQuery & "VPF_Vehicle_Options, "
    SQLQuery = SQLQuery & "RNF_Report_Name"
    'Changed when using the Conventional Vehicles instead of the Log Vehicle
    '11/20/03
    'SQLQuery = SQLQuery + " WHERE (((vefvefCode = 0 AND vefType = 'C') OR vefType = 'L' OR vefType = 'A' OR vefType = 'S')"
    SQLQuery = SQLQuery + " WHERE (((vefType = 'C') OR (vefType = 'A') OR (vefType = 'S') OR (vefType = 'I'))"
    SQLQuery = SQLQuery + " AND vpfvefKCode = vefCode"
    If igCPOrLog = 0 Then
        SQLQuery = SQLQuery + " AND rnfCode = vpfRnfCertCode" & ")"
        SQLQuery = SQLQuery + " ORDER BY vefName, vpfLastCP"
    Else
        SQLQuery = SQLQuery + " AND rnfCode = vpfRnfLogCode" & ")"
        SQLQuery = SQLQuery + " ORDER BY vefName, vpfLastLog"
    End If
    
    ReDim tmCPInfo(0 To 0) As CPINFO

    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        If rst!vefType = "S" Then
            SQLQuery = "Select MAX(attVefCode) from att where attVefCode =" & Str$(rst!vefCode)
            Set rstATT = gSQLSelectCall(SQLQuery)
            If rstATT(0).Value = rst!vefCode Then
                iAdd = True
            Else
                iAdd = False
            End If
        Else
            iAdd = True
        End If
        If iAdd Then
            If igCPOrLog = 0 Then
                sDate = Format$(rst!vpfLastCP, "m/d/yyyy")
                If Len(sDate) > 0 Then
                    sDate = gObtainNextMonday(sDate)
                End If
            Else
                sDate = Format$(rst!vpfLastLog, "m/d/yyyy")
                If Len(sDate) > 0 Then
                    sDate = gObtainNextMonday(sDate)
                End If
            End If
            iUpper = UBound(tmCPInfo)
            tmCPInfo(iUpper).iStatus = 0
            tmCPInfo(iUpper).sVefState = rst!vefState
            tmCPInfo(iUpper).iVefCode = rst!vefCode
            If sgShowByVehType = "Y" Then
                tmCPInfo(iUpper).sVefName = Trim$(rst!vefType) & ":" & rst!vefName
            Else
                tmCPInfo(iUpper).sVefName = rst!vefName
            End If
            tmCPInfo(iUpper).sDate = sDate
            tmCPInfo(iUpper).iCycle = rst!vpfLNoDaysCycle
            tmCPInfo(iUpper).sRnfName = rst!rnfName
            tmCPInfo(iUpper).iRnfPlayCode = rst!vpfRnfPlayCode
            tmCPInfo(iUpper).sRnfOther = ""
            ReDim Preserve tmCPInfo(0 To iUpper + 1) As CPINFO
        End If
        rst.MoveNext
    Wend
    If igCPOrLog = 1 Then
        For iLoop = 0 To UBound(tmCPInfo) - 1 Step 1
            If tmCPInfo(iLoop).iRnfPlayCode > 0 Then
                SQLQuery = "SELECT rnfName "
                SQLQuery = SQLQuery + " FROM RNF_Report_Name"
                SQLQuery = SQLQuery + " WHERE (rnfCode = " & tmCPInfo(iLoop).iRnfPlayCode & ")"
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    tmCPInfo(iLoop).sRnfOther = rst!rnfName
                End If
            End If
        Next iLoop
    End If
    'For iLoop = 0 To UBound(tmCPInfo) - 1 Step 1
    '    SQLQuery = "SELECT cpttPrintStatus " 'vpf.vpfLLeadTime
    '    SQLQuery = SQLQuery + " FROM cptt"
    '    SQLQuery = SQLQuery + " WHERE (cptt.cpttVefCode = " & tmCPInfo(iLoop).iVefCode
    '    SQLQuery = SQLQuery + " AND cptt.cpTTStartDate = '" & tmCPInfo(iLoop).sDate & "')"
    '    Set rst = gSQLSelectCall(SQLQuery)
    '    If rst.EOF Then
    '        tmCPInfo(iLoop).iStatus = 1
    '    End If
    'Next iLoop
    lbcSingle.Clear
    lbcMulti.Clear
    chkAll.Value = 0
    grdLog.Redraw = False
    mClearGrid
    llRow = grdLog.FixedRows
    
    For iLoop = 0 To UBound(tmCPInfo) - 1 Step 1
        If tmCPInfo(iLoop).iStatus = 0 Then
            iAdd = False
            If igCPOrLog = 0 Then
                If rbcVeh(0).Value = True Then
                    If tmCPInfo(iLoop).sVefState = "A" Then
                        iAdd = True
                    End If
                Else
                    iAdd = True
                End If
            Else
                If rbcVeh(0).Value = True Then
                    If tmCPInfo(iLoop).sVefState = "A" Then
                        iAdd = True
                    End If
                Else
                    iAdd = True
                End If
            End If
            If iAdd Then
                If llRow + 1 > grdLog.Rows Then
                    grdLog.AddItem ""
                End If
                grdLog.Row = llRow
                For llCol = 1 To 3 Step 1
                    grdLog.Row = llRow
                    grdLog.Col = llCol
                    grdLog.CellBackColor = LIGHTYELLOW
                Next llCol
                grdLog.Col = 0
                grdLog.CellFontName = "Monotype Sorts"
                If igCPOrLog = 0 Then
                    grdLog.TextMatrix(llRow, 0) = ""
                    grdLog.TextMatrix(llRow, 1) = Trim$(tmCPInfo(iLoop).sVefName)
                    grdLog.TextMatrix(llRow, 2) = Trim$(tmCPInfo(iLoop).sDate)
                    grdLog.TextMatrix(llRow, 3) = tmCPInfo(iLoop).iCycle
                    grdLog.TextMatrix(llRow, 4) = Trim$(tmCPInfo(iLoop).sRnfName)
                    llRow = llRow + 1
                Else
                    grdLog.TextMatrix(llRow, 0) = ""
                    grdLog.TextMatrix(llRow, 1) = Trim$(tmCPInfo(iLoop).sVefName)
                    grdLog.TextMatrix(llRow, 2) = Trim$(tmCPInfo(iLoop).sDate)
                    grdLog.TextMatrix(llRow, 3) = tmCPInfo(iLoop).iCycle
                    grdLog.TextMatrix(llRow, 4) = Trim$(tmCPInfo(iLoop).sRnfName)
                    grdLog.TextMatrix(llRow, 5) = Trim$(tmCPInfo(iLoop).sRnfOther)
                    llRow = llRow + 1
                End If
            End If
        End If
    Next iLoop
    'Don't add extra row
'    If llRow >= grdLog.Rows Then
'        grdLog.AddItem ""
'    End If
    grdLog.Redraw = True
    mFill
    'mFillVehicle
    If igCPOrLog = 0 Then
        If sgUstWin(6) <> "I" Then
            cmdGenerate.Enabled = False
            cmdReprint.Enabled = False
            frcTab(1).Enabled = False
        End If
    Else
        If sgUstWin(4) <> "I" Then
            cmdGenerate.Enabled = False
            cmdReprint.Enabled = False
            frcTab(1).Enabled = False
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-CPMain"
End Sub

Private Sub txtDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    
    Select Case grdLog.Col
        Case 4
            slStr = Trim$(txtDropdown.Text)
            grdLog.Text = slStr
        Case 5
            slStr = Trim$(txtDropdown.Text)
            grdLog.Text = slStr
    End Select
End Sub

Private Sub txtDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Function mLogPrePass(iPrint As Integer, iLog As Integer) As Integer

'
'   iPrint(I)- True = From Log selection; False= From Reprint or Studio selection
'   iLog(I)- True = Generate Log, False = Generate Studio
'
    Dim iRow As Integer
    Dim iVefCode As Integer
    Dim iOutput As Integer
    Dim iSPass As Integer
    Dim iEPass As Integer
    Dim iPass As Integer
    Dim llRow As Long
    Dim iGenL32 As Integer
    Dim iIndex As Integer
    Dim iLoop As Integer
    Dim iSelected As Integer
    Dim iRnf As Integer
    Dim sLogName As String
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    smCntrNo = "0"
    If optRptDest(0).Value Then
        iOutput = 0
    ElseIf optRptDest(2).Value Then
        iOutput = 2
    Else
        iOutput = 1
    End If
    If iLog Then
        iSPass = 4
        iEPass = 5
    Else
        iSPass = 1
        iEPass = 1
    End If
    grdLog.Redraw = False
    llRow = grdLog.FixedRows
    For iRow = 0 To UBound(tmCPInfo) - 1 Step 1
        If (tmCPInfo(iRow).iStatus = 0) And ((rbcVeh(0).Value = False) Or ((rbcVeh(0).Value = True) And (tmCPInfo(iRow).sVefState = "A"))) Then
            iGenL32 = False
            If iLog Then
                For iPass = iSPass To iEPass Step 1
                    If StrComp(Trim$(grdLog.TextMatrix(llRow, iPass)), "L32", 1) = 0 Then
                        iGenL32 = True
                        Exit For
                    End If
                Next iPass
            End If
            ReDim tmCmmlSum(0 To 0) As CMMLSUM
            iVefCode = tmCPInfo(iRow).iVefCode
            iIndex = -1
            For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If tgVehicleInfo(iLoop).iCode = iVefCode Then
                    iIndex = iLoop
                    Exit For
                End If
            Next iLoop
            If iIndex >= 0 Then
                If iPrint Then 'Print Logs
                    If Trim$(grdLog.TextMatrix(llRow, 0)) <> "" Then
                        iSelected = True
                    Else
                        iSelected = False
                    End If
                Else 'Reprint or Studio Logs
                    For iLoop = 0 To lbcMulti.ListCount - 1 Step 1
                        If iVefCode = lbcMulti.ItemData(iLoop) Then
                            If lbcMulti.Selected(iLoop) Then
                                iSelected = True
                                imVefCode = iVefCode
                            Else
                                iSelected = False
                            End If
                            Exit For
                        End If
                    Next iLoop
                End If
                If (iSelected) Then
                    For iLoop = 0 To UBound(tgRnfInfo) - 1 Step 1
                        If (StrComp(Trim$(tgRnfInfo(iLoop).sName), Trim$(grdLog.TextMatrix(llRow, 4)), 1) = 0) Or (StrComp(Trim$(tgRnfInfo(iLoop).sName), Trim$(grdLog.TextMatrix(llRow, 5)), 1) = 0) Then
                            For iPass = iSPass To iEPass Step 1
                                For iRnf = 0 To UBound(tgRnfInfo) - 1 Step 1
                                    If (StrComp(Trim$(tgRnfInfo(iRnf).sName), Trim$(grdLog.TextMatrix(llRow, iPass)), 1) = 0) Or (Not iLog) Then
                                        sLogName = Trim$(grdLog.TextMatrix(llRow, iPass))
                                        If StrComp(sLogName, "L31", 1) <> 0 And StrComp(sLogName, "L32", 1) <> 0 Then
                                            gMsgBox "Only Logs L31 and L32 are allowed at this time.  See Vehicle: " & grdLog.TextMatrix(llRow, 1), vbCritical
                                            Screen.MousePointer = vbDefault
                                            Exit Function
                                        End If
                                    End If
                                Next iRnf
                            Next iPass
                            Exit For
                        End If
                    Next iLoop
                End If
            End If
            llRow = llRow + 1
        End If
    Next iRow
    grdLog.Redraw = True
    Screen.MousePointer = vbDefault
    mLogPrePass = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCP-mLogtPrePass"
    grdLog.Redraw = True
    mLogPrePass = False
End Function

Private Function mPrintPrePass(iPrint As Integer) As Integer
    'iPrint True = Gen CP, False = Reprint CP
    
    Dim iRow As Integer
    Dim iVefCode As Integer
    Dim iLoop As Integer
    Dim llRow As Long
    Dim iSelected As Integer
    Dim sCPName As String
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    grdLog.Redraw = False
    llRow = grdLog.FixedRows
    For iRow = 0 To UBound(tmCPInfo) - 1 Step 1
        If (tmCPInfo(iRow).iStatus = 0) And ((rbcVeh(0).Value = False) Or ((rbcVeh(0).Value = True) And (tmCPInfo(iRow).sVefState = "A"))) Then
            iVefCode = tmCPInfo(iRow).iVefCode
            If iPrint Then
                If Trim$(grdLog.TextMatrix(llRow, 0)) <> "" Then
                    iSelected = True
                Else
                    iSelected = False
                End If
            Else 'Reprint CPs
                If optSort(0).Value Then
                    'Vehicle, then Stations
                    If lbcSingle.ListIndex >= 0 Then
                        If iVefCode = lbcSingle.ItemData(lbcSingle.ListIndex) Then
                            iSelected = True
                            imVefCode = iVefCode
                        Else
                            iSelected = False
                        End If
                    Else
                        iSelected = False
                    End If
                Else
                    'Stations, Then vehicle
                    For iLoop = 0 To lbcMulti.ListCount - 1 Step 1
                        If iVefCode = lbcMulti.ItemData(iLoop) Then
                            If lbcMulti.Selected(iLoop) Then
                                iSelected = True
                                imVefCode = iVefCode
                            Else
                                iSelected = False
                            End If
                            Exit For
                        End If
                    Next iLoop
                End If
            End If
            If (iSelected) Then
                sCPName = Trim$(grdLog.TextMatrix(llRow, 4))
                If StrComp("C17", sCPName, 1) <> 0 Then
                    mPrintPrePass = False
                    gMsgBox "Only CP C17 is allowed at this time.  Please see " & grdLog.TextMatrix(llRow, 1), vbCritical
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
                
            End If
            llRow = llRow + 1
        End If
    Next iRow
    grdLog.Redraw = True
    mPrintPrePass = True
    Screen.MousePointer = vbDefault
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in CP-mPrintPrePass: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    grdLog.Redraw = True
    mPrintPrePass = False
    Screen.MousePointer = vbDefault
End Function

