VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmContactGrid 
   Caption         =   "Results"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9480
   Icon            =   "AffContactGrid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9480
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5025
      Width           =   45
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9285
      Top             =   5460
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   90
      Picture         =   "AffContactGrid.frx":08CA
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.TextBox txtDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4935
      TabIndex        =   9
      Top             =   2190
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcActionSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   5085
      Width           =   60
   End
   Begin VB.PictureBox pbcActionTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   105
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   5505
      Width           =   60
   End
   Begin VB.PictureBox pbcActionFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   240
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5670
      Width           =   60
   End
   Begin VB.PictureBox pbcActionArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   75
      Picture         =   "AffContactGrid.frx":0BD4
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4455
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcPostFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   240
      Width           =   60
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Comments"
      Height          =   375
      Left            =   3975
      TabIndex        =   11
      Top             =   5460
      Width           =   1575
   End
   Begin VB.PictureBox pbcCrystalReport1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1605
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   12
      Top             =   5385
      Width           =   1200
   End
   Begin VB.CommandButton cmdMail 
      Caption         =   "Generate Mail List"
      Height          =   375
      Left            =   5790
      TabIndex        =   1
      Top             =   5460
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7545
      TabIndex        =   0
      Top             =   5460
      Width           =   1575
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   735
      Top             =   5400
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5940
      FormDesignWidth =   9480
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPost 
      Height          =   2730
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4815
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAction 
      Height          =   1425
      Left            =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3600
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2514
      _Version        =   393216
      Cols            =   3
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      Caption         =   "Red = Outstanding"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   15
      Top             =   3195
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Magenta = Partially Posted"
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Index           =   3
      Left            =   1875
      TabIndex        =   14
      Top             =   3195
      Width           =   1965
   End
   Begin VB.Image imcPrt 
      Height          =   480
      Left            =   8445
      Picture         =   "AffContactGrid.frx":0EDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmContactGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private hmMail As Integer
Private tmContactInfo() As CONTACTINFO
Private smToFile As String
Private tmCDate() As CONTACTDATE
Private imCIndex As Integer
Private imCMax As Integer
Private imShttCode As Integer
Private imVefCode As Integer
Private imCommentChgd As Integer
'Dim iLoadedRow As Integer
Private imSIntegralSet As Integer
Private imCIntegralSet As Integer
Private imFirstTime As Integer
Private imHeaderClick As Integer
Private lmPostRow As Long

'Grid Controls
Private imShowGridBox As Integer
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Const STATIONINDEX = 0
Const VEHICLEINDEX = 1
Const CONTACTINDEX = 2
Const TELEPHONEINDEX = 3
Const DATE1INDEX = 4
Const DATE2INDEX = 5
Const DATE3INDEX = 6
Const EXTRAINDEX = 7

Const ACTIONDATEINDEX = 0
Const COMMENTINDEX = 1
Const CCTCODEINDEX = 2


Private Sub mActionSetShow()
    If (lmEnableRow >= grdAction.FixedRows) And (lmEnableRow < grdAction.Rows) Then
        'Set any field that that should only be set after user leaves the cell
    End If
    imShowGridBox = False
    pbcActionArrow.Visible = False
    txtDropdown.Visible = False
End Sub


Private Sub mActionEnableBox()
    If (grdAction.Row >= grdAction.FixedRows) And (grdAction.Row < grdAction.Rows) And (grdAction.Col >= ACTIONDATEINDEX) And (grdAction.Col < grdAction.Cols - 1) Then
        lmEnableRow = grdAction.Row
        imShowGridBox = True
        pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width - 15, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
        pbcActionArrow.Visible = True
        Select Case grdAction.Col
            Case ACTIONDATEINDEX  'Action Date
                txtDropdown.Move grdAction.Left + grdAction.ColPos(grdAction.Col) + 30, grdAction.Top + grdAction.RowPos(grdAction.Row) + 15, grdAction.ColWidth(grdAction.Col) - 30, grdAction.RowHeight(grdAction.Row) - 15
                If grdAction.Text <> "Missing" Then
                    txtDropdown.Text = grdAction.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdAction.RowHeight(grdAction.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdAction.RowHeight(grdAction.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            Case COMMENTINDEX  'Comment
                txtDropdown.Move grdAction.Left + grdAction.ColPos(grdAction.Col) + 30, grdAction.Top + grdAction.RowPos(grdAction.Row) + 15, grdAction.ColWidth(grdAction.Col) - 30, grdAction.RowHeight(grdAction.Row) - 15
                If grdAction.Text <> "Missing" Then
                    txtDropdown.Text = grdAction.Text
                Else
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdAction.RowHeight(grdAction.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdAction.RowHeight(grdAction.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
        End Select
    End If
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    gGrid_Clear grdPost, True
    grdPost.Row = 0
    grdPost.Col = DATE1INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    'grdPost.TextMatrix(0, 2) = Chr$(171)
    grdPost.Row = 0
    grdPost.Col = DATE2INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    'grdPost.TextMatrix(0, 8) = "Dates*"
    grdPost.Row = 0
    grdPost.Col = DATE3INDEX
    grdPost.CellAlignment = flexAlignCenterTop
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        grdPost.Row = llRow
        For llCol = STATIONINDEX To grdPost.Cols - 2 Step 1
            grdPost.Col = llCol
            grdPost.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llRow

End Sub

Private Sub GridPaint(iClear As Integer)
    Dim iTotRec As Integer
    Dim iSIndex As Integer
    Dim iEIndex As Integer
    Dim iRow As Integer
    Dim iLoop As Integer
    Dim iSLoop As Integer
    Dim iELoop As Integer
    Dim iStep As Integer
    Dim iCol As Integer
    Dim iTRow As Integer
    Dim sStation, sVehicle, sACName, sACPhone As String
    Dim sDate(0 To 2) As String
    Dim ilPostingStatus(0 To 2) As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llTRow As Long
    
    grdPost.Redraw = False
    If iClear Then
        mClearGrid
        iSLoop = 0
        iELoop = UBound(tmContactInfo) - 1
        iStep = 1
    Else
        'iTRow = iLoadedRow - grdStation.VisibleRows
'        iELoop = 0
'        iSLoop = UBound(tmContactInfo) - 1
'        iStep = -1
        iSLoop = 0
        iELoop = UBound(tmContactInfo) - 1
        iStep = 1
    End If
    llRow = grdPost.FixedRows
    llTRow = grdPost.TopRow
    For iRow = iSLoop To iELoop Step iStep
        sStation = tmContactInfo(iRow).sStation
        sVehicle = tmContactInfo(iRow).sVehicle
        sACName = tmContactInfo(iRow).sACName
        sACPhone = tmContactInfo(iRow).sACPhone
        iSIndex = tmContactInfo(iRow).iCDateIndex
        If iRow < UBound(tmContactInfo) - 1 Then
            iEIndex = tmContactInfo(iRow + 1).iCDateIndex - 1
        Else
            iEIndex = UBound(tmCDate) - 1
        End If
        iTotRec = iEIndex - iSIndex + 1
        
        iSIndex = 3 * (imCIndex - 1) + iSIndex
        If iSIndex <= iEIndex Then
            If iSIndex + 2 < iEIndex Then
                iEIndex = iSIndex + 2
            End If
            iCol = 0
            sDate(0) = ""
            sDate(1) = ""
            sDate(2) = ""
            ilPostingStatus(0) = 0
            ilPostingStatus(1) = 0
            ilPostingStatus(2) = 0
            For iLoop = iSIndex To iEIndex Step 1
                sDate(iCol) = tmCDate(iLoop).sDate
                ilPostingStatus(iCol) = tmCDate(iLoop).iPostingStatus
                iCol = iCol + 1
            Next iLoop
        Else
            sDate(0) = ""
            sDate(1) = ""
            sDate(2) = ""
            ilPostingStatus(0) = 0
            ilPostingStatus(1) = 0
            ilPostingStatus(2) = 0
        End If
        If llRow + 1 > grdPost.Rows Then
            grdPost.AddItem ""
        End If
        grdPost.Row = llRow
        For llCol = STATIONINDEX To grdPost.Cols - 2 Step 1
            grdPost.Col = llCol
            grdPost.CellBackColor = LIGHTYELLOW
        Next llCol
        If iClear Then
            grdPost.TextMatrix(llRow, STATIONINDEX) = sStation
            grdPost.TextMatrix(llRow, VEHICLEINDEX) = sVehicle
            grdPost.TextMatrix(llRow, CONTACTINDEX) = sACName
            grdPost.TextMatrix(llRow, TELEPHONEINDEX) = sACPhone
            grdPost.Col = DATE1INDEX
            If ilPostingStatus(0) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE1INDEX) = sDate(0)
            grdPost.Col = DATE2INDEX
            If ilPostingStatus(1) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE2INDEX) = sDate(1)
            grdPost.Col = DATE3INDEX
            If ilPostingStatus(2) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE3INDEX) = sDate(2)
            grdPost.TextMatrix(llRow, 7) = iRow
        Else
            grdPost.Col = DATE1INDEX
            If ilPostingStatus(0) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE1INDEX) = sDate(0)
            grdPost.Col = DATE2INDEX
            If ilPostingStatus(1) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE2INDEX) = sDate(1)
            grdPost.Col = DATE3INDEX
            If ilPostingStatus(2) <= 0 Then
                grdPost.CellForeColor = vbRed
            Else
                grdPost.CellForeColor = vbMagenta
            End If
            grdPost.TextMatrix(llRow, DATE3INDEX) = sDate(2)
        End If
        llRow = llRow + 1
    Next iRow
    'Don't add extra row
'    If llRow >= grdPost.Rows Then
'        grdPost.AddItem ""
'    End If
    If Not iClear Then
        grdPost.TopRow = llTRow
    End If
    grdPost.Redraw = True
    Exit Sub
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function OpenMsgFile() As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim sLetter As String

    'On Error GoTo OpenMsgFileErr:
    'slToFile = "Mail.CSV"
    sLetter = "A"
    Do
        ilRet = 0
        smToFile = sgExportDirectory & "C" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & sLetter & ".csv"
        'slDateTime = FileDateTime(smToFile)
        ilRet = gFileExist(smToFile)
        If ilRet = 0 Then
            sLetter = Chr$(Asc(sLetter) + 1)
        End If
    Loop While ilRet = 0
    On Error GoTo 0
    'ilRet = 0
    'On Error GoTo OpenMsgFileErr:
    'hmMail = FreeFile
    'Open smToFile For Output As hmMail
    ilRet = gFileOpen(smToFile, "Output", hmMail)
    If ilRet <> 0 Then
        Close hmMail
        hmMail = -1
        gMsgBox "Open File " & smToFile & " error#" & Str$(Err.Number), vbOKOnly
        OpenMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    OpenMsgFile = True
    Exit Function
'OpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub cmdCancel_Click()
    Unload frmContactGrid
End Sub

Private Sub cmdCancel_GotFocus()
    mActionSetShow
End Sub

Private Sub cmdMail_Click()
    Dim iRet As Integer
    Dim sMail As String
    Dim iRow As Integer
    Dim iGetStation As Integer
    Dim sStationInfo As String
    Dim sStation As String
    Dim llRow As Long
    On Error GoTo ErrHand
    
    If grdPost.Rows - 1 <= grdPost.FixedRows Then
        Exit Sub
    End If
    iRet = OpenMsgFile()
    If iRet = False Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    iGetStation = True
    iRow = 0
    llRow = grdPost.FixedRows
    Do While llRow <= grdPost.Rows - 1
        If iGetStation = True Then
            iRow = Val(grdPost.TextMatrix(llRow, 7))
            sStation = Trim$(grdPost.TextMatrix(llRow, STATIONINDEX))
            SQLQuery = "SELECT shttAddress1, shttAddress2, shttCity, shttState, shttZip , shttFax FROM shtt WHERE (shttCode = " & tmContactInfo(iRow).iShttCode & ")"
            Set rst = gSQLSelectCall(SQLQuery)
            'sMail = "1-" & Trim$(rst(6).Value) & "," & Trim$(grdStations.Columns(3).Text) & "," & Trim$(grdStations.Columns(1).Text)
            'sMail = sMail & "," & Trim$(rst(0).Value) & "," & Trim$(rst(1).Value)
            'sMail = sMail & "," & Trim$(rst(2).Value) & "," & Trim$(rst(3).Value) & "," & Trim$(rst(4).Value)
            sMail = """" & sStation & """" & "," & """" & "1-" & Trim$(rst(5).Value) & """"
            'sStationInfo = sMail
            Print #hmMail, sMail
        End If
        'sMail = sStationInfo & "," & Trim$(grdStations.Columns(2).Text) & "," & Trim$(grdStations.Columns(0).Text)
        'Print #hmMail, sMail
        
        llRow = llRow + 1
        'If iRow <= grdStations.Rows - 1 Then
        '    If sStation <> Trim$(grdStations.Columns(1).Text) Then
        '        iGetStation = True
        '    End If
        'End If
    Loop
    Close hmMail
    Screen.MousePointer = vbDefault
    gMsgBox "Mail File Created Successfully, Its Name is " & smToFile, vbOKOnly
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact Grid-cmdMail"
End Sub


Private Sub cmdMail_GotFocus()
    mActionSetShow
End Sub

Private Sub cmdSave_Click()
    Dim ilRet As Integer
    
    If (imCommentChgd = False) Then
        Exit Sub
    End If
    ilRet = mSave()

End Sub

Private Sub cmdSave_GotFocus()
    mActionSetShow
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        
        grdPost.ColWidth(EXTRAINDEX) = 0
        grdPost.ColWidth(STATIONINDEX) = grdPost.Width * 0.12
        grdPost.ColWidth(VEHICLEINDEX) = grdPost.Width * 0.25
        grdPost.ColWidth(TELEPHONEINDEX) = grdPost.Width * 0.15
        grdPost.ColWidth(DATE1INDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(DATE2INDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(DATE3INDEX) = grdPost.Width * 0.08
        grdPost.ColWidth(CONTACTINDEX) = grdPost.Width - grdPost.ColWidth(STATIONINDEX) - grdPost.ColWidth(VEHICLEINDEX) - grdPost.ColWidth(TELEPHONEINDEX) - grdPost.ColWidth(DATE1INDEX) - grdPost.ColWidth(DATE2INDEX) - grdPost.ColWidth(DATE3INDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdPost
        grdPost.TextMatrix(0, STATIONINDEX) = "Station"
        grdPost.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
        grdPost.TextMatrix(0, CONTACTINDEX) = "Contact"
        grdPost.TextMatrix(0, TELEPHONEINDEX) = "Telephone"
        grdPost.TextMatrix(0, DATE1INDEX) = Chr$(171)
        grdPost.TextMatrix(0, DATE2INDEX) = "Date"
        grdPost.TextMatrix(0, DATE3INDEX) = Chr$(187)
        gGrid_IntegralHeight grdPost
        gGrid_Clear grdPost, True
    
        'Hide column 2
        grdAction.ColWidth(CCTCODEINDEX) = 0
        grdAction.ColWidth(ACTIONDATEINDEX) = grdAction.Width * 0.16
        grdAction.ColWidth(COMMENTINDEX) = grdAction.Width - grdAction.ColWidth(ACTIONDATEINDEX) - GRIDSCROLLWIDTH '(5 * grdStation.Columns(6).Width) / 6
        gGrid_AlignAllColsLeft grdAction
        grdAction.TextMatrix(0, ACTIONDATEINDEX) = "Action Date"
        grdAction.TextMatrix(0, COMMENTINDEX) = "Comment"
        gGrid_IntegralHeight grdAction
        gGrid_Clear grdAction, True
        GridPaint True
        imFirstTime = False
    End If

End Sub

Private Sub Form_Click()
    mActionSetShow
    If Not imCommentChgd Then
        lmPostRow = -1
        pbcArrow.Visible = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.03   '1.3
    Me.Height = Screen.Height / 1.25 '2.2
    Me.Top = (Screen.Height - Me.Height) / 1.7
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim sPrevCall As String
    Dim sPrevDate As String
    Dim sPrevVeh As String
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iUpper As Integer
    Dim iAddDate As Integer
    
    On Error GoTo ErrHand

    'Me.Width = Screen.Width / 1.1   '1.3
    'Me.Height = Screen.Height / 1.5 '2.2
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    Screen.MousePointer = vbHourglass
    frmContactGrid.Caption = "Results - " & sgClientName
    sPrevCall = ""
    sPrevDate = ""
    sPrevVeh = ""
    imCIndex = 1
    imCommentChgd = False
    imSIntegralSet = False
    imHeaderClick = False
    imFirstTime = True
    lmPostRow = -1
    imcPrt.Picture = frmDirectory!imcPrinter.Picture
    imCIntegralSet = False
    'SQLQuery = "SELECT cptt.cpttStartDate, shtt.shttCallLetters, vef.vefName, shttACName, shttACPhone, shttCode"
    'SQLQuery = SQLQuery + " FROM cptt, lst, shtt, vef"
    'SQLQuery = SQLQuery + " WHERE ((cptt.cpttStatus = 0) AND (vef.vefCode = lst.lstLogVefCode)"
    'SQLQuery = SQLQuery + " AND (shtt.shttCode = cptt.cpttShfCode)"
    'SQLQuery = SQLQuery + " AND " & sContracts & ")"
    'SQLQuery = SQLQuery + " ORDER BY shtt.shttCallLetters, vef.vefName, cptt.cpttStartDate"
    ReDim tmContactInfo(0 To 0) As CONTACTINFO
    iRow = 0
    iCol = 0
    iUpper = 0
    imCMax = 1
    ReDim tmCDate(0 To 0) As CONTACTDATE
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        ''If (sPrevCall <> Trim$(rst(1).Value)) Or (sPrevDate <> rst(0).Value) Or (sPrevVeh <> Trim$(rst(2).Value)) Then
        'If (sPrevCall <> Trim$(rst!shttCallLetters)) Or (sPrevDate <> rst!cpttStartDate) Or (sPrevVeh <> Trim$(rst!vefName)) Then
        If (sPrevCall <> Trim$(rst!shttCallLetters)) Or (sPrevVeh <> Trim$(rst!vefName)) Then
            'grdStations.AddItem "" & rst(0).Value & ", " & rst(1).Value & ", " & rst(2).Value & ", " & rst(3).Value & ", " & rst(4).Value & ""
            If iCol \ 3 + 1 > imCMax Then
                imCMax = iCol \ 3 + 1
            End If
            tmContactInfo(iRow).sStation = rst!shttCallLetters
            If sgShowByVehType = "Y" Then
                tmContactInfo(iRow).sVehicle = Trim$(rst!vefType) & ":" & rst!vefName
            Else
                tmContactInfo(iRow).sVehicle = rst!vefName
            End If
            If Trim$(rst!attACName) <> "" Then
                tmContactInfo(iRow).sACName = rst!attACName
                tmContactInfo(iRow).sACPhone = rst!attACPhone
            Else
                tmContactInfo(iRow).sACName = rst!shttACName
                tmContactInfo(iRow).sACPhone = rst!shttACPhone
            End If
            tmContactInfo(iRow).iShttCode = rst!shttCode
            tmContactInfo(iRow).iVefCode = rst!vefCode
            tmContactInfo(iRow).iCDateIndex = iUpper
            iCol = 0
            iRow = iRow + 1
            ReDim Preserve tmContactInfo(0 To iRow) As CONTACTINFO
            sPrevCall = Trim$(rst!shttCallLetters)
            sPrevDate = Format$(Trim$(rst!CpttStartDate), sgShowDateForm)
            sPrevVeh = Trim$(rst!vefName)
            iAddDate = True
        Else
            'When filter is by Advertiser/Contract, multi records for same date is caused by multi-spots within same week
            If DateValue(gAdjYear(sPrevDate)) <> DateValue(gAdjYear(Trim$(rst!CpttStartDate))) Then
                iAddDate = True
                sPrevDate = Format$(Trim$(rst!CpttStartDate), sgShowDateForm)
            Else
                iAddDate = False
            End If
        End If
        If iAddDate Then
            tmCDate(iUpper).iCol = iCol
            tmCDate(iUpper).sDate = Format$(rst!CpttStartDate, sgShowDateForm)
            tmCDate(iUpper).iPostingStatus = rst!cpttPostingStatus
            iUpper = iUpper + 1
            iCol = iCol + 1
            ReDim Preserve tmCDate(0 To iUpper) As CONTACTDATE
        End If
        rst.MoveNext
    Wend
    If iCol \ 3 + 1 > imCMax Then
        imCMax = iCol \ 3 + 1
    End If
    'GridPaint True
    If sgUstWin(8) <> "I" Then
        cmdSave.Enabled = False
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Contact Grid-Form Load"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Erase tmContactInfo
    Erase tmCDate
    Set frmContactGrid = Nothing
End Sub



Private Sub grdAction_Click()
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAction.Col >= grdAction.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAction_EnterCell()
    mActionSetShow
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdAction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdAction.TopRow
    grdAction.Redraw = False
End Sub

Private Sub grdAction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    If sgUstWin(8) <> "I" Then
        grdAction.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (lmPostRow < grdPost.FixedRows) Or (lmPostRow >= grdPost.Rows) Then
        grdAction.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdAction, X, Y)
    If Not ilFound Then
        grdAction.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAction.Col >= grdAction.Cols - 1 Then
        grdAction.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdAction.TopRow
    
    llRow = grdAction.Row
    If grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = "" Then
        grdAction.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = ""
        grdAction.Row = llRow + 1
        grdAction.Col = ACTIONDATEINDEX
        grdAction.Redraw = True
    End If
    grdAction.Redraw = True
    mActionEnableBox
End Sub

Private Sub grdAction_Scroll()
    If grdAction.Redraw = False Then
        grdAction.Redraw = True
        grdAction.TopRow = lmTopRow
        grdAction.Refresh
        grdAction.Redraw = False
    End If
    If (imShowGridBox) And (grdAction.Row >= grdAction.FixedRows) And (grdAction.Col >= ACTIONDATEINDEX) And (grdAction.Col < grdAction.Cols - 1) Then
        If grdAction.RowIsVisible(grdAction.Row) Then
            txtDropdown.Move grdAction.Left + grdAction.ColPos(grdAction.Col) + 30, grdAction.Top + grdAction.RowPos(grdAction.Row) + 30, grdAction.ColWidth(grdAction.Col) - 30, grdAction.RowHeight(grdAction.Row) - 30
            pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
            pbcActionArrow.Visible = True
            txtDropdown.Visible = True
            txtDropdown.SetFocus
        Else
            pbcActionFocus.SetFocus
            txtDropdown.Visible = False
            pbcActionArrow.Visible = False
        End If
    Else
        pbcActionFocus.SetFocus
        pbcActionArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub grdPost_Click()
    Dim ilRet As Integer
    tmcDelay.Enabled = False
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If imCommentChgd Then
        If gMsgBox("Save comment changes?", vbYesNo) = vbYes Then
            ilRet = mSave()
            If ilRet = False Then
                grdPost.Row = lmPostRow
                Exit Sub
            End If
        End If
    End If
    imCommentChgd = False
    DoEvents
    If (grdPost.Row - 1 < STATIONINDEX) Or (grdPost.Row - 1 >= UBound(tmContactInfo)) Then
        lmPostRow = -1
        pbcArrow.Visible = False
        Exit Sub
    End If
    If imHeaderClick Then
        imHeaderClick = False
        Exit Sub
    End If
    tmcDelay.Enabled = True
    Exit Sub
End Sub

Private Sub grdPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Determine if in header
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If Y < grdPost.RowHeight(STATIONINDEX) Then
        imHeaderClick = True
        If (X >= grdPost.ColPos(DATE1INDEX)) And (X <= grdPost.ColPos(DATE1INDEX) + grdPost.ColWidth(DATE1INDEX)) Then
            imCIndex = imCIndex - 1
            If imCIndex < 1 Then
                imCIndex = 1
            Else
                GridPaint False
            End If
        End If
        If (X >= grdPost.ColPos(DATE3INDEX)) And (X <= grdPost.ColPos(DATE3INDEX) + grdPost.ColWidth(DATE3INDEX)) Then
            imCIndex = imCIndex + 1
            If imCIndex > imCMax Then
                imCIndex = imCMax
            Else
                GridPaint False
            End If
        End If
'        If (x >= grdPost.ColPos(8)) And (x <= grdPost.ColPos(8) + grdPost.ColWidth(8)) Then
'            mDateSort
'            mGridPaint False
'        End If
    End If
End Sub


Private Sub grdPost_Scroll()
    If (pbcArrow.Visible) And (grdPost.Row >= grdPost.FixedRows) And (grdPost.Col >= STATIONINDEX) And (grdPost.Col < grdPost.Cols - 1) Then
        If grdPost.RowIsVisible(grdPost.Row) Then
            pbcArrow.Move grdPost.Left - pbcActionArrow.Width - 15, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcActionArrow.Height) / 2
            pbcArrow.Visible = True
        Else
            pbcPostFocus.SetFocus
            pbcArrow.Visible = False
        End If
    Else
        pbcPostFocus.SetFocus
        pbcArrow.Visible = False
    End If

End Sub

Private Sub imcPrt_Click()
    Dim llRow As Long
    Dim sContactPhone As String
    
    mActionSetShow
    If grdPost.Rows - 1 <= grdPost.FixedRows Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Printer.Print ""
    Printer.Print Tab(65); Format$(Now)
    Printer.Print ""
    Printer.Print " " & sAdvtDates
    Printer.Print ""
    Printer.Print "  Call Letters"; Tab(15); "Vehicle"; Tab(37); "Contact"; Tab(69); "Dates"
    
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        sContactPhone = Trim$(grdPost.TextMatrix(llRow, CONTACTINDEX))
        If Len(sContactPhone) > 16 Then
            sContactPhone = Left$(sContactPhone, 30 - Len(Trim$(grdPost.TextMatrix(llRow, TELEPHONEINDEX)))) & " " & Trim$(grdPost.TextMatrix(llRow, TELEPHONEINDEX))
        Else
            sContactPhone = sContactPhone & " " & Trim$(grdPost.TextMatrix(llRow, TELEPHONEINDEX))
        End If
        Printer.Print "  " & Trim$(grdPost.TextMatrix(llRow, STATIONINDEX)); Tab(15); Trim$(grdPost.TextMatrix(llRow, VEHICLEINDEX)); Tab(37); sContactPhone; Tab(69); Trim$(grdPost.TextMatrix(llRow, DATE1INDEX)) & " " & Trim$(grdPost.TextMatrix(llRow, DATE2INDEX)) & " " & Trim$(grdPost.TextMatrix(llRow, DATE3INDEX))
    Next llRow
    Printer.EndDoc
    Screen.MousePointer = vbDefault
End Sub

Private Function mSave() As Integer
    Dim llRow As Long
    Dim sDate As String
    Dim sComment As String
    Dim ilError As Integer
    
    On Error GoTo ErrHand
    
    If sgUstWin(8) <> "I" Then
        gMsgBox "Not Allowed to Save.", vbOKOnly
        mSave = True
        Exit Function
    End If
    
    'D.S. 10/10/02 check for date and comments before saving
    ilError = False
    For llRow = grdAction.FixedRows To grdAction.Rows - 1 Step 1
        If (grdAction.TextMatrix(llRow, ACTIONDATEINDEX) <> "") Or (grdAction.TextMatrix(llRow, COMMENTINDEX) <> "") Then
            If (grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = "") Then
                grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = "Missing"
                ilError = True
            Else
                sDate = grdAction.TextMatrix(llRow, ACTIONDATEINDEX)
                If Not gIsDate(sDate) Then
                    grdAction.Row = llRow
                    grdAction.Col = ACTIONDATEINDEX
                    grdAction.CellForeColor = vbRed
                    ilError = True
                End If
            End If
            If (grdAction.TextMatrix(llRow, COMMENTINDEX) = "") Then
                grdAction.TextMatrix(llRow, COMMENTINDEX) = "Missing"
                ilError = True
            End If
        End If
    Next llRow
    If ilError Then
        mSave = False
        Exit Function
    End If
    SQLQuery = "DELETE "
    SQLQuery = SQLQuery + " FROM cct"
    SQLQuery = SQLQuery + " WHERE (cct.cctShfCode= " & imShttCode & ""
    SQLQuery = SQLQuery + " AND cct.cctVefCode = " & imVefCode & ")"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "ContactGrid-mSave"
        cnn.RollbackTrans
        mSave = False
        Exit Function
    End If
    For llRow = grdAction.FixedRows To grdAction.Rows - 1 Step 1
        If (grdAction.TextMatrix(llRow, ACTIONDATEINDEX) <> "") Or (grdAction.TextMatrix(llRow, COMMENTINDEX) <> "") Then
            sDate = grdAction.TextMatrix(llRow, ACTIONDATEINDEX)
            sDate = Format$(sDate, sgShowDateForm)
            sComment = grdAction.TextMatrix(llRow, COMMENTINDEX)
            sComment = gFixQuote(Trim$(sComment)) & Chr(0)
            If (sDate <> "") And (sComment <> "") Then
                SQLQuery = "INSERT INTO cct (cctShfCode, cctVefCode, cctActionDate, cctComment)"
                SQLQuery = SQLQuery & " VALUES (" & imShttCode & ", " & imVefCode & ", '"
                SQLQuery = SQLQuery & Format$(sDate, sgSQLDateForm) & "','" & sComment & "')"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "ContactGrid-mSave"
                    cnn.RollbackTrans
                    mSave = False
                    Exit Function
                End If
            End If
        End If
    Next llRow
    cnn.CommitTrans
    mSave = True
    imCommentChgd = False
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", ""
    mSave = False
    Exit Function
End Function

Private Sub pbcActionSTab_GotFocus()
    If GetFocus() <> pbcActionSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mActionEnableBox
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mActionSetShow
        If grdAction.Col = ACTIONDATEINDEX Then
            If grdAction.Row > grdAction.FixedRows Then
                lmTopRow = -1
                grdAction.Row = grdAction.Row - 1
                If Not grdAction.RowIsVisible(grdAction.Row) Then
                    grdAction.TopRow = grdAction.TopRow - 1
                End If
                grdAction.Col = COMMENTINDEX
                mActionEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            grdAction.Col = grdAction.Col - 1
            mActionEnableBox
        End If
    Else
        lmTopRow = -1
        grdAction.TopRow = grdAction.FixedRows
        grdAction.Col = ACTIONDATEINDEX
        grdAction.Row = grdAction.FixedRows
        mActionEnableBox
    End If

End Sub

Private Sub pbcActionTab_GotFocus()
    Dim llRow As Long
    
    If GetFocus() <> pbcActionTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(8) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mActionSetShow
        If grdAction.Col = grdAction.Cols - 2 Then
'            If grdAction.Row + 1 < grdAction.Rows Then
            llRow = grdAction.Rows
            Do
                llRow = llRow - 1
            Loop While grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = ""
            llRow = llRow + 1
            If (grdAction.Row + 1 < llRow) Then
                lmTopRow = -1
                grdAction.Row = grdAction.Row + 1
                If Not grdAction.RowIsVisible(grdAction.Row) Then
                    grdAction.TopRow = grdAction.TopRow + 1
                End If
                grdAction.Col = ACTIONDATEINDEX
                If Trim$(grdAction.TextMatrix(grdAction.Row, ACTIONDATEINDEX)) <> "" Then
                    mActionEnableBox
                Else
                    imFromArrow = True
                    pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
                    pbcActionArrow.Visible = True
                    pbcActionArrow.SetFocus
                End If
            Else
                If txtDropdown.Text <> "" Then
                    lmTopRow = -1
                    If grdAction.Row + 1 < grdAction.Rows Then
                        grdAction.AddItem ""
                    End If
                    grdAction.Row = grdAction.Row + 1
                    If Not grdAction.RowIsVisible(grdAction.Row) Then
                        grdAction.TopRow = grdAction.TopRow + 1
                    End If
                    grdAction.Col = ACTIONDATEINDEX
                    'mActionEnableBox
                    imFromArrow = True
                    pbcActionArrow.Move grdAction.Left - pbcActionArrow.Width, grdAction.Top + grdAction.RowPos(grdAction.Row) + (grdAction.RowHeight(grdAction.Row) - pbcActionArrow.Height) / 2
                    pbcActionArrow.Visible = True
                    pbcActionArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdAction.Col = grdAction.Col + 1
            mActionEnableBox
        End If
    Else
        lmTopRow = -1
        grdAction.TopRow = grdAction.FixedRows
        grdAction.Col = ACTIONDATEINDEX
        grdAction.Row = grdAction.FixedRows
        mActionEnableBox
    End If
End Sub

Private Sub tmcDelay_Timer()
    Dim iIndex As Integer
    Dim llRow As Long
    
    tmcDelay.Enabled = False
    Screen.MousePointer = vbHourglass
    pbcArrow.Move grdPost.Left - pbcArrow.Width - 15, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    iIndex = Val(grdPost.TextMatrix(grdPost.Row, 7))
    imShttCode = tmContactInfo(iIndex).iShttCode
    imVefCode = tmContactInfo(iIndex).iVefCode
    SQLQuery = "SELECT cctActionDate, cctComment, cctCode FROM cct WHERE (cctshfCode = " & imShttCode & " And cctVefCode = " & imVefCode & ")" & " ORDER By cctActionDate desc"
    Set rst = gSQLSelectCall(SQLQuery)
    gGrid_Clear grdAction, True
    llRow = grdAction.FixedRows
    While Not rst.EOF
        If llRow + 1 > grdAction.Rows Then
            grdAction.AddItem ""
        End If
        grdAction.Row = llRow
        grdAction.TextMatrix(llRow, ACTIONDATEINDEX) = Trim$(rst!cctActionDate)
        grdAction.TextMatrix(llRow, COMMENTINDEX) = Trim$(rst!cctComment)
        grdAction.TextMatrix(llRow, CCTCODEINDEX) = rst!cctCode
        llRow = llRow + 1
        rst.MoveNext
    Wend
    If llRow >= grdAction.Rows Then
        grdAction.AddItem ""
    End If
    lmPostRow = grdPost.Row
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", ""
End Sub

Private Sub txtDropdown_Change()
    Dim slStr As String
    
    Select Case grdAction.Col
        Case ACTIONDATEINDEX
            slStr = Trim$(txtDropdown.Text)
            If (gIsDate(slStr)) And (slStr <> "") Then
                grdAction.CellForeColor = vbBlack
                If grdAction.Text <> txtDropdown.Text Then
                    imCommentChgd = True
                End If
                grdAction.Text = txtDropdown.Text
            End If
        Case COMMENTINDEX
            slStr = txtDropdown.Text
            grdAction.CellForeColor = vbBlack
            If grdAction.Text <> txtDropdown.Text Then
                imCommentChgd = True
            End If
            grdAction.Text = txtDropdown.Text
    End Select
End Sub

Private Sub txtDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
