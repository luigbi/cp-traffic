VERSION 5.00
Begin VB.Form frmFollowUp 
   Caption         =   "Follow Up"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11580
   Begin VB.PictureBox cbcStartDate 
      Height          =   270
      Left            =   3120
      ScaleHeight     =   210
      ScaleWidth      =   1185
      TabIndex        =   11
      Top             =   6525
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox edcCommentTip 
      Appearance      =   0  'Flat
      Height          =   540
      HideSelection   =   0   'False
      Left            =   4605
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2430
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9600
      Top             =   4680
   End
   Begin VB.Timer tmcComment 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   10125
      Top             =   4035
   End
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   6210
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.ListBox lbcOwner 
      Height          =   255
      Left            =   10455
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   5850
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   2
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
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   15
      Width           =   60
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   11040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   5295
      Width           =   1200
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   10590
      TabIndex        =   3
      Top             =   6495
      Width           =   885
   End
   Begin VB.PictureBox grdComment 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   150
      ScaleHeight     =   975
      ScaleWidth      =   8655
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1125
      Visible         =   0   'False
      Width           =   8685
   End
   Begin VB.PictureBox cbcEndDate 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5355
      ScaleHeight     =   210
      ScaleWidth      =   1185
      TabIndex        =   13
      Top             =   6525
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lacEndDate 
      Caption         =   "End Date"
      Height          =   255
      Left            =   4545
      TabIndex        =   12
      Top             =   6555
      Width           =   825
   End
   Begin VB.Label lacStartDate 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   2250
      TabIndex        =   10
      Top             =   6555
      Width           =   855
   End
   Begin VB.Label lacCommentTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8055
      TabIndex        =   9
      Top             =   2610
      Visible         =   0   'False
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Top             =   6570
      Width           =   480
   End
   Begin VB.Label lacUserOption 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Assigned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   990
      TabIndex        =   5
      Top             =   6525
      Width           =   1095
   End
End
Attribute VB_Name = "frmFollowUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmFollowUp - displays missed spots to be changed to Makegoods
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

Private smWeek1 As String
Private smWeek54 As String
Private imShttCode As Integer
Private lmAttCode As Long
Private imVefCode As Integer

Private smFollowUpStartDate As String
Private lmFollowUpStartDate As Long
Private smFollowUpEndDate As String
Private lmFollowUpEndDate As Long

Private imLastCommentColSorted As Integer
Private imLastCommentSort As Integer

Private imLastPostedInfoColSorted As Integer
Private imLastPostedInfoSort As Integer

Private imLastSpotInfoColSorted As Integer
Private imLastSpotInfoSort As Integer

Private bmAgreementScrollAllowed As Boolean
Private lmAgreementTopRow As Long

Private bmPostedScrollAllowed As Boolean
Private lmPostedTopRow As Long

Private rst_cct As ADODB.Recordset
Private rst_ust As ADODB.Recordset
Private rst_att As ADODB.Recordset

Private imSelectedStations() As Integer
Private tmFilterLink() As FILTERLINK
Private tmAndFilterLink() As FILTERLINK

'Grid Controls

'Comment Grid- grdComment
Const CSELECTINDEX = 0
Const CPOSTEDDATEINDEX = 1
Const CBYINDEX = 2
Const CVEHICLEINDEX = 3
Const CFOLLOWUPINDEX = 4
Const COKINDEX = 5
Const CCOMMENTINDEX = 6
Const CPOSTEDTIMEINDEX = 7
Const CCCTCODEINDEX = 8
Const CSORTINDEX = 9
Const CCHGDINDEX = 10

'Personnel Contact Grid- grdContact
Const PCNAMEINDEX = 0
Const PCTITLEINDEX = 1
Const PCPHONEINDEX = 2
Const PCFAXINDEX = 3
Const PCEMAILINDEX = 4
Const PCWEBINDEX = 5
Const PCISCIINDEX = 6
Const PCARTTCODEINDEX = 7

'Posted Info Grid- grdPostedInfo
Const PSELECTINDEX = 0
Const PWEEKINDEX = 1
Const PNOSCHDINDEX = 2
Const PNOAIREDINDEX = 3
Const PNOCMPLINDEX = 4
Const PPERCENTAINDEX = 5
Const PPERCENTCINDEX = 6
Const PPOSTDATEINDEX = 7
Const PBYINDEX = 8
Const PIPINDEX = 9
Const PSTATUSINDEX = 10
Const PCPTTINDEX = 11
Const PSORTINDEX = 12


'Spot Info Grid- grdSpotInfo
Const DDATEINDEX = 0
Const DFEDINDEX = 1
Const DPLEGDEDDAYINDEX = 2
Const DPLEGDEDTIMEINDEX = 3
Const DAIREDDATEINDEX = 4
Const DAIREDTIMEINDEX = 5
Const DADVTINDEX = 6
Const DPRODINDEX = 7
Const DLENGTHINDEX = 8
Const DISCIINDEX = 9
Const DCARTINDEX = 10
Const DCOMMENTINDEX = 11
Const DSTATUSINDEX = 12
Const DASTCODEINDEX = 13
Const DSORTINDEX = 14

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


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



Private Sub cbcEndDate_CalendarChanged()
    tmcComment.Enabled = False
    tmcComment.Enabled = True
End Sub

Private Sub cbcStartDate_CalendarChanged()
    tmcComment.Enabled = False
    tmcComment.Enabled = True
End Sub


Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim imEnabled As Integer
    
    imEnabled = tmcComment.Enabled
    tmcComment.Enabled = False
    ilRet = MsgBox("OK to Exit Search Station?", vbQuestion + vbYesNo, "Exit")
    If ilRet = vbNo Then
        tmcComment.Enabled = imEnabled
        Exit Sub
    End If
    Unload frmFollowUp
    Exit Sub
   
End Sub




Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        mMousePointer vbHourglass
        tmcStart.Enabled = True
        'mSetGridColumns
        'mSetGridTitles
        'gGrid_IntegralHeight grdStations
        'gGrid_FillWithRows grdStations
        'mPopStations
        'lbcKey.FontBold = False
        'lbcKey.FontName = "Arial"
        'lbcKey.FontBold = False
        'lbcKey.FontSize = 8
        'lbcKey.Height = (lbcKey.ListCount - 1) * 225
        'lbcKey.Move imcKey.Left, imcKey.Top - lbcKey.Height
        'imFirstTime = False
        'Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.15
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmFollowUp
    pbcArrow.Width = 90
    'gCenterForm frmFollowUp
End Sub

Private Sub Form_Load()
    
    mMousePointer vbHourglass
    
    mInit
    mMousePointer vbDefault
    Exit Sub
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    edcCommentTip.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    
    Erase tmFilterLink
    Erase tmAndFilterLink
    
    On Error Resume Next
    rst_cct.Close
    rst_ust.Close
    rst_att.Close
    On Error GoTo 0
    
    Set frmFollowUp = Nothing
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    
    frmFollowUp.Caption = "Follow Up- " & Trim$(sgUserName)
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    pbcFocus.Move -100, -100
    pbcClickFocus.Move -100, -100
    
    ReDim tmFilterLink(0 To 0) As FILTERLINK
    ReDim tmAndFilterLink(0 To 0) As FILTERLINK
    
    
    imLastCommentColSorted = -1
    imLastCommentSort = -1
    
    mPopOwnerList
    
    mClearGrid grdComment

    
End Sub

Private Sub mPopStations()
    Dim ilShtt As Integer
    Dim llRet As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llNext As Long
    Dim llCellColor As Long
    Dim ilIncludeStation As Integer
    Dim llFilterDefIndex As Long
    Dim llNotFilterDefIndex As Long
    Dim llCell As Long
    
    ReDim imSelectedStations(0 To 0) As Integer
    mBuildFilter
    On Error GoTo ErrHand:
    For ilShtt = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(ilShtt).iType = 0 Then
            If (UBound(tmFilterLink) > 0) Then
                For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
                    ilIncludeStation = True
                    'Test if station matches all conditions
                    'Select: 0=DMA; 1=Format; 2=MSA; 3=Owner; 4=Vehicle; 5=Zip
                    'Operator: 0=Contains; 1=Equal; 2=Greater than; 3=Less Than; 4=Not Equal
                    llFilterDefIndex = tmFilterLink(llCell).lFilterDefIndex
                    If llFilterDefIndex >= 0 Then
                        mTestFilter ilShtt, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            llNext = tmFilterLink(llCell).lNextAnd
                            Do While llNext <> -1
                                llFilterDefIndex = tmAndFilterLink(llNext).lFilterDefIndex
                                If llFilterDefIndex >= 0 Then
                                    mTestFilter ilShtt, llFilterDefIndex, ilIncludeStation
                                    If Not ilIncludeStation Then
                                        Exit Do
                                    End If
                                End If
                                llNext = tmAndFilterLink(llNext).lNextAnd
                            Loop
                        End If
                    End If
                    'Test Not array
                    If ilIncludeStation Then
                        llNotFilterDefIndex = tmFilterLink(llCell).lNotFilterDefIndex
                        If llNotFilterDefIndex >= 0 Then
                            mTestFilter ilShtt, llNotFilterDefIndex, ilIncludeStation
                            If ilIncludeStation Then
                                llNext = tmFilterLink(llCell).lNextAnd
                                Do While llNext <> -1
                                    llNotFilterDefIndex = tmAndFilterLink(llNext).lNotFilterDefIndex
                                    If llNotFilterDefIndex >= 0 Then
                                        mTestFilter ilShtt, llNotFilterDefIndex, ilIncludeStation
                                        If Not ilIncludeStation Then
                                            Exit Do
                                        End If
                                    End If
                                    llNext = tmAndFilterLink(llNext).lNextAnd
                                Loop
                            End If
                        End If
                    End If
                    
                    If ilIncludeStation Then
                        Exit For
                    End If
                Next llCell
                
            Else
                ilIncludeStation = True
            End If
            If ilIncludeStation Then
                'Missing
                'If lacUserOption.Caption = "Assigned" Then
                '    'Missing: Treat as All until user defined with station
                'ElseIf lacUserOption.Caption = "All Affiliates" Then
                '    SQLQuery = "SELECT attCode FROM att"
                '    SQLQuery = SQLQuery + " WHERE ("
                '    SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                '    Set rst_att = cnn.Execute(SQLQuery)
                '    If rst_att.EOF Then
                '        ilIncludeStation = False
                '    End If
                'ElseIf lacUserOption.Caption = "Non-Affiliates" Then
                '    SQLQuery = "SELECT attCode FROM att"
                '    SQLQuery = SQLQuery + " WHERE ("
                '    SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                '    Set rst_att = cnn.Execute(SQLQuery)
                '    If Not rst_att.EOF Then
                '        ilIncludeStation = False
                '    End If
                'ElseIf lacUserOption.Caption = "All Stations" Then
                'End If
            End If
            If ilIncludeStation Then
                imSelectedStations(UBound(imSelectedStations)) = tgStationInfo(ilShtt).iCode
                ReDim Preserve imSelectedStations(0 To UBound(imSelectedStations) + 1) As Integer
            End If
        End If
    Next ilShtt
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFollowUp-mPopStations"
    Resume Next
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    
    grdComment.Width = frmFollowUp.Width - 300 - GRIDSCROLLWIDTH
    grdComment.Height = cmcDone.Top - 180
    gGrid_IntegralHeight grdComment
    grdComment.Height = grdComment.Height + 30
    gGrid_FillWithRows grdComment
    grdComment.Move (frmFollowUp.Width - grdComment.Width) \ 2, (cmcDone.Top - grdComment.Height) \ 2
    grdComment.ColWidth(CCCTCODEINDEX) = 0
    grdComment.ColWidth(CSORTINDEX) = 0
    grdComment.ColWidth(CPOSTEDDATEINDEX) = grdComment.Width * 0.08
    grdComment.ColWidth(CBYINDEX) = grdComment.Width * 0.08
    grdComment.ColWidth(CVEHICLEINDEX) = grdComment.Width * 0.12
    grdComment.ColWidth(CFOLLOWUPINDEX) = grdComment.Width * 0.08
    grdComment.ColWidth(COKINDEX) = grdComment.Width * 0.04
           
    grdComment.ColWidth(CCOMMENTINDEX) = grdComment.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To CCOMMENTINDEX Step 1
        If ilCol <> CCOMMENTINDEX Then
            grdComment.ColWidth(CCOMMENTINDEX) = grdComment.ColWidth(CCOMMENTINDEX) - grdComment.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdComment
    
    
End Sub

Private Sub mSetGridTitles()
    

    grdComment.TextMatrix(0, CPOSTEDDATEINDEX) = "Posted"
    grdComment.TextMatrix(0, CBYINDEX) = "By"
    grdComment.TextMatrix(0, CVEHICLEINDEX) = "Vehicle"
    grdComment.TextMatrix(0, CFOLLOWUPINDEX) = "Follow-up"
    grdComment.TextMatrix(0, COKINDEX) = "Ok"
    grdComment.TextMatrix(0, CCOMMENTINDEX) = "Comment"


End Sub


Private Sub grdComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdComment.ToolTipText = ""
    If (grdComment.MouseRow >= grdComment.FixedRows) And (grdComment.TextMatrix(grdComment.MouseRow, grdComment.MouseCol)) <> "" Then
        grdComment.ToolTipText = grdComment.TextMatrix(grdComment.MouseRow, grdComment.MouseCol)
    End If
End Sub



Private Function mPopCommentGrid() As Integer
    Dim llRow As Long
    Dim llVef As Long
    Dim ilCol As Integer
    Dim llYellowRow As Long
    Dim llCol As Long
    Dim llDate As Long
    Dim blDateOk As Boolean
    Dim ilShtt As Integer
    
    mPopCommentGrid = False
    On Error GoTo ErrHand:
    If cbcStartDate.Text = "" Then
        smFollowUpStartDate = "1/1/1970"
    Else
        smFollowUpStartDate = cbcStartDate.Text
    End If
    If cbcEndDate.Text = "" Then
        smFollowUpEndDate = "12/31/2069"
    Else
        smFollowUpEndDate = cbcEndDate.Text
    End If
    lmFollowUpStartDate = DateValue(gAdjYear(smFollowUpStartDate))
    lmFollowUpEndDate = DateValue(gAdjYear(smFollowUpEndDate))
    
    grdComment.Rows = 2
    mClearGrid grdComment
    gGrid_FillWithRows grdComment
    grdComment.Redraw = False
    llRow = grdComment.FixedRows
    For ilShtt = 0 To UBound(imSelectedStations) - 1 Step 1
        SQLQuery = "SELECT * FROM cct"
        SQLQuery = SQLQuery + " WHERE ("
        SQLQuery = SQLQuery & " cctShfCode = " & imSelectedStations(ilShtt) & ")"
        SQLQuery = SQLQuery & " ORDER BY cctEnteredDate Desc, cctEnteredTime Desc"
        Set rst_cct = cnn.Execute(SQLQuery)
        Do While Not rst_cct.EOF
            blDateOk = False
            If Not IsNull(rst_cct!cctActionDate) Then
                If DateValue(rst_cct!cctActionDate) <> DateValue("12/31/2069") Then
                    llDate = DateValue(rst_cct!cctActionDate)
                    If (llDate >= lmFollowUpStartDate) And (llDate <= lmFollowUpEndDate) Then
                        blDateOk = True
                    End If
                End If
            End If
            If blDateOk Then
                If llRow >= grdComment.Rows Then
                    grdComment.AddItem ""
                End If
                grdComment.Row = llRow
                For ilCol = CPOSTEDDATEINDEX To CCOMMENTINDEX Step 1
                    If ilCol <> COKINDEX Then
                        grdComment.Col = ilCol
                        'If (ilCol = CCOMMENTINDEX) And ((StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) Or (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0)) Then
                        'Else
                            grdComment.CellBackColor = LIGHTYELLOW
                        'End If
                    End If
                Next ilCol
                If Not IsNull(rst_cct!cctentereddate) Then
                    grdComment.TextMatrix(llRow, CPOSTEDDATEINDEX) = Format$(Trim$(rst_cct!cctentereddate), "m/d/yy")
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
                    SQLQuery = "SELECT ustname, ustReportName FROM Ust Where ustCode = " & rst_cct!cctUstCode
                    Set rst_ust = cnn.Execute(SQLQuery)
                    If Not rst_ust.EOF Then
                        If Trim$(rst_ust!ustReportName) <> "" Then
                            grdComment.TextMatrix(llRow, CBYINDEX) = Trim$(rst_ust!ustReportName)
                        Else
                            grdComment.TextMatrix(llRow, CBYINDEX) = Trim$(rst_ust!ustname)
                        End If
                    End If
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
                    If DateValue(rst_cct!cctActionDate) <> DateValue("12/31/2069") Then
                        grdComment.TextMatrix(llRow, CFOLLOWUPINDEX) = Format$(Trim$(rst_cct!cctActionDate), "m/d/yy")
                    Else
                        grdComment.TextMatrix(llRow, CFOLLOWUPINDEX) = ""
                    End If
                Else
                    grdComment.TextMatrix(llRow, CFOLLOWUPINDEX) = ""
                End If
                'Missing
                grdComment.TextMatrix(llRow, COKINDEX) = ""
                grdComment.TextMatrix(llRow, CCOMMENTINDEX) = Trim$(rst_cct!cctComment)
                grdComment.TextMatrix(llRow, CCCTCODEINDEX) = rst_cct!cctCode
                llRow = llRow + 1
            End If
            rst_cct.MoveNext
        Loop
    Next ilShtt
    grdComment.Rows = grdComment.Rows + ((cmcDone.Top - grdComment.Top) \ grdComment.RowHeight(1))
    For llYellowRow = llRow To grdComment.Rows - 1 Step 1
        grdComment.Row = llYellowRow
        For llCol = CPOSTEDDATEINDEX To CCOMMENTINDEX Step 1
            grdComment.Col = llCol
            grdComment.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llYellowRow
    grdComment.Row = 0
    grdComment.Col = PCARTTCODEINDEX
    grdComment.Redraw = True
    grdComment.Visible = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFollowUp-mPopCommentGrid"
    grdComment.Redraw = True
    On Error GoTo 0
End Function


Private Function mGetTitle(ilCode As Integer) As String
    Dim rst_tnt As ADODB.Recordset

    mGetTitle = ""
    SQLQuery = "Select tntCode, tntTitle From Tnt where tntCode = " & ilCode
    Set rst_tnt = cnn.Execute(SQLQuery)
    If Not rst_tnt.EOF Then
        mGetTitle = Trim$(rst_tnt!tntTitle)
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmFollowUp-mGetTitle"
End Function


Private Sub mSetGridPosition()
    grdComment.Height = cmcDone.Top - 180
    gGrid_IntegralHeight grdComment
    grdComment.Height = grdComment.Height + 30
End Sub

Private Function mGetCol(grdCtrl As MSHFlexGrid, X As Single) As Long
    Dim llColLeftPos As Long
    Dim llCol As Long
    
    mGetCol = -1
    llColLeftPos = grdCtrl.ColPos(0)
    For llCol = 0 To grdCtrl.Cols - 1 Step 1
        If grdCtrl.ColWidth(llCol) > 0 Then
            If (X >= llColLeftPos) And (X <= llColLeftPos + grdCtrl.ColWidth(llCol)) Then
                mGetCol = llCol
                Exit Function
            End If
            llColLeftPos = llColLeftPos + grdCtrl.ColWidth(llCol)
        End If
    Next llCol
End Function


Private Sub mSetCommands()
    
End Sub


Private Sub lacUserOption_Click()
    tmcComment.Enabled = False
    If lacUserOption.Caption = "Mine" Then
        lacUserOption.Caption = "Market Rep"
    ElseIf lacUserOption.Caption = "Market Rep" Then
        lacUserOption.Caption = "Service Rep"
    ElseIf lacUserOption.Caption = "Service Rep" Then
        lacUserOption.Caption = "Supervisor"
    ElseIf lacUserOption.Caption = "Supervisor" Then
        lacUserOption.Caption = "Mine"
    End If
    tmcComment.Enabled = True
End Sub



Private Sub mBuildFilter()
    Dim llRow As Long
    Dim llTotalCells As Long
    Dim ilLoop As Integer
    Dim llCell As Long
    Dim ilFilter As Integer
    Dim ilRepeat As Integer
    Dim llRepeatCount As Long
    Dim ilMatch As Integer
    Dim blSelectionExist As Boolean
    Dim tlCount(0 To 3) As FILTERCOUNT  '0=Format(1); 1= Owner(3); 2=Vehicle(4); 3=DMA(0) or MSA(2) or Zip(5) or Territory(6)
    
    '
    'Design:  Build array of OR items.  Each OR item to reference a link list of AND items with that OR item
    '
    'To Build the OR array
    'Step  Description
    '  1   Count the number of items that require an OR item
    '      Each Format requires a separate OR item (Group 0)
    '      Each Owner requires a separate OR item (Group 1)
    '      Each Vehicle requires a separate OR item (Group 2)
    '      Each DMA, MSA, Zip, Territory and Station requires separate OR item (Group 3)
    '      These are treated as one group because the each represent a geographic area
    '  2   Sort number of OR items for each group above from largest to smallest counts
    '  3   Create an UDT to hold each OR item (product of the counts)
    '  4   Distribute the groups into the UDT from the largest to smallest
    '      The number of times that an item should be repeated is determined by taking the
    '      number of times item repeated divided by the number of items in the group
    '      The first repeat is determined from the total number of OR divided by the group count
    '
    
    For ilLoop = 0 To 3 Step 1
        tlCount(ilLoop).iCount = 0
        tlCount(ilLoop).iType = ilLoop
    Next ilLoop
    'Compute Counts for each group
    blSelectionExist = False
    For llRow = 0 To UBound(tgFilterDef) - 1 Step 1
        If tgFilterDef(llRow).iOperator <> 4 Then
            Select Case tgFilterDef(llRow).iSelect
                Case 0, 2, 5, 6, 7
                    tlCount(3).iCount = tlCount(3).iCount + 1
                    blSelectionExist = True
                Case 1
                    tlCount(0).iCount = tlCount(0).iCount + 1
                    blSelectionExist = True
                Case 3
                    tlCount(1).iCount = tlCount(1).iCount + 1
                    blSelectionExist = True
                Case 4
                    tlCount(2).iCount = tlCount(2).iCount + 1
                    blSelectionExist = True
            End Select
        End If
    Next llRow
    'Determine total number of OR required
    llTotalCells = 1
    For ilLoop = 0 To 3 Step 1
        If tlCount(ilLoop).iCount > 0 Then
            llTotalCells = llTotalCells * tlCount(ilLoop).iCount
        End If
    Next ilLoop
    ReDim tmFilterLink(0 To llTotalCells) As FILTERLINK
    ReDim tmAndFilterLink(0 To 0) As FILTERLINK
    If blSelectionExist Then
        For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
            tmFilterLink(llCell).lFilterDefIndex = -1
            tmFilterLink(llCell).lNotFilterDefIndex = -1
            tmFilterLink(llCell).lNextAnd = -1
        Next llCell
        'Sort Counts from larest to smallest
        ArraySortTyp fnAV(tlCount(), 0), UBound(tlCount) + 1, 1, LenB(tlCount(0)), 0, -1, 0
        llRepeatCount = llTotalCells / tlCount(0).iCount
        For ilLoop = 0 To 3 Step 1
            llCell = 0
            For ilFilter = 0 To UBound(tgFilterDef) - 1 Step 1
                If tgFilterDef(ilFilter).iOperator <> 4 Then
                    ilMatch = False
                    Select Case tlCount(ilLoop).iType
                        Case 0  'Format
                            If tgFilterDef(ilFilter).iSelect = 1 Then
                                ilMatch = True
                            End If
                        Case 1  'Owner
                            If tgFilterDef(ilFilter).iSelect = 3 Then
                                ilMatch = True
                            End If
                        Case 2  'Vehicle
                            If tgFilterDef(ilFilter).iSelect = 4 Then
                                ilMatch = True
                            End If
                        Case 3  'DMA, MSA, ZIP, Territory and Call letters
                            If (tgFilterDef(ilFilter).iSelect = 0) Or (tgFilterDef(ilFilter).iSelect = 2) Or (tgFilterDef(ilFilter).iSelect = 5) Or (tgFilterDef(ilFilter).iSelect = 6) Or (tgFilterDef(ilFilter).iSelect = 7) Then
                                ilMatch = True
                            End If
                    End Select
                    If ilMatch Then
                        For ilRepeat = 1 To llRepeatCount Step 1
                            If tmFilterLink(llCell).lFilterDefIndex < 0 Then
                                tmFilterLink(llCell).lFilterDefIndex = ilFilter
                                tmFilterLink(llCell).lNextAnd = -1
                            Else
                                tmAndFilterLink(UBound(tmAndFilterLink)).lFilterDefIndex = ilFilter
                                tmAndFilterLink(UBound(tmAndFilterLink)).lNextAnd = tmFilterLink(llCell).lNextAnd
                                tmFilterLink(llCell).lNextAnd = UBound(tmAndFilterLink)
                                ReDim Preserve tmAndFilterLink(0 To UBound(tmAndFilterLink) + 1) As FILTERLINK
                            End If
                            llCell = llCell + 1
                        Next ilRepeat
                    End If
                End If
            Next ilFilter
            If ilLoop = 3 Then
                Exit For
            End If
            If tlCount(ilLoop + 1).iCount = 0 Then
                Exit For
            End If
            llRepeatCount = llRepeatCount / tlCount(ilLoop + 1).iCount
        Next ilLoop
    Else
        ReDim tmFilterLink(0 To 1) As FILTERLINK
        ReDim tmAndFilterLink(0 To 0) As FILTERLINK
        For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
            tmFilterLink(llCell).lFilterDefIndex = -1
            tmFilterLink(llCell).lNotFilterDefIndex = -1
            tmFilterLink(llCell).lNextAnd = -1
        Next llCell
    End If
    'Add Not's
    For ilFilter = 0 To UBound(tgFilterDef) - 1 Step 1
        If tgFilterDef(ilFilter).iOperator = 4 Then
            For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
                If tmFilterLink(llCell).lNotFilterDefIndex < 0 Then
                    tmFilterLink(llCell).lNotFilterDefIndex = ilFilter
                    tmFilterLink(llCell).lNextAnd = -1
                Else
                    tmAndFilterLink(UBound(tmAndFilterLink)).lNotFilterDefIndex = ilFilter
                    tmAndFilterLink(UBound(tmAndFilterLink)).lNextAnd = tmFilterLink(llCell).lNextAnd
                    tmFilterLink(llCell).lNextAnd = UBound(tmAndFilterLink)
                    ReDim Preserve tmAndFilterLink(0 To UBound(tmAndFilterLink) + 1) As FILTERLINK
                End If
            Next llCell
        End If
    Next ilFilter
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mMousePointer vbHourglass
    mSetGridColumns
    mSetGridTitles

    imFirstTime = False
    mMousePointer vbDefault
    tmcComment.Enabled = True
End Sub

Private Sub tmcComment_Timer()
    tmcComment.Enabled = False
    mMousePointer vbHourglass
    grdComment.Visible = True
    mPopStations
    mPopCommentGrid
    mMousePointer vbDefault
End Sub

Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
    gSetMousePointer grdComment, grdComment, ilMousepointer
End Sub

Private Sub mFilterCompare(slStr As String, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Select Case tgFilterDef(llFilterDefIndex).iOperator
        Case 0  'Contains
            If InStr(1, slStr, tgFilterDef(llFilterDefIndex).sValue, vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 1  'Equal
            If StrComp(slStr, tgFilterDef(llFilterDefIndex).sValue, vbBinaryCompare) <> 0 Then
                ilIncludeStation = False
            End If
        Case 2  'Greater Than
            If StrComp(slStr, tgFilterDef(llFilterDefIndex).sValue, vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 3  'Less Than
            If StrComp(slStr, tgFilterDef(llFilterDefIndex).sValue, vbBinaryCompare) >= 0 Then
                ilIncludeStation = False
            End If
        Case 4  'Not Equal
            If StrComp(slStr, tgFilterDef(llFilterDefIndex).sValue, vbBinaryCompare) = 0 Then
                ilIncludeStation = False
            End If
    End Select

End Sub

Private Sub mTestFilter(ilShtt As Integer, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim llValue As Long
    Dim slStr As String
    Dim ilFmt As Integer
    Dim llMSA As Long
    Dim llOwner As Long
    Dim ilVef As Integer
    Dim ilMnt As Integer
    
    Select Case tgFilterDef(llFilterDefIndex).iSelect
        Case 0  'DMA
            llValue = tgStationInfo(ilShtt).iMktCode
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMarket))
            mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
        Case 1  'Format
            llValue = tgStationInfo(ilShtt).iFormatCode
            ilFmt = gBinarySearchFmt(CInt(llValue))
            If ilFmt <> -1 Then
                slStr = UCase$(Trim$(tgFormatInfo(ilFmt).sName))
                mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case 2  'MSA
            llValue = tgStationInfo(ilShtt).iMSAMktCode
            llMSA = gBinarySearchMSAMkt(llValue)
            If llMSA <> -1 Then
                slStr = UCase$(Trim$(tgMSAMarketInfo(llMSA).sName))
                mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case 3  'Owner
            llValue = tgStationInfo(ilShtt).lOwnerCode
            llOwner = mBinarySearchOwner(llValue)
            If llOwner <> -1 Then
                slStr = UCase$(Trim$(tgOwnerInfo(llOwner).sName))
                mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case 4  'Vehicle
            SQLQuery = "SELECT attVefCode FROM att"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_att = cnn.Execute(SQLQuery)
            Do While Not rst_att.EOF
                ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
                If ilVef <> -1 Then
                    slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
                    mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
                Else
                    ilIncludeStation = False
                End If
                If Not ilIncludeStation Then
                    Exit Sub
                End If
                rst_att.MoveNext
            Loop
        Case 5  'Zip
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
            mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
        Case 6  'Territory
            llValue = tgStationInfo(ilShtt).iMntCode
            ilMnt = gBinarySearchMnt(CInt(llValue))
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgTerritoryInfo(ilMnt).sName))
                mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case 7  'Station
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters) & ", " & Trim$(tgStationInfo(ilShtt).sMarket))
            mFilterCompare slStr, llFilterDefIndex, ilIncludeStation
    End Select

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
                    slSort = Trim$(Str$(DateValue(grdComment.TextMatrix(llRow, CPOSTEDDATEINDEX))))
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
            If (ilCol <> imLastCommentColSorted) Or ((ilCol = imLastCommentColSorted) And (imLastCommentSort = flexSortStringNoCaseDescending)) Then
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
    If ilCol = imLastCommentColSorted Then
        imLastCommentColSorted = CSORTINDEX
    Else
        imLastCommentColSorted = -1
        imLastCommentColSorted = -1
    End If
    gGrid_SortByCol grdComment, CVEHICLEINDEX, CSORTINDEX, imLastCommentColSorted, imLastCommentSort
    imLastCommentColSorted = ilCol
End Sub

Public Function mBinarySearchOwner(llCode As Long) As Long
    
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llListCode As Long
    
    llMin = 0
    llMax = lbcOwner.ListCount - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        llListCode = Val(lbcOwner.List(llMiddle))
        If llCode = llListCode Then
            'found the match
            mBinarySearchOwner = lbcOwner.ItemData(llMiddle)
            Exit Function
        ElseIf llCode < llListCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchOwner = -1
    Exit Function
    
End Function

Private Sub mPopOwnerList()
    Dim ilLoop As Integer
    Dim slStr As String
    
    lbcOwner.Clear
    For ilLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
        slStr = Trim$(Str$(tgOwnerInfo(ilLoop).lCode))
        Do While Len(slStr) < 9
            slStr = "0" & slStr
        Loop
        lbcOwner.AddItem slStr
        lbcOwner.ItemData(lbcOwner.NewIndex) = ilLoop
    Next ilLoop
End Sub

