VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRadarProgSchd 
   Caption         =   "Radar Program Schedule"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "AffRadarProgSchd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8910
   Begin VB.ComboBox cboNetVehCode 
      Height          =   315
      ItemData        =   "AffRadarProgSchd.frx":08CA
      Left            =   6930
      List            =   "AffRadarProgSchd.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   1560
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "&Erase"
      Height          =   375
      Left            =   6375
      TabIndex        =   21
      Top             =   5445
      Width           =   1335
   End
   Begin VB.PictureBox pbcLbcVehicleTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   8565
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   570
      Width           =   60
   End
   Begin VB.ListBox lbcDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffRadarProgSchd.frx":08CE
      Left            =   5310
      List            =   "AffRadarProgSchd.frx":08EA
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2610
      Width           =   1410
   End
   Begin VB.TextBox txtDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   5310
      TabIndex        =   13
      Top             =   3510
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
      Left            =   6255
      Picture         =   "AffRadarProgSchd.frx":090F
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcProgSchdFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   9
      Top             =   120
      Width           =   60
   End
   Begin VB.TextBox txtProgSchd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1785
      TabIndex        =   12
      Top             =   2955
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtProgSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   885
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcProgSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   3360
      ScaleHeight     =   180
      ScaleWidth      =   765
      TabIndex        =   6
      Top             =   900
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox pbcProgSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   1110
      Width           =   60
   End
   Begin VB.PictureBox pbcProgSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   345
      Width           =   60
   End
   Begin VB.PictureBox pbcProgSchdSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   11
      Top             =   1275
      Width           =   60
   End
   Begin VB.PictureBox pbcProgSchdTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   16
      Top             =   5130
      Width           =   60
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4590
      TabIndex        =   20
      Top             =   5445
      Width           =   1335
   End
   Begin VB.ComboBox cboSelect 
      Height          =   315
      ItemData        =   "AffRadarProgSchd.frx":0A09
      Left            =   2415
      List            =   "AffRadarProgSchd.frx":0A0B
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   4035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   19
      Top             =   5445
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   135
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5730
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AffRadarProgSchd.frx":0A0D
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   90
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   375
      Top             =   5535
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6060
      FormDesignWidth =   8910
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   1095
      TabIndex        =   18
      Top             =   5445
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdProgSchd 
      Height          =   3615
      Left            =   165
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1410
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdProgSpec 
      Height          =   585
      Left            =   165
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   1032
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7950
      Picture         =   "AffRadarProgSchd.frx":0D17
      Top             =   5385
      Width           =   480
   End
End
Attribute VB_Name = "frmRadarProgSchd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmRadarProgSchd - displays missed spots to be changed to Makegoods
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
Private imFieldChgd As Integer
Private lmRowSelected As Long
Private lmTopRow As Long
Private lmProgSpecEnableRow As Long
Private lmProgSpecEnableCol As Long
Private lmProgSchdEnableRow As Long
Private lmProgSchdEnableCol As Long
Private imProgSchdShowGridBox As Integer
Private imProgSpecShowGridBox As Integer
Private imInChg As Integer
Private imInModel As Integer
Private imVefCode As Integer
Private lmRhtCode As Long
Private imFromArrow As Integer
Private imCallModel As Integer
Private imProgSpecColPos(0 To 4) As Integer 'Save column position because of merge
Private imProgSchdColPos(0 To 5) As Integer 'Save column position because of merge

Private imLastProgColSorted As Integer
Private imLastProgSort As Integer

Private lmDelRetCode() As Long

Private rst_rht As ADODB.Recordset
Private rst_ret As ADODB.Recordset

'Grid Controls

Const NETCODEINDEX = 0
Const VEHCODEINDEX = 1
Const SCHDAYTYPEINDEX = 2
Const CLEARTYPEINDEX = 3
Const RHTCODEINDEX = 4

Const PROGCODEINDEX = 0
Const STARTTIMEINDEX = 1
Const ENDTIMEINDEX = 2
Const DAYINDEX = 3
Const SORTINDEX = 4
Const RETCODEINDEX = 5




Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    llRow = grdProgSpec.FixedRows
    For llCol = NETCODEINDEX To RHTCODEINDEX Step 1
        If llCol <> RHTCODEINDEX Then
            grdProgSpec.TextMatrix(llRow, llCol) = ""
        Else
            grdProgSpec.TextMatrix(llRow, llCol) = "0"
        End If
    Next llCol
    
    grdProgSchd.Rows = grdProgSchd.FixedRows + 1
    gGrid_IntegralHeight grdProgSchd
    gGrid_FillWithRows grdProgSchd
    For llRow = grdProgSchd.FixedRows To grdProgSchd.Rows - 1 Step 1
        For llCol = PROGCODEINDEX To RETCODEINDEX Step 1
            If llCol <> RETCODEINDEX Then
                grdProgSchd.TextMatrix(llRow, llCol) = ""
            Else
                grdProgSchd.TextMatrix(llRow, llCol) = "0"
            End If
        Next llCol
    Next llRow
End Sub


Private Sub cboNetVehCode_Change()
    Dim slName As String
    Dim llRow As Long
    Dim ilLen As Integer
    Dim ilRet As Integer
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    mClearGrid
    imLastProgColSorted = -1
    imLastProgSort = -1
    ReDim lmDelRetCode(0 To 0) As Long
    slName = LTrim$(cboNetVehCode.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboNetVehCode.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboNetVehCode.ListIndex = llRow
        cboNetVehCode.SelStart = ilLen
        cboNetVehCode.SelLength = Len(cboNetVehCode.Text)
        lmRhtCode = cboNetVehCode.ItemData(cboNetVehCode.ListIndex)
        ilRet = mGetRhtRet(False)
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-cboNetVehCode"
End Sub

Private Sub cboNetVehCode_Click()
    cboNetVehCode_Change
End Sub

Private Sub cboNetVehCode_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
    If imInModel = True Then
        imInModel = False
    End If
End Sub

Private Sub cboNetVehCode_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboNetVehCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboNetVehCode.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cboSelect_Change()
    Dim slName As String
    Dim llRow As Long
    Dim ilLen As Integer
    Dim ilRet As Integer
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    mClearGrid
    cboNetVehCode.Clear
    lmRhtCode = -1
    imLastProgColSorted = -1
    imLastProgSort = -1
    ReDim lmDelRetCode(0 To 0) As Long
    slName = LTrim$(cboSelect.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cboSelect.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        cboSelect.ListIndex = llRow
        cboSelect.SelStart = ilLen
        cboSelect.SelLength = Len(cboSelect.Text)
        imVefCode = cboSelect.ItemData(cboSelect.ListIndex)
        mPopNetVehCode
    Else
        imVefCode = -1
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-cboSelect"
End Sub

Private Sub cboSelect_Click()
    cboSelect_Change
End Sub

Private Sub cboSelect_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
    If imInModel = True Then
        imInModel = False
    End If

End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSelect.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cmcDropDown_Click()
    lbcDays.Visible = Not lbcDays.Visible
End Sub

Private Sub cmdCancel_Click()
    Dim ilResponse As Integer
    
    If sgUstWin(12) = "I" Then
        If imFieldChgd Then
            ilResponse = gMsgBox("Changes were made! Are you sure you want to cancel? ", vbYesNo)
            If ilResponse = vbNo Then
                Exit Sub
            End If
        End If
    End If
    Unload frmRadarProgSchd
End Sub

Private Sub cmdCancel_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
End Sub

Private Sub cmdDone_Click()
    Dim ilRet As Integer
    
    If sgUstWin(12) = "I" Then
        Screen.MousePointer = vbHourglass
        If imFieldChgd = True Then
            ilRet = mSave()
            grdProgSpec.Redraw = True
            grdProgSchd.Redraw = True
            If Not ilRet Then
                Screen.MousePointer = vbDefault
                Exit Sub    ' Dont exit until user takes care of whatever fields are invalid or missing.
            End If
        End If
        On Error GoTo 0
        Screen.MousePointer = vbDefault
    End If
    Unload frmRadarProgSchd
    Exit Sub
   
End Sub
Private Sub cmdDone_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
End Sub



Private Sub cmdErase_Click()
    Dim llRhtCode As Long
    Dim ilRet As Integer
    
    ilRet = gMsgBox("Remove?", vbYesNo)
    If ilRet = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand
    llRhtCode = Val(grdProgSpec.TextMatrix(grdProgSpec.FixedRows, RHTCODEINDEX))
    If llRhtCode > 0 Then
        SQLQuery = "DELETE FROM RHT WHERE rhtCode = " & llRhtCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "RadarProgSchd-cmdErase_Click"
            Exit Sub
        End If
        SQLQuery = "DELETE FROM RET WHERE retrhtCode = " & llRhtCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "RadarProgSchd-cmdErase_Click"
            Exit Sub
        End If
    End If
    mClearGrid
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-cmdErase"
 End Sub

Private Sub cmdErase_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
End Sub

Private Sub cmdSave_Click()
    Dim ilRet As Integer
    
    ilRet = mSave()
    grdProgSpec.Redraw = True
    grdProgSchd.Redraw = True
    If ilRet Then
        mClearGrid
        mPopNetVehCode
        'ilRet = mPopulateGrid(lmRhtCode, False)
    End If
End Sub

Private Sub cmdSave_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        bgRadarVisible = True
        mSetGridColumns
        mSetGridTitles
        gGrid_IntegralHeight grdProgSpec
        gGrid_IntegralHeight grdProgSchd
        gGrid_FillWithRows grdProgSchd
        mPopulate
        imFirstTime = False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.55   '1.15
    Me.Height = Screen.Height / 1.55
    Me.Top = (Screen.Height - Me.Height) / 1.75
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmRadarProgSchd
    'gCenterForm frmRadarProgSchd
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    
    mInit
    Screen.MousePointer = vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    bgRadarVisible = False
    Erase lmDelRetCode
    rst_rht.Close
    rst_ret.Close
    Set frmRadarProgSchd = Nothing
End Sub







Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    imVefCode = -1
    lmRhtCode = -1
    imInModel = False
    imLastProgColSorted = -1
    imLastProgSort = -1
    lmProgSpecEnableRow = -1
    lmProgSpecEnableCol = -1
    lmProgSchdEnableRow = -1
    lmProgSchdEnableCol = -1
    imProgSchdShowGridBox = False
    imProgSpecShowGridBox = False
    ReDim lmDelRetCode(0 To 0) As Long
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
    imFromArrow = False
    imFieldChgd = False
    If sgUstWin(12) <> "I" Then
        cmdSave.Enabled = False
        cmdErase.Enabled = False
        imcTrash.Enabled = False
    End If
    mPopulate
    
    mClearGrid

End Sub

Private Sub mPopulate()
    Dim iLoop As Integer
    cboSelect.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            cboSelect.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            cboSelect.ItemData(cboSelect.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdProgSpec.ColWidth(RHTCODEINDEX) = 0
    grdProgSpec.ColWidth(NETCODEINDEX) = grdProgSpec.Width * 0.23
    grdProgSpec.ColWidth(VEHCODEINDEX) = grdProgSpec.Width * 0.2
    grdProgSpec.ColWidth(CLEARTYPEINDEX) = grdProgSpec.Width * 0.25
    
    grdProgSpec.ColWidth(SCHDAYTYPEINDEX) = grdProgSpec.Width - 15
    For ilCol = 0 To CLEARTYPEINDEX Step 1
        If ilCol <> SCHDAYTYPEINDEX Then
            grdProgSpec.ColWidth(SCHDAYTYPEINDEX) = grdProgSpec.ColWidth(SCHDAYTYPEINDEX) - grdProgSpec.ColWidth(ilCol)
        End If
    Next ilCol
    gGrid_AlignAllColsLeft grdProgSpec
    For ilCol = 0 To grdProgSpec.Cols - 1 Step 1
        imProgSpecColPos(ilCol) = grdProgSpec.ColPos(ilCol)
    Next ilCol

    grdProgSchd.ColWidth(SORTINDEX) = 0
    grdProgSchd.ColWidth(RETCODEINDEX) = 0
    grdProgSchd.ColWidth(PROGCODEINDEX) = grdProgSchd.Width * 0.25
    grdProgSchd.ColWidth(STARTTIMEINDEX) = grdProgSchd.Width * 0.25
    grdProgSchd.ColWidth(ENDTIMEINDEX) = grdProgSchd.Width * 0.25
    
    grdProgSchd.ColWidth(DAYINDEX) = grdProgSchd.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To DAYINDEX Step 1
        If ilCol <> DAYINDEX Then
            grdProgSchd.ColWidth(DAYINDEX) = grdProgSchd.ColWidth(DAYINDEX) - grdProgSchd.ColWidth(ilCol)
        End If
    Next ilCol
    gGrid_AlignAllColsLeft grdProgSchd
    For ilCol = 0 To grdProgSchd.Cols - 1 Step 1
        imProgSchdColPos(ilCol) = grdProgSchd.ColPos(ilCol)
    Next ilCol
End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    
    grdProgSpec.TextMatrix(0, NETCODEINDEX) = "Network Code"
    grdProgSpec.TextMatrix(0, VEHCODEINDEX) = "Vehicle Code"
    grdProgSpec.TextMatrix(0, SCHDAYTYPEINDEX) = "Schedule Day Type"
    grdProgSpec.TextMatrix(0, CLEARTYPEINDEX) = "Clearance Type"
    
    grdProgSchd.TextMatrix(0, PROGCODEINDEX) = "Program Code"
    grdProgSchd.TextMatrix(0, STARTTIMEINDEX) = "Start Time"
    grdProgSchd.TextMatrix(0, ENDTIMEINDEX) = "End Time"
    grdProgSchd.TextMatrix(0, DAYINDEX) = "Day(s)"
    grdProgSchd.Row = 0
    For llCol = PROGCODEINDEX To DAYINDEX Step 1
        grdProgSchd.Col = llCol
        grdProgSchd.CellBackColor = LIGHTBLUE
    Next llCol
End Sub

Private Sub mProgSchdSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdProgSchd.FixedRows To grdProgSchd.Rows - 1 Step 1
        slStr = Trim$(grdProgSchd.TextMatrix(llRow, PROGCODEINDEX))
        If slStr <> "" Then
            If (ilCol = STARTTIMEINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = ENDTIMEINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdProgSchd.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdProgSchd.TextMatrix(llRow, SORTINDEX)
            grdProgSchd.TextMatrix(llRow, SORTINDEX) = slSort & slStr
        End If
    Next llRow
    If ilCol = imLastProgColSorted Then
        imLastProgColSorted = SORTINDEX
    Else
        imLastProgColSorted = -1
        imLastProgSort = -1
    End If
    gGrid_SortByCol grdProgSchd, PROGCODEINDEX, SORTINDEX, imLastProgColSorted, imLastProgSort
    imLastProgColSorted = ilCol
End Sub

Private Sub grdProgSchd_EnterCell()
    mProgSchdSetShow
End Sub

Private Sub grdProgSchd_GotFocus()
    mProgSpecSetShow
End Sub

Private Sub grdProgSchd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdProgSchd.TopRow
    grdProgSchd.Redraw = False
End Sub

Private Sub grdProgSchd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    If sgUstWin(12) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imVefCode = -1 Then
        grdProgSchd.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If lmRhtCode = -1 Then
        grdProgSchd.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    mProgSchdSetShow
    If Y < grdProgSchd.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        grdProgSchd.Redraw = True
        grdProgSchd.Col = grdProgSchd.MouseCol
        mProgSchdSortCol grdProgSchd.Col
        grdProgSchd.Row = 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdProgSchd, X, Y)
    If Not ilFound Then
        grdProgSchd.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If grdProgSchd.Col = DAYINDEX Then
        If grdProgSpec.TextMatrix(grdProgSpec.FixedRows, SCHDAYTYPEINDEX) <> "Day Name" Then
            grdProgSchd.Redraw = True
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    End If
    If Trim$(grdProgSchd.TextMatrix(grdProgSchd.Row, PROGCODEINDEX)) <> "" Then
        'If Not mPledgeColAllowed(grdProgSchd.Col) Then
        '    grdProgSchd.Redraw = True
        '    If pbcClickFocus.Enabled Then
        '        pbcClickFocus.SetFocus
        '    End If
        '    Exit Sub
        'End If
    Else
        If grdProgSchd.Col > PROGCODEINDEX Then
            grdProgSchd.Redraw = True
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
            Exit Sub
        End If
    End If
    If grdProgSchd.Col > DAYINDEX Then
        grdProgSchd.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdProgSchd.TopRow
    llRow = grdProgSchd.Row
    If grdProgSchd.TextMatrix(llRow, PROGCODEINDEX) = "" Then
        grdProgSchd.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdProgSchd.TextMatrix(llRow, PROGCODEINDEX) = ""
        grdProgSchd.Row = llRow + 1
        'grdProgSchd.Col = 0
        grdProgSchd.Redraw = True
    End If
    grdProgSchd.Redraw = True
    mProgSchdEnableBox
End Sub

Private Sub mProgSchdEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llCol As Long
    
    If (grdProgSchd.Row >= grdProgSchd.FixedRows) And (grdProgSchd.Row < grdProgSchd.Rows) And (grdProgSchd.Col >= PROGCODEINDEX) And (grdProgSchd.Col <= DAYINDEX) Then
        lmProgSchdEnableRow = grdProgSchd.Row
        lmProgSchdEnableCol = grdProgSchd.Col
        imProgSchdShowGridBox = True
        pbcArrow.Move grdProgSchd.Left - pbcArrow.Width, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + (grdProgSchd.RowHeight(grdProgSchd.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdProgSchd.Col
            Case DAYINDEX
                txtDropdown.Move grdProgSchd.Left + imProgSchdColPos(grdProgSchd.Col) + 15, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + 15, grdProgSchd.ColWidth(grdProgSchd.Col) - cmcDropDown.Width + 30, grdProgSchd.RowHeight(grdProgSchd.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcDays.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcDays, 9
                If lbcDays.Top + lbcDays.Height > cmdDone.Top Then
                    lbcDays.Top = txtDropdown.Top - lbcDays.Height - 15
                End If
                If grdProgSchd.Text = "Missing" Then
                    grdProgSchd.CellForeColor = vbBlack
                    grdProgSchd.Text = ""
                End If
                If grdProgSchd.Text = "" Then
                    If grdProgSchd.Row = grdProgSchd.FixedRows Then
                        grdProgSchd.Text = "M-F"
                    Else
                        grdProgSchd.Text = grdProgSchd.TextMatrix(lmProgSchdEnableRow - 1, DAYINDEX)
                    End If
                End If
                slStr = grdProgSchd.Text
                ilIndex = SendMessageByString(lbcDays.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcDays.ListIndex = ilIndex
                Else
                    lbcDays.ListIndex = 5
                End If
                txtDropdown.Text = lbcDays.List(lbcDays.ListIndex)
                If txtDropdown.Height > grdProgSchd.RowHeight(grdProgSchd.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdProgSchd.RowHeight(grdProgSchd.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcDays.Visible = True
                If txtDropdown.Enabled Then
                    txtDropdown.SetFocus
                End If
            Case PROGCODEINDEX, STARTTIMEINDEX, ENDTIMEINDEX
                If grdProgSchd.Col = PROGCODEINDEX Then
                    txtProgSchd.MaxLength = 3
                Else
                    txtProgSchd.MaxLength = 11
                End If
                txtProgSchd.Move grdProgSchd.Left + imProgSchdColPos(grdProgSchd.Col) + 15, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + 15, grdProgSchd.ColWidth(grdProgSchd.Col) - 30, grdProgSchd.RowHeight(grdProgSchd.Row) - 15
                If grdProgSchd.Text <> "Missing" Then
                    txtProgSchd.Text = grdProgSchd.Text
                Else
                    txtProgSchd.Text = ""
                End If
                If txtProgSchd.Height > grdProgSchd.RowHeight(grdProgSchd.Row) - 15 Then
                    txtProgSchd.FontName = "Arial"
                    txtProgSchd.Height = grdProgSchd.RowHeight(grdProgSchd.Row) - 15
                End If
                txtProgSchd.Visible = True
                If txtProgSchd.Enabled Then
                    txtProgSchd.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub mProgSchdSetShow()
    Dim slStr As String

    If (lmProgSchdEnableRow >= grdProgSchd.FixedRows) And (lmProgSchdEnableRow < grdProgSchd.Rows) Then
        If lmProgSchdEnableCol = PROGCODEINDEX Then
            slStr = grdProgSchd.TextMatrix(lmProgSchdEnableRow, lmProgSchdEnableCol)
            Do While Len(slStr) < 3
                slStr = "0" & slStr
            Loop
            grdProgSchd.TextMatrix(lmProgSchdEnableRow, lmProgSchdEnableCol) = slStr
        End If
    End If
    lmProgSchdEnableRow = -1
    lmProgSchdEnableCol = -1
    imProgSchdShowGridBox = False
    pbcArrow.Visible = False
    txtProgSchd.Visible = False
    txtDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcDays.Visible = False
End Sub

Private Sub grdProgSchd_Scroll()
    If grdProgSchd.Redraw = False Then
        grdProgSchd.Redraw = True
        grdProgSchd.TopRow = lmTopRow
        grdProgSchd.Refresh
        grdProgSchd.Redraw = False
    End If
    If (imProgSchdShowGridBox) And (grdProgSchd.Row >= grdProgSchd.FixedRows) And (grdProgSchd.Col >= PROGCODEINDEX) And (grdProgSchd.Col <= DAYINDEX) Then
        If grdProgSchd.RowIsVisible(grdProgSchd.Row) Then
            pbcArrow.Move grdProgSchd.Left - pbcArrow.Width, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + (grdProgSchd.RowHeight(grdProgSchd.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            If (grdProgSchd.Col = DAYINDEX) Then
                txtDropdown.Move grdProgSchd.Left + imProgSchdColPos(grdProgSchd.Col) + 15, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + 15, grdProgSchd.ColWidth(grdProgSchd.Col) - cmcDropDown.Width + 30, grdProgSchd.RowHeight(grdProgSchd.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcDays.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcDays.Visible = True
                If txtDropdown.Enabled Then
                    txtDropdown.SetFocus
                End If
            ElseIf (grdProgSchd.Col = PROGCODEINDEX) Or (grdProgSchd.Col = STARTTIMEINDEX) Or (grdProgSchd.Col = ENDTIMEINDEX) Then
                txtProgSchd.Move grdProgSchd.Left + imProgSchdColPos(grdProgSchd.Col) + 15, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + 15, grdProgSchd.ColWidth(grdProgSchd.Col) - 30, grdProgSchd.RowHeight(grdProgSchd.Row) - 15
                txtProgSchd.Visible = True
                If txtProgSchd.Enabled Then
                    txtProgSchd.SetFocus
                End If
            End If
        Else
            If pbcProgSchdFocus.Enabled Then
                pbcProgSchdFocus.SetFocus
            End If
            pbcArrow.Visible = False
            txtProgSchd.Visible = False
            txtDropdown.Visible = False
            cmcDropDown.Visible = False
            lbcDays.Visible = False
        End If
    Else
        If pbcProgSchdFocus.Enabled Then
            pbcProgSchdFocus.SetFocus
        End If
        pbcArrow.Visible = False
        imFromArrow = False
    End If

End Sub

Private Sub grdProgSpec_EnterCell()
    mProgSpecSetShow
End Sub

Private Sub grdProgSpec_GotFocus()
    mProgSchdSetShow
End Sub

Private Sub grdProgSpec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    grdProgSpec.Redraw = False
End Sub

Private Sub grdProgSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilFound As Integer
    
    If sgUstWin(12) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imVefCode = -1 Then
        grdProgSpec.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If lmRhtCode = -1 Then
        grdProgSpec.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    mProgSpecSetShow
    If Y < grdProgSpec.RowHeight(0) Then
        grdProgSpec.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdProgSpec, X, Y)
    If Not ilFound Then
        grdProgSpec.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If grdProgSpec.Col > CLEARTYPEINDEX Then
        grdProgSchd.Redraw = True
        Exit Sub
    End If
    grdProgSpec.Redraw = True
    mProgSpecEnableBox
End Sub


Private Sub imcTrash_Click()
    Dim iLoop As Integer
    Dim llRow As Long
    Dim llRows As Long
    
    mProgSchdSetShow
    llRow = grdProgSchd.Row
    llRows = grdProgSchd.Rows
    If (llRow < 0) Or (llRow > grdProgSchd.Rows - 1) Then
        Exit Sub
    End If
    lmTopRow = -1
    If grdProgSchd.TextMatrix(llRow, RETCODEINDEX) <> "0" Then
        imFieldChgd = True
        lmDelRetCode(UBound(lmDelRetCode)) = Val(grdProgSchd.TextMatrix(llRow, RETCODEINDEX))
        ReDim Preserve lmDelRetCode(0 To UBound(lmDelRetCode) + 1) As Long
    End If
    grdProgSchd.RemoveItem llRow
    gGrid_FillWithRows grdProgSchd
    If pbcClickFocus.Enabled Then
        pbcClickFocus.SetFocus
    End If

End Sub

Private Sub lbcDays_Click()
    txtDropdown.Text = lbcDays.List(lbcDays.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        If txtDropdown.Enabled Then
            txtDropdown.SetFocus
        End If
        lbcDays.Visible = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mProgSpecSetShow
    mProgSchdSetShow
End Sub



Private Sub pbcLbcVehicleTab_GotFocus()
    Dim ilRet As Integer
    
    If imVefCode < 0 Then
        cboSelect.SetFocus
        Exit Sub
    End If
    If lmRhtCode < 0 Then
        cboNetVehCode.SetFocus
        Exit Sub
    End If
    If imInModel Then
        Exit Sub
    End If
    imInModel = True
    ilRet = mGetRhtRet(True)
End Sub

Private Sub pbcProgSchdSTab_GotFocus()
    
    If GetFocus() <> pbcProgSchdSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(12) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imVefCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If lmRhtCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mProgSchdEnableBox
        Exit Sub
    End If
    If imProgSchdShowGridBox Then
        mProgSchdSetShow
        If grdProgSchd.Col = PROGCODEINDEX Then
            If grdProgSchd.Row > grdProgSchd.FixedRows Then
                lmTopRow = -1
                grdProgSchd.Row = grdProgSchd.Row - 1
                If Not grdProgSchd.RowIsVisible(grdProgSchd.Row) Then
                    grdProgSchd.TopRow = grdProgSchd.TopRow - 1
                End If
                If grdProgSpec.TextMatrix(grdProgSpec.FixedRows, SCHDAYTYPEINDEX) <> "Day Name" Then
                    grdProgSchd.Col = ENDTIMEINDEX
                Else
                    grdProgSchd.Col = DAYINDEX
                End If
                mProgSchdEnableBox
            Else
                If pbcClickFocus.Enabled Then
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdProgSchd.Col = grdProgSchd.Col - 1
            mProgSchdEnableBox
        End If
    Else
        lmTopRow = -1
        grdProgSchd.TopRow = grdProgSchd.FixedRows
        grdProgSchd.Col = PROGCODEINDEX
        grdProgSchd.Row = grdProgSchd.FixedRows
        mProgSchdEnableBox
    End If

End Sub

Private Sub pbcProgSchdTab_GotFocus()
    Dim llCol As Long
    Dim slStr As String
    
    If GetFocus() <> pbcProgSchdTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(12) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imVefCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If lmRhtCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imProgSchdShowGridBox Then
        mProgSchdSetShow
        If (grdProgSchd.Col = DAYINDEX) Then
            If grdProgSchd.TextMatrix(grdProgSchd.Row, DAYINDEX) <> "" Then
                lmTopRow = -1
                If grdProgSchd.Row + 1 >= grdProgSchd.Rows Then
                    grdProgSchd.AddItem ""
                    grdProgSchd.TextMatrix(grdProgSchd.Row + 1, RETCODEINDEX) = "0"
                End If
                grdProgSchd.Row = grdProgSchd.Row + 1
                If Not grdProgSchd.RowIsVisible(grdProgSchd.Row) Then
                    grdProgSchd.TopRow = grdProgSchd.TopRow + 1
                End If
                grdProgSchd.Col = PROGCODEINDEX
                imFromArrow = True
                pbcArrow.Move grdProgSchd.Left - pbcArrow.Width, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + (grdProgSchd.RowHeight(grdProgSchd.Row) - pbcArrow.Height) / 2
                pbcArrow.Visible = True
                If pbcArrow.Enabled Then
                    pbcArrow.SetFocus
                End If
            Else
                If pbcClickFocus.Enabled Then
                    pbcClickFocus.SetFocus
                End If
            End If
        ElseIf (grdProgSchd.Col = STARTTIMEINDEX) Or (grdProgSchd.Col = ENDTIMEINDEX) Then
            slStr = Trim$(txtProgSchd.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                If grdProgSchd.Col = ENDTIMEINDEX Then
                    If grdProgSpec.TextMatrix(grdProgSpec.FixedRows, SCHDAYTYPEINDEX) <> "Day Name" Then
                        lmTopRow = -1
                        If grdProgSchd.Row + 1 >= grdProgSchd.Rows Then
                            grdProgSchd.AddItem ""
                            grdProgSchd.TextMatrix(grdProgSchd.Row + 1, RETCODEINDEX) = "0"
                        End If
                        grdProgSchd.Row = grdProgSchd.Row + 1
                        If Not grdProgSchd.RowIsVisible(grdProgSchd.Row) Then
                            grdProgSchd.TopRow = grdProgSchd.TopRow + 1
                        End If
                        grdProgSchd.Col = PROGCODEINDEX
                        imFromArrow = True
                        pbcArrow.Move grdProgSchd.Left - pbcArrow.Width, grdProgSchd.Top + grdProgSchd.RowPos(grdProgSchd.Row) + (grdProgSchd.RowHeight(grdProgSchd.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        If pbcArrow.Enabled Then
                            pbcArrow.SetFocus
                        End If
                        Exit Sub
                    End If
                End If
                grdProgSchd.Col = grdProgSchd.Col + 1
                mProgSchdEnableBox
            Else
                Beep
                grdProgSchd.Col = grdProgSchd.Col
                mProgSchdEnableBox
            End If
        Else
            grdProgSchd.Col = grdProgSchd.Col + 1
            mProgSchdEnableBox
        End If
    Else
        lmTopRow = -1
        grdProgSchd.TopRow = grdProgSchd.FixedRows
        grdProgSchd.Col = PROGCODEINDEX
        grdProgSchd.Row = grdProgSchd.FixedRows
        mProgSchdEnableBox
    End If

End Sub

Private Sub pbcProgSpec_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    Select Case grdProgSpec.Col
        Case SCHDAYTYPEINDEX
            If KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
                grdProgSpec.Text = "Day Name"
                imFieldChgd = True
                pbcProgSpec_Paint
            ElseIf KeyAscii = Asc("F") Or (KeyAscii = Asc("f")) Then
                grdProgSpec.Text = "MF1"
                imFieldChgd = True
                pbcProgSpec_Paint
            ElseIf KeyAscii = Asc("S") Or (KeyAscii = Asc("s")) Then
                grdProgSpec.Text = "MS1"
                imFieldChgd = True
                pbcProgSpec_Paint
            End If
            If KeyAscii = Asc(" ") Then
                slStr = grdProgSpec.Text
                If slStr = "Day Name" Then
                    grdProgSpec.Text = "MF1"
                ElseIf slStr = "MF1" Then
                    grdProgSpec.Text = "MS1"
                ElseIf slStr = "MS1" Then
                    grdProgSpec.Text = "Day Name"
                Else
                    grdProgSpec.Text = "Day Name"
                End If
                imFieldChgd = True
                pbcProgSpec_Paint
            End If
        Case CLEARTYPEINDEX
            If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
                grdProgSpec.Text = "Cmml Only"
                imFieldChgd = True
                pbcProgSpec_Paint
            ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
                grdProgSpec.Text = "Prog+Cmml"
                imFieldChgd = True
                pbcProgSpec_Paint
            ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
                grdProgSpec.Text = "Agreement"
                imFieldChgd = True
                pbcProgSpec_Paint
            End If
            If KeyAscii = Asc(" ") Then
                slStr = grdProgSpec.Text
                If slStr = "Cmml Only" Then
                    grdProgSpec.Text = "Prog+Cmml"
                ElseIf slStr = "Prog+Cmml" Then
                    grdProgSpec.Text = "Agreement"
                ElseIf slStr = "Agreement" Then
                    grdProgSpec.Text = "Cmml Only"
                Else
                    grdProgSpec.Text = "Cmml Only"
                End If
                imFieldChgd = True
                pbcProgSpec_Paint
            End If
    End Select

End Sub

Private Sub pbcProgSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slStr As String
    Select Case grdProgSpec.Col
        Case SCHDAYTYPEINDEX
            slStr = grdProgSpec.Text
            If slStr = "Day Name" Then
                grdProgSpec.Text = "MF1"
            ElseIf slStr = "MF1" Then
                grdProgSpec.Text = "MS1"
            ElseIf slStr = "MS1" Then
                grdProgSpec.Text = "Day Name"
            Else
                grdProgSpec.Text = "Day Name"
            End If
            imFieldChgd = True
            pbcProgSpec_Paint
        Case CLEARTYPEINDEX
            slStr = grdProgSpec.Text
            If slStr = "Cmml Only" Then
                grdProgSpec.Text = "Prog+Cmml"
            ElseIf slStr = "Prog+Cmml" Then
                grdProgSpec.Text = "Agreement"
            ElseIf slStr = "Agreement" Then
                grdProgSpec.Text = "Cmml Only"
            Else
                grdProgSpec.Text = "Cmml Only"
            End If
            imFieldChgd = True
            pbcProgSpec_Paint
    End Select

End Sub

Private Sub pbcProgSpec_Paint()
    pbcProgSpec.Cls
    pbcProgSpec.CurrentX = 15
    pbcProgSpec.CurrentY = 0 'fgBoxInsetY
    pbcProgSpec.Print grdProgSpec.Text
End Sub

Private Sub pbcProgSpecSTab_GotFocus()
    If GetFocus() <> pbcProgSpecSTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(12) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imVefCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If lmRhtCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imProgSpecShowGridBox Then
        mProgSpecSetShow
        If grdProgSpec.Col = NETCODEINDEX Then
            If pbcClickFocus.Enabled Then
                pbcClickFocus.SetFocus
            End If
        Else
            grdProgSpec.Col = grdProgSpec.Col - 1
            mProgSpecEnableBox
        End If
    Else
        lmTopRow = -1
        grdProgSpec.TopRow = grdProgSpec.FixedRows
        grdProgSpec.Col = NETCODEINDEX
        grdProgSpec.Row = grdProgSpec.FixedRows
        mProgSpecEnableBox
    End If
End Sub

Private Sub pbcProgSpecTab_GotFocus()
    Dim llCol As Long
    Dim slStr As String
    
    If GetFocus() <> pbcProgSpecTab.hwnd Then
        Exit Sub
    End If
    If sgUstWin(12) <> "I" Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imVefCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If lmRhtCode = -1 Then
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If imProgSpecShowGridBox Then
        mProgSpecSetShow
        If (grdProgSpec.Col = CLEARTYPEINDEX) Then
            pbcProgSchdSTab.SetFocus
            grdProgSpec.Refresh
        Else
            grdProgSpec.Col = grdProgSpec.Col + 1
            mProgSpecEnableBox
        End If
    Else
        lmTopRow = -1
        grdProgSpec.TopRow = grdProgSpec.FixedRows
        grdProgSpec.Col = PROGCODEINDEX
        grdProgSpec.Row = grdProgSpec.FixedRows
        mProgSpecEnableBox
    End If
End Sub

Private Sub mProgSpecEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llCol As Long
    
    If (grdProgSpec.Row >= grdProgSpec.FixedRows) And (grdProgSpec.Row < grdProgSpec.Rows) And (grdProgSpec.Col >= PROGCODEINDEX) And (grdProgSpec.Col <= DAYINDEX) Then
        lmProgSpecEnableRow = grdProgSpec.Row
        lmProgSpecEnableCol = grdProgSpec.Col
        imProgSpecShowGridBox = True
        Select Case grdProgSpec.Col
            Case SCHDAYTYPEINDEX, CLEARTYPEINDEX
                pbcProgSpec.Move grdProgSpec.Left + imProgSpecColPos(grdProgSpec.Col) + 15, grdProgSpec.Top + grdProgSpec.RowPos(grdProgSpec.Row) + 15, grdProgSpec.ColWidth(grdProgSpec.Col) + 30, grdProgSpec.RowHeight(grdProgSpec.Row) - 15
                If grdProgSpec.Text = "Missing" Then
                    grdProgSpec.CellForeColor = vbBlack
                    grdProgSpec.Text = ""
                End If
                If grdProgSpec.Text = "" Then
                    If grdProgSpec.Col = SCHDAYTYPEINDEX Then
                        grdProgSpec.Text = "Day Name"
                    Else
                        grdProgSpec.Text = "Cmml Only"
                    End If
                End If
                If pbcProgSpec.Height > grdProgSpec.RowHeight(grdProgSpec.Row) - 15 Then
                    pbcProgSpec.FontName = "Arial"
                    pbcProgSpec.Height = grdProgSpec.RowHeight(grdProgSpec.Row) - 15
                End If
                pbcProgSpec.Visible = True
                If pbcProgSpec.Enabled Then
                    pbcProgSpec.SetFocus
                End If
            Case NETCODEINDEX, VEHCODEINDEX
                If grdProgSpec.Col = NETCODEINDEX Then
                    txtProgSpec.MaxLength = 2
                Else
                    txtProgSpec.MaxLength = 3
                End If
                txtProgSpec.Move grdProgSpec.Left + imProgSpecColPos(grdProgSpec.Col) + 15, grdProgSpec.Top + grdProgSpec.RowPos(grdProgSpec.Row) + 15, grdProgSpec.ColWidth(grdProgSpec.Col) - 30, grdProgSpec.RowHeight(grdProgSpec.Row) - 15
                If grdProgSpec.Text <> "Missing" Then
                    txtProgSpec.Text = grdProgSpec.Text
                Else
                    txtProgSpec.Text = ""
                End If
                If txtProgSpec.Height > grdProgSpec.RowHeight(grdProgSpec.Row) - 15 Then
                    txtProgSpec.FontName = "Arial"
                    txtProgSpec.Height = grdProgSpec.RowHeight(grdProgSpec.Row) - 15
                End If
                txtProgSpec.Visible = True
                If txtProgSpec.Enabled Then
                    txtProgSpec.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub mProgSpecSetShow()


    If (lmProgSpecEnableRow >= grdProgSpec.FixedRows) And (lmProgSpecEnableRow < grdProgSpec.Rows) Then
    End If
    lmProgSpecEnableRow = -1
    lmProgSpecEnableCol = -1
    imProgSpecShowGridBox = False
    pbcProgSpec.Visible = False
    txtProgSpec.Visible = False
End Sub



Private Sub txtDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    slStr = txtDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(lbcDays.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        lbcDays.ListIndex = llRow
        txtDropdown.Text = lbcDays.List(lbcDays.ListIndex)
        txtDropdown.SelStart = ilLen
        txtDropdown.SelLength = Len(txtDropdown.Text)
        slStr = txtDropdown.Text
        If (slStr <> "") Then
            grdProgSchd.CellForeColor = vbBlack
            If grdProgSchd.Text <> slStr Then
                imFieldChgd = True
            End If
            grdProgSchd.Text = slStr
        End If
    End If

End Sub

Private Sub txtDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub txtDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If txtDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub txtDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        gProcessArrowKey Shift, KeyCode, lbcDays, True ', imLbcArrowSetting
    End If
End Sub

Private Sub txtProgSchd_Change()
    Dim slStr As String
    
    Select Case grdProgSchd.Col
        Case PROGCODEINDEX
            slStr = Trim$(txtProgSchd.Text)
            If (slStr <> "") Then
                grdProgSchd.CellForeColor = vbBlack
                If grdProgSchd.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdProgSchd.Text = slStr
            End If
        Case STARTTIMEINDEX, ENDTIMEINDEX
            slStr = Trim$(txtProgSchd.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                grdProgSchd.CellForeColor = vbBlack
                slStr = gConvertTime(slStr)
                If Second(slStr) = 0 Then
                    slStr = Format$(slStr, sgShowTimeWOSecForm)
                Else
                    slStr = Format$(slStr, sgShowTimeWSecForm)
                End If
                If grdProgSchd.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdProgSchd.Text = slStr
            End If
    End Select

End Sub

Private Sub txtProgSchd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtProgSchd_KeyPress(KeyAscii As Integer)
    Select Case grdProgSchd.Col
        Case PROGCODEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub txtProgSpec_Change()
    Dim slStr As String
    
    Select Case grdProgSpec.Col
        Case NETCODEINDEX, VEHCODEINDEX
            slStr = Trim$(txtProgSpec.Text)
            If (slStr <> "") Then
                grdProgSpec.CellForeColor = vbBlack
                If grdProgSpec.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdProgSpec.Text = slStr
            End If
    End Select

End Sub

Private Sub txtProgSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Function mPopulateGrid(llRhtCode As Long, ilSort As Integer) As Integer
    Dim llRow As Long
    Dim slStr As String
    
    On Error GoTo ErrHand
    mPopulateGrid = True
    grdProgSpec.Redraw = False
    grdProgSchd.Redraw = False
    SQLQuery = "SELECT * FROM rht WHERE (rhtCode = " & llRhtCode & ")"
    Set rst_rht = gSQLSelectCall(SQLQuery)
    If Not rst_rht.EOF Then
        llRow = grdProgSpec.FixedRows
        grdProgSpec.TextMatrix(llRow, NETCODEINDEX) = rst_rht!rhtRadarNetCode
        If llRhtCode = lmRhtCode Then
            grdProgSpec.TextMatrix(llRow, VEHCODEINDEX) = rst_rht!rhtRadarVehCode
        Else
            grdProgSpec.TextMatrix(llRow, VEHCODEINDEX) = ""
        End If
        slStr = rst_rht!rhtSchdDayType
        If slStr = "MF" Then
            grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX) = "MF1"
        ElseIf slStr = "MS" Then
            grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX) = "MS1"
        Else
            grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX) = "Day Name"
        End If
        slStr = rst_rht!rhtClearType
        If slStr = "P" Then
            grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX) = "Prog+Cmml"
        ElseIf slStr = "A" Then
            grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX) = "Agreement"
        Else
            grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX) = "Cmml Only"
        End If
        If llRhtCode = lmRhtCode Then
            grdProgSpec.TextMatrix(llRow, RHTCODEINDEX) = rst_rht!rhtCode
        Else
            grdProgSpec.TextMatrix(llRow, RHTCODEINDEX) = "0"
        End If
        llRow = grdProgSchd.FixedRows
        SQLQuery = "SELECT * FROM ret WHERE (retRhtCode = " & rst_rht!rhtCode & ")"
        Set rst_ret = gSQLSelectCall(SQLQuery)
        Do While Not rst_ret.EOF
            If llRow >= grdProgSchd.Rows Then
                grdProgSchd.AddItem ""
                grdProgSchd.TextMatrix(llRow, RETCODEINDEX) = "0"
            End If
            grdProgSchd.TextMatrix(llRow, PROGCODEINDEX) = rst_ret!retProgCode
            If Second(rst_ret!retStartTime) = 0 Then
                grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX) = Format$(CStr(rst_ret!retStartTime), sgShowTimeWOSecForm)
            Else
                grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX) = Format$(CStr(rst_ret!retStartTime), sgShowTimeWSecForm)
            End If
            If Second(rst_ret!retEndTime) = 0 Then
                grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX) = Format$(CStr(rst_ret!retEndTime), sgShowTimeWOSecForm)
            Else
                grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX) = Format$(CStr(rst_ret!retEndTime), sgShowTimeWSecForm)
            End If
            slStr = rst_ret!retDayType
            If slStr = "MF" Then
                grdProgSchd.TextMatrix(llRow, DAYINDEX) = "M-F"
            Else
                grdProgSchd.TextMatrix(llRow, DAYINDEX) = slStr
            End If
            If llRhtCode = lmRhtCode Then
                grdProgSchd.TextMatrix(llRow, RETCODEINDEX) = rst_ret!retCode
            Else
                grdProgSchd.TextMatrix(llRow, RETCODEINDEX) = "0"
            End If
            llRow = llRow + 1
            rst_ret.MoveNext
        Loop
        rst_ret.Close
        If ilSort Then
            mProgSchdSortCol PROGCODEINDEX
        End If
    Else
        mPopulateGrid = False
    End If
    grdProgSpec.Redraw = True
    grdProgSchd.Redraw = True
    rst_rht.Close
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-mPopulateGrid"
    mPopulateGrid = False
End Function

Private Function mGetRhtRet(ilCallModel As Integer) As Integer
    Dim ilRet As Integer
    mGetRhtRet = True
    On Error GoTo ErrHand
    mClearGrid
    imLastProgColSorted = -1
    imLastProgSort = -1
    SQLQuery = "SELECT * FROM rht WHERE (rhtCode = " & lmRhtCode & ")"
    Set rst_rht = gSQLSelectCall(SQLQuery)
    If rst_rht.EOF Then
        If ilCallModel Then
            SQLQuery = "SELECT * FROM rht "
            Set rst_rht = gSQLSelectCall(SQLQuery)
            If Not rst_rht.EOF Then
                'Model call needs to be on the timer to aviod double calling the model routine
                'The second call occurs when model returns
                igModelType = 1
                frmModel.Show vbModal
                If igModelReturn Then
                    ilRet = mPopulateGrid(lgModelFromCode, True)
                End If
            Else
                ilRet = True
            End If
        Else
            mGetRhtRet = False
            On Error GoTo 0
            Exit Function
        End If
    Else
        ilRet = mPopulateGrid(lmRhtCode, True)
        mGetRhtRet = ilRet
    End If
    If ilCallModel Then
        If ilRet Then
            pbcProgSpecSTab.SetFocus
        Else
            cmdCancel.SetFocus
        End If
    End If
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-mGetRhtRet"
End Function

Private Function mSave() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim llRhtCode As Long
    Dim llRetCode As Long
    Dim slNC As String
    Dim slVC As String
    
    On Error GoTo ErrHand
    llRow = grdProgSpec.FixedRows
    llRhtCode = Val(grdProgSpec.TextMatrix(llRow, RHTCODEINDEX))
    slNC = grdProgSpec.TextMatrix(llRow, NETCODEINDEX)
    slVC = grdProgSpec.TextMatrix(llRow, VEHCODEINDEX)
    If Not mNameOk(slNC, slVC, llRhtCode) Then
        grdProgSpec.Col = VEHCODEINDEX
        grdProgSpec.Row = grdProgSpec.FixedRows
        mProgSpecEnableBox
        mSave = False
        Exit Function
    End If
    If Not mTestGridValues() Then
        mSave = False
        Exit Function
    End If
    slNC = gFixQuote(grdProgSpec.TextMatrix(llRow, NETCODEINDEX))
    slVC = gFixQuote(grdProgSpec.TextMatrix(llRow, VEHCODEINDEX))
    mSave = True
    If llRhtCode <= 0 Then
        llRhtCode = 0
        'Add
        SQLQuery = "Insert Into rht ( "
        SQLQuery = SQLQuery & "rhtCode, "
        SQLQuery = SQLQuery & "rhtVefCode, "
        SQLQuery = SQLQuery & "rhtRadarNetCode, "
        SQLQuery = SQLQuery & "rhtRadarVehCode, "
        SQLQuery = SQLQuery & "rhtSchdDayType, "
        SQLQuery = SQLQuery & "rhtClearType, "
        SQLQuery = SQLQuery & "rhtUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & llRhtCode & ", "
        SQLQuery = SQLQuery & imVefCode & ", "
        SQLQuery = SQLQuery & "'" & slNC & "', "
        SQLQuery = SQLQuery & "'" & slVC & "', "
        slStr = grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX)
        If slStr = "MF1" Then
            SQLQuery = SQLQuery & "'" & gFixQuote("MF") & "', "
        ElseIf slStr = "MS1" Then
            SQLQuery = SQLQuery & "'" & gFixQuote("MS") & "', "
        Else
            SQLQuery = SQLQuery & "'" & gFixQuote("DN") & "', "
        End If
        SQLQuery = SQLQuery & "'" & gFixQuote(Left$(grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX), 1)) & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
    Else
        'Update
        SQLQuery = "Update rht Set "
        SQLQuery = SQLQuery & "rhtVefCode = " & imVefCode & ", "
        SQLQuery = SQLQuery & "rhtRadarNetCode = '" & slNC & "', "
        SQLQuery = SQLQuery & "rhtRadarVehCode = '" & slVC & "', "
        slStr = grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX)
        If slStr = "MF1" Then
            SQLQuery = SQLQuery & "rhtSchdDayType ='" & gFixQuote("MF") & "', "
        ElseIf slStr = "MS1" Then
            SQLQuery = SQLQuery & "rhtSchdDayType ='" & gFixQuote("MS") & "', "
        Else
            SQLQuery = SQLQuery & "rhtSchdDayType ='" & gFixQuote("DN") & "', "
        End If
        SQLQuery = SQLQuery & "rhtClearType = '" & gFixQuote(Left$(grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX), 1)) & "', "
        SQLQuery = SQLQuery & "rhtUnused = '" & "" & "' "
        SQLQuery = SQLQuery & " WHERE rhtCode = " & llRhtCode
    End If
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "RadarProgSchd-mSave"
        mSave = False
        Exit Function
    End If
    If llRhtCode = 0 Then
        SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND rhtRadarNetCode = '" & slNC & "'"
        SQLQuery = SQLQuery & " AND rhtRadarVehCode = '" & slVC & "'" & ")"
        Set rst_rht = gSQLSelectCall(SQLQuery)
        If Not rst_rht.EOF Then
            llRhtCode = rst_rht!rhtCode
            grdProgSpec.TextMatrix(llRow, RHTCODEINDEX) = llRhtCode
        Else
            mSave = False
            Exit Function
        End If
    End If
    For llRow = grdProgSchd.FixedRows To grdProgSchd.Rows - 1 Step 1
        slStr = Trim$(grdProgSchd.TextMatrix(llRow, PROGCODEINDEX))
        If slStr <> "" Then
            llRetCode = Val(grdProgSchd.TextMatrix(llRow, RETCODEINDEX))
            If llRetCode <= 0 Then
                llRetCode = 0
                'Add
                SQLQuery = "Insert Into ret ( "
                SQLQuery = SQLQuery & "retCode, "
                SQLQuery = SQLQuery & "retRhtCode, "
                SQLQuery = SQLQuery & "retProgCode, "
                SQLQuery = SQLQuery & "retStartTime, "
                SQLQuery = SQLQuery & "retEndTime, "
                SQLQuery = SQLQuery & "retDayType, "
                SQLQuery = SQLQuery & "retUnused "
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & 0 & ", "
                SQLQuery = SQLQuery & llRhtCode & ", "
                SQLQuery = SQLQuery & "'" & gFixQuote(grdProgSchd.TextMatrix(llRow, PROGCODEINDEX)) & "', "
                SQLQuery = SQLQuery & "'" & Format$(grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX), sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX), sgSQLTimeForm) & "', "
                slStr = grdProgSchd.TextMatrix(llRow, DAYINDEX)
                If slStr = "M-F" Then
                    SQLQuery = SQLQuery & "'" & gFixQuote("MF") & "', "
                Else
                    SQLQuery = SQLQuery & "'" & gFixQuote(slStr) & "', "
                End If
                SQLQuery = SQLQuery & "'" & "" & "' "
                SQLQuery = SQLQuery & ") "
            Else
                'Update
                SQLQuery = "Update ret Set "
                SQLQuery = SQLQuery & "retRhtCode = " & llRhtCode & ", "
                SQLQuery = SQLQuery & "retProgCode = '" & gFixQuote(grdProgSchd.TextMatrix(llRow, PROGCODEINDEX)) & "', "
                SQLQuery = SQLQuery & "retStartTime = '" & Format$(grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX), sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "retEndTime = '" & Format$(grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX), sgSQLTimeForm) & "', "
                slStr = grdProgSchd.TextMatrix(llRow, DAYINDEX)
                If slStr = "M-F" Then
                    SQLQuery = SQLQuery & "retDayType ='" & gFixQuote("MF") & "', "
                Else
                    SQLQuery = SQLQuery & "retDayType ='" & gFixQuote(slStr) & "', "
                End If
                SQLQuery = SQLQuery & "retUnused = '" & "" & "' "
                SQLQuery = SQLQuery & " WHERE retCode = " & llRetCode
            End If
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "RadarProgSchd-mSave"
                mSave = False
                Exit Function
            End If
            If llRetCode = 0 Then
                SQLQuery = "SELECT MAX(retCode) from ret"
                Set rst_ret = gSQLSelectCall(SQLQuery)
                llRetCode = rst_ret(0).Value
                grdProgSchd.TextMatrix(llRow, RETCODEINDEX) = llRetCode
            End If
        End If
    Next llRow
    For llRow = 0 To UBound(lmDelRetCode) - 1 Step 1
        SQLQuery = "DELETE FROM RET WHERE retCode = " & lmDelRetCode(llRow)
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "RadarProgSchd-mSave"
            mSave = False
            Exit Function
        End If
    Next llRow
    lmRhtCode = llRhtCode
    imFieldChgd = False
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-mSave"
    mSave = False
End Function

Private Function mTestGridValues()
    Dim iLoop As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilStartTimeOk As Integer
    Dim ilEndTimeOk As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilError As Integer
    
    grdProgSchd.Redraw = False
    'Test if fields defined
    llRow = grdProgSpec.FixedRows
    If Trim$(grdProgSpec.TextMatrix(llRow, NETCODEINDEX)) = "" Then
        ilError = True
        grdProgSpec.TextMatrix(llRow, NETCODEINDEX) = "Missing"
        grdProgSpec.Row = llRow
        grdProgSpec.Col = NETCODEINDEX
        grdProgSpec.CellForeColor = vbRed
    End If
    If Trim$(grdProgSpec.TextMatrix(llRow, VEHCODEINDEX)) = "" Then
        ilError = True
        grdProgSpec.TextMatrix(llRow, VEHCODEINDEX) = "Missing"
        grdProgSpec.Row = llRow
        grdProgSpec.Col = VEHCODEINDEX
        grdProgSpec.CellForeColor = vbRed
    End If
    If Trim$(grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX)) = "" Then
        ilError = True
        grdProgSpec.TextMatrix(llRow, SCHDAYTYPEINDEX) = "Missing"
        grdProgSpec.Row = llRow
        grdProgSpec.Col = SCHDAYTYPEINDEX
        grdProgSpec.CellForeColor = vbRed
    End If
    If Trim$(grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX)) = "" Then
        ilError = True
        grdProgSpec.TextMatrix(llRow, CLEARTYPEINDEX) = "Missing"
        grdProgSpec.Row = llRow
        grdProgSpec.Col = CLEARTYPEINDEX
        grdProgSpec.CellForeColor = vbRed
    End If
    ilError = False
    For llRow = grdProgSchd.FixedRows To grdProgSchd.Rows - 1 Step 1
        slStr = Trim$(grdProgSchd.TextMatrix(llRow, PROGCODEINDEX))
        If slStr <> "" Then
            ilStartTimeOk = True
            slStartTime = grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX)
            If (gIsTime(slStartTime) = False) Or (Len(Trim$(slStartTime)) = 0) Then    'Time not valid.
                ilStartTimeOk = False
                ilError = True
                If Len(Trim$(slStartTime)) = 0 Then
                    grdProgSchd.TextMatrix(llRow, STARTTIMEINDEX) = "Missing"
                End If
                grdProgSchd.Row = llRow
                grdProgSchd.Col = STARTTIMEINDEX
                grdProgSchd.CellForeColor = vbRed
            End If
            ilEndTimeOk = True
            slEndTime = grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX)
            If (gIsTime(slEndTime) = False) Or (Len(Trim$(slEndTime)) = 0) Then    'Time not valid.
                ilEndTimeOk = False
                ilError = True
                If Len(Trim$(slEndTime)) = 0 Then
                    grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX) = "Missing"
                End If
                grdProgSchd.Row = llRow
                grdProgSchd.Col = ENDTIMEINDEX
                grdProgSchd.CellForeColor = vbRed
            End If
            If ilStartTimeOk And ilEndTimeOk Then
                If gTimeToLong(slEndTime, True) < gTimeToLong(slStartTime, False) Then
                    ilError = True
                    If Len(Trim$(slEndTime)) = 0 Then
                        grdProgSchd.TextMatrix(llRow, ENDTIMEINDEX) = "Missing"
                    End If
                    grdProgSchd.Row = llRow
                    grdProgSchd.Col = ENDTIMEINDEX
                    grdProgSchd.CellForeColor = vbRed
                End If
            End If
            If grdProgSpec.TextMatrix(grdProgSpec.FixedRows, SCHDAYTYPEINDEX) = "Day Name" Then
                If Trim$(grdProgSchd.TextMatrix(llRow, DAYINDEX)) = "" Then
                    ilError = True
                    grdProgSchd.TextMatrix(llRow, DAYINDEX) = "Missing"
                    grdProgSchd.Row = llRow
                    grdProgSchd.Col = DAYINDEX
                    grdProgSchd.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    If ilError Then
        grdProgSchd.Redraw = True
        mTestGridValues = False
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        mTestGridValues = True
        Exit Function
    End If
End Function


Private Function mNameOk(slNC As String, slVC As String, llRhtCode As Long) As Integer
    On Error GoTo ErrHand
    'Allow same newtork code and Vehicle code
    mNameOk = True
    'SQLQuery = "SELECT * FROM rht WHERE (rhtRadarNetCode = '" & slNC & "'" & " AND rhtRadarVehCode = '" & slVC & "'" & " AND rhtCode <> " & llRhtCode & ")"
    'Set rst_rht = gSQLSelectCall(SQLQuery)
    'If Not rst_rht.EOF Then
    '    gMsgBox "Network Code plus Vehicle Code already defined, enter either a different Network Code or Vehicle Code", vbOKOnly + vbExclamation, "Error"
    '    mNameOk = False
    'Else
    '    mNameOk = True
    'End If
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-mNameOk"
    mNameOk = False
End Function

Private Sub mPopNetVehCode()

    On Error GoTo ErrHand
    cboNetVehCode.Clear
    lmRhtCode = -1
    SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & imVefCode & ")"
    Set rst_rht = gSQLSelectCall(SQLQuery)
    Do While Not rst_rht.EOF
        cboNetVehCode.AddItem Trim$(rst_rht!rhtRadarNetCode) & "-" & Trim$(rst_rht!rhtRadarVehCode)
        cboNetVehCode.ItemData(cboNetVehCode.NewIndex) = rst_rht!rhtCode
        rst_rht.MoveNext
    Loop
    rst_rht.Close
    cboNetVehCode.AddItem "[New]", 0
    cboNetVehCode.ItemData(cboNetVehCode.NewIndex) = 0
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarProgSchd-mPopNetVehCode"
End Sub
