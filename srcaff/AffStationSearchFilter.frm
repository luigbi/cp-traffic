VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStationSearchFilter 
   Caption         =   "Station Filter"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffStationSearchFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   Begin VB.ListBox lbcAudioDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08CA
      Left            =   5970
      List            =   "AffStationSearchFilter.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2970
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcLogDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08CE
      Left            =   5550
      List            =   "AffStationSearchFilter.frx":08D0
      Sorted          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmcCancelExit 
      Caption         =   "Cancel && Exit"
      Height          =   375
      Left            =   7485
      TabIndex        =   37
      Top             =   4545
      Width           =   1395
   End
   Begin VB.ListBox lbcTimeZone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08D2
      Left            =   4980
      List            =   "AffStationSearchFilter.frx":08D4
      Sorted          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2715
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcStateLic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08D6
      Left            =   4695
      List            =   "AffStationSearchFilter.frx":08D8
      Sorted          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2595
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcServiceRep 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08DA
      Left            =   4395
      List            =   "AffStationSearchFilter.frx":08DC
      Sorted          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2565
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08DE
      Left            =   4245
      List            =   "AffStationSearchFilter.frx":08E0
      Sorted          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcMoniker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08E2
      Left            =   3990
      List            =   "AffStationSearchFilter.frx":08E4
      Sorted          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcMarketRep 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08E6
      Left            =   3675
      List            =   "AffStationSearchFilter.frx":08E8
      Sorted          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCounty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08EA
      Left            =   3435
      List            =   "AffStationSearchFilter.frx":08EC
      Sorted          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcToggle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3255
      ScaleHeight     =   150
      ScaleWidth      =   765
      TabIndex        =   29
      Top             =   1665
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ListBox lbcCity 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08EE
      Left            =   3150
      List            =   "AffStationSearchFilter.frx":08F0
      Sorted          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08F2
      Left            =   2880
      List            =   "AffStationSearchFilter.frx":08F4
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2295
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08F6
      Left            =   2625
      List            =   "AffStationSearchFilter.frx":08F8
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox edcName 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1035
      MaxLength       =   30
      TabIndex        =   2
      Top             =   495
      Width           =   2670
   End
   Begin VB.ComboBox cbcSelection 
      Height          =   315
      ItemData        =   "AffStationSearchFilter.frx":08FA
      Left            =   6195
      List            =   "AffStationSearchFilter.frx":08FC
      TabIndex        =   0
      Top             =   105
      Width           =   2670
   End
   Begin VB.CommandButton cmcDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6060
      TabIndex        =   35
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4590
      TabIndex        =   34
      Top             =   4560
      Width           =   1110
   End
   Begin VB.ListBox lbcStation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":08FE
      Left            =   2355
      List            =   "AffStationSearchFilter.frx":0900
      Sorted          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcTerritory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":0902
      Left            =   2055
      List            =   "AffStationSearchFilter.frx":0904
      Sorted          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3120
      TabIndex        =   33
      Top             =   4560
      Width           =   1110
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   30
      Top             =   4290
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   165
      Width           =   60
   End
   Begin VB.ListBox lbcMatchOn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":0906
      Left            =   1860
      List            =   "AffStationSearchFilter.frx":0908
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":090A
      Left            =   1560
      List            =   "AffStationSearchFilter.frx":090C
      Sorted          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":090E
      Left            =   1245
      List            =   "AffStationSearchFilter.frx":0910
      Sorted          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2025
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":0912
      Left            =   945
      List            =   "AffStationSearchFilter.frx":0914
      Sorted          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1995
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
      Left            =   1365
      Picture         =   "AffStationSearchFilter.frx":0916
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   255
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":0A10
      Left            =   720
      List            =   "AffStationSearchFilter.frx":0A12
      Sorted          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1935
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcDMA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":0A14
      Left            =   465
      List            =   "AffStationSearchFilter.frx":0A16
      Sorted          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcMSA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearchFilter.frx":0A18
      Left            =   285
      List            =   "AffStationSearchFilter.frx":0A1A
      Sorted          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1650
      TabIndex        =   32
      Top             =   4545
      Width           =   1110
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   36
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
      Picture         =   "AffStationSearchFilter.frx":0A1C
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
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
      Left            =   180
      TabIndex        =   31
      Top             =   4560
      Width           =   1110
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFilter 
      Height          =   3345
      Left            =   210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   990
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   5900
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacName 
      Caption         =   "Name:"
      Height          =   315
      Left            =   165
      TabIndex        =   1
      Top             =   525
      Width           =   1875
   End
End
Attribute VB_Name = "frmStationSearchFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmStationSearchFilter - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'New Items:
'  PopSubs:
'     Add item to gPopStations array tgStationInfo
'  frmStation:
'     Add item to tgStationInfo (2 places) and tgStationInfoByCode (1 place)
'  modStationSearch:
'     Add Constant Names
'     Increase tgFilterTypes array size
'  frmStationSearchFilter:
'     If List:  Add Pop call to mInit; Add Case to mEnableBox; Add Case to mSave; Add Case to mStoreFilter
'     If Toggle: Add Case to mEnableBox; Add Case to mSetToggleGridControl
'     If Edit: Add Case to mEnableBox; Maybe add Case to mStoreFilter
'  frmSerachStation:
'     Add items to tgFilterTypes in mPopFilterTypes
'     Increase tlCount array size if required in mBuildFilter
'     Add Case to mTestFilter

Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private lmRowSelected As Long
Private lmEnableRow As Long
Private lmEnableCol As Long
Private imCtrlVisible As Integer
Private imSelectIndex As Integer
Private imSelectItemData As Integer
Private imFilterTypeIndex As Integer
Private imMatchOnIndex As Integer
Private imMatchOnItemData As Integer
Private imFieldChgd As Integer
Private lmTopRow As Integer
Private imFromArrow As Integer
Private lmFhtCode As Long

Private smAllowedToggleValues() As String
Private smCurrentToggle As String

Private tmVendors() As VendorInfo

Private rst_fht As ADODB.Recordset
Private rst_fit As ADODB.Recordset
Private rst_att As ADODB.Recordset

'Grid Controls

Const SELECTINDEX = 0
Const MATCHONINDEX = 1
Const FROMVALUEINDEX = 2
Const TOVALUEINDEX = 3
Const FITCODEINDEX = 4




Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    gGrid_Clear grdFilter, True
    'Set color within cells
    'For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        'For llCol = 0 To SORTINDEX Step 1
        '    grdFilter.Row = llRow
        '    grdFilter.Col = llCol
        '    grdFilter.CellBackColor = LIGHTYELLOW
        'Next llCol
    'Next llRow
    
End Sub


Private Sub cbcSelection_Change()
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilLen As Integer
    Dim ilSel As Integer
    Dim llRow As Long

    If imFirstTime Then
        Exit Sub
    End If
    mMousePointer vbHourglass
    slName = LTrim$(cbcSelection.Text)
    ilLen = Len(slName)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slName = Left$(slName, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(cbcSelection.hwnd, CB_FINDSTRING, -1, slName)
    If llRow >= 0 Then
        On Error GoTo ErrHand
        cbcSelection.ListIndex = llRow
        cbcSelection.SelStart = ilLen
        cbcSelection.SelLength = Len(cbcSelection.Text)
        If cbcSelection.ListIndex < 0 Then
            lmFhtCode = -1
        'ElseIf (cbcSelection.ListIndex = 0) Then
        '    lmFhtCode = -1
        'ElseIf (cbcSelection.ListIndex = 1) Then
        ElseIf (cbcSelection.ListIndex = 0) Then
            lmFhtCode = 0
        Else
            lmFhtCode = CLng(cbcSelection.ItemData(cbcSelection.ListIndex))
        End If
        If lmFhtCode < 0 Then
            mClearControls
        ElseIf (lmFhtCode = 0) Then
            mClearControls
        Else
            mBindControls
        End If
    End If
    imFieldChgd = False
    mSetCommands
    mMousePointer vbDefault
    Exit Sub

ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearchFilter-cbcSelection_Change"
End Sub

Private Sub cbcSelection_Click()
    cbcSelection_Change
End Sub

Private Sub cmcCancelExit_Click()
    igFilterReturn = False
    ReDim tgFilterDef(0 To 0) As FILTERDEF
    ReDim tgNotFilterDef(0 To 0) As FILTERDEF
    Unload frmStationSearchFilter
    Exit Sub
End Sub

Private Sub cmcCancelExit_GotFocus()
    mSetShow
End Sub

Private Sub cmcDelete_Click()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If lmFhtCode <= 0 Then
        Exit Sub
    End If
    ilRet = MsgBox("This will permanently remove " & edcName.Text & ", are you sure", vbYesNo + vbQuestion, "Remove")
    If ilRet = vbYes Then
        cnn.BeginTrans
        SQLQuery = "DELETE FROM fit"
        SQLQuery = SQLQuery & " WHERE (fitfhtCode = " & lmFhtCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationSearchFilter-cmcDelete_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        SQLQuery = "DELETE FROM fht"
        SQLQuery = SQLQuery & " WHERE (fhtCode = " & lmFhtCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationSearchFilter-cmcDelete_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        mPopulate
    End If
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearchFilter-cmcDelete"
End Sub

Private Sub cmcDelete_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    Select Case lmEnableCol
        Case SELECTINDEX
            'lbcSelect.Visible = Not lbcSelect.Visible
            lbcDropdown.Visible = Not lbcDropdown.Visible
        Case MATCHONINDEX
            'lbcMatchOn.Visible = Not lbcMatchOn.Visible
            lbcDropdown.Visible = Not lbcDropdown.Visible
        Case FROMVALUEINDEX
            'Select Case imSelectItemData
            '    Case SFAREA
            '        lbcArea.Visible = Not lbcArea.Visible
            '    Case SFDMA  'DMA
            '        lbcDMA.Visible = Not lbcDMA.Visible
            '    Case SFFORMAT  'Format
            '        lbcFormat.Visible = Not lbcFormat.Visible
            '    Case SFMSA  'MSA
            '        lbcMSA.Visible = Not lbcMSA.Visible
            '    Case SFOWNER  'Owner
            '        lbcOwner.Visible = Not lbcOwner.Visible
            '    Case SFVEHICLE  'Vehicle
            '        lbcVehicle.Visible = Not lbcVehicle.Visible
            '    Case SFZIP  'Zip
            '    Case SFTERRITORY 'Territory
            '        lbcTerritory.Visible = Not lbcTerritory.Visible
            '    Case SFCALLLETTERS  'Station
            '        lbcStation.Visible = Not lbcStation.Visible
            'End Select
            lbcDropdown.Visible = Not lbcDropdown.Visible
    End Select
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilRow As Integer
    
    mMousePointer vbHourglass
    slName = UCase$(edcName.Text)
    ilRet = mSave()
    mPopulate
    For ilRow = 0 To cbcSelection.ListCount - 1 Step 1
        If StrComp(slName, UCase$(cbcSelection.List(ilRow)), vbBinaryCompare) = 0 Then
            cbcSelection.ListIndex = ilRow
            lmFhtCode = cbcSelection.ItemData(ilRow)
            imFieldChgd = False
            Exit For
        End If
    Next ilRow
    mMousePointer vbDefault
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub


Private Sub cmdClear_Click()
    Dim ilRet As Integer
    Dim llEnableRow As Long
    
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        'ilRet = MsgBox("Clear Row " & llEnableRow, vbYesNo + vbQuestion, "Clear")
        'If ilRet = vbYes Then
        '    imFieldChgd = True
        '    grdFilter.RemoveItem llEnableRow
        '    mSetShow
        'End If
        sgGenMsg = "Clear All Rows or Only Clear Row " & llEnableRow
        sgCMCTitle(0) = "Clear All"
        sgCMCTitle(1) = "Clear Row"
        sgCMCTitle(2) = "Cancel"
        sgCMCTitle(3) = ""
        igDefCMC = 1
        igEditBox = 0
        sgEditValue = ""
        frmGenMsg.Show vbModal
        If igAnsCMC = 0 Then
            mSetShow
            mClearGrid
            cbcSelection.ListIndex = 0
            imFieldChgd = False
        ElseIf igAnsCMC = 1 Then
            imFieldChgd = True
            grdFilter.RemoveItem llEnableRow
            mSetShow
        End If
    Else
        ilRet = MsgBox("Clear All Rows", vbYesNo + vbQuestion, "Clear")
        If ilRet = vbYes Then
            mSetShow
            mClearGrid
            cbcSelection.ListIndex = 0
            imFieldChgd = False
        End If
    End If
End Sub

Private Sub cmdDone_Click()
    Dim blFilterDefined As Boolean
    Dim bmItemsDefined As Boolean
    Dim slStr As String
    Dim llRow As Long
    Dim llCol As Long
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim blServiceAgreementFilter As Boolean
    
    mMousePointer vbHourglass
    igFilterReturn = True
    blFilterDefined = False
    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        If (Trim$(grdFilter.TextMatrix(llRow, SELECTINDEX)) <> "") Then
            bmItemsDefined = True
            If (Trim$(grdFilter.TextMatrix(llRow, MATCHONINDEX)) = "") Then
                bmItemsDefined = False
            End If
            'If (Trim$(grdFilter.TextMatrix(llRow, FROMVALUEINDEX)) = "") And (Trim$(grdFilter.TextMatrix(llRow, TOVALUEINDEX)) = "") Then
            '    If (Trim$(grdFilter.TextMatrix(llRow, MATCHONINDEX)) = "Equal") Then
            '        slStr = grdFilter.TextMatrix(llRow, SELECTINDEX)
            '        imSelectItemData = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
            '        If imSelectItemData >= 0 Then
            '            If tgFilterTypes(imSelectItemData).sCntrlType <> "E" Then
            '                bmItemsDefined = False
            '            Else
            '                If (tgFilterTypes(imSelectItemData).iSelect = SFCALLLETTERSCHGDATE) Or (tgFilterTypes(imSelectItemData).iSelect = SFZIP) Then
            '                    bmItemsDefined = False
            '                End If
            '            End If
            '        Else
            '            bmItemsDefined = False
            '        End If
            '    Else
            '        bmItemsDefined = False
            '    End If
            'End If
            If bmItemsDefined Then
                blFilterDefined = True
                Exit For
            End If
        End If
    Next llRow
    If blFilterDefined Then
    
        For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
            If (Trim$(grdFilter.TextMatrix(llRow, SELECTINDEX)) <> "") Then
                'For llCol = MATCHONINDEX To TOVALUEINDEX Step 1
                    'If (Trim$(grdFilter.TextMatrix(llRow, llCol)) = "") Then
                    '    mMousePointer vbDefault
                    '    MsgBox "Not all Columns are filled in on Row " & llRow & ", please Fix", vbCritical + vbOKOnly, "Fix"
                    '    Exit Sub
                    'End If
                    If (Trim$(grdFilter.TextMatrix(llRow, MATCHONINDEX)) = "") Then
                        mMousePointer vbDefault
                        MsgBox "Not all Columns are filled in on Row " & llRow & ", please Fix", vbCritical + vbOKOnly, "Fix"
                        Exit Sub
                    End If
                    If (Trim$(grdFilter.TextMatrix(llRow, FROMVALUEINDEX)) = "") And (Trim$(grdFilter.TextMatrix(llRow, TOVALUEINDEX)) = "") Then
                        bmItemsDefined = True
                        If (Trim$(grdFilter.TextMatrix(llRow, MATCHONINDEX)) = "Equal") Then
                            slStr = grdFilter.TextMatrix(llRow, SELECTINDEX)
                            imSelectIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
                            If imSelectIndex >= 0 Then
                                imFilterTypeIndex = lbcSelect.ItemData(imSelectIndex) - 1
                                If tgFilterTypes(imFilterTypeIndex).sCntrlType <> "E" Then
                                    bmItemsDefined = False
                                Else
                                    'If (tgFilterTypes(imFilterTypeIndex).iSelect = SFCALLLETTERSCHGDATE) Or (tgFilterTypes(imFilterTypeIndex).iSelect = SFPHONE) Or (tgFilterTypes(imFilterTypeIndex).iSelect = SFZIP) Then
                                    If (tgFilterTypes(imFilterTypeIndex).iSelect = SFCALLLETTERSCHGDATE) Then
                                        bmItemsDefined = False
                                    End If
                                End If
                            Else
                                bmItemsDefined = False
                            End If
                        Else
                            bmItemsDefined = False
                        End If
                        If Not bmItemsDefined Then
                            mMousePointer vbDefault
                            MsgBox "Not all Columns are filled in on Row " & llRow & ", please Fix", vbCritical + vbOKOnly, "Fix"
                            Exit Sub
                        End If
                    End If
                'Next llCol
            End If
        Next llRow
        '12/10/10: Jim request that the question be removed
        'If (lmFhtCode >= 0) And (imFieldChgd) Then
        '    ilRet = MsgBox("Save Filter", vbYesNo + vbQuestion, "Filter")
        '    If ilRet = vbYes Then
        '        ilRet = mSave()
        '        If Not ilRet Then
        '            mMousePointer vbDefault
        '            Exit Sub
        '        End If
        '    End If
        'End If
        lgFhtCode = lmFhtCode
        sgFilterName = Trim$(edcName.Text)
        igFilterChgd = imFieldChgd
        ReDim tgFilterDef(0 To 0) As FILTERDEF
        ReDim tgNotFilterDef(0 To 0) As FILTERDEF
        For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
            If (Trim$(grdFilter.TextMatrix(llRow, SELECTINDEX)) <> "") Then
                slStr = grdFilter.TextMatrix(llRow, SELECTINDEX)
                imSelectIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
                If imSelectIndex >= 0 Then
                    imFilterTypeIndex = lbcSelect.ItemData(imSelectIndex) - 1
                    mPopMatchOn tgFilterTypes(imFilterTypeIndex)
                    slStr = grdFilter.TextMatrix(llRow, MATCHONINDEX)
                    imMatchOnIndex = SendMessageByString(lbcMatchOn.hwnd, LB_FINDSTRING, -1, slStr)
                    If imMatchOnIndex >= 0 Then
                        'If (lbcMatchOn.ItemData(imMatchOnItemData) = 4) Then
                        '    mStoreFilter llRow, tgNotFilterDef()
                        'Else
                            mStoreFilter llRow, tgFilterTypes(imFilterTypeIndex), tgFilterDef()
                        'End If
                    End If
                End If
            End If
        Next llRow
    Else
        ReDim tgFilterDef(0 To 0) As FILTERDEF
        ReDim tgNotFilterDef(0 To 0) As FILTERDEF
    End If
    If sgUsingServiceAgreement = "Y" Then
        'Add Service Agreement filter as None if not already defined
        blServiceAgreementFilter = False
        For llRow = 0 To UBound(tgFilterDef) Step 1
            If tgFilterDef(llRow).iSelect = SFSERVICEAGREEMENT Then
                blServiceAgreementFilter = True
                Exit For
            End If
        Next llRow
        If Not blServiceAgreementFilter Then
            'bgServiceAgreementExist = False
            ''Test if any service agreements exist
            'SQLQuery = "SELECT Count(1) as SACount FROM att"
            'SQLQuery = SQLQuery + " WHERE ("
            'SQLQuery = SQLQuery & " attServiceAgreement = '" & "Y" & "')"
            'Set rst_att = gSQLSelectCall(SQLQuery)
            'If Not rst_att.EOF Then
            '    If rst_att!SACount > 0 Then
            '        bgServiceAgreementExist = True
            '        ilUpper = UBound(tgFilterDef)
            '        tgFilterDef(ilUpper).iSelect = SFSERVICEAGREEMENT
            '        tgFilterDef(ilUpper).iOperator = 1  'equal
            '        tgFilterDef(ilUpper).iCountGroup = 0
            '        tgFilterDef(ilUpper).sCntrlType = "T"
            '        tgFilterDef(ilUpper).sFromValue = "NONE"
            '        tgFilterDef(ilUpper).lFromValue = 0
            '        tgFilterDef(ilUpper).sToValue = ""
            '        tgFilterDef(ilUpper).lToValue = 0
            '        tgFilterDef(ilUpper).iFirstFilterLink = -1
            '        ReDim Preserve tgFilterDef(0 To ilUpper + 1) As FILTERDEF
            '    End If
            'End If
        Else
            bgServiceAgreementExist = True
        End If
    Else
        bgServiceAgreementExist = False
    End If
    'lgFhtCode = lmFhtCode
    'sgFilterName = Trim$(edcName.Text)
    mMousePointer vbDefault
    Unload frmStationSearchFilter
    Exit Sub
   
End Sub






Private Sub cmdDone_GotFocus()
    mSetShow
End Sub

Private Sub cmdUndo_Click()
    If lmFhtCode > 0 Then
        mMousePointer vbHourglass
        mBindControls
        imFieldChgd = False
        mSetCommands
        mMousePointer vbDefault
    End If
End Sub

Private Sub cmdUndo_GotFocus()
    mSetShow
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    Select Case lmEnableCol
        Case SELECTINDEX
            'mDropdownChangeEvent lbcSelect
            mDropdownChangeEvent lbcDropdown
        Case MATCHONINDEX
            'mDropdownChangeEvent lbcMatchOn
            mDropdownChangeEvent lbcDropdown
        Case FROMVALUEINDEX
            If tgFilterTypes(imFilterTypeIndex).sCntrlType = "L" Then
                mDropdownChangeEvent lbcDropdown
            End If
    End Select
    mSetCommands
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
        Select Case lmEnableCol
            Case SELECTINDEX
                gProcessArrowKey Shift, KeyCode, lbcSelect, True
            Case MATCHONINDEX
                gProcessArrowKey Shift, KeyCode, lbcMatchOn, True
            Case FROMVALUEINDEX
                If tgFilterTypes(imFilterTypeIndex).sCntrlType = "L" Then
                    gProcessArrowKey Shift, KeyCode, lbcDropdown, True
                End If
        End Select
    End If
End Sub

Private Sub edcName_Change()
    mSetCommands
End Sub

Private Sub edcName_GotFocus()
    mSetShow
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    If imFirstTime Then
        Me.Width = Screen.Width / 1.55
        Me.Height = Screen.Height / 1.55
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
        gSetFonts frmStationSearchFilter
        gCenterForm frmStationSearchFilter
        
        mMousePointer vbHourglass
        mSetGridColumns
        mSetGridTitles
        gGrid_IntegralHeight grdFilter
        gGrid_FillWithRows grdFilter
        mPopulate
        If UBound(tgFilterDef) > LBound(tgFilterDef) Then
            If lgFhtCode > 0 Then
                For ilRow = 0 To cbcSelection.ListCount - 1 Step 1
                    If cbcSelection.ItemData(ilRow) = lgFhtCode Then
                        cbcSelection.ListIndex = ilRow
                        lmFhtCode = cbcSelection.ItemData(ilRow)
                        Exit For
                    End If
                Next ilRow
            ElseIf lgFhtCode = 0 Then
                'cbcSelection.ListIndex = 1
                cbcSelection.ListIndex = 0
                lmFhtCode = 0
            'ElseIf lgFhtCode < 0 Then
            '    cbcSelection.ListIndex = 0
            '    lmFhtCode = -1
            End If
            edcName.Text = sgFilterName
            imFieldChgd = igFilterChgd
            mPopulateGrid
        Else
            lgFhtCode = -1
            sgFilterName = ""
            igFilterChgd = False
            cbcSelection.ListIndex = 0
            'lmFhtCode = -1
            lmFhtCode = 0
            edcName.Text = ""
        End If
        imFirstTime = False
        cbcSelection.SetFocus
        mSetCommands
        
        Me.Visible = True
        mMousePointer vbDefault
    End If
    
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.55
'    Me.Height = Screen.Height / 1.55
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
'    gSetFonts frmStationSearchFilter
'    gCenterForm frmStationSearchFilter
End Sub

Private Sub Form_Load()
    Me.Visible = False
    mMousePointer vbHourglass
    
    mInit
    mMousePointer vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase smAllowedToggleValues
    rst_fht.Close
    rst_fit.Close
    rst_att.Close
    Set frmStationSearchFilter = Nothing
End Sub

Private Sub grdFilter_Click()
    Dim llRow As Long
    
'    If grdFilter.Row >= grdFilter.FixedRows Then
'        If grdFilter.TextMatrix(grdFilter.Row, GAMENOINDEX) <> "" Then
'            If (lmRowSelected = grdFilter.Row) Then
'                If imCtrlKey Then
'                    lmRowSelected = -1
'                    grdFilter.Row = 0
'                    grdFilter.Col = GSFFITCODEINDEX
'                End If
'            Else
'                lmRowSelected = grdFilter.Row
'            End If
'        Else
'            lmRowSelected = -1
'            grdFilter.Row = 0
'            grdFilter.Col = GSFFITCODEINDEX
'        End If
'    End If

End Sub

Private Sub grdFilter_EnterCell()
    mSetShow
End Sub

Private Sub grdFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    imSelectIndex = -1
    imSelectItemData = -1
    imFilterTypeIndex = -1
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    imFieldChgd = False
    imFromArrow = False
    
    mPopSelect
    mPopVehicle
    mPopDMA
    mPopMSA
    mPopOwner
    mPopFormat
    mPopTerritory
    mPopArea
    mPopCity
    mPopCounty
    mPopStation
    mPopMarketRep
    mPopMoniker
    mPopOperator
    mPopServiceRep
    mPopStateLic
    mPopTimeZone
    tmVendors = gGetAvailableVendors()
    mPopLogDelivery
    mPopAudioDelivery
    
    pbcSTab.Left = -100
    pbcTab.Left = -100
    pbcClickFocus.Left = -100
    
    mClearGrid
    grdFilter.Visible = True
End Sub

Private Sub mPopulateGrid()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFilter As Integer
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand:
    grdFilter.Redraw = False
    grdFilter.Row = 0
    llRow = grdFilter.FixedRows
    For ilFilter = 0 To UBound(tgFilterDef) - 1 Step 1
        If llRow >= grdFilter.Rows Then
            grdFilter.AddItem ""
        End If
        'Select
        For ilLoop = 0 To lbcSelect.ListCount - 1 Step 1
            If tgFilterDef(ilFilter).iSelect = lbcSelect.ItemData(ilLoop) Then
                lbcSelect.ListIndex = ilLoop
                grdFilter.TextMatrix(llRow, SELECTINDEX) = lbcSelect.List(ilLoop)
                mPopMatchOn tgFilterTypes(lbcSelect.ItemData(ilLoop) - 1)
                Exit For
            End If
        Next ilLoop
        'MatchOn
        For ilLoop = 0 To lbcMatchOn.ListCount - 1 Step 1
            If tgFilterDef(ilFilter).iOperator = lbcMatchOn.ItemData(ilLoop) Then
                lbcMatchOn.ListIndex = ilLoop
                grdFilter.TextMatrix(llRow, MATCHONINDEX) = lbcMatchOn.List(ilLoop)
                Exit For
            End If
        Next ilLoop
        grdFilter.TextMatrix(llRow, FROMVALUEINDEX) = Trim$(tgFilterDef(ilFilter).sFromValue)
        If tgFilterDef(ilFilter).iOperator = 3 Then
            grdFilter.TextMatrix(llRow, TOVALUEINDEX) = Trim$(tgFilterDef(ilFilter).sToValue)
        Else
            grdFilter.Row = llRow
            grdFilter.Col = TOVALUEINDEX
            grdFilter.CellBackColor = LIGHTYELLOW
            grdFilter.TextMatrix(llRow, TOVALUEINDEX) = ""
        End If
        llRow = llRow + 1
    Next ilFilter
    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        If grdFilter.TextMatrix(llRow, SELECTINDEX) <> "" Then
            If grdFilter.TextMatrix(llRow, MATCHONINDEX) <> "Range" Then
                grdFilter.Row = llRow
                grdFilter.Col = TOVALUEINDEX
                grdFilter.CellBackColor = LIGHTYELLOW
                grdFilter.TextMatrix(llRow, TOVALUEINDEX) = ""
            End If
        End If
    Next llRow
    grdFilter.Redraw = True
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearchFilter-mPopulateGrid"
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdFilter.ColWidth(FITCODEINDEX) = 0
    grdFilter.ColWidth(SELECTINDEX) = grdFilter.Width * 0.3
    grdFilter.ColWidth(MATCHONINDEX) = grdFilter.Width * 0.18

    grdFilter.ColWidth(TOVALUEINDEX) = (grdFilter.Width - grdFilter.ColWidth(SELECTINDEX) - grdFilter.ColWidth(MATCHONINDEX) - GRIDSCROLLWIDTH - 15) / 2
    grdFilter.ColWidth(FROMVALUEINDEX) = grdFilter.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To TOVALUEINDEX Step 1
        If ilCol <> FROMVALUEINDEX Then
            grdFilter.ColWidth(FROMVALUEINDEX) = grdFilter.ColWidth(FROMVALUEINDEX) - grdFilter.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdFilter
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdFilter.TextMatrix(0, SELECTINDEX) = "Select"
    grdFilter.TextMatrix(0, MATCHONINDEX) = "Command"
    grdFilter.TextMatrix(0, FROMVALUEINDEX) = "From Value"
    grdFilter.TextMatrix(0, TOVALUEINDEX) = "To Value"

End Sub

Private Sub mFilterSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
'    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
'        slStr = Trim$(grdFilter.TextMatrix(llRow, GAMENOINDEX))
'        If slStr <> "" Then
'            If ilCol = AIRDATEINDEX Then
'                slSort = Trim$(Str$(DateValue(grdFilter.TextMatrix(llRow, AIRDATEINDEX))))
'                Do While Len(slSort) < 6
'                    slSort = "0" & slSort
'                Loop
'            ElseIf (ilCol = AIRTIMEINDEX) Then
'                slSort = Trim$(Str$(gTimeToLong(grdFilter.TextMatrix(llRow, AIRTIMEINDEX), False)))
'                Do While Len(slSort) < 6
'                    slSort = "0" & slSort
'                Loop
'            ElseIf (ilCol = GAMENOINDEX) Then
'                slSort = Trim$(grdFilter.TextMatrix(llRow, GAMENOINDEX))
'                Do While Len(slSort) < 8
'                    slSort = "0" & slSort
'                Loop
'            Else
'                slSort = UCase$(Trim$(grdFilter.TextMatrix(llRow, ilCol)))
'            End If
'            slStr = grdFilter.TextMatrix(llRow, SORTINDEX)
'            ilPos = InStr(1, slStr, "|", vbTextCompare)
'            If ilPos > 1 Then
'                slStr = Left$(slStr, ilPos - 1)
'            End If
'            If (ilCol <> imLastGameColSorted) Or ((ilCol = imLastGameColSorted) And (imLastGameSort = flexSortStringNoCaseDescending)) Then
'                slRow = Trim$(Str$(llRow))
'                Do While Len(slRow) < 4
'                    slRow = "0" & slRow
'                Loop
'                grdFilter.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
'            Else
'                slRow = Trim$(Str$(llRow))
'                Do While Len(slRow) < 4
'                    slRow = "0" & slRow
'                Loop
'                grdFilter.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
'            End If
'        End If
'    Next llRow
'    If ilCol = imLastGameColSorted Then
'        imLastGameColSorted = SORTINDEX
'    Else
'        imLastGameColSorted = -1
'        imLastGameSort = -1
'    End If
'    gGrid_SortByCol grdFilter, GAMENOINDEX, SORTINDEX, imLastGameColSorted, imLastGameSort
'    imLastGameColSorted = ilCol
End Sub

Private Sub grdFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    Dim slStr As String
    
    mSetShow
    If Y < grdFilter.RowHeight(0) Then
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdFilter, X, Y)
    If Not ilFound Then
        grdFilter.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdFilter.Col >= grdFilter.Cols - 1 Then
        grdFilter.Redraw = True
        Exit Sub
    End If
    If grdFilter.Col = TOVALUEINDEX Then
        If imSelectIndex < 0 Then
            grdFilter.Redraw = True
            Exit Sub
        End If
        slStr = grdFilter.TextMatrix(grdFilter.Row, MATCHONINDEX)
        imMatchOnIndex = SendMessageByString(lbcMatchOn.hwnd, LB_FINDSTRING, -1, slStr)
        If imMatchOnIndex < 0 Then
            grdFilter.Redraw = True
            Exit Sub
        End If
        If (lbcMatchOn.ItemData(imMatchOnIndex) <> 3) Then
            grdFilter.Redraw = True
            Exit Sub
        End If
    End If
    lmTopRow = grdFilter.TopRow
    
    llRow = grdFilter.Row
    If Trim(grdFilter.TextMatrix(llRow, SELECTINDEX)) = "" Then
        grdFilter.Redraw = False
        Do
            llRow = llRow - 1
            If llRow < grdFilter.FixedRows Then
                Exit Do
            End If
        Loop While Trim(grdFilter.TextMatrix(llRow, SELECTINDEX)) = ""
        grdFilter.Row = llRow + 1
        grdFilter.Col = SELECTINDEX
        grdFilter.Redraw = True
    End If
    grdFilter.Redraw = True
    mEnableBox

End Sub

Private Sub mEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilType As Integer
    
    If (grdFilter.Row >= grdFilter.FixedRows) And (grdFilter.Row < grdFilter.Rows) And (grdFilter.Col >= SELECTINDEX) And (grdFilter.Col < grdFilter.Cols - 1) Then
        lmEnableRow = grdFilter.Row
        lmEnableCol = grdFilter.Col
        imCtrlVisible = True
        pbcArrow.Move grdFilter.Left - pbcArrow.Width, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdFilter.Col
            Case SELECTINDEX
                If (lmEnableRow > grdFilter.FixedRows) And (Trim(grdFilter.Text) = "") Then
                    grdFilter.Text = grdFilter.TextMatrix(lmEnableRow - 1, SELECTINDEX)
                End If
                mSetLbcGridControl lbcSelect
            Case MATCHONINDEX
                slStr = grdFilter.TextMatrix(lmEnableRow, SELECTINDEX)
                imSelectIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
                If imSelectIndex >= 0 Then
                    imFilterTypeIndex = lbcSelect.ItemData(imSelectIndex) - 1
                    mPopMatchOn tgFilterTypes(imFilterTypeIndex)
                    If (lmEnableRow > grdFilter.FixedRows) And (Trim(grdFilter.Text) = "") And (grdFilter.TextMatrix(lmEnableRow, SELECTINDEX) = grdFilter.TextMatrix(lmEnableRow - 1, SELECTINDEX)) Then
                        grdFilter.Text = grdFilter.TextMatrix(lmEnableRow - 1, MATCHONINDEX)
                    End If
                End If
                mSetLbcGridControl lbcMatchOn
            Case FROMVALUEINDEX, TOVALUEINDEX
                slStr = grdFilter.TextMatrix(lmEnableRow, SELECTINDEX)
                imSelectIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
                If imSelectIndex >= 0 Then
                    imSelectItemData = lbcSelect.ItemData(imSelectIndex)
                    imFilterTypeIndex = lbcSelect.ItemData(imSelectIndex) - 1
                    mPopMatchOn tgFilterTypes(imFilterTypeIndex)
                Else
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                slStr = grdFilter.TextMatrix(lmEnableRow, MATCHONINDEX)
                imMatchOnIndex = SendMessageByString(lbcMatchOn.hwnd, LB_FINDSTRING, -1, slStr)
                If imMatchOnIndex >= 0 Then
                    imMatchOnItemData = lbcMatchOn.ItemData(imMatchOnIndex)
                    If imMatchOnItemData = 0 Then
                        mSetEdcGridControl
                    Else
                        'slStr = grdFilter.TextMatrix(lmEnableRow, SELECTINDEX)
                        'ilIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
                        'If ilIndex >= 0 Then
                        '    imSelectItemData = lbcSelect.ItemData(ilIndex)
                            Select Case imSelectItemData
                                Case SFAREA
                                    mSetLbcGridControl lbcArea
                                Case SFCALLLETTERSCHGDATE
                                    mSetEdcGridControl
                                Case SFCALLLETTERS  'Station
                                    mSetLbcGridControl lbcStation
                                Case SFCITYLIC
                                    mSetLbcGridControl lbcCity
                                Case SFCOMMERCIAL
                                    mSetToggleGridControl
                                Case SFCOUNTYLIC
                                    mSetLbcGridControl lbcCounty
                                Case SFDAYLIGHT
                                    mSetToggleGridControl
                                Case SFDMA  'DMA
                                    mSetLbcGridControl lbcDMA
                                Case SFDMARANK
                                    mSetEdcGridControl
                                Case SFFORMAT  'Format
                                    mSetLbcGridControl lbcFormat
                                Case SFFREQ
                                    mSetEdcGridControl
                                Case SFHISTSTARTDATE
                                    mSetEdcGridControl
                                Case SFPERMID
                                    mSetEdcGridControl
                                Case SFMAILADDRESS
                                    mSetEdcGridControl
                                Case SFMARKETREP
                                    mSetLbcGridControl lbcMarketRep
                                Case SFMONIKER
                                    mSetLbcGridControl lbcMoniker
                                Case SFMSA  'MSA
                                    mSetLbcGridControl lbcMSA
                                Case SFMSARANK
                                    mSetEdcGridControl
                                Case SFONAIR
                                    mSetToggleGridControl
                                Case SFOPERATOR
                                    mSetLbcGridControl lbcOperator
                                Case SFOWNER  'Owner
                                    mSetLbcGridControl lbcOwner
                                Case SFP12PLUS
                                    mSetEdcGridControl
                                Case SFPHONE
                                    mSetEdcGridControl
                                Case SFPHYADDRESS
                                    mSetEdcGridControl
                                Case SFSERIAL
                                    mSetEdcGridControl
                                Case SFSERVICEREP
                                    mSetLbcGridControl lbcServiceRep
                                Case SFSTATELIC
                                    mSetLbcGridControl lbcStateLic
                                Case SFEMAIL
                                    mSetToggleGridControl
                                Case SFISCI
                                    mSetToggleGridControl
                                Case SFLABEL
                                    mSetToggleGridControl
                                Case SFPERSONNEL
                                    mSetEdcGridControl
                                Case SFAGREEMENT
                                    mSetToggleGridControl
                                Case SFWEGENER
                                    mSetToggleGridControl
                                Case SFXDS
                                    mSetToggleGridControl
                                Case SFZONE
                                    mSetLbcGridControl lbcTimeZone
                                Case SFENTERPRISEID
                                    mSetEdcGridControl
                                Case SFVEHICLEACTIVE  'Vehicle
                                    mSetLbcGridControl lbcVehicle
                                Case SFVEHICLEALL  'Vehicle
                                    mSetLbcGridControl lbcVehicle
                                Case SFWEBADDRESS
                                    mSetEdcGridControl
                                Case SFWEBPW
                                    mSetEdcGridControl
                                Case SFXDSID
                                    mSetEdcGridControl
                                Case SFZIP  'Zip
                                    mSetEdcGridControl
                                Case SFTERRITORY  'Territory
                                    mSetLbcGridControl lbcTerritory
                                Case SFMULTICAST
                                    mSetToggleGridControl
                                Case SFSISTER
                                    mSetToggleGridControl
                                Case SFWATTS
                                    mSetEdcGridControl
                                '6048
                                Case SFEMAILADDRESS
                                    mSetEdcGridControl
                                Case SFDUE
                                    mSetEdcGridControl
                                '5/6/18: Enable ListBox
                                Case SFLOGDELIVERY
                                    mSetLbcGridControl lbcLogDelivery
                                Case SFAUDIODELIVERY
                                    mSetLbcGridControl lbcAudioDelivery
                                Case SFSERVICEAGREEMENT
                                    mSetToggleGridControl
                                    
                            End Select
                        'Else
                        '    imSelectItemData = -1
                        '    pbcClickFocus.SetFocus
                        'End If
                    End If
                Else
                    imSelectIndex = -1
                    imSelectItemData = -1
                    pbcClickFocus.SetFocus
                End If
        End Select
    End If
End Sub


Private Sub mSetShow()
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilType As Integer
    Dim llSvRow As Long
    Dim llSvCol As Long
    
    If (lmEnableRow >= grdFilter.FixedRows) And (lmEnableRow < grdFilter.Rows) Then
        llSvRow = grdFilter.Row
        llSvCol = grdFilter.Col
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case SELECTINDEX
                slStr = edcDropdown.Text
                If grdFilter.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdFilter.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    grdFilter.TextMatrix(lmEnableRow, MATCHONINDEX) = ""
                    grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = ""
                End If
                slStr = grdFilter.TextMatrix(lmEnableRow, SELECTINDEX)
                imSelectIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
                If imSelectIndex >= 0 Then
                    imFilterTypeIndex = lbcSelect.ItemData(imSelectIndex) - 1
                Else
                    imFilterTypeIndex = -1
                End If
            Case MATCHONINDEX
                slStr = edcDropdown.Text
                If grdFilter.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdFilter.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    grdFilter.Row = lmEnableRow
                    grdFilter.Col = lmEnableCol
                    If (Trim$(grdFilter.TextMatrix(lmEnableRow, MATCHONINDEX)) <> "Range") Then
                        grdFilter.TextMatrix(lmEnableRow, TOVALUEINDEX) = ""
                        grdFilter.Col = TOVALUEINDEX
                        grdFilter.CellBackColor = LIGHTYELLOW
                    Else
                        grdFilter.Col = TOVALUEINDEX
                        grdFilter.CellBackColor = vbWhite
                    End If
                    grdFilter.Row = llSvRow
                    grdFilter.Col = llSvCol
                End If
            Case FROMVALUEINDEX
                If pbcToggle.Visible Then
                    slStr = smCurrentToggle
                Else
                    slStr = edcDropdown.Text
                End If
                If grdFilter.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdFilter.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case TOVALUEINDEX
                slStr = edcDropdown.Text
                If grdFilter.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdFilter.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    pbcArrow.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcDropdown.Visible = False
    lbcSelect.Visible = False
    lbcMatchOn.Visible = False
    pbcToggle.Visible = False
    'lbcDMA.Visible = False
    'lbcFormat.Visible = False
    'lbcMSA.Visible = False
    'lbcOwner.Visible = False
    'lbcVehicle.Visible = False
    'lbcTerritory.Visible = False
    'lbcStation.Visible = False
    'lbcArea.Visible = False
    mSetCommands
End Sub


Private Sub mPopVehicle()
    Dim ilLoop As Integer
    lbcVehicle.Clear
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicle.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgVehicleInfo(ilLoop).iCode
    Next ilLoop

End Sub

Private Sub mPopTerritory()
    Dim ilLoop As Integer
    lbcTerritory.Clear
    For ilLoop = 0 To UBound(tgTerritoryInfo) - 1 Step 1
        lbcTerritory.AddItem Trim$(tgTerritoryInfo(ilLoop).sName)
        lbcTerritory.ItemData(lbcTerritory.NewIndex) = tgTerritoryInfo(ilLoop).lCode
    Next ilLoop
    lbcTerritory.AddItem "[Defined]", 0
    lbcTerritory.ItemData(lbcTerritory.NewIndex) = -1
End Sub

Private Sub mPopDMA()
    Dim ilLoop As Integer
    lbcDMA.Clear
    For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        lbcDMA.AddItem Trim$(tgMarketInfo(ilLoop).sName)
        lbcDMA.ItemData(lbcDMA.NewIndex) = tgMarketInfo(ilLoop).lCode
    Next ilLoop
    lbcDMA.AddItem "[Defined]", 0
    lbcDMA.ItemData(lbcDMA.NewIndex) = -1

End Sub

Private Sub mPopMSA()
    Dim ilLoop As Integer
    lbcMSA.Clear
    For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        lbcMSA.AddItem Trim$(tgMSAMarketInfo(ilLoop).sName)
        lbcMSA.ItemData(lbcMSA.NewIndex) = tgMSAMarketInfo(ilLoop).lCode
    Next ilLoop
    lbcMSA.AddItem "[Defined]", 0
    lbcMSA.ItemData(lbcMSA.NewIndex) = -1

End Sub

Private Sub mPopOwner()
    Dim ilLoop As Integer
    lbcOwner.Clear
    For ilLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
        lbcOwner.AddItem Trim$(tgOwnerInfo(ilLoop).sName)
        lbcOwner.ItemData(lbcOwner.NewIndex) = tgOwnerInfo(ilLoop).lCode
    Next ilLoop
    lbcOwner.AddItem "[Defined]", 0
    lbcOwner.ItemData(lbcOwner.NewIndex) = -1

End Sub

Private Sub mPopFormat()
    Dim ilLoop As Integer
    lbcFormat.Clear
    For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        lbcFormat.AddItem Trim$(tgFormatInfo(ilLoop).sName)
        lbcFormat.ItemData(lbcFormat.NewIndex) = tgFormatInfo(ilLoop).lCode
    Next ilLoop
    lbcFormat.AddItem "[Defined]", 0
    lbcFormat.ItemData(lbcFormat.NewIndex) = -1

End Sub
Private Sub mPopLogDelivery()
    Dim ilLoop As Integer
    lbcLogDelivery.Clear
    For ilLoop = 0 To UBound(tmVendors) - 1 Step 1
        If tmVendors(ilLoop).sDeliveryType = "L" Then
            lbcLogDelivery.AddItem Trim$(tmVendors(ilLoop).sName)
            lbcLogDelivery.ItemData(lbcLogDelivery.NewIndex) = tmVendors(ilLoop).iIdCode
        End If
    Next ilLoop

End Sub
Private Sub mPopAudioDelivery()
    Dim ilLoop As Integer
    lbcAudioDelivery.Clear
    For ilLoop = 0 To UBound(tmVendors) - 1 Step 1
        If tmVendors(ilLoop).sDeliveryType = "A" Then
            lbcAudioDelivery.AddItem Trim$(tmVendors(ilLoop).sName)
            lbcAudioDelivery.ItemData(lbcAudioDelivery.NewIndex) = tmVendors(ilLoop).iIdCode
        End If
    Next ilLoop

End Sub
Private Sub mSetLbcGridControl(lbcCtrl As ListBox)
    Dim slStr As String
    Dim ilIndex As Integer
    
    mCopyList lbcCtrl
    edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
    cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
    lbcDropdown.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
    gSetListBoxHeight lbcDropdown, 6
    slStr = grdFilter.Text
    ilIndex = SendMessageByString(lbcDropdown.hwnd, LB_FINDSTRING, -1, slStr)
    If ilIndex >= 0 Then
        lbcDropdown.ListIndex = ilIndex
        edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
    Else
        If lbcDropdown.ListCount = 1 Then
            lbcDropdown.ListIndex = 0
            edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
        Else
            lbcDropdown.ListIndex = -1
            edcDropdown.Text = ""
        End If
    End If
    If edcDropdown.Height > grdFilter.RowHeight(grdFilter.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdFilter.RowHeight(grdFilter.Row) - 15
    End If
    edcDropdown.Visible = True
    cmcDropDown.Visible = True
    lbcDropdown.Visible = True
    edcDropdown.SetFocus
End Sub

Private Sub mSetEdcGridControl()
    edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
    edcDropdown.Text = grdFilter.Text
    If edcDropdown.Height > grdFilter.RowHeight(grdFilter.Row) - 15 Then
        edcDropdown.FontName = "Arial"
        edcDropdown.Height = grdFilter.RowHeight(grdFilter.Row) - 15
    End If
    edcDropdown.Visible = True
    edcDropdown.SetFocus
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

Private Sub grdFilter_Scroll()
    mSetShow
End Sub

Private Sub lbcArea_Click()
    edcDropdown.Text = lbcArea.List(lbcArea.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcArea.Visible = False
    End If
End Sub

Private Sub lbcAudioDelivery_Click()
    edcDropdown.Text = lbcAudioDelivery.List(lbcAudioDelivery.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcAudioDelivery.Visible = False
    End If
End Sub

Private Sub lbcDMA_Click()
    edcDropdown.Text = lbcDMA.List(lbcDMA.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcDMA.Visible = False
    End If
End Sub

Private Sub lbcDropdown_Click()
    edcDropdown.Text = lbcDropdown.List(lbcDropdown.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcDropdown.Visible = False
    End If
End Sub

Private Sub lbcFormat_Click()
    edcDropdown.Text = lbcFormat.List(lbcFormat.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcFormat.Visible = False
    End If
End Sub

Private Sub lbcLogDelivery_Click()
    edcDropdown.Text = lbcLogDelivery.List(lbcLogDelivery.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcLogDelivery.Visible = False
    End If
End Sub

Private Sub lbcMSA_Click()
    edcDropdown.Text = lbcMSA.List(lbcMSA.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcMSA.Visible = False
    End If
End Sub

Private Sub lbcMatchOn_Click()
    edcDropdown.Text = lbcMatchOn.List(lbcMatchOn.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcMatchOn.Visible = False
    End If
End Sub

Private Sub lbcOwner_Click()
    edcDropdown.Text = lbcOwner.List(lbcOwner.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcOwner.Visible = False
    End If
End Sub

Private Sub lbcSelect_Click()
    edcDropdown.Text = lbcSelect.List(lbcSelect.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcSelect.Visible = False
    End If
End Sub

Private Sub lbcStation_Click()
    edcDropdown.Text = lbcStation.List(lbcStation.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcStation.Visible = False
    End If
End Sub

Private Sub lbcTerritory_Click()
    edcDropdown.Text = lbcTerritory.List(lbcTerritory.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcTerritory.Visible = False
    End If
End Sub

Private Sub lbcVehicle_Click()
    edcDropdown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcVehicle.Visible = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        grdFilter.Col = SELECTINDEX
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        mSetShow
        Do
            ilNext = False
            Select Case grdFilter.Col
                Case SELECTINDEX
                    If grdFilter.Row = grdFilter.FixedRows Then
                        mSetShow
                        cmdDone.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdFilter.Row = grdFilter.Row - 1
                    If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                        grdFilter.TopRow = grdFilter.TopRow - 1
                    End If
                    grdFilter.Col = FROMVALUEINDEX
                Case Else
                    grdFilter.Col = grdFilter.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        lmTopRow = -1
        grdFilter.TopRow = grdFilter.FixedRows
        grdFilter.Row = grdFilter.FixedRows
        grdFilter.Col = SELECTINDEX
        Do
            If mColOk() Then
                Exit Do
            End If
            If grdFilter.Row + 1 >= grdFilter.Rows Then
                cmdDone.SetFocus
                Exit Sub
            End If
            grdFilter.Row = grdFilter.Row + 1
            Do
                If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                    grdFilter.TopRow = grdFilter.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
        Loop
    End If
    lmTopRow = grdFilter.TopRow
    mEnableBox
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdFilter.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function

Private Sub pbcTab_GotFocus()
    Dim slStr As String
    Dim ilNext As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim ilRange As Integer

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        llEnableCol = lmEnableCol
        mSetShow
        grdFilter.Row = llEnableRow
        grdFilter.Col = llEnableCol
        'Branch
        Do
            ilNext = False
            Select Case grdFilter.Col
                Case FROMVALUEINDEX, TOVALUEINDEX
                    'If mGridFieldsOk(CInt(lmEnableRow)) = False Then
                    '    mEnableBox
                    '    Exit Sub
                    'End If
                    ilRange = False
                    If (imSelectIndex >= 0) And (grdFilter.Col = FROMVALUEINDEX) Then
                        slStr = grdFilter.TextMatrix(grdFilter.Row, MATCHONINDEX)
                        imMatchOnIndex = SendMessageByString(lbcMatchOn.hwnd, LB_FINDSTRING, -1, slStr)
                        If imMatchOnIndex >= 0 Then
                            If (lbcMatchOn.ItemData(imMatchOnIndex) = 3) Or (grdFilter.Col = TOVALUEINDEX) Then
                                ilRange = True
                            End If
                        End If
                    End If
                    If Not ilRange Then
                        If (grdFilter.Row + 1 >= grdFilter.Rows) Then
                            grdFilter.Rows = grdFilter.Rows + 1
                            grdFilter.Row = grdFilter.Row + 1
                            If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                                grdFilter.TopRow = grdFilter.TopRow + 1
                            End If
                            imFromArrow = True
                            pbcArrow.Move grdFilter.Left - pbcArrow.Width, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
                            pbcArrow.Visible = True
                            pbcArrow.SetFocus
                            Exit Sub
                        End If
                        If (grdFilter.Row + 1 < grdFilter.Rows) Then
                            If (Trim$(grdFilter.TextMatrix(grdFilter.Row + 1, SELECTINDEX)) = "") Then
                                grdFilter.Row = grdFilter.Row + 1
                                grdFilter.Col = SELECTINDEX
                                If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                                    grdFilter.TopRow = grdFilter.TopRow + 1
                                End If
                                imFromArrow = True
                                pbcArrow.Move grdFilter.Left - pbcArrow.Width, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
                                pbcArrow.Visible = True
                                pbcArrow.SetFocus
                                Exit Sub
                            End If
                        End If
                        grdFilter.Row = grdFilter.Row + 1
                        grdFilter.Col = SELECTINDEX
                        If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                            grdFilter.TopRow = grdFilter.TopRow + 1
                        End If
                    Else
                        grdFilter.Col = grdFilter.Col + 1
                    End If
                Case Else
                    grdFilter.Col = grdFilter.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        grdFilter.TopRow = grdFilter.FixedRows
        grdFilter.Col = FROMVALUEINDEX
        Do
            If grdFilter.Row <= grdFilter.FixedRows Then
                cmdDone.SetFocus
                Exit Sub
            End If
            grdFilter.Row = grdFilter.Rows - 1
            Do
                If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                    grdFilter.TopRow = grdFilter.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
            If mColOk() Then
                Exit Do
            End If
        Loop
    End If
    lmTopRow = grdFilter.TopRow
    mEnableBox

End Sub

Private Sub mStoreFilter(llRow As Long, tlFilterTypes As FILTERTYPES, tlFilterDef() As FILTERDEF)
    Dim ilUpper As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    
    ilUpper = UBound(tlFilterDef)
    tlFilterDef(ilUpper).iSelect = lbcSelect.ItemData(imSelectIndex)
    tlFilterDef(ilUpper).iOperator = lbcMatchOn.ItemData(imMatchOnIndex)
    tlFilterDef(ilUpper).iCountGroup = tlFilterTypes.iCountGroup
    tlFilterDef(ilUpper).sCntrlType = tlFilterTypes.sCntrlType
    If (tlFilterDef(ilUpper).iOperator = 0) Then
        tlFilterDef(ilUpper).sFromValue = Trim$(grdFilter.TextMatrix(llRow, FROMVALUEINDEX))
        tlFilterDef(ilUpper).lFromValue = 0
        tlFilterDef(ilUpper).sToValue = ""
        tlFilterDef(ilUpper).lToValue = 0
    ElseIf tlFilterTypes.sCntrlType = "E" Then
        tlFilterDef(ilUpper).sFromValue = Trim$(grdFilter.TextMatrix(llRow, FROMVALUEINDEX))
        tlFilterDef(ilUpper).lFromValue = 0
        If tlFilterDef(ilUpper).iOperator = 3 Then
            tlFilterDef(ilUpper).sToValue = Trim$(grdFilter.TextMatrix(llRow, TOVALUEINDEX))
            tlFilterDef(ilUpper).lToValue = 0
        Else
            tlFilterDef(ilUpper).sToValue = ""
            tlFilterDef(ilUpper).lToValue = 0
        End If
        Select Case tlFilterDef(ilUpper).iSelect
            Case SFCALLLETTERSCHGDATE
                If tlFilterDef(ilUpper).sFromValue <> "" Then
                    tlFilterDef(ilUpper).lFromValue = gDateValue(tlFilterDef(ilUpper).sFromValue)
                Else
                    tlFilterDef(ilUpper).lFromValue = 0
                End If
                If tlFilterDef(ilUpper).sToValue <> "" Then
                    tlFilterDef(ilUpper).lToValue = gDateValue(tlFilterDef(ilUpper).sToValue)
                Else
                    tlFilterDef(ilUpper).lToValue = 1000000
                End If
            Case SFHISTSTARTDATE
                If tlFilterDef(ilUpper).sFromValue <> "" Then
                    tlFilterDef(ilUpper).lFromValue = gDateValue(tlFilterDef(ilUpper).sFromValue)
                Else
                    tlFilterDef(ilUpper).lFromValue = gDateValue("1/1/1970")
                End If
                If tlFilterDef(ilUpper).sToValue <> "" Then
                    tlFilterDef(ilUpper).lToValue = gDateValue(tlFilterDef(ilUpper).sToValue)
                Else
                    tlFilterDef(ilUpper).lToValue = 1000000
                End If
            Case SFPERMID, SFP12PLUS, SFXDSID, SFWATTS, SFDMARANK, SFMSARANK
                If tlFilterDef(ilUpper).sFromValue <> "" Then
                    tlFilterDef(ilUpper).lFromValue = Val(gRemoveChar(tlFilterDef(ilUpper).sFromValue, ","))
                Else
                    tlFilterDef(ilUpper).lFromValue = 0
                End If
                If tlFilterDef(ilUpper).sToValue <> "" Then
                    tlFilterDef(ilUpper).lToValue = Val(gRemoveChar(tlFilterDef(ilUpper).sToValue, ","))
                Else
                    tlFilterDef(ilUpper).lToValue = 2000000000
                End If
            Case SFDUE
                tlFilterDef(ilUpper).lFromValue = Val(tlFilterDef(ilUpper).sFromValue)
        End Select
    ElseIf tlFilterTypes.sCntrlType = "T" Then
        tlFilterDef(ilUpper).sFromValue = Trim$(grdFilter.TextMatrix(llRow, FROMVALUEINDEX))
        tlFilterDef(ilUpper).lFromValue = 0
        tlFilterDef(ilUpper).sToValue = ""
        tlFilterDef(ilUpper).lToValue = 0
    Else
        tlFilterDef(ilUpper).sFromValue = ""
        tlFilterDef(ilUpper).lFromValue = 0
        tlFilterDef(ilUpper).sToValue = ""
        tlFilterDef(ilUpper).lToValue = 0
        slStr = grdFilter.TextMatrix(llRow, FROMVALUEINDEX)
        Select Case tlFilterDef(ilUpper).iSelect
            Case SFAREA
                mSetFilterForList lbcArea, slStr, tlFilterDef(ilUpper)
            Case SFCITYLIC
                mSetFilterForList lbcCity, slStr, tlFilterDef(ilUpper)
            Case SFCOUNTYLIC
                mSetFilterForList lbcCounty, slStr, tlFilterDef(ilUpper)
            Case SFDMA  'DMA
                mSetFilterForList lbcDMA, slStr, tlFilterDef(ilUpper)
            Case SFFORMAT  'Format
                mSetFilterForList lbcFormat, slStr, tlFilterDef(ilUpper)
            Case SFMARKETREP
                mSetFilterForList lbcMarketRep, slStr, tlFilterDef(ilUpper)
            Case SFMONIKER
                mSetFilterForList lbcMoniker, slStr, tlFilterDef(ilUpper)
            Case SFMSA  'MSA
                mSetFilterForList lbcMSA, slStr, tlFilterDef(ilUpper)
            Case SFOPERATOR
                mSetFilterForList lbcOperator, slStr, tlFilterDef(ilUpper)
            Case SFOWNER  'Owner
                mSetFilterForList lbcOwner, slStr, tlFilterDef(ilUpper)
            Case SFSERVICEREP
                mSetFilterForList lbcServiceRep, slStr, tlFilterDef(ilUpper)
            Case SFSTATELIC
                mSetFilterForList lbcStateLic, slStr, tlFilterDef(ilUpper)
            Case SFZONE
                mSetFilterForList lbcTimeZone, slStr, tlFilterDef(ilUpper)
            Case SFVEHICLEACTIVE  'Vehicle
                mSetFilterForList lbcVehicle, slStr, tlFilterDef(ilUpper)
            Case SFVEHICLEALL  'Vehicle
                mSetFilterForList lbcVehicle, slStr, tlFilterDef(ilUpper)
            Case SFZIP  'Zip
            Case SFTERRITORY  'Territory
                mSetFilterForList lbcTerritory, slStr, tlFilterDef(ilUpper)
            Case SFCALLLETTERS  'Station
                mSetFilterForList lbcStation, slStr, tlFilterDef(ilUpper)
            '5/6/18: Save filter
            Case SFLOGDELIVERY  'Log Delivery
                mSetFilterForList lbcLogDelivery, slStr, tlFilterDef(ilUpper)
            Case SFAUDIODELIVERY  'Audio Delivery
                mSetFilterForList lbcAudioDelivery, slStr, tlFilterDef(ilUpper)
        End Select
    End If
    tlFilterDef(ilUpper).iFirstFilterLink = -1
    ReDim Preserve tlFilterDef(0 To ilUpper + 1) As FILTERDEF
End Sub

Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
    gSetMousePointer grdFilter, grdFilter, ilMousepointer
End Sub


Private Sub mPopStation()
    Dim ilLoop As Integer
    lbcStation.Clear
    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(ilLoop).sUsedForATT = "Y" Then
            If tgStationInfo(ilLoop).iType = 0 Then
                lbcStation.AddItem Trim$(tgStationInfo(ilLoop).sCallLetters) & ", " & Trim$(tgStationInfo(ilLoop).sMarket)
                lbcStation.ItemData(lbcStation.NewIndex) = tgStationInfo(ilLoop).iCode
            End If
        End If
    Next ilLoop
End Sub

Private Sub mPopulate()
    Dim slName As String
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    mClearControls
    cbcSelection.Clear
    SQLQuery = "SELECT * FROM Fht WHERE fhtUstCode = " & igUstCode
    Set rst_fht = gSQLSelectCall(SQLQuery)
    While Not rst_fht.EOF
        cbcSelection.AddItem Trim$(rst_fht!fhtName)
        cbcSelection.ItemData(cbcSelection.NewIndex) = rst_fht!fhtCode
        rst_fht.MoveNext
    Wend
    'cbcSelection.AddItem "[New]", 0
    'cbcSelection.ItemData(cbcSelection.NewIndex) = 0
    cbcSelection.AddItem "[None]", 0
    cbcSelection.ItemData(cbcSelection.NewIndex) = 0

    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearchFilter-mPopulate"
End Sub

Private Function mSave()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String
    Dim llFhtCode As Long
    Dim llRow As Long
    Dim llCode As Long
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    
    mSave = False
    slName = Trim$(edcName.Text)
    If Trim$(slName) = "" Then
        MsgBox "Name must be Defined.", vbOKOnly
        Exit Function
    End If
    
    ilRow = SendMessageByString(cbcSelection.hwnd, CB_FINDSTRING, -1, slName)
    If ilRow >= 0 Then
        If UCase$(slName) = UCase$(cbcSelection.List(ilRow)) Then
            llFhtCode = cbcSelection.ItemData(ilRow)
        Else
            llFhtCode = 0
            For ilLoop = ilRow + 1 To cbcSelection.ListCount - 1 Step 1
                If UCase$(slName) = UCase$(cbcSelection.List(ilLoop)) Then
                    llFhtCode = cbcSelection.ItemData(ilRow)
                    Exit For
                End If
            Next ilLoop
        End If
    Else
        llFhtCode = 0
    End If
    If (lmFhtCode <> llFhtCode) And (llFhtCode <> 0) Then
        MsgBox "Filter Name previously defined.", vbOKOnly
        Exit Function
    End If
    mSave = False
    'cnn.BeginTrans
    If lmFhtCode = 0 Then
        Do
            SQLQuery = "SELECT MAX(fhtCode) from fht"
            Set rst_fht = gSQLSelectCall(SQLQuery)
            If IsNull(rst_fht(0).Value) Then
                llCode = 1
            Else
                If Not rst_fht.EOF Then
                    llCode = rst_fht(0).Value + 1
                Else
                    llCode = 1
                End If
            End If
            ilRet = 0
            SQLQuery = "Insert into FHT "
            SQLQuery = SQLQuery & "(fhtCode, fhtUstCode, fhtName, fhtUnused) "
            SQLQuery = SQLQuery & " VALUES (" & llCode & ", " & igUstCode & ", '" & gFixQuote(slName) & "'," & "''" & ")"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                If Not gHandleError4994("AffErrorLog.txt", "StationSearchFilter-mSave") Then
                    'cnn.RollbackTrans
                    mSave = False
                    Exit Function
                End If
                ilRet = 1
            End If
        Loop While ilRet <> 0
        lmFhtCode = llCode
    Else
        SQLQuery = "UPDATE FHT"
        SQLQuery = SQLQuery & " SET FhtName = '" & gFixQuote(Trim$(edcName.Text)) & "'"
        SQLQuery = SQLQuery & " WHERE FhtCode = " & lmFhtCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationSearchFilter-mSave"
            'cnn.RollbackTrans
            mSave = False
            Exit Function
        End If
        SQLQuery = "DELETE FROM fit"
        SQLQuery = SQLQuery & " WHERE (fitfhtCode = " & lmFhtCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationSearchFilter-mSave"
            'cnn.RollbackTrans
            mSave = False
            Exit Function
        End If
    End If
    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        If (Trim$(grdFilter.TextMatrix(llRow, SELECTINDEX)) <> "") Then
            slStr = grdFilter.TextMatrix(llRow, SELECTINDEX)
            imSelectIndex = SendMessageByString(lbcSelect.hwnd, LB_FINDSTRING, -1, slStr)
            If imSelectIndex >= 0 Then
                imSelectItemData = lbcSelect.ItemData(imSelectIndex)
                imFilterTypeIndex = lbcSelect.ItemData(imSelectIndex) - 1
                mPopMatchOn tgFilterTypes(imFilterTypeIndex)
                slStr = grdFilter.TextMatrix(llRow, MATCHONINDEX)
                imMatchOnIndex = SendMessageByString(lbcMatchOn.hwnd, LB_FINDSTRING, -1, slStr)
                If imMatchOnIndex >= 0 Then
                    SQLQuery = "Insert into FIT "
                    SQLQuery = SQLQuery & "(fitfhtCode, "
                    SQLQuery = SQLQuery & "fitType, "
                    SQLQuery = SQLQuery & "fitSelectName,"
                    SQLQuery = SQLQuery & "fitOperator,"
                    SQLQuery = SQLQuery & "fitFromValue,"
                    SQLQuery = SQLQuery & "fitToValue, "
                    SQLQuery = SQLQuery & "fitFromLValue,"
                    SQLQuery = SQLQuery & "fitToLValue,"
                    SQLQuery = SQLQuery & "fitUnused) "
                    SQLQuery = SQLQuery & " VALUES ("
                    SQLQuery = SQLQuery & lmFhtCode & ","
                    SQLQuery = SQLQuery & "'" & "S" & "',"
                    SQLQuery = SQLQuery & "'" & lbcSelect.List(imSelectIndex) & "',"
                    SQLQuery = SQLQuery & "'" & lbcMatchOn.List(imMatchOnIndex) & "',"
                    If (lbcMatchOn.ItemData(imMatchOnIndex) = 0) Then
                        SQLQuery = SQLQuery & "'" & gFixQuote(grdFilter.TextMatrix(llRow, FROMVALUEINDEX)) & "',"
                        SQLQuery = SQLQuery & "'" & "" & "',"   'To
                        SQLQuery = SQLQuery & 0 & ","   'From Long Value
                        SQLQuery = SQLQuery & 0 & ","   'To Long Value
                    ElseIf tgFilterTypes(imFilterTypeIndex).sCntrlType = "E" Then
                        SQLQuery = SQLQuery & "'" & gFixQuote(grdFilter.TextMatrix(llRow, FROMVALUEINDEX)) & "',"
                        SQLQuery = SQLQuery & "'" & gFixQuote(grdFilter.TextMatrix(llRow, TOVALUEINDEX)) & "',"
                        SQLQuery = SQLQuery & 0 & ","   'From Long Value
                        SQLQuery = SQLQuery & 0 & ","   'To Long Value
                    ElseIf tgFilterTypes(imFilterTypeIndex).sCntrlType = "T" Then
                        SQLQuery = SQLQuery & "'" & gFixQuote(grdFilter.TextMatrix(llRow, FROMVALUEINDEX)) & "',"
                        SQLQuery = SQLQuery & "'" & "" & "',"   'To
                        SQLQuery = SQLQuery & 0 & ","   'From Long Value
                        SQLQuery = SQLQuery & 0 & ","   'To Long Value
                    Else
                        slStr = grdFilter.TextMatrix(llRow, FROMVALUEINDEX)
                        Select Case imSelectItemData
                            Case SFAREA
                                ilIndex = SendMessageByString(lbcArea.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcArea.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcArea.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFCALLLETTERS  'Station
                                ilIndex = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcStation.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcStation.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFCITYLIC
                                ilIndex = SendMessageByString(lbcCity.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcCity.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcCity.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFCOUNTYLIC
                                ilIndex = SendMessageByString(lbcCounty.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcCounty.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcCounty.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFDMA  'DMA
                                ilIndex = SendMessageByString(lbcDMA.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcDMA.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcDMA.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFFORMAT  'Format
                                ilIndex = SendMessageByString(lbcFormat.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcFormat.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcFormat.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFMARKETREP
                                ilIndex = SendMessageByString(lbcMarketRep.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcMarketRep.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcMarketRep.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFMONIKER
                                ilIndex = SendMessageByString(lbcMoniker.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcMoniker.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcMoniker.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFMSA  'MSA
                                ilIndex = SendMessageByString(lbcMSA.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcMSA.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcMSA.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFOPERATOR
                                ilIndex = SendMessageByString(lbcOperator.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcOperator.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcOperator.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFOWNER  'Owner
                                ilIndex = SendMessageByString(lbcOwner.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcOwner.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcOwner.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFSERVICEREP
                                ilIndex = SendMessageByString(lbcServiceRep.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcServiceRep.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcServiceRep.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFSTATELIC
                                ilIndex = SendMessageByString(lbcStateLic.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcStateLic.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcStateLic.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFZONE
                                ilIndex = SendMessageByString(lbcTimeZone.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcTimeZone.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcTimeZone.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFVEHICLEACTIVE  'Vehicle
                                ilIndex = SendMessageByString(lbcVehicle.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcVehicle.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcVehicle.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFVEHICLEALL  'Vehicle
                                ilIndex = SendMessageByString(lbcVehicle.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcVehicle.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcVehicle.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFTERRITORY  'Territory
                                ilIndex = SendMessageByString(lbcTerritory.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcTerritory.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcTerritory.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                End If
                            '5/6/18:
                            Case SFLOGDELIVERY
                                ilIndex = SendMessageByString(lbcLogDelivery.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcLogDelivery.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcLogDelivery.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                            Case SFAUDIODELIVERY
                                ilIndex = SendMessageByString(lbcAudioDelivery.hwnd, LB_FINDSTRING, -1, slStr)
                                If ilIndex >= 0 Then
                                    SQLQuery = SQLQuery & "'" & Trim$(lbcAudioDelivery.List(ilIndex)) & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & lbcAudioDelivery.ItemData(ilIndex) & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                Else
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'From
                                    SQLQuery = SQLQuery & "'" & "" & "',"   'To
                                    SQLQuery = SQLQuery & 0 & ","
                                    SQLQuery = SQLQuery & 0 & ","   'To
                                End If
                        End Select
                    End If
                    SQLQuery = SQLQuery & "'" & "" & "')"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "StationSearchFilter-mSave"
                        'cnn.RollbackTrans
                        mSave = False
                        Exit Function
                    End If
                End If
            End If
        End If
    Next llRow
    'cnn.CommitTrans
    imFieldChgd = False
    mSave = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearchFilter-mSave"
End Function


Private Sub mClearControls()
    edcName.Text = ""
    mClearGrid
End Sub

Private Sub mBindControls()
    Dim llRow As Long
    
    On Error GoTo ErrHand:
    
    mClearGrid
    grdFilter.Redraw = False
    SQLQuery = "SELECT * FROM Fht WHERE fhtCode = " & lmFhtCode
    Set rst_fht = gSQLSelectCall(SQLQuery)
    If Not rst_fht.EOF Then
        edcName.Text = Trim$(rst_fht!fhtName)
    Else
        edcName.Text = ""
    End If
    llRow = grdFilter.FixedRows
    SQLQuery = "SELECT * FROM Fit WHERE fitfhtCode = " & lmFhtCode
    Set rst_fit = gSQLSelectCall(SQLQuery)
    Do While Not rst_fit.EOF
        If llRow >= grdFilter.Rows Then
            grdFilter.AddItem ""
        End If
        If InStr(1, Trim$(rst_fit!fitSelectName), "Vehicle", vbBinaryCompare) = 1 Then
            If InStr(1, Trim$(rst_fit!fitSelectName), "-", vbBinaryCompare) <> 0 Then
                grdFilter.TextMatrix(llRow, SELECTINDEX) = Trim$(rst_fit!fitSelectName)
            Else
                grdFilter.TextMatrix(llRow, SELECTINDEX) = tgFilterTypes(35).sFieldName
            End If
        Else
            grdFilter.TextMatrix(llRow, SELECTINDEX) = Trim$(rst_fit!fitSelectName)
        End If
        grdFilter.TextMatrix(llRow, MATCHONINDEX) = Trim$(rst_fit!fitOperator)
        If StrComp(Trim$(rst_fit!fitSelectName), Trim$(tgFilterTypes(SFAGREEMENT).sFieldName), vbBinaryCompare) = 0 Then
            If Trim$(rst_fit!fitFromValue) = "Yes" Then
                grdFilter.TextMatrix(llRow, FROMVALUEINDEX) = "Active"
            ElseIf Trim$(rst_fit!fitFromValue) = "No" Then
                grdFilter.TextMatrix(llRow, FROMVALUEINDEX) = "None"
            Else
                grdFilter.TextMatrix(llRow, FROMVALUEINDEX) = Trim$(rst_fit!fitFromValue)
            End If
        Else
            grdFilter.TextMatrix(llRow, FROMVALUEINDEX) = Trim$(rst_fit!fitFromValue)
        End If
        If grdFilter.TextMatrix(llRow, MATCHONINDEX) = "Range" Then
            grdFilter.TextMatrix(llRow, TOVALUEINDEX) = Trim$(rst_fit!fitToValue)
        Else
            grdFilter.Row = llRow
            grdFilter.Col = TOVALUEINDEX
            grdFilter.CellBackColor = LIGHTYELLOW
            grdFilter.TextMatrix(llRow, TOVALUEINDEX) = ""
        End If
        grdFilter.TextMatrix(llRow, FITCODEINDEX) = Trim$(rst_fit!fitCode)
        llRow = llRow + 1
        rst_fit.MoveNext
    Loop
    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        If grdFilter.TextMatrix(llRow, SELECTINDEX) = "" Then
            grdFilter.Row = llRow
            grdFilter.Col = TOVALUEINDEX
            grdFilter.CellBackColor = LIGHTYELLOW
            grdFilter.TextMatrix(llRow, TOVALUEINDEX) = ""
        End If
    Next llRow
    grdFilter.Redraw = True
    Exit Sub
ErrHand:
    mMousePointer vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearchFilter-mBindControls"
End Sub

Private Sub mSetCommands()
    Dim ilSave As Integer
    If lmFhtCode < 0 Then
        cmcSave.Enabled = False
        cmcDelete.Enabled = False
        edcName.Enabled = False
        lacName.Enabled = False
    Else
        edcName.Enabled = True
        lacName.Enabled = True
        ilSave = False
        If lmEnableRow = grdFilter.FixedRows Then
            If lmEnableCol = FROMVALUEINDEX Then
                If edcDropdown.Text <> "" Then
                    ilSave = True
                End If
            End If
        Else
            If grdFilter.TextMatrix(grdFilter.FixedRows, FROMVALUEINDEX) <> "" Then
                ilSave = True
            End If
        End If
        If Not ilSave Then
            cmdDone.Enabled = True  'False
            cmcSave.Enabled = False
            cmcDelete.Enabled = False
        Else
            cmdDone.Enabled = True
            cmcSave.Enabled = True
            If lmFhtCode > 0 Then
                cmcDelete.Enabled = True
            Else
                cmcDelete.Enabled = False
            End If
        End If
    End If
    If (imFieldChgd) And (lmFhtCode > 0) Then
        cmdUndo.Enabled = True
    Else
        cmdUndo.Enabled = False
    End If
End Sub



Private Sub mPopMatchOn(tlFilterType As FILTERTYPES)
    lbcMatchOn.Clear
    If tlFilterType.sContainAllowed = "Y" Then
        lbcMatchOn.AddItem "Contains"
        lbcMatchOn.ItemData(lbcMatchOn.NewIndex) = 0
    End If
    If tlFilterType.sEqualAllowed = "Y" Then
        lbcMatchOn.AddItem "Equal"
        lbcMatchOn.ItemData(lbcMatchOn.NewIndex) = 1
    End If
    If tlFilterType.sNoEqualAllowed = "Y" Then
        lbcMatchOn.AddItem "Not Equal"
        lbcMatchOn.ItemData(lbcMatchOn.NewIndex) = 2
    End If
    If tlFilterType.sRangeAllowed = "Y" Then
        lbcMatchOn.AddItem "Range"
        lbcMatchOn.ItemData(lbcMatchOn.NewIndex) = 3
    End If
    If tlFilterType.sGreaterOrEqual = "Y" Then
        lbcMatchOn.AddItem "Greater or Equal"
        lbcMatchOn.ItemData(lbcMatchOn.NewIndex) = 4
    End If
End Sub

Private Sub mPopSelect()
    Dim ilLoop As Integer
    lbcSelect.Clear
    For ilLoop = 0 To UBound(tgFilterTypes) - 1 Step 1
        If ((sgUsingServiceAgreement = "Y") Or ((sgUsingServiceAgreement <> "Y") And (Trim(tgFilterTypes(ilLoop).sFieldName) <> "Service Agreements"))) Then
            lbcSelect.AddItem Trim$(tgFilterTypes(ilLoop).sFieldName)
            lbcSelect.ItemData(lbcSelect.NewIndex) = tgFilterTypes(ilLoop).iSelect
        End If
    Next ilLoop
End Sub

Private Sub mPopArea()
    Dim ilLoop As Integer
    

    lbcArea.Clear
    For ilLoop = 0 To UBound(tgAreaInfo) - 1 Step 1
        lbcArea.AddItem Trim$(tgAreaInfo(ilLoop).sName)
        lbcArea.ItemData(lbcArea.NewIndex) = tgAreaInfo(ilLoop).lCode
    Next ilLoop
    lbcArea.AddItem "[Defined]", 0
    lbcArea.ItemData(lbcArea.NewIndex) = -1
    
End Sub

Private Sub mPopCity()
    Dim ilLoop As Integer
    

    lbcCity.Clear
    For ilLoop = 0 To UBound(tgCityInfo) - 1 Step 1
        lbcCity.AddItem Trim$(tgCityInfo(ilLoop).sName)
        lbcCity.ItemData(lbcCity.NewIndex) = tgCityInfo(ilLoop).lCode
    Next ilLoop
    lbcCity.AddItem "[Defined]", 0
    lbcCity.ItemData(lbcCity.NewIndex) = -1
    
End Sub
Private Sub mPopCounty()
    Dim ilLoop As Integer
    

    lbcCounty.Clear
    For ilLoop = 0 To UBound(tgCountyInfo) - 1 Step 1
        lbcCounty.AddItem Trim$(tgCountyInfo(ilLoop).sName)
        lbcCounty.ItemData(lbcCounty.NewIndex) = tgCountyInfo(ilLoop).lCode
    Next ilLoop
    lbcCounty.AddItem "[Defined]", 0
    lbcCounty.ItemData(lbcCounty.NewIndex) = -1
    
End Sub

Private Sub mPopMarketRep()
    Dim ilLoop As Integer
    

    lbcMarketRep.Clear
    For ilLoop = 0 To UBound(tgMarketRepInfo) - 1 Step 1
        lbcMarketRep.AddItem Trim$(tgMarketRepInfo(ilLoop).sName)
        lbcMarketRep.ItemData(lbcMarketRep.NewIndex) = tgMarketRepInfo(ilLoop).iUstCode
    Next ilLoop
    lbcMarketRep.AddItem "[Defined]", 0
    lbcMarketRep.ItemData(lbcMarketRep.NewIndex) = -1
    
End Sub
Private Sub mPopMoniker()
    Dim ilLoop As Integer
    

    lbcMoniker.Clear
    For ilLoop = 0 To UBound(tgMonikerInfo) - 1 Step 1
        lbcMoniker.AddItem Trim$(tgMonikerInfo(ilLoop).sName)
        lbcMoniker.ItemData(lbcMoniker.NewIndex) = tgMonikerInfo(ilLoop).lCode
    Next ilLoop
    
End Sub

Private Sub mPopOperator()
    Dim ilLoop As Integer
    

    lbcOperator.Clear
    'Moved to StationSearch
    'gPopMntInfo "O", tgOperatorInfo()
    For ilLoop = 0 To UBound(tgOperatorInfo) - 1 Step 1
        lbcOperator.AddItem Trim$(tgOperatorInfo(ilLoop).sName)
        lbcOperator.ItemData(lbcOperator.NewIndex) = tgOperatorInfo(ilLoop).lCode
    Next ilLoop
    lbcOperator.AddItem "[Defined]", 0
    lbcOperator.ItemData(lbcOperator.NewIndex) = -1
    
End Sub
Private Sub mSetFilterForList(lbcList As ListBox, slStr As String, tlFilterDef As FILTERDEF)
    Dim ilIndex As Integer
    
    ilIndex = SendMessageByString(lbcList.hwnd, LB_FINDSTRING, -1, slStr)
    'If ilIndex < 0 Then
    '    ilIndex = 0
    'End If
    If ilIndex >= 0 Then
        tlFilterDef.lFromValue = lbcList.ItemData(ilIndex)
        tlFilterDef.sFromValue = Trim$(lbcList.List(ilIndex))
    End If
End Sub

Private Sub mPopServiceRep()
    Dim ilLoop As Integer
    

    lbcServiceRep.Clear
    For ilLoop = 0 To UBound(tgServiceRepInfo) - 1 Step 1
        lbcServiceRep.AddItem Trim$(tgServiceRepInfo(ilLoop).sName)
        lbcServiceRep.ItemData(lbcServiceRep.NewIndex) = tgServiceRepInfo(ilLoop).iUstCode
    Next ilLoop
    lbcServiceRep.AddItem "[Defined]", 0
    lbcServiceRep.ItemData(lbcServiceRep.NewIndex) = -1
    
End Sub

Private Sub mPopStateLic()
    Dim ilLoop As Integer
    Dim ilRet As Integer

    lbcStateLic.Clear
    For ilLoop = 0 To UBound(tgStateInfo) - 1 Step 1
        lbcStateLic.AddItem Trim$(tgStateInfo(ilLoop).sPostalName) '& " (" & Trim$(tgStateInfo(ilRow).sName) & ")"
        lbcStateLic.ItemData(lbcStateLic.NewIndex) = tgStateInfo(ilLoop).iCode
    Next ilLoop
    lbcStateLic.AddItem "[Defined]", 0
    lbcStateLic.ItemData(lbcStateLic.NewIndex) = -1
    
End Sub

Private Sub mPopTimeZone()
    Dim ilLoop As Integer
    Dim ilRet As Integer

    lbcTimeZone.Clear
    For ilLoop = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
        lbcTimeZone.AddItem Trim$(tgTimeZoneInfo(ilLoop).sName) '& " (" & Trim$(tgTimeZoneInfo(ilRow).sName) & ")"
        lbcTimeZone.ItemData(lbcTimeZone.NewIndex) = tgTimeZoneInfo(ilLoop).iCode
    Next ilLoop
    lbcTimeZone.AddItem "[Defined]", 0
    lbcTimeZone.ItemData(lbcTimeZone.NewIndex) = -1
    
End Sub

Private Sub mCopyList(lbcFrom As ListBox)
    Dim llLoop As Long
    lbcDropdown.Clear
    For llLoop = 0 To lbcFrom.ListCount - 1 Step 1
        lbcDropdown.AddItem lbcFrom.List(llLoop)
        lbcDropdown.ItemData(lbcDropdown.NewIndex) = lbcFrom.ItemData(llLoop)
    Next llLoop
End Sub

Private Sub mSetToggleGridControl()
    Select Case imSelectItemData
        Case SFCOMMERCIAL
            ReDim smAllowedToggleValues(0 To 1) As String
            smAllowedToggleValues(0) = "Commercial"
            smAllowedToggleValues(1) = "Non-Commercial"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "Commercial"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
        Case SFDAYLIGHT
            ReDim smAllowedToggleValues(0 To 1) As String
            smAllowedToggleValues(0) = "Honor Daylight Savings"
            smAllowedToggleValues(1) = "Ignore Daylight Savings"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "Honor Daylight Savings"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
        Case SFONAIR
            ReDim smAllowedToggleValues(0 To 1) As String
            smAllowedToggleValues(0) = "On Air"
            smAllowedToggleValues(1) = "Off Air"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "On Air"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
        Case SFEMAIL, SFISCI, SFLABEL
            ReDim smAllowedToggleValues(0 To 1) As String
            smAllowedToggleValues(0) = "Checked"
            smAllowedToggleValues(1) = "Not Checked"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "Not Checked"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
        Case SFWEGENER, SFXDS, SFMULTICAST, SFSISTER
            ReDim smAllowedToggleValues(0 To 1) As String
            smAllowedToggleValues(0) = "Yes"
            smAllowedToggleValues(1) = "No"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "Yes"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
        Case SFAGREEMENT
            ReDim smAllowedToggleValues(0 To 3) As String
            smAllowedToggleValues(0) = "Active"
            smAllowedToggleValues(1) = "None"
            smAllowedToggleValues(2) = "All"
            smAllowedToggleValues(3) = "None Active"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "Active"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
        Case SFSERVICEAGREEMENT
            ReDim smAllowedToggleValues(0 To 2) As String
            smAllowedToggleValues(0) = "None"
            smAllowedToggleValues(1) = "Only"
            smAllowedToggleValues(2) = "Both"
            If grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX) = "" Then
                smCurrentToggle = "None"
            Else
                smCurrentToggle = grdFilter.TextMatrix(lmEnableRow, FROMVALUEINDEX)
            End If
    End Select
    pbcToggle.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
    If pbcToggle.Height > grdFilter.RowHeight(grdFilter.Row) - 15 Then
        pbcToggle.FontName = "Arial"
        pbcToggle.Height = grdFilter.RowHeight(grdFilter.Row) - 15
    End If
    pbcToggle.Visible = True
    pbcToggle.SetFocus
End Sub

Private Sub pbcToggle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    For ilLoop = 0 To UBound(smAllowedToggleValues) Step 1
        If smCurrentToggle = smAllowedToggleValues(ilLoop) Then
            If ilLoop = UBound(smAllowedToggleValues) Then
                smCurrentToggle = smAllowedToggleValues(0)
            Else
                smCurrentToggle = smAllowedToggleValues(ilLoop + 1)
            End If
            pbcToggle_Paint
            Exit Sub
        End If
    Next ilLoop
    smCurrentToggle = smAllowedToggleValues(0)
    pbcToggle_Paint
End Sub

Private Sub pbcToggle_Paint()
    pbcToggle.Cls
    pbcToggle.CurrentX = 15
    pbcToggle.CurrentY = 0 'fgBoxInsetY
    pbcToggle.Print smCurrentToggle
End Sub
