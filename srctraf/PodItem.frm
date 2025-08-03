VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PodItem 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5580
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9210
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   9210
   Begin VB.PictureBox pbcTest 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   4440
      ScaleHeight     =   135
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmcImport 
      Caption         =   "Import"
      Height          =   285
      Left            =   6480
      TabIndex        =   21
      Top             =   4920
      Width           =   945
   End
   Begin V81Traffic.CSI_ComboBoxList cbcPodCategoryCombo 
      Height          =   330
      Left            =   195
      TabIndex        =   2
      Top             =   360
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   582
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin VB.PictureBox plcHeader 
      BorderStyle     =   0  'None
      Height          =   1055
      Left            =   200
      ScaleHeight     =   1050
      ScaleWidth      =   8955
      TabIndex        =   19
      Top             =   360
      Width           =   8950
      Begin VB.TextBox edcItemName 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4340
         TabIndex        =   5
         Top             =   630
         Width           =   4455
      End
      Begin VB.PictureBox pbcItemName 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4320
         Picture         =   "PodItem.frx":0000
         ScaleHeight     =   375
         ScaleWidth      =   4650
         TabIndex        =   20
         Top             =   480
         Width           =   4650
      End
      Begin VB.ComboBox cbcItem 
         BackColor       =   &H00FFFF00&
         Height          =   330
         ItemData        =   "PodItem.frx":7768
         Left            =   4440
         List            =   "PodItem.frx":776A
         TabIndex        =   4
         Top             =   0
         Width           =   4515
      End
   End
   Begin VB.PictureBox pbcTab 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   2280
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   10
      Top             =   100
      Width           =   135
   End
   Begin VB.PictureBox pbcCatTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1800
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   100
      Width           =   135
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8640
      Top             =   4920
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "PodItem.frx":776C
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   345
      Width           =   30
   End
   Begin VB.TextBox edcAPICode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4640
      TabIndex        =   13
      Top             =   4920
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3005
      TabIndex        =   12
      Top             =   4920
      Width           =   945
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   705
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   9045
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   14
      Top             =   100
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1235
      TabIndex        =   11
      Top             =   4920
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdServerVendorCode 
      Height          =   3330
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   5874
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
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
   End
   Begin VB.Label lacScreen 
      Caption         =   "Ad Server Item"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   1965
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "PodItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PodItem.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PodItem.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim tmThf As THF        'Header record image
Dim tmTif() As TIF
Dim tmThfSrchKey As THFKEY0
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastTaxColSorted As Integer
Dim imLastTaxSort As Integer
Dim lmServerVendorCodeRowSelected As Long
Dim imApiCodeChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer
Dim imSelectedPodItem As Integer
Dim imSelectedCategory As Integer
Dim tmCtrls(0)  As FIELDAREA
Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim serverVendorRows As Integer
Dim imInitNoRows As Integer
Dim AdServerVendor() As AVF
Dim imTifHeaderNew As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imChgMode As Integer

'control index
Const NAMEINDEX = 0
Const GRIDINDEX = 1
'grid index
Const SERVERVENDORINDEX = 0
Const APICODEINDEX = 1
Const AVFCODEINDEX = 2
Const TIFTARGETINDEX = 3
Const SORTINDEX = 4

Private Sub cmcImport_Click()
    ImptPodItem.Show vbModal
End Sub

Private Sub edcItemName_GotFocus()
    edcItemName.BackColor = &HFFFF00
End Sub

Private Sub edcItemName_LostFocus()
    edcItemName.BackColor = &H80000005
    mSetShow
End Sub

Private Sub edcAPICode_LostFocus()
    edcAPICode.Visible = False
    edcAPICode.BackColor = &H80000005
End Sub

Private Sub mClear()
 edcItemName.Text = ""
 Erase tmTif
End Sub

Private Sub cbcItem_Change()
    Dim slStr As String
    
    lmEnableRow = -1
    edcAPICode.Text = ""
    edcAPICode.Visible = False
    mSetShow
    If cbcItem.ListIndex > 0 Then
        mLoadLookupData
        edcItemName.SelStart = 0
        edcItemName.SelLength = Len(edcItemName.Text)
        'edcItemName.SetFocus
        mMoveRecToCtrl
    Else
        edcItemName.Text = ""
        tmThfSrchKey.lCode = 0
        Erase tmTif
        mMoveRecToCtrl
    End If
     If Not imUpdateAllowed Then
        slStr = Trim$(edcItemName.Text)
        edcItemName.Visible = False
        pbcItemName.Cls
        pbcItemName.CurrentX = tmCtrls(0).fBoxX + fgBoxInsetX
        pbcItemName.CurrentY = tmCtrls(0).fBoxY + fgBoxInsetY
        gSetShow pbcItemName, slStr, tmCtrls(0)
        pbcItemName.Print tmCtrls(0).sShow
    End If
End Sub

Private Sub cbcItem_Click()
    cbcItem_Change
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
   Dim ilRet As Integer
    
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    
    If imApiCodeChg Then
        If MsgBox("Save all changes?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer
    ilRet = mSaveRec()
    
    If ilRet Then
        mMoveRecToCtrl
    End If
End Sub

Private Sub cmcUpdate_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAPICode_Change()
    Select Case lmEnableCol
        Case SERVERVENDORINDEX
        Case APICODEINDEX
    End Select
    grdServerVendorCode.CellForeColor = vbBlack

End Sub

Private Sub edcAPICode_GotFocus()
    Select Case lmEnableCol
       Case SERVERVENDORINDEX
       Case APICODEINDEX
    End Select
    edcAPICode.Visible = True
    gCtrlGotFocus ActiveControl
    edcAPICode.BackColor = &HFFFF00
End Sub

Private Sub edcItemName_KeyPress(KeyAscii As Integer)
    mSetShow
End Sub

Private Sub edcAPICode_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String

    ilKey = KeyAscii

    Select Case lmEnableCol
       
    End Select

End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub Form_Activate()
        If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(PODITEMSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcTab.Enabled = False
        pbcSTab.Enabled = False
        pbcCatTab.Enabled = False
        edcItemName.Visible = False
        imUpdateAllowed = False
    Else
        pbcTab.Enabled = True
        pbcSTab.Enabled = True
        pbcCatTab.Enabled = True
        imUpdateAllowed = True
        edcItemName.Visible = True
    End If
    gShowBranner imUpdateAllowed
    PodItem.Refresh
    Me.KeyPreview = True
    cbcPodCategoryCombo.SetListIndex = 0
    cbcPodCategoryCombo_OnChange
    cbcPodCategoryCombo.SetFocus
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    Set PodItem = Nothing   'Remove data segment
End Sub


Private Sub grdServerVendorCode_EnterCell()
    mSetShow
End Sub

Private Sub grdServerVendorCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Not imUpdateAllowed Then
        Exit Sub
    End If
    lmTopRow = grdServerVendorCode.TopRow
    grdServerVendorCode.Redraw = False
End Sub

Private Sub grdServerVendorCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    If Not imUpdateAllowed Then
        Exit Sub
    End If

    imIgnoreScroll = False
    If Y < grdServerVendorCode.rowHeight(0) Then
        grdServerVendorCode.Col = grdServerVendorCode.MouseCol
        mPodItemSortCol grdServerVendorCode.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdServerVendorCode.MouseCol
    ilRow = grdServerVendorCode.MouseRow
    If ilRow > serverVendorRows Then
        Exit Sub
    End If
    
    If ilCol < grdServerVendorCode.FixedCols Then
        grdServerVendorCode.Redraw = True
        Exit Sub
    End If
    If ilRow < grdServerVendorCode.FixedRows Then
        grdServerVendorCode.Redraw = True
        Exit Sub
    End If
    grdServerVendorCode.Redraw = True
    lmTopRow = grdServerVendorCode.TopRow
    mEnableBox
End Sub

Private Sub grdServerVendorCode_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdServerVendorCode.Redraw = False Then
        grdServerVendorCode.Redraw = True
        If lmTopRow < grdServerVendorCode.FixedRows Then
            grdServerVendorCode.TopRow = grdServerVendorCode.FixedRows
        Else
            grdServerVendorCode.TopRow = lmTopRow
        End If
        grdServerVendorCode.Refresh
        grdServerVendorCode.Redraw = False
    End If
    If (imCtrlVisible) And (grdServerVendorCode.Row >= grdServerVendorCode.FixedRows) And (grdServerVendorCode.Col >= 0) And (grdServerVendorCode.Col < grdServerVendorCode.Cols - 1) Then
        If grdServerVendorCode.RowIsVisible(grdServerVendorCode.Row) Then
            pbcArrow.Move grdServerVendorCode.Left - pbcArrow.Width - 30, grdServerVendorCode.Top + grdServerVendorCode.RowPos(grdServerVendorCode.Row) + (grdServerVendorCode.rowHeight(grdServerVendorCode.Row) - pbcArrow.height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            edcAPICode.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*            Created: 12/10/20      By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slNameCode                    slCode                    *
'*                                                                                        *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer

    imFirstActivate = True
    imTerminate = False
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False
    imTifHeaderNew = 0

    Screen.MousePointer = vbHourglass
    cmcDone.Top = cmcDone.Top + 2500
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcImport.Top = cmcDone.Top
    
    PodItem.height = cmcDone.Top + 8 * cmcDone.height / 3
    gCenterStdAlone PodItem
    Screen.MousePointer = vbHourglass
    lmServerVendorCodeRowSelected = -1
    imApiCodeChg = False
    imLbcArrowSetting = False
    imDoubleClickName = False
    imChgMode = False
    imLbcMouseDown = False
    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    cbcPodCategoryCombo.BackColor = &HFFFF00
    cbcItem.BackColor = &HFFFF00
    'imTrfRecLen = Len(tmTrf)
    On Error GoTo mInitErr
    On Error GoTo 0
    mInitBox
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*            Created:12/10/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload PodItem
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    If imFirstFocus Then
        imFirstFocus = False
    End If
    If grdServerVendorCode.Visible Then
        lmServerVendorCodeRowSelected = -1
        grdServerVendorCode.Row = 0
        grdServerVendorCode.Col = AVFCODEINDEX
        mSetCommands
    End If
End Sub


Private Sub mPopulate()
    mPopCategory
    mLoadServerVendor
    mMoveRecToCtrl
End Sub


Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  cmcEraseErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer

    If imApiCodeChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    Exit Sub
cmcEraseErr: 'VBC NR
    ilRet = 1
    Resume Next
End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*            Created:12/10/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilLoop                        ilRow                     *
'*  ilCol                                                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35

    mGridPodItemLayout
    mGridPodItemColumnWidths
    mGridPodItemColumns
    
    grdServerVendorCode.Move 180, lacScreen.Top + plcHeader.height + 400, grdServerVendorCode.Width, cmcDone.Top - (lacScreen.Top + plcHeader.height) - 480
    imInitNoRows = (cmcDone.Top - 360 - grdServerVendorCode.Top) \ fgBoxGridH
    grdServerVendorCode.height = grdServerVendorCode.RowPos(0) + imInitNoRows * (fgBoxGridH) + fgPanelAdj - 15
    
    cbcPodCategoryCombo.Move plcHeader.Left + 4300, plcHeader.Top, pbcItemName.Width - 50, 330
    cbcItem.Move 0, cbcPodCategoryCombo.height + 180, pbcItemName.Width - 650
    pbcItemName.Move plcHeader.Left + 4100, cbcPodCategoryCombo.height + 100
    edcItemName.Move plcHeader.Left + 4100, cbcPodCategoryCombo.height + 100, pbcItemName.Width - 120
    
    'pbcItemName.Move plcHeader.Width - pbcItemName.Width + 80, plcHeader.Top + fgBevelY * 2
    gSetCtrl tmCtrls(NAMEINDEX), 20, 30, 4545, fgBoxStH
    'edcItemName.Width = tmCtrls(NAMEINDEX).fBoxW
    edcItemName.MaxLength = 255
    gMoveFormCtrl pbcItemName, edcItemName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
   ' cbcItem.Move plcHeader.Width - cbcItem.Width - 80, 0, edcItemName.Width + 80
   
    
End Sub

Private Sub mGridPodItemLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer
    For ilRow = 0 To grdServerVendorCode.Rows - 1 Step 1
        grdServerVendorCode.rowHeight(ilRow) = fgBoxGridH
    Next ilRow
    For ilCol = 0 To grdServerVendorCode.Cols - 1 Step 1
        grdServerVendorCode.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridPodItemColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdServerVendorCode.Row = grdServerVendorCode.FixedRows - 1
    grdServerVendorCode.Col = SERVERVENDORINDEX
    grdServerVendorCode.CellFontBold = False
    grdServerVendorCode.CellFontName = "Arial"
    grdServerVendorCode.CellFontSize = 6.75
    grdServerVendorCode.CellForeColor = vbBlue
    grdServerVendorCode.CellBackColor = LIGHTBLUE
    grdServerVendorCode.TextMatrix(grdServerVendorCode.Row, grdServerVendorCode.Col) = "Ad Server Vendor"
    grdServerVendorCode.Col = APICODEINDEX
    grdServerVendorCode.CellFontBold = False
    grdServerVendorCode.CellFontName = "Arial"
    grdServerVendorCode.CellFontSize = 6.75
    grdServerVendorCode.CellForeColor = vbBlue
    grdServerVendorCode.CellBackColor = LIGHTBLUE
    grdServerVendorCode.TextMatrix(grdServerVendorCode.Row, grdServerVendorCode.Col) = "API Code"
    grdServerVendorCode.Col = AVFCODEINDEX
    grdServerVendorCode.CellFontBold = False
    grdServerVendorCode.CellFontName = "Arial"
    grdServerVendorCode.CellFontSize = 6.75
    grdServerVendorCode.CellForeColor = vbBlue
    grdServerVendorCode.CellBackColor = LIGHTBLUE
    grdServerVendorCode.TextMatrix(grdServerVendorCode.Row, grdServerVendorCode.Col) = "Avf code"
    grdServerVendorCode.Col = TIFTARGETINDEX
    grdServerVendorCode.CellFontBold = False
    grdServerVendorCode.CellFontName = "Arial"
    grdServerVendorCode.CellFontSize = 6.75
    grdServerVendorCode.CellForeColor = vbBlue
    grdServerVendorCode.CellBackColor = LIGHTBLUE
    grdServerVendorCode.TextMatrix(grdServerVendorCode.Row, grdServerVendorCode.Col) = "Tif Target Code"
    
    grdServerVendorCode.Col = SORTINDEX
    grdServerVendorCode.CellFontBold = False
    grdServerVendorCode.CellFontName = "Arial"
    grdServerVendorCode.CellFontSize = 6.75
    grdServerVendorCode.CellForeColor = vbBlue
    grdServerVendorCode.CellBackColor = LIGHTBLUE
    grdServerVendorCode.TextMatrix(grdServerVendorCode.Row, grdServerVendorCode.Col) = "Sort"
End Sub

Private Sub mGridPodItemColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdServerVendorCode.ColWidth(AVFCODEINDEX) = 0
    grdServerVendorCode.ColWidth(TIFTARGETINDEX) = 0
    grdServerVendorCode.ColWidth(SORTINDEX) = 0
    grdServerVendorCode.ColWidth(SERVERVENDORINDEX) = 0.35 * grdServerVendorCode.Width
    grdServerVendorCode.ColWidth(APICODEINDEX) = 0.65 * grdServerVendorCode.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdServerVendorCode.Width
    For ilCol = 0 To grdServerVendorCode.Cols - 1 Step 1
        llWidth = llWidth + grdServerVendorCode.ColWidth(ilCol)
        If (grdServerVendorCode.ColWidth(ilCol) > 15) And (grdServerVendorCode.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdServerVendorCode.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdServerVendorCode.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdServerVendorCode.Width
            For ilCol = 0 To grdServerVendorCode.Cols - 1 Step 1
                If (grdServerVendorCode.ColWidth(ilCol) > 15) And (grdServerVendorCode.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdServerVendorCode.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdServerVendorCode.FixedCols To grdServerVendorCode.Cols - 1 Step 1
                If grdServerVendorCode.ColWidth(ilCol) > 15 Then
                    ilColInc = grdServerVendorCode.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdServerVendorCode.ColWidth(ilCol) = grdServerVendorCode.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*            Created: 12/10/20      By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdServerVendorCode.Row < grdServerVendorCode.FixedRows) Or (grdServerVendorCode.Row >= grdServerVendorCode.Rows) Or (grdServerVendorCode.Col < grdServerVendorCode.FixedCols) Or (grdServerVendorCode.Col >= grdServerVendorCode.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdServerVendorCode.Row
    lmEnableCol = grdServerVendorCode.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdServerVendorCode.Left - pbcArrow.Width - 30, grdServerVendorCode.Top + grdServerVendorCode.RowPos(grdServerVendorCode.Row) + (grdServerVendorCode.rowHeight(grdServerVendorCode.Row) - pbcArrow.height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    Select Case grdServerVendorCode.Col
        Case APICODEINDEX
            edcAPICode.MaxLength = 255
            slStr = grdServerVendorCode.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdServerVendorCode.Row > grdServerVendorCode.FixedRows Then
                    slStr = grdServerVendorCode.TextMatrix(grdServerVendorCode.Row, grdServerVendorCode.Col)
                End If
            End If
            edcAPICode.Text = slStr
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*            Created:12/16/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If ((cbcPodCategoryCombo.ListIndex = 0) Or (grdServerVendorCode.Row < grdServerVendorCode.FixedRows) Or (grdServerVendorCode.Row >= grdServerVendorCode.Rows) Or (grdServerVendorCode.Col < grdServerVendorCode.FixedCols) Or (grdServerVendorCode.Col >= grdServerVendorCode.Cols - 1)) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdServerVendorCode.Left - pbcArrow.Width - 30, grdServerVendorCode.Top + grdServerVendorCode.RowPos(grdServerVendorCode.Row) + (grdServerVendorCode.rowHeight(grdServerVendorCode.Row) - pbcArrow.height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdServerVendorCode.Col - 1 Step 1
        llColPos = llColPos + grdServerVendorCode.ColWidth(ilCol)
    Next ilCol
    Select Case grdServerVendorCode.Col
        Case APICODEINDEX
            edcAPICode.Move grdServerVendorCode.Left + llColPos + 30, grdServerVendorCode.Top + grdServerVendorCode.RowPos(grdServerVendorCode.Row) + 30, grdServerVendorCode.ColWidth(grdServerVendorCode.Col), grdServerVendorCode.rowHeight(grdServerVendorCode.Row) - 15
            edcAPICode.Visible = True
            edcAPICode.SetFocus
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*            Created:12/16/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                                                                                 *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    slStr = Trim$(edcItemName.Text)
    If slStr = "" Then
        imApiCodeChg = False
        mSetCommands
        Exit Sub
    ElseIf StrComp(Trim$(tmThf.sName), slStr, vbTextCompare) <> 0 Then
        imApiCodeChg = True
    End If
    
    pbcArrow.Visible = False
    If (lmEnableRow >= grdServerVendorCode.FixedRows) And (lmEnableRow < grdServerVendorCode.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case APICODEINDEX
                edcAPICode.Visible = False  'Set visibility
                slStr = edcAPICode.Text
                If StrComp(grdServerVendorCode.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imApiCodeChg = True
                End If
                grdServerVendorCode.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End Select
    End If
    pbcArrow.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub

Private Sub pbcCatTab_GotFocus()
    If GetFocus() <> pbcCatTab.hWnd Then
        Exit Sub
    End If
    
    If mPodItemCategoryBranch() Then
        Exit Sub
    End If
    cbcItem.SetFocus
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    
    imTabDirection = -1  'Set-right to left
    
    If imCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdServerVendorCode.Row > grdServerVendorCode.FixedRows Then
                lmTopRow = -1
                grdServerVendorCode.Row = grdServerVendorCode.Row - 1
                If Not grdServerVendorCode.RowIsVisible(grdServerVendorCode.Row) Then
                    grdServerVendorCode.TopRow = grdServerVendorCode.TopRow - 1
                End If
                grdServerVendorCode.Col = APICODEINDEX
                mEnableBox
            Else
                edcItemName.SetFocus
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdServerVendorCode.TopRow = grdServerVendorCode.FixedRows
        grdServerVendorCode.Col = APICODEINDEX
        grdServerVendorCode.Row = grdServerVendorCode.FixedRows
        mEnableBox
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    
    If lmEnableRow > serverVendorRows - 1 Then
        If cmcUpdate.Enabled = True Then
            cmcUpdate.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        
        If (llEnableRow > 12) Then
            imIgnoreScroll = True
            grdServerVendorCode.TopRow = grdServerVendorCode.TopRow + 1
        End If
        
        mSetShow
        grdServerVendorCode.Row = grdServerVendorCode.Row + 1
        grdServerVendorCode.Col = APICODEINDEX
        mEnableBox
    Else
        lmTopRow = -1
        grdServerVendorCode.TopRow = grdServerVendorCode.FixedRows
        grdServerVendorCode.Col = APICODEINDEX
        grdServerVendorCode.Row = grdServerVendorCode.FixedRows
        mEnableBox
    End If
End Sub

Private Function mSaveRec() As Integer
    Dim ilRow As Integer
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim slStr As String
    
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdServerVendorCode, grdServerVendorCode, vbHourglass
    For ilRow = grdServerVendorCode.FixedRows To grdServerVendorCode.Rows - 1 Step 1
        If mGridFieldsOk(ilRow) = False Then
            ilError = True
        End If
    Next ilRow
    If ilError Then
        gSetMousePointer grdServerVendorCode, grdServerVendorCode, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    mMoveCtrlToRec
    slStr = Trim$(edcItemName)
    mPopItems (imSelectedCategory)
    gFindMatch slStr, 0, cbcItem    'Determine if name exist
        If gLastFound(cbcItem) <> -1 Then
            cbcItem.ListIndex = gLastFound(cbcItem)
        End If
    imApiCodeChg = False
    mSaveRec = True
    gSetMousePointer grdServerVendorCode, grdServerVendorCode, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    gSetMousePointer grdServerVendorCode, grdServerVendorCode, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*            Created:12/23/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer control values to     *
'*                      records                        *
'*                                                     *
'*******************************************************


Private Sub mMoveCtrlToRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilTrf                                                   *
'******************************************************************************************

    Dim llRow As Long
    Dim slStr As String
    Dim slQuery As String
    Dim llCount As Long
    Dim itifRet As Long
    Dim tifCode As Long
    Dim tifThfCode As Integer
    Dim tifAvfCode As Integer
    Dim tifUrfCode As Integer
    Dim tifAPICode As String
    
    tifUrfCode = tgUrf(0).iCode
    
    tmThf.iUrfCode = tgUrf(0).iCode
    tmThf.sName = edcItemName.Text
    tmThf.iCategoryMnfCode = tgPodCategoryItems(cbcPodCategoryCombo.ListIndex - 1).iCode
    
    If tmThfSrchKey.lCode <> 0 Then
        slQuery = "Update thf_Target_Header Set thfName = '" & tmThf.sName & "',thfUrfCode =" & tmThf.iUrfCode & " WHERE thfCode = " & tmThfSrchKey.lCode
        If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
                    gHandleError "TrafficErrors.txt", "PodItem-mMoveCtrlToRec"
                    Exit Sub
        End If
        tifThfCode = tmThfSrchKey.lCode
        imTifHeaderNew = 0
    Else
        slQuery = "INSERT INTO thf_Target_Header(thfCode, thfName, thfCategoryMnfCode, thfUrfCode) Values("
        slQuery = slQuery & "replace" & ","
        slQuery = slQuery & "'" & gFixQuote(Trim(tmThf.sName)) & "',"
        slQuery = slQuery & tmThf.iCategoryMnfCode & ","
        slQuery = slQuery & tmThf.iUrfCode & ")"
        tifThfCode = gInsertAndReturnCode(slQuery, "thf_Target_Header", "thfCode", "replace")
        imTifHeaderNew = 1
    End If
    
    For llRow = grdServerVendorCode.FixedRows To serverVendorRows Step 1
        tifCode = Val(grdServerVendorCode.TextMatrix(llRow, TIFTARGETINDEX))
        tifAPICode = Trim$(grdServerVendorCode.TextMatrix(llRow, APICODEINDEX))
        tifAvfCode = gStrDecToInt(Trim$(grdServerVendorCode.TextMatrix(llRow, AVFCODEINDEX)), 0)
            
            If tifCode <> 0 And tifAPICode = "" Then
               slQuery = "Delete from tif_Target_Items WHERE tifCode = " & tifCode
               If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
                    gHandleError "TrafficErrors.txt", "PodItem-mMoveCtrlToRec"
                    Exit Sub
                End If
            ElseIf tifCode <> 0 Then
               slQuery = "Update tif_Target_Items Set tifApiCode = '" & gFixQuote(Trim(tifAPICode)) & "', tifUrfCode = " & tifUrfCode & " WHERE tifCode = " & tifCode
               If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
                    gHandleError "TrafficErrors.txt", "PodItem-mMoveCtrlToRec"
                    Exit Sub
                End If
            ElseIf Trim(tifAPICode) <> "" Then
                slQuery = "INSERT INTO tif_Target_Items(tifCode, tifTHfCode, tifAvfCode, tifApiCode,tifUrfCode) Values("
                slQuery = slQuery & "replace" & ","
                slQuery = slQuery & tifThfCode & ","
                slQuery = slQuery & tifAvfCode & ","
                slQuery = slQuery & "'" & gFixQuote(Trim(tifAPICode)) & "',"
                slQuery = slQuery & tifUrfCode & ")"
                itifRet = gInsertAndReturnCode(slQuery, "tif_Target_Items", "tifCode", "replace")
            End If
            
            
    Next llRow
    Exit Sub
End Sub





'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*            Created:12/16/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                                                                                 *
'******************************************************************************************

    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    
   If serverVendorRows = 0 Then
        ReDim Preserve tmTif(serverVendorRows)
    Else
        ReDim Preserve tmTif(serverVendorRows - 1)
    End If
    grdServerVendorCode.Redraw = False
    grdServerVendorCode.Rows = imInitNoRows
    For llRow = grdServerVendorCode.FixedRows To grdServerVendorCode.Rows - 1 Step 1
        grdServerVendorCode.rowHeight(llRow) = fgBoxGridH
        grdServerVendorCode.Col = 0
        grdServerVendorCode.Row = llRow
        grdServerVendorCode.CellBackColor = &HC0FFFF
    Next llRow
    llRow = grdServerVendorCode.FixedRows

    If serverVendorRows > 12 Then
        grdServerVendorCode.ColWidth(APICODEINDEX) = 0.62 * grdServerVendorCode.Width
    End If

    For ilLoop = 0 To serverVendorRows - 1 Step 1
        If llRow >= grdServerVendorCode.Rows Then
            grdServerVendorCode.AddItem ""
            grdServerVendorCode.rowHeight(llRow) = fgBoxGridH
            grdServerVendorCode.Col = 0
            grdServerVendorCode.Row = llRow
            grdServerVendorCode.CellBackColor = &HC0FFFF
        End If
        ilIndex = gBinarySearchTifAvCode(AdServerVendor(ilLoop).iCode)
        If ilIndex >= 0 Then
            grdServerVendorCode.TextMatrix(llRow, APICODEINDEX) = Trim$(tmTif(ilIndex).sAPICode)
            grdServerVendorCode.TextMatrix(llRow, TIFTARGETINDEX) = tmTif(ilIndex).lCode
        Else
            grdServerVendorCode.TextMatrix(llRow, APICODEINDEX) = ""
            grdServerVendorCode.TextMatrix(llRow, TIFTARGETINDEX) = 0
        End If
        grdServerVendorCode.TextMatrix(llRow, SERVERVENDORINDEX) = Trim$(AdServerVendor(ilLoop).sName)
        grdServerVendorCode.TextMatrix(llRow, AVFCODEINDEX) = Trim$(AdServerVendor(ilLoop).iCode)
        llRow = llRow + 1
    Next ilLoop
    
    grdServerVendorCode.Row = 0
    grdServerVendorCode.Col = AVFCODEINDEX
    grdServerVendorCode.Redraw = True
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*            Created:12/16/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk(ilRowNo As Integer) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
   
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mLoadServerVendor               *
'*                                                     *
'*            Created:12/16/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: populate Pod Target selection  *
'*                                                     *
'*******************************************************
Private Sub mLoadServerVendor()
    Dim Index As Integer
    Dim vendorCount As Integer
    Dim hasRecord As Integer
    On Error GoTo ErrHand
    SQLQuery = "SELECT avfCode, avfName FROM avf_AdVendor"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        hasRecord = 1
        ReDim Preserve AdServerVendor(Index)
        AdServerVendor(Index).iCode = Val(rst!avfCode)
        AdServerVendor(Index).sName = Trim$(rst!avfName)
        Index = Index + 1
        rst.MoveNext
    Loop
    If (hasRecord = 1) Then
        serverVendorRows = UBound(AdServerVendor) + 1
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
     gHandleError "TrafficErrors.txt", "mLoadServerVendor"
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mPopCategory                    *
'*                                                     *
'*            Created:12/16/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: populate Pod Target selection  *
'*                                                     *
'*******************************************************
Private Sub mPopCategory()
    Dim i As Integer
    On Error GoTo ErrHand
    gPopCategoryItems
    cbcPodCategoryCombo.Clear
    cbcPodCategoryCombo.AddItem ("[New]")
    cbcPodCategoryCombo.SetItemData = 0

    For i = 0 To UBound(tgPodCategoryItems) - 1 Step 1
        cbcPodCategoryCombo.AddItem Trim$(tgPodCategoryItems(i).sName)
        cbcPodCategoryCombo.SetItemData = tgPodCategoryItems(i).iCode
    Next i
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
     gHandleError "TrafficErrors.txt", "poditem-mPopCategory"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopItems                       *
'*                                                     *
'*            Created:12/17/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: populate Pod Item selection    *
'*                                                     *
'*******************************************************
Private Sub mPopItems(imSelectedCategory As Integer)
    On Error GoTo ErrHand
    gPodItems (imSelectedCategory)
    cbcItem.Clear
    cbcItem.AddItem ("[New]")
    cbcItem.ItemData(cbcItem.NewIndex) = -1
    
    Dim i As Integer
    For i = 0 To UBound(tgPodItems) - 1 Step 1
        cbcItem.AddItem Trim$(tgPodItems(i).ItemName)
        cbcItem.ItemData(cbcItem.NewIndex) = tgPodItems(i).iCode
    Next i
    cbcItem.Text = "Item Selection"
    
    If imTifHeaderNew > 0 Then
        cbcItem.ListIndex = cbcItem.ListCount - 1
    Else
        cbcItem.ListIndex = imSelectedPodItem
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
     gHandleError "TrafficErrors.txt", "poditem-mPopItems"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mLoadLookupData                 *
'*                                                     *
'*            Created:12/22/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get targetItems by ilCode      *
'*                                                     *
'*******************************************************
Private Sub mLoadLookupData()
    imSelectedPodItem = cbcItem.ListIndex
    tmThfSrchKey.lCode = tgPodItems(cbcItem.ListIndex - 1).iCode
    edcItemName.Text = cbcItem.Text
    If tmThfSrchKey.lCode > 0 Then
        tmThf.sName = Trim$(cbcItem.Text)
        mLoadTargetItems (tmThfSrchKey.lCode)
        'mMoveRecToCtrl
    End If
End Sub


Private Sub mLoadTargetItems(iThfCode As Integer)
    Dim llUpper As Long
    Dim rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slStamp As String
    Dim llMax As Long

    slStamp = gFileDateTime(sgDBPath & "tif.mkd")
    If Not gFileChgd("tif.mkd") Then
        Exit Sub
    End If
    
    SQLQuery = "select * from tif_Target_Items WHERE tifTHfCOde = " & iThfCode & " order by tifAvfCode"
    Set rst = gSQLSelectCall(SQLQuery)
    llUpper = 0
    
    If (Not AdServerVendor) = -1 Then
        Exit Sub
    End If
    
    llMax = UBound(AdServerVendor)
    ReDim tmTif(0 To llMax) As TIF
    
    While Not rst.EOF
        tmTif(llUpper).lCode = rst!tifCode
        tmTif(llUpper).iAvfCode = rst!tifAvfCode
        tmTif(llUpper).sAPICode = rst!tifAPICode
        tmTif(llUpper).iThfCode = rst!tifThfCode
        tmTif(llUpper).iUrfCode = rst!tifUrfCode
        llUpper = llUpper + 1
        rst.MoveNext
    Wend
    gFileChgdUpdate "tif.mkd", False
    rst.Close
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchTifAvCode          *
'*                                                     *
'*            Created:12/22/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get targetItems by ilCode      *
'*                                                     *
'*******************************************************
Private Function gBinarySearchTifAvCode(iAvfCode As Integer) As Integer

   Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tmTif)
    llMax = UBound(tmTif)
    Do While llMin <= llMax
        
        If iAvfCode = tmTif(llMin).iAvfCode Then
            'found the match
            gBinarySearchTifAvCode = llMin
            Exit Function
        End If
        llMin = llMin + 1
    Loop
    gBinarySearchTifAvCode = -1
    Exit Function

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchPodCategory        *
'*                                                     *
'*            Created:12/22/20      By:L. Bianchi      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get targetItems by ilCode      *
'*                                                     *
'*******************************************************
Private Function gBinarySearchPodCategory(Name As String) As Integer

   Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tgPodCategoryItems)
    llMax = UBound(tgPodCategoryItems)
    Do While llMin <= llMax
        
        If Name = Trim$(tgPodCategoryItems(llMin).sName) Then
            'found the match
            gBinarySearchPodCategory = llMin
            Exit Function
        End If
        llMin = llMin + 1
    Loop
    gBinarySearchPodCategory = -1
    Exit Function

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchTif                *
'*                                                     *
'*            Created:12/22/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get targetItems by ilCode      *
'*                                                     *
'*******************************************************
Private Function gBinarySearchTif(lCode As Long) As Long

    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tmTif)
    llMax = UBound(tmTif) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If lCode = tmTif(llMiddle).lCode Then
            'found the match
            gBinarySearchTif = llMiddle
            Exit Function
        ElseIf lCode < tmTif(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchTif = -1
    Exit Function

End Function


Private Function mPodItemCategoryBranch() As Integer
'
'   ilRet = mPodItemCategoryBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    Dim ifoundIndex As Integer
    
    
    ilRet = cbcPodCategoryCombo.ListIndex
    If Not imUpdateAllowed Then
        Exit Function
    End If
    
    If (ilRet > 0) And (Not imDoubleClickName) Then
        mPodItemCategoryBranch = False
        Exit Function
    End If
   
    sgMnfCallType = "5"
    igMNmCallSource = CALLSOURCEPODITEM
    If cbcPodCategoryCombo.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = cbcPodCategoryCombo.Text
    End If
    ilUpdateAllowed = imUpdateAllowed

        If igTestSystem Then
            slStr = "PodItem^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "PodItem^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    
    sgCommandStr = slStr
    MultiNm.Show vbModal
    
   
    imDoubleClickName = False
    mPodItemCategoryBranch = True
    imUpdateAllowed = ilUpdateAllowed
    
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
        cbcPodCategoryCombo.Clear
        mPopCategory
        If imTerminate Then
            mPodItemCategoryBranch = False
            Exit Function
        End If
        ifoundIndex = gBinarySearchPodCategory(sgMNmName)
        sgMNmName = ""
        If ifoundIndex + 1 > 0 Then
            imChgMode = True
            cbcPodCategoryCombo.SetListIndex = ifoundIndex + 1
            imChgMode = False
            cbcPodCategoryCombo.SetFocus
        Else
            imChgMode = True
            cbcPodCategoryCombo.SetListIndex = 0
            imChgMode = False
            cbcPodCategoryCombo.SetFocus
            Exit Function
        End If
    End If
    
     If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        cbcPodCategoryCombo.SetFocus
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        cbcPodCategoryCombo.SetFocus
        Exit Function
    End If
    Exit Function
  
    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*            Created:12/25/20       By:L. Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
        slStr = Trim$(edcItemName)
        gFindMatch slStr, 0, cbcItem    'Determine if name exist
        If gLastFound(cbcItem) <> -1 Then   'Name found
            If gLastFound(cbcItem) <> imSelectedPodItem Then
                slStr = Trim$(edcItemName)
                If slStr = cbcItem.List(gLastFound(cbcItem)) Then
                    Beep
                    MsgBox "Postcast item name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    mOKName = True
End Function

Private Sub cbcPodCategoryCombo_Click()
    'cbcPodCategoryCombo.BackColor = &HFFFF00
End Sub

Private Sub cbcPodCategoryCombo_Change()
    'cbcPodCategoryCombo.BackColor = &HFFFF00
End Sub

Private Sub cbcPodCategoryCombo_GotFocus()
    'cbcPodCategoryCombo.BackColor = &HFFFF00
End Sub



    Private Sub mPodItemSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    For llRow = grdServerVendorCode.FixedRows To grdServerVendorCode.Rows - 1 Step 1
        slStr = Trim$(grdServerVendorCode.TextMatrix(llRow, AVFCODEINDEX))
        If slStr <> "" Then
            If ilCol = SERVERVENDORINDEX Then
                slSort = grdServerVendorCode.TextMatrix(llRow, SERVERVENDORINDEX)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            ElseIf ilCol = APICODEINDEX Then
                slSort = grdServerVendorCode.TextMatrix(llRow, APICODEINDEX)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            End If
            slStr = grdServerVendorCode.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastTaxColSorted) Or ((ilCol = imLastTaxColSorted) And (imLastTaxSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdServerVendorCode.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdServerVendorCode.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastTaxColSorted Then
        imLastTaxColSorted = SORTINDEX
    Else
        imLastTaxColSorted = -1
        imLastTaxSort = -1
    End If
    gGrid_SortByCol grdServerVendorCode, AVFCODEINDEX, SORTINDEX, imLastTaxColSorted, imLastTaxSort
    imLastTaxColSorted = ilCol
End Sub

Private Sub cbcPodCategoryCombo_OnChange()
 'cbcPodCategoryCombo.BackColor = &H80000005
    imSelectedPodItem = 0
    imTifHeaderNew = 0
    If cbcPodCategoryCombo.ListIndex > 0 Then
        imChgMode = True
        cbcItem.Enabled = True
        edcItemName.Enabled = True
        grdServerVendorCode.Enabled = True
        If imUpdateAllowed Then
            pbcTab.Enabled = True
            pbcSTab.Enabled = True
        End If
        
        imSelectedCategory = tgPodCategoryItems(cbcPodCategoryCombo.ListIndex - 1).iCode
        If Not imDoubleClickName Then
            cbcItem.Enabled = True
            cbcItem.SetFocus
        End If
    Else
        cbcItem.Enabled = False
        edcItemName.Enabled = False
        grdServerVendorCode.Enabled = False
        pbcTab.Enabled = False
        pbcSTab.Enabled = False
    End If
    mPopItems (imSelectedCategory)
End Sub

Private Sub cbcPodCategoryCombo_DblClick()
imDoubleClickName = True

If mPodItemCategoryBranch() Then
        Exit Sub
End If
    
End Sub

