VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ResearchBaseDupl 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9315
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
   ScaleHeight     =   5400
   ScaleWidth      =   9315
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7845
      Top             =   3675
   End
   Begin V81Research.CSI_ComboBoxList cbcSelect 
      Height          =   300
      Left            =   4995
      TabIndex        =   1
      Top             =   30
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   529
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   3480
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   3
      Top             =   345
      Width           =   30
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Duplicate"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5730
      TabIndex        =   8
      Top             =   4890
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4005
      TabIndex        =   7
      Top             =   4890
      Width           =   945
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
      TabIndex        =   9
      Top             =   3885
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2235
      TabIndex        =   6
      Top             =   4890
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDupl 
      Height          =   3090
      Left            =   210
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   435
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   15
      Cols            =   9
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
      _Band(0).Cols   =   9
   End
   Begin VB.Label lacScreen 
      Caption         =   "Research Base Duplicate"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   2190
   End
End
Attribute VB_Name = "ResearchBaseDupl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ResearchBaseDupl.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer

Private tmRateCard() As SORTCODE
Private smRateCardTag As String

Private imTerminate As Integer  'True = terminating task, False= OK
Private imFirstFocus As Integer
Private imLastSelectRow As Integer
Private imCtrlKey As Integer
Private lmDuplRowSelected As Long
Private imDuplChg As Integer

Private lmUpperDrf As Long
Private lmUpperDpf As Long

Dim hmDpf As Integer 'Demo plus data file handle
Dim tmDpf As DPF        'DPF record image
Dim tmDpfSrchKey As LONGKEY0    'DPF key record image
Dim tmDpfSrchKey1 As DPFKEY1    'DPF key record image
Dim tmDpfSrchKey2 As DPFKEY2    'DPF key record image
Dim imDpfRecLen As Integer        'DPF record length

Dim tmMnf As MNF        'Mnf record image
Dim tmMnfSrchKey As INTKEY0    'Mnf key record image
Dim hmMnf As Integer    'Multi-Name file handle
Dim imMnfRecLen As Integer        'MNF record length

Private lmTopRow As Long
Private imInitNoRows As Integer

Private bmInGrid As Boolean

Private Rif_rst As ADODB.Recordset
Private drf_rst As ADODB.Recordset

Private Type DUPLRESEARCHINFO
    iVefCode As Integer
    sVefName As String * 40
    iRdfCode As Integer
    sRdfName As String * 20
    sBase As String * 1
End Type

Private tmBase() As DUPLRESEARCHINFO
Private tmNonBase() As DUPLRESEARCHINFO

Const VEHICLEINDEX = 0
Const BASEDAYPARTINDEX = 1
Const NONBASEDAYPARTINDEX = 2
Const DUPLICATEINDEX = 3
Const SELECTEDINDEX = 4
Const VEFCODEINDEX = 5
Const COLORINDEX = 6
Const BASERDFCODEINDEX = 7
Const NONBASERDFCODEINDEX = 8


Private Sub cbcSelect_OnChange()
    tmcDelay.Enabled = False
    DoEvents
    tmcDelay.Enabled = True
End Sub

Private Sub cmcCancel_Click()
    igReturn = 0
    ReDim Preserve tgAllDrf(0 To lmUpperDrf)
    ReDim Preserve tgAllDpf(0 To lmUpperDpf)
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imDuplChg Then
        If MsgBox("Duplicate all changes?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    igReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer

    ilRet = mSaveRec()
    If ilRet Then
        tmcDelay_Timer
    End If
End Sub

Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    ResearchBaseDupl.Refresh
    Me.KeyPreview = True
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
    If Screen.Width * 15 = 640 Then
        'fmAdjFactorW = 1#
        'fmAdjFactorH = 1#
    Else
        'fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        'Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmRateCard
    Erase tmBase
    Erase tmNonBase
    Rif_rst.Close
    drf_rst.Close
    
    btrDestroy hmDpf
    
    Set ResearchBaseDupl = Nothing   'Remove data segment
End Sub


Private Sub grdDupl_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdDupl_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdDupl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdDupl.RowHeight(0) Then
        Exit Sub
    End If
    bmInGrid = True
    llCurrentRow = grdDupl.MouseRow
    llCol = grdDupl.MouseCol
    If llCurrentRow < grdDupl.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdDupl.FixedRows Then
        'If grdDupl.TextMatrix(llCurrentRow, PRODUCTINDEX) <> "" Then
        If grdDupl.TextMatrix(llCurrentRow, BASEDAYPARTINDEX) <> "" Then
            If llCol = DUPLICATEINDEX Then
                llTopRow = grdDupl.TopRow
                If grdDupl.TextMatrix(grdDupl.Row, SELECTEDINDEX) <> "1" Then
                    grdDupl.TextMatrix(grdDupl.Row, SELECTEDINDEX) = "1"
                    'Remove any of click for same vehicle/dyapart
                    For llRow = llCurrentRow - 1 To grdDupl.FixedRows Step -1
                        If grdDupl.TextMatrix(grdDupl.Row, VEFCODEINDEX) = grdDupl.TextMatrix(llRow, VEFCODEINDEX) Then
                            If grdDupl.TextMatrix(grdDupl.Row, NONBASEDAYPARTINDEX) = grdDupl.TextMatrix(llRow, NONBASEDAYPARTINDEX) Then
                                If grdDupl.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                                    grdDupl.TextMatrix(llRow, SELECTEDINDEX) = ""
                                    mPaintRowColor llRow
                                    grdDupl.Row = llCurrentRow
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next llRow
                     For llRow = llCurrentRow + 1 To grdDupl.Rows - 1 Step 1
                        If grdDupl.TextMatrix(grdDupl.Row, VEFCODEINDEX) = grdDupl.TextMatrix(llRow, VEFCODEINDEX) Then
                            If grdDupl.TextMatrix(grdDupl.Row, NONBASEDAYPARTINDEX) = grdDupl.TextMatrix(llRow, NONBASEDAYPARTINDEX) Then
                                If grdDupl.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                                    grdDupl.TextMatrix(llRow, SELECTEDINDEX) = ""
                                    mPaintRowColor llRow
                                    grdDupl.Row = llCurrentRow
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next llRow
                Else
                    grdDupl.TextMatrix(grdDupl.Row, SELECTEDINDEX) = ""
                End If
                imDuplChg = True
                mPaintRowColor grdDupl.Row
                grdDupl.TopRow = llTopRow
                grdDupl.Row = llCurrentRow
            End If
        End If
    End If
    grdDupl.Row = 0
    grdDupl.Col = SELECTEDINDEX
    bmInGrid = False
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()

'
'   mInit
'   Where:
'
    Dim ilRet As Integer

    imFirstActivate = True
    imTerminate = False
    lmUpperDrf = UBound(tgAllDrf)
    lmUpperDpf = UBound(tgAllDpf)
    gSetMousePointer grdDupl, grdDupl, vbHourglass
    'mParseCmmdLine
    'ResearchBaseDupl.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone ResearchBaseDupl
    'ResearchBaseDupl.Show
    gSetMousePointer grdDupl, grdDupl, vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    
    hmDpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", ResearchBaseDupl
    On Error GoTo 0
    imDpfRecLen = Len(tmDpf)
    hmMnf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", ResearchBaseDupl
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    
    
    lmDuplRowSelected = -1
    imDuplChg = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    
    mInitBox

    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    gSetMousePointer grdDupl, grdDupl, vbDefault
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
'*             Created:5/17/93       By:D. LeVine      *
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
    gSetMousePointer grdDupl, grdDupl, vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload ResearchBaseDupl
    igManUnload = NO
End Sub




Private Sub pbcClickFocus_GotFocus()

    If imFirstFocus Then
        imFirstFocus = False
    End If
    If grdDupl.Visible Then
        lmDuplRowSelected = -1
        grdDupl.Row = 0
        grdDupl.Col = VEFCODEINDEX
        mSetCommands
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub mSetCommands()

    Dim ilRet As Integer

    If imDuplChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    Exit Sub
End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35

    mGridDuplLayout
    mGridDuplColumnWidths
    mGridDuplColumns
    cmcDone.Top = Me.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    
    grdDupl.Move 180, lacScreen.Top + lacScreen.Height + 120, grdDupl.Width, cmcDone.Top - (lacScreen.Top + lacScreen.Height) - 240
    ''grdDupl.Height = grdDupl.RowPos(0) + 14 * grdDupl.RowHeight(0) + fgPanelAdj - 15
    'imInitNoRows = (cmcDone.Top - 120 - grdDupl.Top) \ fgFlexGridRowH
    'grdDupl.Height = grdDupl.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
    gGrid_IntegralHeight grdDupl, CInt(fgBoxGridH + 30) ' + 15
    gGrid_FillWithRows grdDupl, fgBoxGridH + 15
    grdDupl.Height = grdDupl.Height + 15
    mClearGrid
End Sub

Private Sub mGridDuplLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdDupl.Rows - 1 Step 1
        grdDupl.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdDupl.Cols - 1 Step 1
        grdDupl.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridDuplColumns()

    grdDupl.Row = grdDupl.FixedRows - 1
    grdDupl.Col = VEHICLEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Vehicle"
    grdDupl.Col = BASEDAYPARTINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Base Daypart"
    grdDupl.Col = NONBASEDAYPARTINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Non-Base Daypart"
    grdDupl.Col = DUPLICATEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Duplicate"
    grdDupl.Col = SELECTEDINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = ""
    grdDupl.Col = VEFCODEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Vef Code"
    grdDupl.Col = COLORINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Color"
    grdDupl.Col = BASERDFCODEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Base RDF Code"
    grdDupl.Col = NONBASERDFCODEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    grdDupl.CellForeColor = vbBlue
    'grdDupl.CellBackColor = LIGHTBLUE
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Non-Base RDF Code"

End Sub

Private Sub mGridDuplColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdDupl.ColWidth(VEFCODEINDEX) = 0
    grdDupl.ColWidth(COLORINDEX) = 0
    grdDupl.ColWidth(BASERDFCODEINDEX) = 0
    grdDupl.ColWidth(NONBASERDFCODEINDEX) = 0
    grdDupl.ColWidth(SELECTEDINDEX) = 0
    grdDupl.ColWidth(VEHICLEINDEX) = 0.4 * grdDupl.Width
    grdDupl.ColWidth(BASEDAYPARTINDEX) = 0.2 * grdDupl.Width
    grdDupl.ColWidth(NONBASEDAYPARTINDEX) = 0.2 * grdDupl.Width
    grdDupl.ColWidth(DUPLICATEINDEX) = 0.1 * grdDupl.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdDupl.Width
    For ilCol = 0 To grdDupl.Cols - 1 Step 1
        llWidth = llWidth + grdDupl.ColWidth(ilCol)
        If (grdDupl.ColWidth(ilCol) > 15) And (grdDupl.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdDupl.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdDupl.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdDupl.Width
            For ilCol = 0 To grdDupl.Cols - 1 Step 1
                If (grdDupl.ColWidth(ilCol) > 15) And (grdDupl.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdDupl.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdDupl.FixedCols To grdDupl.Cols - 1 Step 1
                If grdDupl.ColWidth(ilCol) > 15 Then
                    ilColInc = grdDupl.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdDupl.ColWidth(ilCol) = grdDupl.ColWidth(ilCol) + 15
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




Private Sub pbcSetFocus_GotFocus()
    tmcDelay.Enabled = False
    tmcDelay_Timer
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
End Sub

Private Function mSaveRec() As Integer
    Dim llRow As Long
    Dim llLoop As Long
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim tlDrf As DRF
    Dim llUpper As Long
    Dim llDpf As Long
    Dim ilFound As Integer
    Dim slDemoStr As String
    Dim slPopStr As String

    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdDupl, grdDupl, vbHourglass
    For llRow = grdDupl.FixedRows To grdDupl.Rows - 1 Step 1
        If grdDupl.TextMatrix(llRow, BASEDAYPARTINDEX) <> "" Then
            If grdDupl.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                'Find record to duplicate
                tlDrf.lCode = 0
                For llLoop = 1 To UBound(tgAllDrf) - 1 Step 1
                    If tgAllDrf(llLoop).iStatus >= 0 Then
                        If (tgAllDrf(llLoop).tDrf.iVefCode > 0) And (tgAllDrf(llLoop).tDrf.sDemoDataType <> "P") And (tgAllDrf(llLoop).tDrf.sInfoType = "D") Then
                            If tgAllDrf(llLoop).tDrf.iVefCode = Val(grdDupl.TextMatrix(llRow, VEFCODEINDEX)) Then
                                If tgAllDrf(llLoop).tDrf.iRdfCode = Val(grdDupl.TextMatrix(llRow, BASERDFCODEINDEX)) Then
                                    tlDrf = tgAllDrf(llLoop).tDrf
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next llLoop
                If tlDrf.lCode > 0 Then
                    llUpper = UBound(tgAllDrf)
                    tgAllDrf(llUpper).tDrf = tlDrf
                    tgAllDrf(llUpper).tDrf.lCode = -tlDrf.lCode
                    tgAllDrf(llUpper).tDrf.iRdfCode = Val(grdDupl.TextMatrix(llRow, NONBASERDFCODEINDEX))
                    tgAllDrf(llUpper).sKey = Trim$(grdDupl.TextMatrix(llRow, VEHICLEINDEX)) & Trim$(grdDupl.TextMatrix(llRow, NONBASEDAYPARTINDEX))
                    tgAllDrf(llUpper).iStatus = 0   'New
                    tgAllDrf(llUpper).lIndex = 0
                    tgAllDrf(llUpper).iModel = False
                    tgAllDrf(llUpper).lModelDrfCode = 0
                    tgAllDrf(llUpper).lLink = -1
                    ilFound = False
                    For llDpf = 0 To UBound(tgAllDpf) - 1 Step 1
                        If tgAllDpf(llDpf).lDrfCode = tlDrf.lCode Then
                            ilFound = True
                            tgAllDpf(UBound(tgAllDpf)) = tgAllDpf(llDpf)
                            tgAllDpf(UBound(tgAllDpf)).lDrfCode = -tlDrf.lCode
                            tgAllDpf(UBound(tgAllDpf)).iStatus = 0
                            tgAllDpf(UBound(tgAllDpf)).lDpfCode = 0
                            tgAllDpf(UBound(tgAllDpf)).lIndex = UBound(tgAllDpf)
                            tgAllDpf(UBound(tgAllDpf)).sSource = "B"
                            tgAllDpf(UBound(tgAllDpf)).iRdfCode = tgAllDrf(llUpper).tDrf.iRdfCode
                            
                            'tgDpfRec(UBound(tgDpfRec)) = tgAllDpf(UBound(tgAllDpf))
                            'tgDpfRec(UBound(tgDpfRec)).lIndex = UBound(tgAllDpf)
                            'ReDim Preserve tgDpfRec(0 To UBound(tgDpfRec) + 1) As DPFREC
                            ReDim Preserve tgAllDpf(0 To UBound(tgAllDpf) + 1) As DPFREC
                        End If
                    Next llDpf
                    If Not ilFound Then
                        tmDpfSrchKey1.lDrfCode = tlDrf.lCode
                        tmDpfSrchKey1.iMnfDemo = 0
                        ilRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = tlDrf.lCode)
                            tmMnfSrchKey.iCode = tmDpf.iMnfDemo
                            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If tgSpf.sSAudData = "H" Then
                                slDemoStr = gLongToStrDec(tmDpf.lDemo, 1)
                                slPopStr = gLongToStrDec(tmDpf.lPop, 1)
                            ElseIf tgSpf.sSAudData = "N" Then
                                slDemoStr = gLongToStrDec(tmDpf.lDemo, 2)
                                slPopStr = gLongToStrDec(tmDpf.lPop, 2)
                            ElseIf tgSpf.sSAudData = "U" Then
                                slDemoStr = gLongToStrDec(tmDpf.lDemo, 3)
                                slPopStr = gLongToStrDec(tmDpf.lPop, 3)
                            Else
                                slDemoStr = Trim$(Str$(tmDpf.lDemo))
                                slPopStr = Trim$(Str$(tmDpf.lPop))
                            End If
                            'lbcPlus.AddItem Trim$(tmMnf.sName) & "|" & slDemoStr & "|" & slPopStr
                            tgAllDpf(UBound(tgAllDpf)).sKey = Trim$(tmMnf.sName)
                            tgAllDpf(UBound(tgAllDpf)).iStatus = 0
                            tgAllDpf(UBound(tgAllDpf)).lDpfCode = 0
                            tgAllDpf(UBound(tgAllDpf)).lDrfCode = -tmDpf.lDrfCode
                            tgAllDpf(UBound(tgAllDpf)).sDemo = slDemoStr
                            tgAllDpf(UBound(tgAllDpf)).sPop = slPopStr
                            tgAllDpf(UBound(tgAllDpf)).lIndex = llUpper
                            tgAllDpf(UBound(tgAllDpf)).sSource = "B"
                            tgAllDpf(UBound(tgAllDpf)).iRdfCode = tgAllDrf(llUpper).tDrf.iRdfCode
                            
                            'tgDpfRec(UBound(tgDpfRec)) = tgAllDpf(UBound(tgAllDpf))
                            'tgDpfRec(UBound(tgDpfRec)).lIndex = UBound(tgAllDpf)
                            'ReDim Preserve tgDpfRec(0 To UBound(tgDpfRec) + 1) As DPFREC
                            ReDim Preserve tgAllDpf(0 To UBound(tgAllDpf) + 1) As DPFREC
                            ilRet = btrGetNext(hmDpf, tmDpf, imDpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                    llUpper = llUpper + 1
                    ReDim Preserve tgAllDrf(0 To llUpper) As DRFREC
                End If
            End If
        End If
    Next llRow
    
    
    imDuplChg = False
    mSaveRec = True
    gSetMousePointer grdDupl, grdDupl, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    gSetMousePointer grdDupl, grdDupl, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function





'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
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
    slStr = Trim$(grdDupl.TextMatrix(ilRowNo, BASEDAYPARTINDEX))
    If slStr <> "" Then
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    cbcSelect.FontBold = True
    cbcSelect.SetDropDownWidth (cbcSelect.Width)
    ilRet = gPopRateCardBox(ResearchBaseDupl, 0, cbcSelect, tmRateCard(), smRateCardTag, -1)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopRateCardBox)", ResearchBaseDupl
        On Error GoTo 0
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdDupl.Row = llRow
    For llCol = VEHICLEINDEX To DUPLICATEINDEX Step 1
        grdDupl.Col = llCol
        If grdDupl.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            If llCol = DUPLICATEINDEX Then
                grdDupl.CellFontName = "Monotype Sorts"
                grdDupl.TextMatrix(llRow, DUPLICATEINDEX) = ""
            End If
        Else
            If llCol = DUPLICATEINDEX Then
                grdDupl.CellFontName = "Monotype Sorts"
                grdDupl.TextMatrix(llRow, DUPLICATEINDEX) = "4"
            End If
        End If
        If llCol >= VEHICLEINDEX And llCol < DUPLICATEINDEX Then
            If grdDupl.TextMatrix(llRow, COLORINDEX) <> "G" Then
                grdDupl.CellBackColor = LIGHTYELLOW
            Else
                grdDupl.CellBackColor = LIGHTERGREEN
            End If
        End If
    Next llCol
End Sub

Private Sub tmcDelay_Timer()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slSQLQuery As String
    Dim ilVefCode As Integer
    Dim llRow As Long
    Dim blAddRow As Boolean
    Dim ilUpper As Integer
    Dim ilBase As Integer
    Dim ilNonBase As Integer
    Dim llLoop As Long
    Dim blAdd As Boolean
    Dim slColor As String
    
    tmcDelay.Enabled = False
    gSetMousePointer grdDupl, grdDupl, vbHourglass
    grdDupl.Redraw = False
    mClearGrid
    ReDim tmBase(0 To 0) As DUPLRESEARCHINFO
    ReDim tmNonBase(0 To 0) As DUPLRESEARCHINFO
    If cbcSelect.ListIndex >= 0 Then
        slNameCode = tmRateCard(cbcSelect.ListIndex).sKey
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        slSQLQuery = "Select vefName, rifVefCode, rifBase, rdfName, rifRdfCode from RIF_Rate_Card_Items "
        slSQLQuery = slSQLQuery & " Left Outer Join VEF_Vehicles On rifvefCode = vefCode "
        slSQLQuery = slSQLQuery & " Left Outer Join RDF_Standard_Daypart On rifrdfCode = rdfCode "
        slSQLQuery = slSQLQuery & " Where rifrcfCode = " & slCode
        slSQLQuery = slSQLQuery & " And vefType <> " & "'P'"
        slSQLQuery = slSQLQuery & " Order By vefName, rdfName"
        Set Rif_rst = gSQLSelectCall(slSQLQuery)
        Do While Not Rif_rst.EOF
            
            slSQLQuery = "Select drfCode from DRF_Demo_Rsrch_Data "
            slSQLQuery = slSQLQuery & " Where drfdnfCode = " & igDuplDnfCode
            slSQLQuery = slSQLQuery & " And drfDemoDataType <> 'P' "
            slSQLQuery = slSQLQuery & " And drfmnfSocEco = 0 "
            slSQLQuery = slSQLQuery & " And drfvefCode = " & Rif_rst!rifVefCode
            slSQLQuery = slSQLQuery & " And drfInfotype ='D'"
            slSQLQuery = slSQLQuery & " And drfrdfCode = " & Rif_rst!rifRdfCode
            Set drf_rst = gSQLSelectCall(slSQLQuery)
            If Not drf_rst.EOF Then
                If Rif_rst!rifBase = "Y" Then
                    'Add to base array
                    ilUpper = UBound(tmBase)
                    tmBase(ilUpper).iVefCode = Rif_rst!rifVefCode
                    tmBase(ilUpper).sVefName = Rif_rst!VEFNAME
                    tmBase(ilUpper).iRdfCode = Rif_rst!rifRdfCode
                    tmBase(ilUpper).sRdfName = Rif_rst!rdfName
                    tmBase(ilUpper).sBase = "Y"
                    blAdd = False
                    For llLoop = 1 To UBound(tgAllDrf) - 1 Step 1
                        If tgAllDrf(llLoop).iStatus >= 0 Then
                            If (tgAllDrf(llLoop).tDrf.iVefCode > 0) And (tgAllDrf(llLoop).tDrf.sDemoDataType <> "P") And (tgAllDrf(llLoop).tDrf.sInfoType = "D") Then
                                If tgAllDrf(llLoop).tDrf.iVefCode = tmBase(ilUpper).iVefCode Then
                                    If tgAllDrf(llLoop).tDrf.iRdfCode = tmBase(ilUpper).iRdfCode Then
                                        blAdd = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next llLoop
                    If blAdd Then
                        ReDim Preserve tmBase(0 To ilUpper + 1) As DUPLRESEARCHINFO
                    End If
                End If
            Else
                If Rif_rst!rifBase <> "Y" Then
                    'Add to non-base array
                    ilUpper = UBound(tmNonBase)
                    tmNonBase(ilUpper).iVefCode = Rif_rst!rifVefCode
                    tmNonBase(ilUpper).sVefName = Rif_rst!VEFNAME
                    tmNonBase(ilUpper).iRdfCode = Rif_rst!rifRdfCode
                    tmNonBase(ilUpper).sRdfName = Rif_rst!rdfName
                    tmNonBase(ilUpper).sBase = "N"
                    blAdd = True
                    For llLoop = 1 To UBound(tgAllDrf) - 1 Step 1
                        If tgAllDrf(llLoop).iStatus >= 0 Then
                            If (tgAllDrf(llLoop).tDrf.iVefCode > 0) And (tgAllDrf(llLoop).tDrf.sDemoDataType <> "P") And (tgAllDrf(llLoop).tDrf.sInfoType = "D") Then
                                If tgAllDrf(llLoop).tDrf.iVefCode = tmNonBase(ilUpper).iVefCode Then
                                    If tgAllDrf(llLoop).tDrf.iRdfCode = tmNonBase(ilUpper).iRdfCode Then
                                        blAdd = False
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next llLoop
                    If blAdd Then
                        ReDim Preserve tmNonBase(0 To ilUpper + 1) As DUPLRESEARCHINFO
                    End If
                End If
            End If
            Rif_rst.MoveNext
        Loop
    End If
    ilVefCode = -1
    slColor = "G"
    llRow = grdDupl.FixedRows
    For ilBase = 0 To UBound(tmBase) - 1 Step 1
        For ilNonBase = 0 To UBound(tmNonBase) - 1 Step 1
            If tmBase(ilBase).iVefCode = tmNonBase(ilNonBase).iVefCode Then
                If llRow >= grdDupl.Rows Then
                    grdDupl.AddItem ""
                    grdDupl.RowHeight(llRow) = fgBoxGridH + 15
                End If
                If ilVefCode <> tmBase(ilBase).iVefCode Then
                    grdDupl.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tmBase(ilBase).sVefName)
                    ilVefCode = tmBase(ilBase).iVefCode
                    If slColor = "G" Then
                        slColor = "Y"
                    Else
                        slColor = "G"
                    End If
                End If
                grdDupl.TextMatrix(llRow, BASEDAYPARTINDEX) = Trim$(tmBase(ilBase).sRdfName)
                grdDupl.TextMatrix(llRow, NONBASEDAYPARTINDEX) = Trim$(tmNonBase(ilNonBase).sRdfName)
                grdDupl.TextMatrix(llRow, SELECTEDINDEX) = ""
                grdDupl.TextMatrix(llRow, VEFCODEINDEX) = Trim$(tmNonBase(ilNonBase).iVefCode)
                grdDupl.TextMatrix(llRow, COLORINDEX) = slColor
                grdDupl.TextMatrix(llRow, BASERDFCODEINDEX) = Trim$(tmBase(ilBase).iRdfCode)
                grdDupl.TextMatrix(llRow, NONBASERDFCODEINDEX) = Trim$(tmNonBase(ilNonBase).iRdfCode)
                mPaintRowColor llRow
                llRow = llRow + 1
            End If
        Next ilNonBase
    Next ilBase
    grdDupl.Redraw = True
    gSetMousePointer grdDupl, grdDupl, vbDefault
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long

    'Blank rows within grid
    grdDupl.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdDupl.FixedRows To grdDupl.Rows - 1 Step 1
        grdDupl.Row = llRow
        For llCol = VEHICLEINDEX To NONBASERDFCODEINDEX Step 1
            grdDupl.Col = llCol
            If llCol <= NONBASEDAYPARTINDEX Then
                grdDupl.CellBackColor = LIGHTYELLOW
            Else
                grdDupl.CellBackColor = WHITE
            End If
            grdDupl.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdDupl.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub
