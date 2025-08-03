VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form GetCopyGrid 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   4410
   ClientWidth     =   9045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "GetCopyGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmcErase 
      Caption         =   "Erase"
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
      Left            =   7500
      TabIndex        =   10
      Top             =   4530
      Width           =   1335
   End
   Begin VB.CheckBox ckcIncludeCopy 
      Caption         =   "Include Copy"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   4365
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -30
      ScaleHeight     =   45
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   285
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   4860
      Width           =   45
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
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
      Left            =   5730
      TabIndex        =   8
      Top             =   4530
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
      TabIndex        =   6
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
      Picture         =   "GetCopyGrid.frx":08CA
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   0
      Width           =   60
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Ok"
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
      Left            =   3900
      TabIndex        =   7
      Top             =   4530
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCopyGrid 
      Height          =   3810
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6720
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
   Begin VB.Label plcScreen 
      Caption         =   "Copy Grid Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   0
      Top             =   30
      Width           =   6930
   End
End
Attribute VB_Name = "GetCopyGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GetCopyGrid.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  lmRowSelected                 tmGhfSrchKey0                 tmGsfSrchKey0             *
'*  tmIhfSrchKey1                 tmIhfSrchKey2                 tmIsfSrchKey0             *
'*  tmIsfSrchKey1                 tmIsfSrchKey2                                           *
'******************************************************************************************

'******************************************************
'*  GetCopyGrid - displays missed spots to be changed to Makegoods
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
Private imShiftKey As Integer
Private imTerminate As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private lmEnableRow As Long
Private lmEnableCol As Long
Private smNowDate As String
Private lmNowDate As Long

Private imCopyColSorted As Integer
Private imCopySort As Integer

'Private rst_Gsf As ADODB.Recordset
'Private rst_Ast As ADODB.Recordset

'Grid Name
Dim tmGnf() As GNF            'Gpf record image
Dim tmGnfSrchKey As LONGKEY0  'Gpf key record image
Dim tmGnfSrchKey1 As GNFKEY1  'Gpf key record image
Dim tmGnfSrchKey2 As GNFKEY2  'Gpf key record image
Dim hmGnf As Integer        'Gpf Handle
Dim imGnfRecLen As Integer      'Gpf record length
Dim lmGnfCode As Long
'Grid Position
Dim tmGpf As GPF            'Gpf record image
Dim tmGpfSrchKey As LONGKEY0  'Gpf key record image
Dim tmGpfSrchKey1 As GPFKEY1  'Gpf key record image
Dim hmGpf As Integer        'Gpf Handle
Dim imGpfRecLen As Integer      'Gpf record length
'Grid Copy
Dim tmGcf As GCF            'GCF record image
Dim tmGcfSrchKey As LONGKEY0  'GCF key record image
Dim tmGcfSrchKey1 As GCFKEY1  'GCF key record image
Dim hmGcf As Integer        'GCF Handle
Dim imGcfRecLen As Integer      'GCF record length
'Grid Region Name
Dim tmGtf As GTF            'Gtf record image
Dim tmGtfSrchKey As LONGKEY0  'Gtf key record image
Dim tmGtfSrchKey1 As GTFKEY1  'Gtf key record image
Dim hmGtf As Integer        'Gtf Handle
Dim imGtfRecLen As Integer      'GCF record length

'Comment
Dim tmCsf As CSF            'CSF record image
Dim tmCsfSrchKey As LONGKEY0  'CSF key record image
Dim hmCsf As Integer        'CSF Handle
Dim imCsfRecLen As Integer      'CSF record length

Const GRIDNAMEINDEX = 0
Const ENTEREDDATEINDEX = 1
Const SORTINDEX = 2
Const SELECTEDINDEX = 3
Const GNFCODEINDEX = 4






Private Sub mClearGrid()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llCol                                                                                 *
'******************************************************************************************

    Dim llRow As Long

    'Blank rows within grid
'    gGrid_Clear grdCopyGrid, True
    'Set color within cells
    grdCopyGrid.rowHeight(0) = fgBoxGridH + 15
    For llRow = grdCopyGrid.FixedRows To grdCopyGrid.Rows - 1 Step 1
        'For llCol = 0 To SORTINDEX Step 1
        '    grdCopyGrid.Row = llRow
        '    grdCopyGrid.Col = llCol
        '    grdCopyGrid.CellBackColor = LIGHTYELLOW
        'Next llCol
        grdCopyGrid.rowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub

Private Sub cmcCancel_Click()
    igGetCopyGridReturn = False
    mTerminate
End Sub

Private Sub cmcDone_Click()
    Dim llRow As Long
    Dim slStr As String

    igGetCopyGridReturn = True
    bgGetCopyGridIncludeCopy = False
    If ckcIncludeCopy.Value = vbChecked Then
        bgGetCopyGridIncludeCopy = True
    End If
    For llRow = grdCopyGrid.FixedRows To grdCopyGrid.Rows - 1 Step 1
        slStr = Trim$(grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX))
        If (slStr <> "") Then
            If grdCopyGrid.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                igGetCopyGridGnfCode = Val(grdCopyGrid.TextMatrix(llRow, GNFCODEINDEX))
            End If
        End If
    Next llRow
    mTerminate
End Sub

Private Sub cmcErase_Click()

    Dim llRow As Long
    Dim slStr As String
    Dim ilRet As Integer
    
    For llRow = grdCopyGrid.FixedRows To grdCopyGrid.Rows - 1 Step 1
        slStr = Trim$(grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX))
        If (slStr <> "") Then
            If grdCopyGrid.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                ilRet = MsgBox("Ok to Remove " & grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX) & "?", vbQuestion + vbYesNo, "Erase")
                
                ' TTP 10588 JJB 2023-01-11 BEGIN
                
                'If ilRet = vbYes Then
                '    ilRet = mErase(llRow)
                '    grdCopyGrid.RemoveItem llRow
                '    Exit Sub
                'End If
                
                If ilRet = vbYes Then
                    ilRet = mErase(llRow)
                    If grdCopyGrid.Rows = 2 Then
                        grdCopyGrid.Rows = 1
                    Else
                        grdCopyGrid.RemoveItem llRow
                    End If
                    Exit Sub
                End If
                
                ' TTP 10588 JJB END

            End If
        End If
    Next llRow
    
End Sub

Private Sub Form_Activate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                                                                                 *
'******************************************************************************************


    If imFirstTime Then
        imFirstTime = False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (CLng(50) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.height = (CLng(50) * ((Screen.height) / (480 * 15 / Me.height))) / 100
    gCenterStdAlone GetCopyGrid
    DoEvents
    mSetControls
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    mInit
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmGnf


    ilRet = btrClose(hmGnf)
    btrDestroy hmGnf
    ilRet = btrClose(hmGpf)
    btrDestroy hmGpf
    ilRet = btrClose(hmGcf)
    btrDestroy hmGcf
    ilRet = btrClose(hmGtf)
    btrDestroy hmGtf
    ilRet = btrClose(hmCsf)
    btrDestroy hmCsf
    
    Set GetCopyGrid = Nothing
End Sub






Private Sub grdCopyGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
    If (Shift And SHIFTMASK) > 0 Then
        imShiftKey = True
    Else
        imShiftKey = False
    End If
End Sub

Private Sub grdCopyGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim ilAdf As Integer

    gSetMousePointer grdCopyGrid, grdCopyGrid, vbHourglass
    ilAdf = gBinarySearchAdf(igGetCopyGridAdfCode)
    If ilAdf <> -1 Then
        plcScreen.Caption = "Copy Grid Selection-" & Trim$(tgCommAdf(ilAdf).sName)
    Else
        plcScreen.Caption = "Copy Grid Selection"
    End If
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdCopyGrid.FixedRows
    imCopyColSorted = -1
    imCopySort = -1

    ReDim tmGnf(0 To 0) As GNF
    hmGnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGnf, "", sgDBPath & "Gnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetCopyGrid
    On Error GoTo 0
    imGnfRecLen = Len(tmGnf(0))  'Get and save ARF record length

    hmGpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGpf, "", sgDBPath & "Gpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetCopyGrid
    On Error GoTo 0
    imGpfRecLen = Len(tmGpf)  'Get and save ARF record length
    hmGcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGcf, "", sgDBPath & "Gcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetCopyGrid
    On Error GoTo 0
    imGcfRecLen = Len(tmGcf)  'Get and save ARF record length
    hmGtf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGtf, "", sgDBPath & "Gtf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetCopyGrid
    On Error GoTo 0
    imGtfRecLen = Len(tmGtf)  'Get and save ARF record length
    hmCsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCsf, "", sgDBPath & "Csf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetCopyGrid
    On Error GoTo 0
    imCsfRecLen = Len(tmCsf)  'Get and save ARF record length


    mClearGrid
    ilRet = mReadRec()
    mPopulate

    Screen.MousePointer = vbDefault
    gSetMousePointer grdCopyGrid, grdCopyGrid, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdCopyGrid, grdCopyGrid, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilGsf As Integer
    Dim ilGsfIndex As Integer
    Dim ilGame As Integer
    Dim llUnitTotal As Long
    Dim llTypeUnitsOrdered As Long
    Dim llTypeUnitsProp As Long
    Dim ilCol As Integer

    On Error GoTo ErrHand:

    grdCopyGrid.Redraw = False
    grdCopyGrid.Row = 0
    For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
        grdCopyGrid.Col = llCol
        grdCopyGrid.CellBackColor = vbBlack
        grdCopyGrid.CellBackColor = LIGHTBLUE
    Next llCol
    grdCopyGrid.rowHeight(0) = fgBoxGridH + 15
    llRow = grdCopyGrid.FixedRows
    For ilLoop = 0 To UBound(tmGnf) - 1 Step 1
        If llRow >= grdCopyGrid.Rows Then
            grdCopyGrid.AddItem ""
        End If
        grdCopyGrid.rowHeight(llRow) = fgBoxGridH + 15
        grdCopyGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
        grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX) = Trim$(tmGnf(ilLoop).sGridName)
        gUnpackDate tmGnf(ilLoop).iEnteredDate(0), tmGnf(ilLoop).iEnteredDate(1), slStr
        grdCopyGrid.TextMatrix(llRow, ENTEREDDATEINDEX) = slStr
        grdCopyGrid.TextMatrix(llRow, GNFCODEINDEX) = tmGnf(ilLoop).lCode
        llRow = llRow + 1
    Next ilLoop
    For llRow = grdCopyGrid.FixedRows To grdCopyGrid.Rows - 1 Step 1
        slStr = Trim$(grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX))
        If slStr <> "" Then
            grdCopyGrid.Row = llRow
            grdCopyGrid.Col = GRIDNAMEINDEX
            grdCopyGrid.CellBackColor = LIGHTYELLOW
            grdCopyGrid.Col = ENTEREDDATEINDEX
            grdCopyGrid.CellBackColor = LIGHTYELLOW
        End If
    Next llRow
    'rst_Gsf.Close
    gGrid_AlignAllColsLeft grdCopyGrid
    mCopyGridSortCol ENTEREDDATEINDEX
    mCopyGridSortCol ENTEREDDATEINDEX
    grdCopyGrid.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer

    grdCopyGrid.ColWidth(GNFCODEINDEX) = 0
    grdCopyGrid.ColWidth(SORTINDEX) = 0
    grdCopyGrid.ColWidth(SELECTEDINDEX) = 0
    grdCopyGrid.ColWidth(ENTEREDDATEINDEX) = grdCopyGrid.Width * 0.2
    grdCopyGrid.ColWidth(GRIDNAMEINDEX) = grdCopyGrid.Width - grdCopyGrid.ColWidth(ENTEREDDATEINDEX) - GRIDSCROLLWIDTH - 15
    'Align columns to left
    'gGrid_AlignAllColsLeft grdCopyGrid
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdCopyGrid.TextMatrix(0, GRIDNAMEINDEX) = "Name"
    grdCopyGrid.TextMatrix(0, ENTEREDDATEINDEX) = "Entered Date"
    'Set height of grid

End Sub

Private Sub mCopyGridSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdCopyGrid.FixedRows To grdCopyGrid.Rows - 1 Step 1
        slStr = Trim$(grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX))
        If slStr <> "" Then
            If ilCol = ENTEREDDATEINDEX Then
                slStr = grdCopyGrid.TextMatrix(llRow, ENTEREDDATEINDEX)
                If slStr <> "" Then
                    slSort = Trim$(str$(gDateValue(slStr)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdCopyGrid.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdCopyGrid.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imCopyColSorted) Or ((ilCol = imCopyColSorted) And (imCopySort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCopyGrid.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdCopyGrid.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imCopyColSorted Then
        imCopyColSorted = SORTINDEX
    Else
        imCopyColSorted = -1
        imCopySort = -1
    End If
    gGrid_SortByCol grdCopyGrid, GRIDNAMEINDEX, SORTINDEX, imCopyColSorted, imCopySort
    imCopyColSorted = ilCol
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

    Screen.MousePointer = vbDefault
    gSetMousePointer grdCopyGrid, grdCopyGrid, vbDefault
    igManUnload = YES
    Unload GetCopyGrid
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mReadRecErr                                                                           *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llUpper As Long

    ReDim tmGsf(0 To 0) As GSF
    llUpper = 0
    tmGnfSrchKey1.iAdfCode = igGetCopyGridAdfCode
    ilRet = btrGetEqual(hmGnf, tmGnf(llUpper), imGnfRecLen, tmGnfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmGnf(llUpper).iAdfCode = igGetCopyGridAdfCode)
        If Trim$(tmGnf(llUpper).sGridName) <> "Current Working" Then
            ReDim Preserve tmGnf(0 To UBound(tmGnf) + 1) As GNF
            llUpper = UBound(tmGnf)
        End If
        ilRet = btrGetNext(hmGnf, tmGnf(llUpper), imGnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    mReadRec = True
    Exit Function
mReadRecErr: 'VBC NR
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function

Private Sub mSetControls()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilCol                                                   *
'******************************************************************************************

    Dim ilGap As Integer

    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.height - cmcDone.height - 120
    cmcCancel.Top = cmcDone.Top
    cmcErase.Top = cmcDone.Top
    ckcIncludeCopy.Top = cmcDone.Top + cmcDone.height - ckcIncludeCopy.height
    grdCopyGrid.Move 180, 255, GetCopyGrid.Width - 360, cmcDone.Top - 255 - 120
    ckcIncludeCopy.Left = grdCopyGrid.Left
    cmcErase.Left = grdCopyGrid.Left + grdCopyGrid.Width - cmcErase.Width
    cmcCancel.Left = cmcErase.Left - cmcCancel.Width - ilGap
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdCopyGrid, fgBoxGridH + 15

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdCopyGrid.Row = llRow
    For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
        grdCopyGrid.Col = llCol
        If grdCopyGrid.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdCopyGrid.CellBackColor = LIGHTYELLOW
        Else
            grdCopyGrid.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub

Private Sub grdCopyGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdCopyGrid.rowHeight(0) Then
        grdCopyGrid.Col = grdCopyGrid.MouseCol
        mCopyGridSortCol grdCopyGrid.Col
        grdCopyGrid.Row = 0
        grdCopyGrid.Col = GNFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdCopyGrid, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdCopyGrid.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdCopyGrid.FixedRows Then
        If grdCopyGrid.TextMatrix(llCurrentRow, GRIDNAMEINDEX) <> "" Then
            llTopRow = grdCopyGrid.TopRow
            For llRow = grdCopyGrid.FixedRows To grdCopyGrid.Rows - 1 Step 1
                If grdCopyGrid.TextMatrix(llRow, GRIDNAMEINDEX) <> "" Then
                    If llRow = llCurrentRow Then
                        grdCopyGrid.TextMatrix(llRow, SELECTEDINDEX) = "1"
                    Else
                        grdCopyGrid.TextMatrix(llRow, SELECTEDINDEX) = "0"
                    End If
                    mPaintRowColor llRow
                End If
            Next llRow
            grdCopyGrid.TopRow = llTopRow
            grdCopyGrid.Row = llCurrentRow
        End If
    End If

End Sub

Private Function mErase(llRow As Long) As Integer
    Dim llGnfCode As Long
    Dim llGpfCode As Long
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    llGnfCode = Val(grdCopyGrid.TextMatrix(llRow, GNFCODEINDEX))
    tmGnfSrchKey.lCode = llGnfCode
    ilRet = btrGetEqual(hmGnf, tmGnf(0), imGnfRecLen, tmGnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmGpfSrchKey1.lGnfCode = llGnfCode
        tmGpfSrchKey1.iRowNo = 1
        ilRet = btrGetGreaterOrEqual(hmGpf, tmGpf, imGpfRecLen, tmGpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGpf.lGnfCode = llGnfCode)
            llGpfCode = tmGpf.lCode
            If tmGpf.lCsfCode > 0 Then
                tmCsfSrchKey.lCode = tmGpf.lCsfCode
                imCsfRecLen = Len(tmCsf) '5011
                ilRet = btrGetEqual(hmCsf, tmCsf, imCsfRecLen, tmCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmCsf)
                End If
            End If
            tmGcfSrchKey1.lGpfCode = llGpfCode
            tmGcfSrchKey1.iSeqNo = 0
            ilRet = btrGetGreaterOrEqual(hmGcf, tmGcf, imGcfRecLen, tmGcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmGcf.lGpfCode = llGpfCode)
                ilRet = btrDelete(hmGcf)
                tmGcfSrchKey1.lGpfCode = llGpfCode
                tmGcfSrchKey1.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hmGcf, tmGcf, imGcfRecLen, tmGcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Loop
            'Delete then add
            tmGtfSrchKey1.lGpfCode = llGpfCode
            ilRet = btrGetGreaterOrEqual(hmGtf, tmGtf, imGtfRecLen, tmGtfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmGtf.lGpfCode = llGpfCode)
                ilRet = btrDelete(hmGtf)
                tmGtfSrchKey1.lGpfCode = llGpfCode
                ilRet = btrGetGreaterOrEqual(hmGtf, tmGtf, imGtfRecLen, tmGcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Loop
            ilRet = btrDelete(hmGpf)
            tmGpfSrchKey1.lGnfCode = llGnfCode
            tmGpfSrchKey1.iRowNo = 1
            ilRet = btrGetGreaterOrEqual(hmGpf, tmGpf, imGpfRecLen, tmGpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Loop
        ilRet = btrDelete(hmGnf)
    End If
    
    mErase = True
End Function

