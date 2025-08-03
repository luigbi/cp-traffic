VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form CntrVerify 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CntrVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbcVerifyStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   960
      ScaleHeight     =   210
      ScaleWidth      =   1170
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      TabIndex        =   9
      Top             =   4530
      Width           =   1335
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
      TabIndex        =   4
      Top             =   4860
      Width           =   45
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
      Picture         =   "CntrVerify.frx":030A
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdVerify 
      Height          =   3810
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label plcScreen 
      Caption         =   "Contract Verification"
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
Attribute VB_Name = "CntrVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CntrVerify.FRM on Wed 6/17/09 @ 12:56 P
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
'*  CntrVerify - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imTerminate As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private lmEnableRow As Long
Private lmEnableCol As Long
Private smNowDate As String
Private lmNowDate As Long
Private imSetCtrlVisible As Integer
Private imVerifyStatus As Integer
Private imChg As Integer

Private imVerifyColSorted As Integer
Private imVerifySort As Integer

'Contract
Dim tmChf As CHF        'Chf record image
Dim tmChfSrchKey0 As LONGKEY0    'Chf key record image
Dim hmChf As Integer    'Contract file handle
Dim imChfRecLen As Integer        'CHF record length

Const STATUSINDEX = 0
Const STATUSDATEINDEX = 1
Const CNTRNOINDEX = 2
Const ADVTPRODINDEX = 3
Const AGENCYINDEX = 4
Const SALESPERSONINDEX = 5
Const GROSSINDEX = 6
Const DATESINDEX = 7
Const SORTINDEX = 8
Const CHFCODEINDEX = 9






Private Sub mClearGrid()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llCol                                                                                 *
'******************************************************************************************

    Dim llRow As Long

    'Blank rows within grid
'    gGrid_Clear grdVerify, True
    'Set color within cells
    grdVerify.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdVerify.FixedRows To grdVerify.Rows - 1 Step 1
        'For llCol = 0 To SORTINDEX Step 1
        '    grdVerify.Row = llRow
        '    grdVerify.Col = llCol
        '    grdVerify.CellBackColor = LIGHTYELLOW
        'Next llCol
        grdVerify.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imChg Then
        ilRet = mSaveRec()
    End If
    mTerminate
End Sub



Private Sub cmcSave_Click()
    Dim ilRet As Integer
    
    ilRet = mSaveRec()
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
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
    gSetMousePointer grdVerify, grdVerify, vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (CLng(75) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(75) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone CntrVerify
    DoEvents
    mSetControls
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdVerify, grdVerify, vbHourglass

    mInit
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CntrVerify = Nothing
End Sub






Private Sub mInit()
    Dim ilRet As Integer
    Dim ilAdf As Integer

    gSetMousePointer grdVerify, grdVerify, vbHourglass
    plcScreen.Caption = "Contract Verification: "
    If sgCntrVerifyType = "N" Then
        plcScreen.Caption = plcScreen.Caption & "Not Verified"
    ElseIf sgCntrVerifyType = "S" Then
        plcScreen.Caption = plcScreen.Caption & "Sent to Agency"
    ElseIf sgCntrVerifyType = "V" Then
        plcScreen.Caption = plcScreen.Caption & "Verified"
    End If
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdVerify.FixedRows
    imVerifyColSorted = -1
    imVerifySort = -1
    imChg = False

    hmChf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CntrVerify
    On Error GoTo 0
    imChfRecLen = Len(tmChf)  'Get and save ARF record length


    mClearGrid
    ilRet = mReadRec()

    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate(ilFlag As Integer, llCurrentRow As Long)
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim ilAgf As Integer
    Dim ilSlf As Integer
    Dim slStartDate As String
    Dim slEndDate As String

    On Error GoTo ErrHand:

    If ilFlag = 0 Then
        grdVerify.Redraw = False
        grdVerify.Row = 0
        For llCol = STATUSDATEINDEX To DATESINDEX Step 1
            grdVerify.Col = llCol
            grdVerify.CellBackColor = vbBlack
            grdVerify.CellBackColor = LIGHTBLUE
        Next llCol
        grdVerify.RowHeight(0) = fgBoxGridH + 15
        llCurrentRow = grdVerify.FixedRows
        Exit Sub
    End If
    If ilFlag = 1 Then
        llRow = llCurrentRow
        If llRow >= grdVerify.Rows Then
            grdVerify.AddItem ""
        End If
        grdVerify.RowHeight(llRow) = fgBoxGridH + 15
        If tmChf.sVerifyFlag = "S" Then
            grdVerify.TextMatrix(llRow, STATUSINDEX) = "Sent"
        ElseIf tmChf.sVerifyFlag = "N" Then
            grdVerify.TextMatrix(llRow, STATUSINDEX) = "Not Sent"
        ElseIf tmChf.sVerifyFlag = "V" Then
            grdVerify.TextMatrix(llRow, STATUSINDEX) = "Verified"
        End If
        gUnpackDate tmChf.iVerifyDate(0), tmChf.iVerifyDate(1), slStr
        grdVerify.TextMatrix(llRow, STATUSDATEINDEX) = slStr
        grdVerify.TextMatrix(llRow, CNTRNOINDEX) = Trim$(Str(tmChf.lCntrNo))
        ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
        If ilAdf <> -1 Then
            slStr = Trim$(tgCommAdf(ilAdf).sName) & "/" & Trim$(tmChf.sProduct)
        Else
            slStr = "Advertiser " & tmChf.iAdfCode & " missing"
        End If
        grdVerify.TextMatrix(llRow, ADVTPRODINDEX) = slStr
        If tmChf.iAgfCode > 0 Then
            ilAgf = gBinarySearchAgf(tmChf.iAgfCode)
            If ilAgf <> -1 Then
                slStr = Trim$(tgCommAgf(ilAgf).sName)
            Else
                slStr = "Agency " & tmChf.iAgfCode & " missing"
            End If
        End If
        grdVerify.TextMatrix(llRow, AGENCYINDEX) = slStr
        ilSlf = gBinarySearchSlf(tmChf.iSlfCode(0))
        If ilSlf <> -1 Then
            slStr = Trim$(tgMSlf(ilSlf).sFirstName) & " " & Trim$(tgMSlf(ilSlf).sLastName)
        Else
            slStr = "Salesperson " & tmChf.iSlfCode(0) & " missing"
        End If
        grdVerify.TextMatrix(llRow, SALESPERSONINDEX) = slStr
        slStr = gLongToStrDec(tmChf.lInputGross, 2)
        grdVerify.TextMatrix(llRow, GROSSINDEX) = slStr
        gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStartDate
        gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slEndDate
        grdVerify.TextMatrix(llRow, DATESINDEX) = slStartDate & "-" & slEndDate
        grdVerify.TextMatrix(llRow, CHFCODEINDEX) = tmChf.lCode
        llCurrentRow = llCurrentRow + 1
        Exit Sub
    End If
    If ilFlag = 2 Then
        For llRow = grdVerify.FixedRows To grdVerify.Rows - 1 Step 1
            slStr = Trim$(grdVerify.TextMatrix(llRow, CNTRNOINDEX))
            If slStr <> "" Then
                grdVerify.Row = llRow
                For llCol = STATUSDATEINDEX To DATESINDEX Step 1
                    grdVerify.Col = llCol
                    grdVerify.CellBackColor = LIGHTYELLOW
                Next llCol
            End If
        Next llRow
        gGrid_AlignAllColsLeft grdVerify
        mVerifyGridSortCol CNTRNOINDEX
        mVerifyGridSortCol CNTRNOINDEX
        grdVerify.Redraw = True
        Exit Sub
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer

    grdVerify.ColWidth(CHFCODEINDEX) = 0
    grdVerify.ColWidth(SORTINDEX) = 0
    grdVerify.ColWidth(STATUSINDEX) = grdVerify.Width * 0.08
    grdVerify.ColWidth(STATUSDATEINDEX) = grdVerify.Width * 0.08
    grdVerify.ColWidth(CNTRNOINDEX) = grdVerify.Width * 0.08
    grdVerify.ColWidth(AGENCYINDEX) = grdVerify.Width * 0.15
    grdVerify.ColWidth(SALESPERSONINDEX) = grdVerify.Width * 0.15
    grdVerify.ColWidth(GROSSINDEX) = grdVerify.Width * 0.08
    grdVerify.ColWidth(DATESINDEX) = grdVerify.Width * 0.14
    
    grdVerify.ColWidth(ADVTPRODINDEX) = grdVerify.Width - GRIDSCROLLWIDTH - 15
    For ilCol = STATUSINDEX To DATESINDEX Step 1
        If ilCol <> ADVTPRODINDEX Then
            grdVerify.ColWidth(ADVTPRODINDEX) = grdVerify.ColWidth(ADVTPRODINDEX) - grdVerify.ColWidth(ilCol)
        End If
    Next ilCol
    
    'Align columns to left
    gGrid_AlignAllColsLeft grdVerify
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdVerify.TextMatrix(0, STATUSINDEX) = "Verified"
    grdVerify.TextMatrix(0, STATUSDATEINDEX) = "Verify Date"
    grdVerify.TextMatrix(0, CNTRNOINDEX) = "Contract"
    grdVerify.TextMatrix(0, ADVTPRODINDEX) = "Advertiser/Product"
    grdVerify.TextMatrix(0, AGENCYINDEX) = "Agency"
    grdVerify.TextMatrix(0, SALESPERSONINDEX) = "Salesperson"
    grdVerify.TextMatrix(0, GROSSINDEX) = "Gross $'s"
    grdVerify.TextMatrix(0, DATESINDEX) = "Date Range"
    'Set height of grid

End Sub

Private Sub mVerifyGridSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdVerify.FixedRows To grdVerify.Rows - 1 Step 1
        slStr = Trim$(grdVerify.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            If ilCol = DATESINDEX Then
                slStr = grdVerify.TextMatrix(llRow, DATESINDEX)
                If slStr <> "" Then
                    ilPos = InStr(1, slStr, " ", vbBinaryCompare)
                    If ilPos > 0 Then
                        slStr = Left(slStr, ilPos)
                    End If
                    slSort = Trim$(Str$(gDateValue(slStr)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = STATUSDATEINDEX Then
                slStr = grdVerify.TextMatrix(llRow, STATUSDATEINDEX)
                If slStr <> "" Then
                    ilPos = InStr(1, slStr, " ", vbBinaryCompare)
                    If ilPos > 0 Then
                        slStr = Left(slStr, ilPos)
                    End If
                    slSort = Trim$(Str$(gDateValue(slStr)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = GROSSINDEX Then
                slStr = grdVerify.TextMatrix(llRow, GROSSINDEX)
                slStr = gStrDecToLong(slStr, 2)
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdVerify.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdVerify.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imVerifyColSorted) Or ((ilCol = imVerifyColSorted) And (imVerifySort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVerify.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVerify.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imVerifyColSorted Then
        imVerifyColSorted = SORTINDEX
    Else
        imVerifyColSorted = -1
        imVerifySort = -1
    End If
    gGrid_SortByCol grdVerify, CNTRNOINDEX, SORTINDEX, imVerifyColSorted, imVerifySort
    imVerifyColSorted = ilCol
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

    ilRet = btrClose(hmChf)
    btrDestroy hmChf

    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    igManUnload = YES
    Unload CntrVerify
    Set CntrVerify = Nothing   'Remove data segment
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
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffset As Integer
    Dim ilExtLen As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim llRow As Long

    mPopulate 0, llRow
    ilRet = btrGetFirst(hmChf, tmChf, imChfRecLen, INDEXKEY0, BTRV_LOCK_NONE, 0)

    ' Prepare to execute an extended operation.
    btrExtClear hmChf   'Clear any previous extend operation

    ilExtLen = Len(tmChf)
    llNoRec = gExtNoRec(ilExtLen)
    Call btrExtSetBounds(hmChf, llNoRec, -1, "UC", "Chf", "") '"EG") 'Set extract limits (all records)

    ilOffset = gFieldOffset("Chf", "chfVerifyFlag")
    tlCharTypeBuff.sType = sgCntrVerifyType
    ilRet = btrExtAddLogicConst(hmChf, BTRV_KT_STRING, ilOffset, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

    ilOffset = gFieldOffset("Chf", "chfDelete")
    tlCharTypeBuff.sType = "N"
    ilRet = btrExtAddLogicConst(hmChf, BTRV_KT_STRING, ilOffset, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)


    ilRet = btrExtAddField(hmChf, 0, ilExtLen) 'Extract the whole record
    ilRet = btrExtGetNext(hmChf, tmChf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tmChf)  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmChf, tmChf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            mPopulate 1, llRow
            ilExtLen = Len(tmChf)  'Extract operation record size
            ilRet = btrExtGetNext(hmChf, tmChf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmChf, tmChf, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mPopulate 2, llRow
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
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcCancel.Top = cmcDone.Top
    cmcSave.Top = cmcDone.Top
    grdVerify.Move 180, 255, CntrVerify.Width - 360, cmcDone.Top - 255 - 120
    cmcSave.Left = grdVerify.Left + grdVerify.Width - cmcSave.Width
    cmcCancel.Left = cmcSave.Left - cmcCancel.Width - ilGap
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdVerify, fgBoxGridH + 15

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

    'grdVerify.Row = llRow
    'For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
    '    grdVerify.Col = llCol
    '    If grdVerify.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
    '        grdVerify.CellBackColor = LIGHTYELLOW
    '    Else
    '        grdVerify.CellBackColor = GRAY    'vbBlue
    '    End If
    'Next llCol

End Sub

Private Sub grdVerify_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdVerify.RowHeight(0) Then
        grdVerify.Col = grdVerify.MouseCol
        mVerifyGridSortCol grdVerify.Col
        grdVerify.Row = 0
        grdVerify.Col = CHFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdVerify, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdVerify.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdVerify.FixedRows Then
        If grdVerify.TextMatrix(llCurrentRow, STATUSINDEX) <> "" Then
            If llCol = STATUSINDEX Then
                grdVerify.Row = llCurrentRow
                grdVerify.Col = llCol
                mEnableBox
                Exit Sub
            End If
        End If
    End If

End Sub



Private Sub pbcClickFocus_Click()
    mSetShow
End Sub

Private Sub pbcSTab_GotFocus()
    pbcClickFocus.SetFocus
End Sub

Private Sub pbcTab_GotFocus()
    pbcClickFocus.SetFocus
End Sub

Private Sub pbcVerifyStatus_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcVerifyStatus_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
        If imVerifyStatus <> 0 Then
            imChg = True
        End If
        imVerifyStatus = 0
        pbcVerifyStatus_Paint
    ElseIf (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
        If imVerifyStatus <> 1 Then
            imChg = True
        End If
        imVerifyStatus = 1
        pbcVerifyStatus_Paint
    ElseIf KeyAscii = Asc("V") Or (KeyAscii = Asc("v")) Then
        If imVerifyStatus <> 2 Then
            imChg = True
        End If
        imVerifyStatus = 2
        pbcVerifyStatus_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imVerifyStatus = 0 Then
            imVerifyStatus = 1
            pbcVerifyStatus_Paint
            imChg = True
        ElseIf imVerifyStatus = 1 Then
            imVerifyStatus = 2
            pbcVerifyStatus_Paint
            imChg = True
        ElseIf imVerifyStatus = 2 Then
            imVerifyStatus = 0
            pbcVerifyStatus_Paint
            imChg = True
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcVerifyStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imVerifyStatus = 0 Then
        imVerifyStatus = 1
        pbcVerifyStatus_Paint
        imChg = True
    ElseIf imVerifyStatus = 1 Then
        imVerifyStatus = 2
        pbcVerifyStatus_Paint
        imChg = True
    ElseIf imVerifyStatus = 2 Then
        imVerifyStatus = 0
        pbcVerifyStatus_Paint
        imChg = True
    End If
    mSetCommands
End Sub

Private Sub pbcVerifyStatus_Paint()
    pbcVerifyStatus.Cls
    pbcVerifyStatus.CurrentX = fgBoxInsetX
    pbcVerifyStatus.CurrentY = 0 'fgBoxInsetY
    If imVerifyStatus = 0 Then
        pbcVerifyStatus.Print "Not Sent"
    ElseIf imVerifyStatus = 1 Then
        pbcVerifyStatus.Print "Sent"
    ElseIf imVerifyStatus = 2 Then
        pbcVerifyStatus.Print "Verified"
    End If
End Sub

Private Sub mEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilLang                        slNameCode                *
'*  slCode                        ilCode                        ilRet                     *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (grdVerify.Row < grdVerify.FixedRows) Or (grdVerify.Row >= grdVerify.Rows) Or (grdVerify.Col < grdVerify.FixedCols) Or (grdVerify.Col >= grdVerify.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdVerify.Row
    lmEnableCol = grdVerify.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdVerify.Left - pbcArrow.Width - 30, grdVerify.Top + grdVerify.RowPos(grdVerify.Row) + (grdVerify.RowHeight(grdVerify.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True

    Select Case grdVerify.Col
        Case STATUSINDEX
            If grdVerify.Text = "" Then
                imVerifyStatus = 0
            Else
                If grdVerify.Text = "Sent" Then
                    imVerifyStatus = 1
                ElseIf grdVerify.Text = "Verified" Then
                    imVerifyStatus = 2
                Else
                    imVerifyStatus = 0
                End If
            End If
            pbcVerifyStatus_Paint
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
    Dim ilUnits As Integer

    pbcArrow.Visible = False
    If (lmEnableRow >= grdVerify.FixedRows) And (lmEnableRow < grdVerify.Rows) Then
        Select Case lmEnableCol
            Case STATUSINDEX
                pbcVerifyStatus.Visible = False
                If imVerifyStatus = 1 Then
                    grdVerify.TextMatrix(lmEnableRow, lmEnableCol) = "Sent"
                ElseIf imVerifyStatus = 2 Then
                    grdVerify.TextMatrix(lmEnableRow, lmEnableCol) = "Verified"
                Else
                    grdVerify.TextMatrix(lmEnableRow, lmEnableCol) = "Not Sent"
                End If
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imSetCtrlVisible = False
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim llColWidth As Long

    If (grdVerify.Row < grdVerify.FixedRows) Or (grdVerify.Row >= grdVerify.Rows) Or (grdVerify.Col < grdVerify.FixedCols) Or (grdVerify.Col >= grdVerify.Cols - 1) Then
        Exit Sub
    End If
    imSetCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdVerify.Col - 1 Step 1
        llColPos = llColPos + grdVerify.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdVerify.ColWidth(grdVerify.Col)
    ilCol = grdVerify.Col
    Do While ilCol < grdVerify.Cols - 1
        If (Trim$(grdVerify.TextMatrix(grdVerify.Row - 1, grdVerify.Col)) <> "") And (Trim$(grdVerify.TextMatrix(grdVerify.Row - 1, grdVerify.Col)) = Trim$(grdVerify.TextMatrix(grdVerify.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdVerify.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdVerify.Col
        Case STATUSINDEX
            pbcVerifyStatus.Move grdVerify.Left + llColPos + 30, grdVerify.Top + grdVerify.RowPos(grdVerify.Row) + 15, grdVerify.ColWidth(grdVerify.Col), grdVerify.RowHeight(grdVerify.Row) - 15
            pbcVerifyStatus.Visible = True
            pbcVerifyStatus.SetFocus
    End Select
End Sub

Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim llRow As Integer   'For loop control
    Dim ilRet As Integer
    Dim slVerifyStatus As String
    
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdVerify, grdVerify, vbHourglass
    For llRow = grdVerify.FixedRows To grdVerify.Rows - 1 Step 1
        If grdVerify.TextMatrix(llRow, CNTRNOINDEX) <> "" Then
            tmChfSrchKey0.lCode = Val(grdVerify.TextMatrix(llRow, CHFCODEINDEX))
            ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If grdVerify.TextMatrix(llRow, STATUSINDEX) = "Sent" Then
                    slVerifyStatus = "S"
                ElseIf grdVerify.TextMatrix(llRow, STATUSINDEX) = "Verified" Then
                    slVerifyStatus = "V"
                Else
                    slVerifyStatus = "N"
                End If
                If tmChf.sVerifyFlag <> slVerifyStatus Then
                    tmChf.sVerifyFlag = slVerifyStatus
                    gPackDate Format(gNow(), "m/d/yy"), tmChf.iVerifyDate(0), tmChf.iVerifyDate(1)
                    gPackTime Format(gNow(), "h:mm:ssAM/PM"), tmChf.iVerifyTime(0), tmChf.iVerifyTime(1)
                    ilRet = btrUpdate(hmChf, tmChf, imChfRecLen)
                End If
            End If
        End If
    Next llRow
    imChg = False
    mSaveRec = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVerify, grdVerify, vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    If imChg Then
        cmcSave.Enabled = True
    Else
        cmcSave.Enabled = False
    End If
End Sub

