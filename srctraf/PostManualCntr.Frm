VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PostManualCntr 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PostManualCntr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "PostManualCntr.frx":08CA
      Left            =   330
      List            =   "PostManualCntr.frx":08CC
      TabIndex        =   11
      Top             =   495
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.ListBox lbcCntr 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "PostManualCntr.frx":08CE
      Left            =   2040
      List            =   "PostManualCntr.frx":08D0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4230
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   795
      Top             =   4680
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "Save-&Changes"
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
      Left            =   3900
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
      Picture         =   "PostManualCntr.frx":08D2
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Save-&All"
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
      TabIndex        =   7
      Top             =   4530
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdManualPost 
      Height          =   3690
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6509
      _Version        =   393216
      Rows            =   3
      Cols            =   30
      FixedRows       =   2
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
      _Band(0).Cols   =   30
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   255
      Picture         =   "PostManualCntr.frx":0BDC
      Top             =   240
      Width           =   480
   End
   Begin VB.Label plcScreen 
      Caption         =   "Manual Posting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   285
      TabIndex        =   0
      Top             =   30
      Width           =   6930
   End
End
Attribute VB_Name = "PostManualCntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PostManualCntr.FRM on Wed 6/17/09 @ 12:56 P
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
'*  PostManualCntr - displays missed spots to be changed to Makegoods
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
Private imCtrlVisible As Integer
Private imManualPostStatus As Integer
Private imChg As Integer
Private imManualPostColSorted As Integer
Private lmInvStartDate As Long
Private lmInvEndDate As Long
Private imManualPostSort As Integer
Private imFromArrow As Integer
Private lmTopRow As Long
Private imInitNoRows As Integer
Private imIgnoreScroll As Integer
Private imLastWeekNo As Integer
Private lmMaxHeight As Long
Private imUpdateAllowed As Integer
Private lmLastStdMnthBilled As Long

'Contract
Dim tmChf As CHF        'Chf record image
Dim tmChfSrchKey0 As LONGKEY0    'Chf key record image
Dim hmCHF As Integer    'Contract file handle
Dim imCHFRecLen As Integer        'CHF record length

Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey0 As CLFKEY0 'CLF key record image
Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
Dim imClfRecLen As Integer     'CLF record length

Dim hmCff As Integer        'Contract Flight file handle
Dim tmCff As CFF            'CFF record image
Dim tmCffSrchKey0 As CFFKEY0 'CFF key record image
Dim imCffRecLen As Integer     'CFF record length

Dim hmCgf As Integer        'Contract Flight file handle
Dim tmCgf As CGF            'CFF record image
Dim tmCgfSrchKey As CFFKEY0 'CFF key record image
Dim imCgfRecLen As Integer     'CFF record length

Dim hmSdf As Integer    'file handle
Dim imSdfRecLen As Integer  'Record length
Dim tmSdf As SDF
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey3 As LONGKEY0

Dim hmCif As Integer
Dim tmCif As CIF        'Rvf record image
Dim tmCifSrchKey As LONGKEY0
Dim imCifRecLen As Integer        'RvF record length

Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length

Dim hmSsf As Integer        'Spot summary file handle
Dim lmSsfDate(0 To 6) As Long    'Dates of the days stored into tmSsf
Dim lmSsfRecPos(0 To 6) As Long  'Record positions
Dim tmSsf(0 To 6) As SSF         'Spot summary for one week (0 index for monday; 1 for tuesday;...; 6 for sunday)
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim imSsfRecLen As Integer     'SSF record length
Dim imSelectedDay As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim tmClfList() As CLFLIST            'CLF record image
Dim tmCfflist() As CFFLIST            'CFF record image
Dim tmCgfList() As CGFLIST

Dim hmSmf As Integer
Dim hmCrf As Integer
Dim hmGsf As Integer
Dim hmGhf As Integer
Dim hmSxf As Integer

Dim tmRdf As RDF
Dim tmSmf As SMF

Private tmCntr() As SORTCODE
Private tmSdfExtSort() As SDFEXTSORT
Private tmSdfExt() As SDFEXT
Private Type SDFDATETIMEINFO
    lSdfCode As Long
    sSpotSchStatus As String * 1
    sTimeSchStatus As String * 1
    sAirDate As String * 10
    sAirTime As String * 11
    sISCI As String * 20
    lCifCode As Long
    iWeek As Integer
    lNext As Long
End Type
Private tmSdfDateTimeInfo() As SDFDATETIMEINFO

Dim hmIihf As Integer
Dim tmIihf As IIHF        'IIHF record image
Dim tmIihfSrchKey0 As LONGKEY0    'IIHF key record image
Dim tmIihfSrchKey1 As IIHFKEY1    'IIHF key record image
Dim tmIihfSrchKey2 As IIHFKEY2    'IIHF key record image
Dim tmIihfSrchKey3 As IIHFKEY3    'IIHF key record image
Dim imIihfRecLen As Integer        'IIHF record length

Dim hmApf As Integer
Dim tmApf As APF        'CFF record image
Dim tmApfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmApfSrchKey4 As APFKEY4
Dim imApfRecLen As Integer        'CFF record length

Const ADVTNAMEINDEX = 0
Const CNTRNOINDEX = 1
Const AGYESTNOINDEX = 2
Const DAYPARTTIMESINDEX = 3
Const DAYPARTDAYSINDEX = 4
Const LENGTHINDEX = 5
Const ACQCOSTINDEX = 6
Const SOLDSPOTSINDEX = 7
Const STATIONCNTRNOINDEX = 8
Const STAIONINVNOINDEX = 9
Const STAIONESTNOINDEX = 10
Const WEEK1AIREDINDEX = 11
Const WEEK2AIREDINDEX = 12
Const WEEK3AIREDINDEX = 13
Const WEEK4AIREDINDEX = 14
Const WEEK5AIREDINDEX = 15
Const TOTALINDEX = 16
Const AFFIDAVITINDEX = 17
Const POSTTIMESINDEX = 18
Const SORTINDEX = 19
Const CHFCODEINDEX = 20
Const IIHFCODEINDEX = 21
Const WEEK1SOLDINDEX = 22
Const WEEK2SOLDINDEX = 23
Const WEEK3SOLDINDEX = 24
Const WEEK4SOLDINDEX = 25
Const WEEK5SOLDINDEX = 26
Const SOURCEFORMINDEX = 27
Const CHGINDEX = 28
Const TIMESINDEX = 29



Private Sub mClearGrid()

    Dim llRow As Long
    Dim llCol As Long

    grdManualPost.Redraw = False
    grdManualPost.RowHeight(0) = fgFlexGridRowH
    For llRow = grdManualPost.FixedRows To grdManualPost.Rows - 1 Step 1
        For llCol = ADVTNAMEINDEX To TIMESINDEX Step 1
            imIgnoreScroll = False
            grdManualPost.Row = llRow
            grdManualPost.Col = llCol
            grdManualPost.TextMatrix(llRow, llCol) = ""
            If (llCol <= SOLDSPOTSINDEX) Or (llCol = TOTALINDEX) Or (llCol = AFFIDAVITINDEX) Then
                grdManualPost.CellBackColor = LIGHTYELLOW
            'ElseIf (llCol = POSTTIMESINDEX) Then
            '    If grdManualPost.TextMatrix(llRow, CNTRNOINDEX) <> grdManualPost.TextMatrix(llRow - 1, CNTRNOINDEX) Then
            '        grdManualPost.CellBackColor = GRAY
            '    End If
            Else
                grdManualPost.CellBackColor = vbWhite
            End If
        Next llCol
        grdManualPost.RowHeight(llRow) = fgFlexGridRowH
    Next llRow
    'grdManualPost.Redraw = True
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim llRow As Long
    
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdManualPost, grdManualPost, vbHourglass
    For llRow = grdManualPost.FixedRows To grdManualPost.Rows - 1 Step 1
        If Trim(grdManualPost.TextMatrix(llRow, CNTRNOINDEX)) <> "" Then
            grdManualPost.TextMatrix(llRow, CHGINDEX) = "Y"
        End If
    Next llRow
    ilRet = mSaveRec()
    If Not ilRet Then
        Exit Sub
    End If
    ilRet = mReadRec()
    mSetCommands
    cmcCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    'mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdManualPost, grdManualPost, vbHourglass
    imManualPostColSorted = -1
    imManualPostSort = -1
    ilRet = mSaveRec()
    If Not ilRet Then
        Exit Sub
    End If
    ilRet = mReadRec()
    mSetCommands
    cmcCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slMaxValue As String
    Dim slStr As String
    Dim ilSoldAdj As Integer
    
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case lmEnableCol
        Case WEEK1AIREDINDEX, WEEK2AIREDINDEX, WEEK3AIREDINDEX, WEEK4AIREDINDEX, WEEK5AIREDINDEX, TOTALINDEX
            ilSoldAdj = WEEK1SOLDINDEX - WEEK1AIREDINDEX
            slMaxValue = grdManualPost.TextMatrix(lmEnableRow, lmEnableCol + ilSoldAdj)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, slMaxValue) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select

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
        gSetMousePointer grdManualPost, grdManualPost, vbDefault
    End If
    If (igWinStatus(POSTLOGSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    'Me.Width = (CLng(75) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    'Me.Height = (CLng(75) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    'gCenterStdAlone PostManualCntr
    'DoEvents
    'mSetControls
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdManualPost, grdManualPost, vbHourglass
    Me.Width = (CLng(75) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(75) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    DoEvents
    mSetControls
    mInit
    gCenterStdAlone PostManualCntr
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    If imTerminate Then
        tmcStart.Enabled = True
    End If
    Exit Sub

End Sub

Private Sub Form_Terminate()
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmSdfDateTimeInfo
    Erase tmClfList
    Erase tmCfflist
    Erase tmCgfList
    Erase tmSdfExtSort
    Erase tmSdfExt
    Erase tgPostManualTimesInfo
    
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmApf)
    btrDestroy hmApf
    ilRet = btrClose(hmIihf)
    btrDestroy hmIihf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    igManUnload = YES
    Set PostManualCntr = Nothing   'Remove data segment
    igManUnload = NO
    
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim ilVef As Integer

    gSetMousePointer grdManualPost, grdManualPost, vbHourglass
    imTerminate = False
    plcScreen.Caption = "Manual Posting: "
    ilVef = gBinarySearchVef(igManualPostVefCode)
    If ilVef = -1 Then
        imTerminate = True
        Exit Sub
    End If
    plcScreen.Caption = plcScreen.Caption & Trim$(tgMVef(ilVef).sName)
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), lmLastStdMnthBilled
    lmLastClickedRow = -1
    lmScrollTop = grdManualPost.FixedRows
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False
    imManualPostColSorted = -1
    imManualPostSort = -1
    imChg = False

    hmCHF = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ARF record length

    hmClf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imClfRecLen = Len(tmClf)  'Get and save ARF record length
    
    hmCff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imCffRecLen = Len(tmCff)  'Get and save ARF record length
    
    hmCgf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)  'Get and save ARF record length

    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save ARF record length
    
    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save ARF record length
    
    hmCpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save ARF record length

    hmIihf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imIihfRecLen = Len(tmIihf)  'Get and save ARF record length

    hmApf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmApf, "", sgDBPath & "Apf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    imApfRecLen = Len(tmApf)  'Get and save ARF record length

    hmSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0

    hmCrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    hmSxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PostManualCntr
    On Error GoTo 0
    
    mPopListKey
    
    ilRet = mReadRec()
    mSetCommands
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate(ilFlag As Integer, llCurrentRow As Long, ilClf As Integer)
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim ilRdf As Integer
    Dim ilCff As Integer
    Dim slClfStartTime As String
    Dim slClfEndTime As String
    Dim ilWeek As Integer
    Dim llWeekStartDate As Long
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilDay As Integer
    Dim ilNoSpots As Integer
    Dim ilSchSpots As Integer
    Dim ilTotalAiredSpots As Integer
    Dim ilTotalSoldSpots As Integer
    Dim llSdfIndex As Long
    Dim slWkSDate As String
    Dim slWkEDate As String
    Dim llForeColor As Long
    Dim ilTest As Integer
    Dim slAllowedDays As String
    Dim ilCount As Integer
    Dim llMissedTime As Long
    Dim ilTBIndex As Integer
    Dim llTBStartTime(0 To 6) As Long  'Allowed times if time buy
    Dim llTBEndTime(0 To 6) As Long
    
    On Error GoTo ErrHand:

    If ilFlag = 0 Then
        grdManualPost.Redraw = False
        grdManualPost.Row = 0
        'For llCol = ADVTNAMEINDEX To AGYESTNOINDEX Step 1
        '    grdManualPost.Col = llCol
        '    grdManualPost.CellBackColor = vbBlack
        '    grdManualPost.CellBackColor = LIGHTBLUE
        'Next llCol
        grdManualPost.RowHeight(0) = fgFlexGridRowH
        grdManualPost.RowHeight(1) = fgFlexGridRowH
        llCurrentRow = grdManualPost.FixedRows
        lmMaxHeight = 2 * (fgFlexGridRowH)
        imManualPostColSorted = -1
        imManualPostSort = -1
        Exit Sub
    End If
    If ilFlag = 1 Then
        llRow = llCurrentRow
        If llRow >= grdManualPost.Rows Then
            grdManualPost.AddItem ""
        End If
        grdManualPost.RowHeight(llRow) = fgFlexGridRowH
        lmMaxHeight = lmMaxHeight + fgFlexGridRowH
        ilAdf = gBinarySearchAdf(tmChf.iAdfCode)
        If ilAdf <> -1 Then
            grdManualPost.TextMatrix(llRow, ADVTNAMEINDEX) = Trim$(tgCommAdf(ilAdf).sName)
        Else
            grdManualPost.TextMatrix(llRow, ADVTNAMEINDEX) = Trim$(Str(tmChf.iAdfCode))
        End If
        grdManualPost.TextMatrix(llRow, CNTRNOINDEX) = tmChf.lCntrNo
        grdManualPost.TextMatrix(llRow, AGYESTNOINDEX) = Trim$(tmChf.sAgyEstNo) & Trim$(tmChf.sTitle)
        If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
            ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
            If ilRdf <> -1 Then
                grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) = ""
                For ilTest = UBound(tgMRdf(ilRdf).iStartTime, 2) To LBound(tgMRdf(ilRdf).iStartTime, 2) Step -1
                    If (tgMRdf(ilRdf).iStartTime(0, ilTest) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilTest) <> 0) Then
                        gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilTest), tgMRdf(ilRdf).iStartTime(1, ilTest), "A", "1", slClfStartTime
                        gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilTest), tgMRdf(ilRdf).iEndTime(1, ilTest), "A", "1", slClfEndTime
                        If grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) = "" Then
                            grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) = slClfStartTime & "-" & slClfEndTime
                            ilTBIndex = 0
                            llTBStartTime(ilTBIndex) = gTimeToLong(slClfStartTime, False)
                            llTBEndTime(ilTBIndex) = gTimeToLong(slClfEndTime, True)
                        Else
                            grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) = grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) & "; " & slClfStartTime & "-" & slClfEndTime
                            ilTBIndex = ilTBIndex + 1
                            llTBStartTime(ilTBIndex) = gTimeToLong(slClfStartTime, False)
                            llTBEndTime(ilTBIndex) = gTimeToLong(slClfEndTime, True)
                        End If
                    End If
                Next ilTest
            Else
                grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) = Trim$(Str(tmClf.iRdfCode))
            End If
        Else
            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slClfStartTime
            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slClfEndTime
            grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX) = slClfStartTime & "-" & slClfEndTime
            ilTBIndex = 0
            llTBStartTime(ilTBIndex) = gTimeToLong(slClfStartTime, False)
            llTBEndTime(ilTBIndex) = gTimeToLong(slClfEndTime, True)
        End If
        grdManualPost.TextMatrix(llRow, LENGTHINDEX) = tmClf.iLen
        grdManualPost.TextMatrix(llRow, ACQCOSTINDEX) = gLongToStrDec(tmClf.lAcquisitionCost, 2)
        If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Then
            grdManualPost.TextMatrix(llRow, STATIONCNTRNOINDEX) = Trim$(tmIihf.sStnContractNo)
            grdManualPost.TextMatrix(llRow, STAIONINVNOINDEX) = Trim$(tmIihf.sStnInvoiceNo)
            grdManualPost.TextMatrix(llRow, STAIONESTNOINDEX) = Trim$(tmIihf.sStnEstimateNo)
        Else
            grdManualPost.TextMatrix(llRow, STATIONCNTRNOINDEX) = ""
            grdManualPost.TextMatrix(llRow, STAIONINVNOINDEX) = ""
            grdManualPost.TextMatrix(llRow, STAIONESTNOINDEX) = ""
        End If
        ilTotalAiredSpots = 0
        ilTotalSoldSpots = 0
        grdManualPost.TextMatrix(llRow, WEEK1AIREDINDEX) = ""
        grdManualPost.TextMatrix(llRow, WEEK2AIREDINDEX) = ""
        grdManualPost.TextMatrix(llRow, WEEK3AIREDINDEX) = ""
        grdManualPost.TextMatrix(llRow, WEEK4AIREDINDEX) = ""
        grdManualPost.TextMatrix(llRow, WEEK5AIREDINDEX) = ""
        grdManualPost.TextMatrix(llRow, WEEK1SOLDINDEX) = "0"
        grdManualPost.TextMatrix(llRow, WEEK2SOLDINDEX) = "0"
        grdManualPost.TextMatrix(llRow, WEEK3SOLDINDEX) = "0"
        grdManualPost.TextMatrix(llRow, WEEK4SOLDINDEX) = "0"
        grdManualPost.TextMatrix(llRow, WEEK5SOLDINDEX) = "0"
        grdManualPost.TextMatrix(llRow, DAYPARTDAYSINDEX) = ""
        grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX) = "C"
        grdManualPost.TextMatrix(llRow, CHGINDEX) = "N"
        grdManualPost.TextMatrix(llRow, TIMESINDEX) = ""
        grdManualPost.Row = llRow
        For ilWeek = 1 To 5 Step 1
            llWeekStartDate = gDateValue(DateAdd("d", 7 * (ilWeek - 1), sgManualPostDate))
            If (llWeekStartDate >= lmInvStartDate) And (llWeekStartDate <= lmInvEndDate) Then
                ilCff = tmClfList(ilClf).iFirstCff
                Do While ilCff <> -1
                    tmCff = tmCfflist(ilCff).CffRec
                    gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llCffStartDate
                    gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llCffEndDate
                    If (llWeekStartDate >= llCffStartDate) And (llWeekStartDate <= llCffEndDate) Then
                        If tmCff.sDyWk <> "D" Then
                            ilNoSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                            slAllowedDays = ""
                            For ilDay = 0 To 6 Step 1
                                If (tmCff.iDay(ilDay) > 0) Or (tmCff.sXDay(ilDay) = "1") Then
                                    slAllowedDays = slAllowedDays & "Y"
                                Else
                                    slAllowedDays = slAllowedDays & "N"
                                End If
                            Next ilDay
                            If grdManualPost.TextMatrix(llRow, DAYPARTDAYSINDEX) = "" Then
                                grdManualPost.TextMatrix(llRow, DAYPARTDAYSINDEX) = gDayStrToNames(slAllowedDays, slAllowedDays, 1)
                            End If
                       Else    'Daily
                            ilNoSpots = 0
                            For ilDay = 0 To 6 Step 1
                                ilNoSpots = ilNoSpots + tmCff.iDay(ilDay)
                            Next ilDay
                            slAllowedDays = ""
                            For ilDay = 0 To 6 Step 1
                                If (tmCff.iDay(ilDay) > 0) Then
                                    slAllowedDays = slAllowedDays & "Y"
                                Else
                                    slAllowedDays = slAllowedDays & "N"
                                End If
                            Next ilDay
                            If grdManualPost.TextMatrix(llRow, DAYPARTDAYSINDEX) = "" Then
                                grdManualPost.TextMatrix(llRow, DAYPARTDAYSINDEX) = gDayStrToNames(slAllowedDays, slAllowedDays, 1)
                            End If
                        End If
                        ilTotalSoldSpots = ilTotalSoldSpots + ilNoSpots
                        Select Case ilWeek
                            Case 1
                                grdManualPost.TextMatrix(llRow, WEEK1SOLDINDEX) = ilNoSpots
                            Case 2
                                grdManualPost.TextMatrix(llRow, WEEK2SOLDINDEX) = ilNoSpots
                            Case 3
                                grdManualPost.TextMatrix(llRow, WEEK3SOLDINDEX) = ilNoSpots
                            Case 4
                                grdManualPost.TextMatrix(llRow, WEEK4SOLDINDEX) = ilNoSpots
                            Case 5
                                grdManualPost.TextMatrix(llRow, WEEK5SOLDINDEX) = ilNoSpots
                        End Select
                        'Get schedule spot count
                        ilSchSpots = 0
                        slWkSDate = DateAdd("d", 7 * (ilWeek - 1), sgManualPostDate)
                        slWkEDate = DateAdd("d", 6, slWkSDate)
                        ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
                        ReDim tmSdfExt(0 To 0) As SDFEXT
                        ilRet = gObtainCntrSpot(CLng(tmClf.iVefCode), False, tmClf.lChfCode, tmClf.iLine, "S", slWkSDate, slWkEDate, tmSdfExtSort(), tmSdfExt(), 0, False)
                        For llSdfIndex = LBound(tmSdfExt) To UBound(tmSdfExt) - 1 Step 1
                            If tmSdfExt(llSdfIndex).iLineNo = tmClf.iLine Then
                                If (tmSdfExt(llSdfIndex).sSpotType <> "X") And (tmSdfExt(llSdfIndex).sSpotType <> "O") And (tmSdfExt(llSdfIndex).sSpotType <> "C") Then
                                    If grdManualPost.TextMatrix(llRow, TIMESINDEX) = "" Then
                                        grdManualPost.TextMatrix(llRow, TIMESINDEX) = UBound(tmSdfDateTimeInfo)
                                    Else
                                        tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo) - 1).lNext = UBound(tmSdfDateTimeInfo)
                                    End If
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).lSdfCode = tmSdfExt(llSdfIndex).lCode
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sSpotSchStatus = tmSdfExt(llSdfIndex).sSchStatus
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sTimeSchStatus = tmSdfExt(llSdfIndex).sSchStatus
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).lCifCode = 0
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sISCI = ""
                                    If (tmSdfExt(llSdfIndex).sSchStatus = "S") Or (tmSdfExt(llSdfIndex).sSchStatus = "O") Or (tmSdfExt(llSdfIndex).sSchStatus = "G") Then
                                        ilSchSpots = ilSchSpots + 1
                                        gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sAirDate
                                        gUnpackTime tmSdfExt(llSdfIndex).iTime(0), tmSdfExt(llSdfIndex).iTime(1), "A", "1", tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sAirTime
                                        tmSdfSrchKey3.lCode = tmSdfExt(llSdfIndex).lCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If tmSdf.sPtType = "1" Then
                                                tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).lCifCode = tmSdf.lCopyCode
                                                tmCifSrchKey.lCode = tmSdf.lCopyCode
                                                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                If (ilRet = BTRV_ERR_NONE) Then
                                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).lCifCode = tmSdf.lCopyCode
                                                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                                                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                        tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sISCI = tmCpf.sISCI
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        'Radomily pick day and time
                                        Do
                                            ilDay = Int((7) * Rnd)  '0 to 6
                                            If Mid(slAllowedDays, ilDay + 1, 1) = "Y" Then
                                                tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sAirDate = DateAdd("d", ilDay, slWkSDate)
                                                Exit Do
                                            End If
                                        Loop
                                        ilCount = 0
                                        Do
                                            If ((ilTBIndex = 0) And (llTBEndTime(0) - llTBStartTime(0) < 3600)) Or (ilCount >= 100) Then
                                                llMissedTime = llTBStartTime(0)
                                                Exit Do
                                            Else
                                                llMissedTime = 60 * CLng(Int(1440 * Rnd))  '0-1439
                                                ilCount = ilCount + 1
                                                For ilLoop = 0 To ilTBIndex Step 1
                                                    If (llMissedTime >= llTBStartTime(ilLoop)) And (llMissedTime < llTBEndTime(ilLoop)) Then
                                                        Exit Do
                                                    End If
                                                Next ilLoop
                                            End If
                                        Loop
                                        tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).sAirTime = gFormatTimeLong(llMissedTime, "A", "1")
                                    End If
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).iWeek = ilWeek
                                    tmSdfDateTimeInfo(UBound(tmSdfDateTimeInfo)).lNext = -1
                                    ReDim Preserve tmSdfDateTimeInfo(0 To UBound(tmSdfDateTimeInfo) + 1) As SDFDATETIMEINFO
                                End If
                            End If
                        Next llSdfIndex
                        If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Then
                            llForeColor = vbBlack
                            ilNoSpots = ilSchSpots
                            grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX) = tmIihf.sSourceForm
                        Else
                            llForeColor = CYAN
                            grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX) = "C"
                        End If
                        Select Case ilWeek
                            Case 1
                                grdManualPost.Col = WEEK1AIREDINDEX
                                grdManualPost.CellForeColor = llForeColor
                                grdManualPost.TextMatrix(llRow, WEEK1AIREDINDEX) = ilNoSpots
                            Case 2
                                grdManualPost.Col = WEEK2AIREDINDEX
                                grdManualPost.CellForeColor = llForeColor
                                grdManualPost.TextMatrix(llRow, WEEK2AIREDINDEX) = ilNoSpots
                            Case 3
                                grdManualPost.Col = WEEK3AIREDINDEX
                                grdManualPost.CellForeColor = llForeColor
                                grdManualPost.TextMatrix(llRow, WEEK3AIREDINDEX) = ilNoSpots
                            Case 4
                                grdManualPost.Col = WEEK4AIREDINDEX
                                grdManualPost.CellForeColor = llForeColor
                                grdManualPost.TextMatrix(llRow, WEEK4AIREDINDEX) = ilNoSpots
                            Case 5
                                grdManualPost.Col = WEEK5AIREDINDEX
                                grdManualPost.CellForeColor = llForeColor
                                grdManualPost.TextMatrix(llRow, WEEK5AIREDINDEX) = ilNoSpots
                        End Select
                        ilTotalAiredSpots = ilTotalAiredSpots + ilNoSpots
                        Exit Do
                    End If
                    ilCff = tmCfflist(ilCff).iNextCff
                Loop
            End If
        Next ilWeek
        If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Then
            grdManualPost.Col = TOTALINDEX
            grdManualPost.CellForeColor = vbBlack
        Else
            grdManualPost.Col = TOTALINDEX
            grdManualPost.CellForeColor = CYAN
        End If
        grdManualPost.TextMatrix(llRow, TOTALINDEX) = ilTotalAiredSpots
        grdManualPost.TextMatrix(llRow, SOLDSPOTSINDEX) = ilTotalSoldSpots
        If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Then
            grdManualPost.TextMatrix(llRow, IIHFCODEINDEX) = tmIihf.lCode
        Else
            grdManualPost.TextMatrix(llRow, IIHFCODEINDEX) = ""
        End If
        If (Trim$(tmIihf.sSourceForm) = "T") Then
            grdManualPost.TextMatrix(llRow, AFFIDAVITINDEX) = "Date/Time"
        Else
            grdManualPost.TextMatrix(llRow, AFFIDAVITINDEX) = "Count"
        End If
        grdManualPost.TextMatrix(llRow, CHFCODEINDEX) = tmChf.lCode
        If ilTotalSoldSpots > 0 Then
            llCurrentRow = llCurrentRow + 1
        End If
        'mGridSortCol DAYPARTTIMESINDEX
        'mGridSortCol CNTRNOINDEX
        'mGridSortCol ADVTNAMEINDEX
        Exit Sub
    End If
    If ilFlag = 2 Then
        mGridSortCol DAYPARTTIMESINDEX
        mGridSortCol CNTRNOINDEX
        mGridSortCol ADVTNAMEINDEX
        For llRow = grdManualPost.FixedRows To grdManualPost.Rows - 1 Step 1
            slStr = Trim$(grdManualPost.TextMatrix(llRow, CNTRNOINDEX))
            If slStr <> "" Then
                grdManualPost.Row = llRow
                For llCol = ADVTNAMEINDEX To SOLDSPOTSINDEX Step 1
                    grdManualPost.Col = llCol
                    grdManualPost.CellBackColor = LIGHTYELLOW
                Next llCol
                grdManualPost.Col = TOTALINDEX
                grdManualPost.CellBackColor = LIGHTYELLOW
                grdManualPost.Col = AFFIDAVITINDEX
                grdManualPost.CellBackColor = LIGHTYELLOW
                grdManualPost.Col = POSTTIMESINDEX
                If grdManualPost.TextMatrix(llRow, CNTRNOINDEX) <> grdManualPost.TextMatrix(llRow - 1, CNTRNOINDEX) Then
                    grdManualPost.CellBackColor = GRAY
                Else
                    grdManualPost.CellBackColor = LIGHTYELLOW
                    grdManualPost.Col = STATIONCNTRNOINDEX
                    grdManualPost.CellBackColor = LIGHTYELLOW
                    grdManualPost.Col = STAIONINVNOINDEX
                    grdManualPost.CellBackColor = LIGHTYELLOW
                    grdManualPost.Col = STAIONESTNOINDEX
                    grdManualPost.CellBackColor = LIGHTYELLOW
                End If
            End If
        Next llRow
        gGrid_AlignAllColsLeft grdManualPost
        'mGridSortCol CNTRNOINDEX
        grdManualPost.Redraw = True
    Exit Sub
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    Dim slDate As String
    Dim slDateWk2 As String
    Dim slDateWk5 As String
    
    slDate = sgManualPostDate
    slDateWk2 = Format(DateAdd("d", 7, slDate), "m/d/yy")
    slDateWk5 = Format(DateAdd("d", 34, slDate), "m/d/yy")

    grdManualPost.ColWidth(TIMESINDEX) = 0
    grdManualPost.ColWidth(CHGINDEX) = 0
    grdManualPost.ColWidth(SOURCEFORMINDEX) = 0
    grdManualPost.ColWidth(IIHFCODEINDEX) = 0
    grdManualPost.ColWidth(CHFCODEINDEX) = 0
    grdManualPost.ColWidth(SORTINDEX) = 0
    grdManualPost.ColWidth(WEEK1SOLDINDEX) = 0
    grdManualPost.ColWidth(WEEK2SOLDINDEX) = 0
    grdManualPost.ColWidth(WEEK3SOLDINDEX) = 0
    grdManualPost.ColWidth(WEEK4SOLDINDEX) = 0
    grdManualPost.ColWidth(WEEK5SOLDINDEX) = 0
    grdManualPost.ColWidth(CNTRNOINDEX) = grdManualPost.Width * 0.05
    grdManualPost.ColWidth(AGYESTNOINDEX) = grdManualPost.Width * 0.05
    grdManualPost.ColWidth(DAYPARTTIMESINDEX) = grdManualPost.Width * 0.06
    grdManualPost.ColWidth(DAYPARTDAYSINDEX) = grdManualPost.Width * 0.05
    grdManualPost.ColWidth(LENGTHINDEX) = grdManualPost.Width * 0.04
    grdManualPost.ColWidth(ACQCOSTINDEX) = grdManualPost.Width * 0.06
    grdManualPost.ColWidth(SOLDSPOTSINDEX) = grdManualPost.Width * 0.04
    grdManualPost.ColWidth(STATIONCNTRNOINDEX) = grdManualPost.Width * 0.06
    grdManualPost.ColWidth(STAIONINVNOINDEX) = grdManualPost.Width * 0.06
    grdManualPost.ColWidth(STAIONESTNOINDEX) = grdManualPost.Width * 0.06
    grdManualPost.ColWidth(WEEK1AIREDINDEX) = grdManualPost.Width * 0.04
    grdManualPost.ColWidth(WEEK2AIREDINDEX) = grdManualPost.Width * 0.04
    grdManualPost.ColWidth(WEEK3AIREDINDEX) = grdManualPost.Width * 0.04
    grdManualPost.ColWidth(WEEK4AIREDINDEX) = grdManualPost.Width * 0.04
    If Format(slDateWk2, "m") = Format(slDateWk5, "m") Then
        grdManualPost.ColWidth(WEEK5AIREDINDEX) = grdManualPost.Width * 0.04
        imLastWeekNo = WEEK5AIREDINDEX
    Else
        grdManualPost.ColWidth(WEEK5AIREDINDEX) = 0
        imLastWeekNo = WEEK4AIREDINDEX
    End If
    grdManualPost.ColWidth(TOTALINDEX) = grdManualPost.Width * 0.04
    grdManualPost.ColWidth(AFFIDAVITINDEX) = grdManualPost.Width * 0.05
    grdManualPost.ColWidth(POSTTIMESINDEX) = grdManualPost.Width * 0.04
    
    grdManualPost.ColWidth(ADVTNAMEINDEX) = grdManualPost.Width - GRIDSCROLLWIDTH - 15
    For ilCol = ADVTNAMEINDEX To POSTTIMESINDEX Step 1
        If ilCol <> ADVTNAMEINDEX Then
            grdManualPost.ColWidth(ADVTNAMEINDEX) = grdManualPost.ColWidth(ADVTNAMEINDEX) - grdManualPost.ColWidth(ilCol)
        End If
    Next ilCol
    
    'Align columns to left
    gGrid_AlignAllColsLeft grdManualPost
End Sub

Private Sub mSetGridTitles()
    Dim slDate As String
    slDate = sgManualPostDate
    'Set column titles
    grdManualPost.Row = 0
    grdManualPost.Col = ADVTNAMEINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, ADVTNAMEINDEX) = "Advertiser"
    grdManualPost.Row = 0
    grdManualPost.Col = CNTRNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, CNTRNOINDEX) = "Contract"
    grdManualPost.Row = 1
    grdManualPost.Col = CNTRNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, CNTRNOINDEX) = "#"
    grdManualPost.Row = 0
    grdManualPost.Col = AGYESTNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, AGYESTNOINDEX) = "Estimate"
    grdManualPost.Row = 1
    grdManualPost.Col = AGYESTNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, AGYESTNOINDEX) = "#"
    grdManualPost.Row = 0
    grdManualPost.Col = DAYPARTTIMESINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, DAYPARTTIMESINDEX) = "Daypart"
    grdManualPost.Row = 1
    grdManualPost.Col = DAYPARTTIMESINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, DAYPARTTIMESINDEX) = "Times"
    grdManualPost.Row = 0
    grdManualPost.Col = DAYPARTDAYSINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, DAYPARTDAYSINDEX) = "Daypart"
    grdManualPost.Row = 1
    grdManualPost.Col = DAYPARTDAYSINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, DAYPARTDAYSINDEX) = "Days"
    grdManualPost.Row = 0
    grdManualPost.Col = LENGTHINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, LENGTHINDEX) = "Len"
    grdManualPost.Row = 0
    grdManualPost.Col = ACQCOSTINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, ACQCOSTINDEX) = "Acquisition"
    grdManualPost.Row = 1
    grdManualPost.Col = ACQCOSTINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, ACQCOSTINDEX) = "Cost"
    grdManualPost.Row = 0
    grdManualPost.Col = SOLDSPOTSINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, SOLDSPOTSINDEX) = "Sold"
    grdManualPost.Row = 1
    grdManualPost.Col = SOLDSPOTSINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, SOLDSPOTSINDEX) = "#"
    grdManualPost.Row = 0
    grdManualPost.Col = STATIONCNTRNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, STATIONCNTRNOINDEX) = "Station"
    grdManualPost.Row = 1
    grdManualPost.Col = STATIONCNTRNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, STATIONCNTRNOINDEX) = "Contract #"
    grdManualPost.Row = 0
    grdManualPost.Col = STAIONINVNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, STAIONINVNOINDEX) = "Station"
    grdManualPost.Row = 1
    grdManualPost.Col = STAIONINVNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, STAIONINVNOINDEX) = "Invoice #"
    grdManualPost.Row = 0
    grdManualPost.Col = STAIONESTNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, STAIONESTNOINDEX) = "Station"
    grdManualPost.Row = 1
    grdManualPost.Col = STAIONESTNOINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, STAIONESTNOINDEX) = "Estimate #"
    grdManualPost.Row = 0
    grdManualPost.Col = WEEK1AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, WEEK1AIREDINDEX) = Format(slDate, "m/d/yy")
    grdManualPost.Row = 1
    grdManualPost.Col = WEEK1AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, WEEK1AIREDINDEX) = Format(DateAdd("d", 6, slDate), "m/d/yy")
    grdManualPost.Row = 0
    grdManualPost.Col = WEEK2AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, WEEK2AIREDINDEX) = Format(DateAdd("d", 7, slDate), "m/d/yy")
    grdManualPost.Row = 1
    grdManualPost.Col = WEEK2AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, WEEK2AIREDINDEX) = Format(DateAdd("d", 13, slDate), "m/d/yy")
    grdManualPost.Row = 0
    grdManualPost.Col = WEEK3AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, WEEK3AIREDINDEX) = Format(DateAdd("d", 14, slDate), "m/d/yy")
    grdManualPost.Row = 1
    grdManualPost.Col = WEEK3AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, WEEK3AIREDINDEX) = Format(DateAdd("d", 20, slDate), "m/d/yy")
    grdManualPost.Row = 0
    grdManualPost.Col = WEEK4AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, WEEK4AIREDINDEX) = Format(DateAdd("d", 21, slDate), "m/d/yy")
    grdManualPost.Row = 1
    grdManualPost.Col = WEEK4AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, WEEK4AIREDINDEX) = Format(DateAdd("d", 27, slDate), "m/d/yy")
    grdManualPost.Row = 0
    grdManualPost.Col = WEEK5AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, WEEK5AIREDINDEX) = Format(DateAdd("d", 28, slDate), "m/d/yy")
    grdManualPost.Row = 1
    grdManualPost.Col = WEEK5AIREDINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, WEEK5AIREDINDEX) = Format(DateAdd("d", 34, slDate), "m/d/yy")
    grdManualPost.Row = 0
    grdManualPost.Col = TOTALINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, TOTALINDEX) = "Total"
    grdManualPost.Row = 1
    grdManualPost.Col = TOTALINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, TOTALINDEX) = "Aired"
    grdManualPost.Row = 0
    grdManualPost.Col = AFFIDAVITINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, AFFIDAVITINDEX) = "Affidavit"
    grdManualPost.Row = 1
    grdManualPost.Col = AFFIDAVITINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, AFFIDAVITINDEX) = "Time"
    grdManualPost.Row = 0
    grdManualPost.Col = POSTTIMESINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(0, POSTTIMESINDEX) = "Post Aff."
    grdManualPost.Row = 1
    grdManualPost.Col = POSTTIMESINDEX
    grdManualPost.CellFontBold = False
    grdManualPost.CellFontName = "Arial"
    grdManualPost.CellFontSize = 6.75
    grdManualPost.CellForeColor = vbBlue
    grdManualPost.TextMatrix(1, POSTTIMESINDEX) = "Times"
    grdManualPost.TextMatrix(0, SORTINDEX) = "Sort"
    grdManualPost.TextMatrix(0, CHFCODEINDEX) = "ChfCode"
    'Set height of grid

End Sub

Private Sub mGridSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slCntrNo As String
    Dim llStartTime As Long

    For llRow = grdManualPost.FixedRows To grdManualPost.Rows - 1 Step 1
        slStr = Trim$(grdManualPost.TextMatrix(llRow, CNTRNOINDEX))
        If slStr <> "" Then
            slCntrNo = grdManualPost.TextMatrix(llRow, CNTRNOINDEX)
            Do While Len(slCntrNo) < 8
                slCntrNo = "0" & slCntrNo
            Loop
            If ilCol = CNTRNOINDEX Then
                slSort = grdManualPost.TextMatrix(llRow, CNTRNOINDEX)
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = AGYESTNOINDEX Then
                slSort = grdManualPost.TextMatrix(llRow, AGYESTNOINDEX)
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
                slSort = slSort & slCntrNo
            ElseIf ilCol = DAYPARTTIMESINDEX Then
                slSort = grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX)
                ilPos = InStr(1, slSort, "-", vbBinaryCompare)
                If ilPos > 0 Then
                    llStartTime = gTimeToLong(Left(slSort, ilPos - 1), False)
                    slSort = llStartTime
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                End If
            Else
                slSort = UCase$(Trim$(grdManualPost.TextMatrix(llRow, ilCol)))
                slSort = slSort & slCntrNo
            End If
            slStr = ""
            'slStr = grdManualPost.TextMatrix(llRow, SORTINDEX)
            'ilPos = InStr(1, slStr, "|", vbTextCompare)
            'If ilPos > 1 Then
            '    slStr = Left$(slStr, ilPos - 1)
            'End If
            If (ilCol <> imManualPostColSorted) Or ((ilCol = imManualPostColSorted) And (imManualPostSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdManualPost.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdManualPost.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imManualPostColSorted Then
        imManualPostColSorted = SORTINDEX
    Else
        imManualPostColSorted = -1
        imManualPostSort = -1
    End If
    gGrid_SortByCol grdManualPost, CNTRNOINDEX, SORTINDEX, imManualPostColSorted, imManualPostSort
    imManualPostColSorted = ilCol
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


    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    igManUnload = YES
    Unload PostManualCntr
    'Set PostManualCntr = Nothing   'Remove data segment
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
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim llRow As Long

    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilCurrent As Integer
    Dim ilShow As Integer
    Dim ilState As Integer
    Dim ilAAS As Integer
    Dim ilAASCode As Integer
    Dim slDateTag As String
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    
    grdManualPost.Visible = False
    mClearGrid
    mPopulate 0, llRow, 0
'    ilRet = btrGetFirst(hmCHF, tmChf, imChfRecLen, INDEXKEY0, BTRV_LOCK_NONE, 0)
'
'    ' Prepare to execute an extended operation.
'    btrExtClear hmCHF   'Clear any previous extend operation
'
'    ilExtLen = Len(tmChf)
'    llNoRec = gExtNoRec(ilExtLen)
'    Call btrExtSetBounds(hmCHF, llNoRec, -1, "UC", "Chf", "") '"EG") 'Set extract limits (all records)
'
'    ilOffset = gFieldOffset("Chf", "chfVerifyFlag")
'    tlCharTypeBuff.sType = sgPostManualCntrType
'    ilRet = btrExtAddLogicConst(hmCHF, BTRV_KT_STRING, ilOffset, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
'
'    ilOffset = gFieldOffset("Chf", "chfDelete")
'    tlCharTypeBuff.sType = "N"
'    ilRet = btrExtAddLogicConst(hmCHF, BTRV_KT_STRING, ilOffset, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
'
'
'    ilRet = btrExtAddField(hmCHF, 0, ilExtLen) 'Extract the whole record
'    ilRet = btrExtGetNext(hmCHF, tmChf, ilExtLen, llRecPos)
'    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'        ilExtLen = Len(tmChf)  'Extract operation record size
'        Do While ilRet = BTRV_ERR_REJECT_COUNT
'            ilRet = btrExtGetNext(hmCHF, tmChf, ilExtLen, llRecPos)
'        Loop
'        Do While ilRet = BTRV_ERR_NONE
'            mPopulate 1, llRow
'            ilExtLen = Len(tmChf)  'Extract operation record size
'            ilRet = btrExtGetNext(hmCHF, tmChf, ilExtLen, llRecPos)
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hmCHF, tmChf, ilExtLen, llRecPos)
'            Loop
'        Loop
'    End If
    ReDim tmSdfDateTimeInfo(0 To 0) As SDFDATETIMEINFO
    lmInvStartDate = gDateValue(sgManualPostDate)
    lmInvEndDate = gDateValue(gObtainEndStd(sgManualPostDate))
    slCntrStatus = "OH" 'Hold and Orders
    slCntrType = "CTRQSM" 'Standard only
    ilCurrent = 4
    ilShow = 0
    ilState = 1
    ilAAS = 3
    ilAASCode = igManualPostVefCode
    slDateTag = sgManualPostDate
    ilRet = gPopCntrForAASBox(PostManualCntr, ilAAS, ilAASCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, lbcCntr, tmCntr(), slDateTag)
    For ilLoop = 0 To UBound(tmCntr) - 1 Step 1
        slNameCode = tmCntr(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            llChfCode = Val(slCode)
            tmIihfSrchKey2.lChfCode = llChfCode
            tmIihfSrchKey2.iVefCode = igManualPostVefCode
            gPackDateLong lmInvStartDate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
            ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If (ilRet <> BTRV_ERR_NONE) Then
                tmIihf.sSourceForm = ""
            End If
            If (Trim$(tmIihf.sSourceForm) = "T") Or (Trim$(tmIihf.sSourceForm) = "C") Or (Trim$(tmIihf.sSourceForm) = "") Then
                ilRet = gObtainCntrPlusGame(hmCHF, hmClf, hmCff, hmCgf, llChfCode, False, tmChf, tmClfList(), tmCfflist(), tmCgfList())
                If ilRet Then
                    For ilClf = LBound(tmClfList) To UBound(tmClfList) - 1 Step 1
                        tmClf = tmClfList(ilClf).ClfRec
                        If (tmClf.sType <> "O") And (tmClf.sType <> "A") And (tmClf.sType <> "E") Then
                            If tmClf.iVefCode = igManualPostVefCode Then
                                gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llStartDate
                                If llStartDate <= lmInvEndDate Then
                                    gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llEndDate
                                    If llEndDate >= lmInvStartDate Then
                                        mPopulate 1, llRow, ilClf
                                    End If
                                End If
                            End If
                        End If
                    Next ilClf
                End If
            End If
        End If
    Next ilLoop
    mPopulate 2, llRow, 0
    grdManualPost.Visible = True
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

    ilGap = cmcDone.Left - (cmcCancel.Left + cmcCancel.Width)
    cmcCancel.Top = Me.Height - cmcCancel.Height - 120
    cmcDone.Top = cmcCancel.Top
    cmcSave.Top = cmcCancel.Top
    grdManualPost.Move 180, 2 * plcScreen.Height + plcScreen.Top + 60, PostManualCntr.Width - 360, cmcDone.Top - (2 * plcScreen.Height + plcScreen.Top + 60)
    plcScreen.Left = 120
    imcKey.Move plcScreen.Left, plcScreen.Top + plcScreen.Height
    lbcKey.Move plcScreen.Left, imcKey.Top + plcScreen.Height + 60
    cmcCancel.Left = PostManualCntr.Width / 2 - cmcCancel.Width - cmcDone / 2 - ilGap
    cmcDone.Left = cmcCancel.Left + cmcCancel.Width + ilGap
    cmcSave.Left = cmcDone.Left + cmcDone.Width + ilGap
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdManualPost, CInt(fgFlexGridRowH)
    grdManualPost.Height = grdManualPost.Height + 30

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

    'grdManualPost.Row = llRow
    'For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
    '    grdManualPost.Col = llCol
    '    If grdManualPost.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
    '        grdManualPost.CellBackColor = LIGHTYELLOW
    '    Else
    '        grdManualPost.CellBackColor = GRAY    'vbBlue
    '    End If
    'Next llCol

End Sub

Private Sub grdManualPost_EnterCell()
    cmcCancel.Caption = "&Cancel"
    mSetShow
End Sub

Private Sub grdManualPost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long
    Dim ilSoldAdj As Integer

    llHeight = 0
    For llRow = 0 To grdManualPost.FixedRows - 1 Step 1
        llHeight = llHeight + grdManualPost.RowHeight(llRow)
    Next llRow
    grdManualPost.ToolTipText = ""
    If Y <= llHeight Then
        Exit Sub
    End If
    If Y > lmMaxHeight Then
        Exit Sub
    End If
    If grdManualPost.MouseRow >= grdManualPost.Rows Then
        Exit Sub
    End If
    ilSoldAdj = WEEK1SOLDINDEX - WEEK1AIREDINDEX
    If (grdManualPost.MouseCol >= WEEK1AIREDINDEX) And (grdManualPost.MouseCol <= WEEK5AIREDINDEX) Then
        'If Trim$(grdManualPost.TextMatrix(grdManualPost.MouseRow, grdManualPost.MouseCol)) <> "" Then
        '    grdManualPost.ToolTipText = Trim$(grdManualPost.TextMatrix(grdManualPost.MouseRow, grdManualPost.MouseCol)) & " of " & Trim$(grdManualPost.TextMatrix(grdManualPost.MouseRow, grdManualPost.MouseCol + ilSoldAdj))
        'Else
            grdManualPost.ToolTipText = "Sold " & Trim$(grdManualPost.TextMatrix(grdManualPost.MouseRow, grdManualPost.MouseCol + ilSoldAdj))
        'End If
    Else
        grdManualPost.ToolTipText = Trim$(grdManualPost.TextMatrix(grdManualPost.MouseRow, grdManualPost.MouseCol))
    End If
End Sub

Private Sub grdManualPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    imIgnoreScroll = False
    mSetShow
    If Y < grdManualPost.RowHeight(0) Then
        'If grdManualPost.MouseCol <= AGYESTNOINDEX Then
        '    grdManualPost.Col = grdManualPost.MouseCol
        '    mGridSortCol grdManualPost.Col
        'End If
        grdManualPost.Row = 0
        grdManualPost.Col = CHFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdManualPost, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdManualPost.FixedRows Then
        Exit Sub
    End If
    If grdManualPost.TextMatrix(llCurrentRow, ADVTNAMEINDEX) = "" Then
        Exit Sub
    End If
    If llCurrentRow >= grdManualPost.FixedRows Then
        If llCol = POSTTIMESINDEX Then
            grdManualPost.Row = llCurrentRow
            grdManualPost.Col = llCol
            If grdManualPost.CellBackColor = GRAY Then
                mGetSpotTimes llCurrentRow
            End If
            pbcClickFocus.SetFocus
        Else
            If grdManualPost.TextMatrix(llCurrentRow, ADVTNAMEINDEX) <> "" Then
                If (llCol >= STATIONCNTRNOINDEX) And (llCol <= imLastWeekNo) Then
                    grdManualPost.Row = llCurrentRow
                    grdManualPost.Col = llCol
                    If Not mColOk() Then
                        pbcClickFocus.SetFocus
                    Else
                        mEnableBox
                    End If
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub



Private Sub grdManualPost_Scroll()
    If grdManualPost.Visible = False Then
        Exit Sub
    End If
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdManualPost.Redraw = False Then
        grdManualPost.Redraw = True
        If lmTopRow < grdManualPost.FixedRows Then
            grdManualPost.TopRow = grdManualPost.FixedRows
        Else
            grdManualPost.TopRow = lmTopRow
        End If
        grdManualPost.Refresh
        grdManualPost.Redraw = False
    End If
    If (imCtrlVisible) And (grdManualPost.Row >= grdManualPost.FixedRows) And (grdManualPost.Col >= 0) And (grdManualPost.Col < grdManualPost.Cols - 1) Then
        If grdManualPost.RowIsVisible(grdManualPost.Row) Then
            pbcArrow.Move grdManualPost.Left - pbcArrow.Width - 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + (grdManualPost.RowHeight(grdManualPost.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcClickFocus.SetFocus
            edcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If

End Sub

Private Sub imcKey_Click()
    lbcKey.Visible = Not lbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = False
End Sub

Private Sub pbcClickFocus_Click()
    mSetShow
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdManualPost.Col = STATIONCNTRNOINDEX Then
                If grdManualPost.Row > grdManualPost.FixedRows Then
                    lmTopRow = -1
                    grdManualPost.Row = grdManualPost.Row - 1
                    If Not grdManualPost.RowIsVisible(grdManualPost.Row) Then
                        grdManualPost.TopRow = grdManualPost.TopRow - 1
                    End If
                    grdManualPost.Col = imLastWeekNo
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdManualPost.Col = grdManualPost.Col - 1
                ''If gColOk(grdManualPost, grdManualPost.Row, grdManualPost.Col) Then
                '    mEnableBox
                ''Else
                ''    ilPrev = True
                ''End If
                If Not mColOk() Then
                    ilPrev = True
                Else
                    mEnableBox
                End If
            End If
        Loop While ilPrev
    Else
        If grdManualPost.TextMatrix(grdManualPost.FixedRows, CNTRNOINDEX) = "" Then
            cmcCancel.SetFocus
            Exit Sub
        End If
        lmTopRow = -1
        grdManualPost.TopRow = grdManualPost.FixedRows
        grdManualPost.Col = STATIONCNTRNOINDEX
        grdManualPost.Row = grdManualPost.FixedRows
        'If gColOk(grdManualPost, grdManualPost.Row, grdManualPost.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdManualPost.Col = imLastWeekNo Then
                llRow = grdManualPost.Rows
                Do
                    llRow = llRow - 1
                Loop While grdManualPost.TextMatrix(llRow, CNTRNOINDEX) = ""
                llRow = llRow + 1
                If (grdManualPost.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdManualPost.Row = grdManualPost.Row + 1
                    If Not grdManualPost.RowIsVisible(grdManualPost.Row) Or (grdManualPost.Row - (grdManualPost.TopRow - grdManualPost.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdManualPost.TopRow = grdManualPost.TopRow + 1
                    End If
                    grdManualPost.Col = STATIONCNTRNOINDEX
                    'grdManualPost.TextMatrix(grdManualPost.Row, CODEINDEX) = 0
                    If Trim$(grdManualPost.TextMatrix(grdManualPost.Row, CNTRNOINDEX)) <> "" Then
                        ''If gColOk(grdManualPost, grdManualPost.Row, grdManualPost.Col) Then
                        '    mEnableBox
                        ''Else
                        ''    cmcCancel.SetFocus
                        ''End If
                        If Not mColOk() Then
                            grdManualPost.Col = WEEK1AIREDINDEX
                        End If
                        mEnableBox
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdManualPost.Left - pbcArrow.Width - 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + (grdManualPost.RowHeight(grdManualPost.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdManualPost.Col = grdManualPost.Col + 1
                'If gColOk(grdManualPost, grdManualPost.Row, grdManualPost.Col) Then
                    mEnableBox
                'Else
                '    ilNext = True
                'End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdManualPost.TopRow = grdManualPost.FixedRows
        grdManualPost.Col = STATIONCNTRNOINDEX
        grdManualPost.Row = grdManualPost.FixedRows
        'If gColOk(grdManualPost, grdManualPost.Row, grdManualPost.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If

End Sub

Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (grdManualPost.Row < grdManualPost.FixedRows) Or (grdManualPost.Row >= grdManualPost.Rows) Or (grdManualPost.Col < grdManualPost.FixedCols) Or (grdManualPost.Col >= grdManualPost.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdManualPost.Row
    lmEnableCol = grdManualPost.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdManualPost.Left - pbcArrow.Width - 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + (grdManualPost.RowHeight(grdManualPost.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    
    Select Case grdManualPost.Col

        Case STATIONCNTRNOINDEX
            edcDropDown.MaxLength = 20
            edcDropDown.Text = grdManualPost.Text
        Case STAIONINVNOINDEX
            edcDropDown.MaxLength = 20
            edcDropDown.Text = grdManualPost.Text
        Case STAIONESTNOINDEX
            edcDropDown.MaxLength = 20
            edcDropDown.Text = grdManualPost.Text
        Case WEEK1AIREDINDEX
            grdManualPost.CellForeColor = vbBlack
            edcDropDown.MaxLength = 3
            edcDropDown.Text = grdManualPost.Text
        Case WEEK2AIREDINDEX
            grdManualPost.CellForeColor = vbBlack
            edcDropDown.MaxLength = 3
            edcDropDown.Text = grdManualPost.Text
        Case WEEK3AIREDINDEX
            grdManualPost.CellForeColor = vbBlack
            edcDropDown.MaxLength = 3
            edcDropDown.Text = grdManualPost.Text
        Case WEEK4AIREDINDEX
            grdManualPost.CellForeColor = vbBlack
            edcDropDown.MaxLength = 3
            edcDropDown.Text = grdManualPost.Text
        Case WEEK5AIREDINDEX
            grdManualPost.CellForeColor = vbBlack
            edcDropDown.MaxLength = 3
            edcDropDown.Text = grdManualPost.Text
        Case TOTALINDEX
            edcDropDown.MaxLength = 4
            edcDropDown.Text = grdManualPost.Text
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
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim llRow As Long
    Dim blFound As Boolean

    pbcArrow.Visible = False
    If (lmEnableRow >= grdManualPost.FixedRows) And (lmEnableRow < grdManualPost.Rows) Then
        If (grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropDown.Text) Or ((lmEnableCol >= WEEK1AIREDINDEX) And (lmEnableCol <= WEEK5AIREDINDEX)) Then
            imChg = True
            llEnableRow = grdManualPost.Row
            llEnableCol = grdManualPost.Col
            If (lmEnableCol >= WEEK1AIREDINDEX) And (lmEnableCol <= WEEK5AIREDINDEX) Then
                grdManualPost.Row = lmEnableRow
                grdManualPost.Col = lmEnableCol
                grdManualPost.CellForeColor = vbBlack
                grdManualPost.Col = TOTALINDEX
                grdManualPost.CellForeColor = vbBlack
            End If
            grdManualPost.Row = llEnableRow
            grdManualPost.Col = llEnableCol
            grdManualPost.TextMatrix(lmEnableRow, CHGINDEX) = "Y"
            blFound = False
            For llRow = grdManualPost.FixedRows To grdManualPost.Rows - 1 Step 1
                If Val(grdManualPost.TextMatrix(llRow, CNTRNOINDEX)) = Val(grdManualPost.TextMatrix(lmEnableRow, CNTRNOINDEX)) Then
                   blFound = True
                   grdManualPost.TextMatrix(llRow, CHGINDEX) = "Y"
                    Select Case lmEnableCol
                        Case STATIONCNTRNOINDEX
                            grdManualPost.TextMatrix(llRow, lmEnableCol) = edcDropDown.Text
                        Case STAIONINVNOINDEX
                            grdManualPost.TextMatrix(llRow, lmEnableCol) = edcDropDown.Text
                        Case STAIONESTNOINDEX
                            grdManualPost.TextMatrix(llRow, lmEnableCol) = edcDropDown.Text
                    End Select
                ElseIf blFound Then
                    Exit For
                End If
            Next llRow
        End If
        Select Case lmEnableCol
            Case STATIONCNTRNOINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case STAIONINVNOINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case STAIONESTNOINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case WEEK1AIREDINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX) = Val(grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX)) + Val(edcDropDown.Text) - Val(grdManualPost.TextMatrix(lmEnableRow, lmEnableCol))
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case WEEK2AIREDINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX) = Val(grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX)) + Val(edcDropDown.Text) - Val(grdManualPost.TextMatrix(lmEnableRow, lmEnableCol))
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case WEEK3AIREDINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX) = Val(grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX)) + Val(edcDropDown.Text) - Val(grdManualPost.TextMatrix(lmEnableRow, lmEnableCol))
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case WEEK4AIREDINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX) = Val(grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX)) + Val(edcDropDown.Text) - Val(grdManualPost.TextMatrix(lmEnableRow, lmEnableCol))
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case WEEK5AIREDINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX) = Val(grdManualPost.TextMatrix(lmEnableRow, TOTALINDEX)) + Val(edcDropDown.Text) - Val(grdManualPost.TextMatrix(lmEnableRow, lmEnableCol))
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
            Case TOTALINDEX
                edcDropDown.Visible = False
                grdManualPost.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
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

    If (grdManualPost.Row < grdManualPost.FixedRows) Or (grdManualPost.Row >= grdManualPost.Rows) Or (grdManualPost.Col < grdManualPost.FixedCols) Or (grdManualPost.Col >= grdManualPost.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdManualPost.Col - 1 Step 1
        llColPos = llColPos + grdManualPost.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdManualPost.ColWidth(grdManualPost.Col)
    ilCol = grdManualPost.Col
    Do While ilCol < grdManualPost.Cols - 1
        If (Trim$(grdManualPost.TextMatrix(grdManualPost.Row - 1, grdManualPost.Col)) <> "") And (Trim$(grdManualPost.TextMatrix(grdManualPost.Row - 1, grdManualPost.Col)) = Trim$(grdManualPost.TextMatrix(grdManualPost.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdManualPost.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdManualPost.Col
        Case STATIONCNTRNOINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STAIONINVNOINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STAIONESTNOINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case WEEK1AIREDINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case WEEK2AIREDINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case WEEK3AIREDINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case WEEK4AIREDINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case WEEK5AIREDINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TOTALINDEX
            edcDropDown.Move grdManualPost.Left + llColPos + 30, grdManualPost.Top + grdManualPost.RowPos(grdManualPost.Row) + 15, grdManualPost.ColWidth(grdManualPost.Col) - 30, grdManualPost.RowHeight(grdManualPost.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
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
    Dim llNext As Long
    Dim ilWeek As Integer
    Dim ilAired As Integer
    Dim ilSold As Integer
    Dim ilMissedCount As Integer
    Dim ilSchdCount As Integer
    Dim ilMakeMissedCount As Integer
    Dim ilMakeSchdCount As Integer
    Dim slSchDate As String
    Dim slSchTime As String
    Dim slSourceForm As String
    Dim slStnEstNo As String
    Dim slStnInvNo As String
    Dim slStnCntrNo As String
    Dim ilPass As Integer
    Dim llChfCode As Long
    
    If Not imUpdateAllowed Then
        mSaveRec = False
        Exit Function
    End If
    
    llChfCode = -1
    For llRow = grdManualPost.FixedRows To grdManualPost.Rows - 1 Step 1
        If (grdManualPost.TextMatrix(llRow, CNTRNOINDEX) <> "") And (grdManualPost.TextMatrix(llRow, CHGINDEX) = "Y") Then
            slSourceForm = grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX)
            slStnEstNo = grdManualPost.TextMatrix(llRow, STAIONESTNOINDEX)
            slStnInvNo = grdManualPost.TextMatrix(llRow, STAIONINVNOINDEX)
            slStnCntrNo = grdManualPost.TextMatrix(llRow, STATIONCNTRNOINDEX)
            ilRet = mAddOrChgIIHF(igManualPostVefCode, Val(grdManualPost.TextMatrix(llRow, CHFCODEINDEX)), slSourceForm, slStnEstNo, slStnInvNo, slStnCntrNo)
            For ilWeek = 1 To 5 Step 1
                If ilWeek = 5 Then
                    If grdManualPost.ColWidth(WEEK5AIREDINDEX) = 0 Then
                        Exit For
                    End If
                End If
                Select Case ilWeek
                    Case 1
                        ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK1AIREDINDEX))
                        ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK1SOLDINDEX))
                    Case 2
                        ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK2AIREDINDEX))
                        ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK2SOLDINDEX))
                    Case 3
                        ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK3AIREDINDEX))
                        ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK3SOLDINDEX))
                    Case 4
                        ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK4AIREDINDEX))
                        ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK4SOLDINDEX))
                    Case 5
                        ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK5AIREDINDEX))
                        ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK5SOLDINDEX))
                End Select
                ilSchdCount = 0
                ilMissedCount = 0
                llNext = Val(grdManualPost.TextMatrix(llRow, TIMESINDEX))
                Do While llNext <> -1
                    If tmSdfDateTimeInfo(llNext).iWeek = ilWeek Then
                        If tmSdfDateTimeInfo(llNext).sSpotSchStatus = "M" Then
                            ilMissedCount = ilMissedCount + 1
                        Else
                            ilSchdCount = ilSchdCount + 1
                        End If
                    End If
                    llNext = tmSdfDateTimeInfo(llNext).lNext
                Loop
                If ilSchdCount = 0 Then 'All missed, None scheduled
                    ilMakeSchdCount = ilAired
                    ilMakeMissedCount = 0
                ElseIf ilMissedCount = 0 Then   'all schedule, none missed
                    ilMakeMissedCount = ilSold - ilAired
                    ilMakeSchdCount = 0
                Else
                    If ilAired = ilSold Then
                        ilMakeSchdCount = ilMissedCount
                        ilMakeMissedCount = 0
                    Else
                        ilMakeMissedCount = ilSold - ilAired
                        ilMakeSchdCount = 0
                        If ilMissedCount > ilMakeMissedCount Then
                            ilMakeSchdCount = ilMissedCount - ilMakeMissedCount
                            ilMakeMissedCount = 0
                        ElseIf ilMissedCount < ilMakeMissedCount Then
                            ilMakeMissedCount = ilMakeMissedCount - ilMissedCount
                        Else
                            ilMakeMissedCount = 0
                        End If
                    End If
                End If
                For ilPass = 1 To 2 Step 1
                    If (ilMakeMissedCount > 0) Or (ilMakeSchdCount > 0) Then
                        llNext = Val(grdManualPost.TextMatrix(llRow, TIMESINDEX))
                        Do While llNext <> -1
                            If tmSdfDateTimeInfo(llNext).iWeek = ilWeek Then
                                If ilMakeSchdCount > 0 Then
                                    If ((ilPass = 1) And (tmSdfDateTimeInfo(llNext).sTimeSchStatus <> "M") And (tmSdfDateTimeInfo(llNext).sSpotSchStatus = "M")) Or ((ilPass = 2) And (tmSdfDateTimeInfo(llNext).sSpotSchStatus = "M")) Then
                                        'Book spot
                                        tmSdfSrchKey3.lCode = tmSdfDateTimeInfo(llNext).lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                        If ilRet = BTRV_ERR_NONE Then
                                            'gPackDate tmSdfDateTimeInfo(llNext).sAirDate, tmSdf.iDate(0), tmSdf.iDate(1)
                                            'gPackTime tmSdfDateTimeInfo(llNext).sAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                                            ilRet = mBookSpot(tmSdfDateTimeInfo(llNext), tmSdf, tmSdfDateTimeInfo(llNext).lCifCode)
                                            If ilRet Then
                                                ilMakeSchdCount = ilMakeSchdCount - 1
                                                If ilMakeSchdCount <= 0 Then
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If ((ilPass = 1) And (tmSdfDateTimeInfo(llNext).sSpotSchStatus <> "M") And (tmSdfDateTimeInfo(llNext).sTimeSchStatus = "M")) Or ((ilPass = 2) And (tmSdfDateTimeInfo(llNext).sSpotSchStatus <> "M")) Then
                                        'Miss spot
                                        tmSdfSrchKey3.lCode = tmSdfDateTimeInfo(llNext).lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                        If ilRet = BTRV_ERR_NONE Then
                                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                                            imSelectedDay = gWeekDayStr(slSchDate)
                                            ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                                            If ilRet Then
                                                ilMakeMissedCount = ilMakeMissedCount - 1
                                                If ilMakeMissedCount <= 0 Then
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            llNext = tmSdfDateTimeInfo(llNext).lNext
                        Loop
                    End If
                Next ilPass
                If (slSourceForm = "T") And (ilSchdCount <> 0) Then
                    'Check if date/time changed
                    llNext = Val(grdManualPost.TextMatrix(llRow, TIMESINDEX))
                    Do While llNext <> -1
                        If tmSdfDateTimeInfo(llNext).iWeek = ilWeek Then
                            If tmSdfDateTimeInfo(llNext).sSpotSchStatus <> "M" Then
                                tmSdfSrchKey3.lCode = tmSdfDateTimeInfo(llNext).lSdfCode
                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                If ilRet = BTRV_ERR_NONE Then
                                    If tmSdf.sSchStatus <> "M" Then
                                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slSchTime
                                        If gDateValue(tmSdfDateTimeInfo(llNext).sAirDate) <> gDateValue(slSchDate) Then
                                            imSelectedDay = gWeekDayStr(slSchDate)
                                            ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                                            If ilRet Then
                                                tmSdfSrchKey3.lCode = tmSdfDateTimeInfo(llNext).lSdfCode
                                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'gPackDate tmSdfDateTimeInfo(llNext).sAirDate, tmSdf.iDate(0), tmSdf.iDate(1)
                                                    'gPackTime tmSdfDateTimeInfo(llNext).sAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                                                    ilRet = mBookSpot(tmSdfDateTimeInfo(llNext), tmSdf, tmSdfDateTimeInfo(llNext).lCifCode)
                                                End If
                                            End If
                                        ElseIf gTimeToLong(tmSdfDateTimeInfo(llNext).sAirTime, False) <> gTimeToLong(slSchTime, False) Then
                                            tmSdf.sAffChg = "Y"
                                            If tmSdfDateTimeInfo(llNext).lCifCode > 0 Then
                                                If tmSdf.lCopyCode <= 0 Then
                                                    tmSdf.sPtType = "1"
                                                    tmSdf.lCopyCode = tmSdfDateTimeInfo(llNext).lCifCode
                                                End If
                                            ElseIf tmSdfDateTimeInfo(llNext).lCifCode <= 0 Then
                                                If tmSdf.lCopyCode > 0 Then
                                                    tmSdf.sPtType = ""
                                                    tmSdf.lCopyCode = 0
                                                End If
                                            End If
                                            gPackTime slSchTime, tmSdf.iTime(0), tmSdf.iTime(1)
                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                        Else
                                            If tmSdfDateTimeInfo(llNext).lCifCode > 0 Then
                                                If tmSdf.lCopyCode <= 0 Then
                                                    tmSdf.sPtType = "1"
                                                    tmSdf.lCopyCode = tmSdfDateTimeInfo(llNext).lCifCode
                                                    tmSdf.sAffChg = "Y"
                                                End If
                                            ElseIf tmSdfDateTimeInfo(llNext).lCifCode <= 0 Then
                                                If tmSdf.lCopyCode > 0 Then
                                                    tmSdf.sPtType = ""
                                                    tmSdf.lCopyCode = 0
                                                    tmSdf.sAffChg = "Y"
                                                End If
                                            End If
                                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        llNext = tmSdfDateTimeInfo(llNext).lNext
                    Loop
                End If
            Next ilWeek
            'If llChfCode = -1 Then
            '    llChfCode = Val(grdManualPost.TextMatrix(llRow, CHFCODEINDEX))
            'Else
            '    If (llRow = grdManualPost.Rows - 1) Then
            '        ilRet = mUpdateApf(llChfCode, igManualPostVefCode, slStnInvNo, slStnCntrNo)
            '        llChfCode = -1
            '    ElseIf (llChfCode <> Val(grdManualPost.TextMatrix(llRow + 1, CNTRNOINDEX))) Then
            '        ilRet = mUpdateApf(llChfCode, igManualPostVefCode, slStnInvNo, slStnCntrNo)
            '        llChfCode = -1
            '    End If
            'End If
            If llChfCode = -1 Then
                llChfCode = Val(grdManualPost.TextMatrix(llRow, CHFCODEINDEX))
            End If
            If (llRow = grdManualPost.Rows - 1) Then
                ilRet = mUpdateApf(llChfCode, igManualPostVefCode, slStnInvNo, slStnCntrNo)
                llChfCode = -1
            ElseIf (llChfCode <> Val(grdManualPost.TextMatrix(llRow + 1, CHFCODEINDEX))) Then
                ilRet = mUpdateApf(llChfCode, igManualPostVefCode, slStnInvNo, slStnCntrNo)
                llChfCode = -1
            End If
        
        End If
    Next llRow
    If llChfCode <> -1 Then
        ilRet = mUpdateApf(llChfCode, igManualPostVefCode, slStnInvNo, slStnCntrNo)
    End If
    imChg = False
    mSaveRec = True
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    gSetMousePointer grdManualPost, grdManualPost, vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    If Not imUpdateAllowed Then
        cmcSave.Enabled = False
        cmcDone.Enabled = False
    Else
        cmcDone.Enabled = True
        If imChg Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    End If
End Sub

Private Sub tmcStart_Timer()
    
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
End Sub

Private Sub mGetSpotTimes(llCntrRow As Long)
    Dim ilWeek As Integer
    Dim ilAired As Integer
    Dim ilSold As Integer
    Dim ilCount As Integer
    Dim ilTotal As Integer
    Dim ilTimes As Integer
    Dim llNext As Long
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llIndex As Long
    Dim llRow As Long
    Dim llCntr As Long
    Dim blRowFound As Boolean
    Dim ilRet As Integer
    
    pbcArrow.Visible = False
    pbcArrow.Move grdManualPost.Left - pbcArrow.Width - 30, grdManualPost.Top + grdManualPost.RowPos(llRow) + (grdManualPost.RowHeight(llRow) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    ReDim tgPostManualCntrInfo(0 To 0) As POSTMANUALCNTRINFO
    ReDim tgPostManualTimesInfo(0 To 0) As POSTMANUALTIMESINFO
    llRow = llCntrRow
    Do
        tgPostManualCntrInfo(UBound(tgPostManualCntrInfo)).lRow = llRow
        tgPostManualCntrInfo(UBound(tgPostManualCntrInfo)).sDaypartTimes = grdManualPost.TextMatrix(llRow, DAYPARTTIMESINDEX)
        tgPostManualCntrInfo(UBound(tgPostManualCntrInfo)).sDaypartDays = grdManualPost.TextMatrix(llRow, DAYPARTDAYSINDEX)
        tgPostManualCntrInfo(UBound(tgPostManualCntrInfo)).sLength = grdManualPost.TextMatrix(llRow, LENGTHINDEX)
        For ilWeek = 1 To 5 Step 1
            If ilWeek = 5 Then
                If grdManualPost.ColWidth(WEEK5AIREDINDEX) = 0 Then
                    Exit For
                End If
            End If
            Select Case ilWeek
                Case 1
                    ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK1AIREDINDEX))
                    ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK1SOLDINDEX))
                Case 2
                    ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK2AIREDINDEX))
                    ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK2SOLDINDEX))
                Case 3
                    ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK3AIREDINDEX))
                    ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK3SOLDINDEX))
                Case 4
                    ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK4AIREDINDEX))
                    ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK4SOLDINDEX))
                Case 5
                    ilAired = Val(grdManualPost.TextMatrix(llRow, WEEK5AIREDINDEX))
                    ilSold = Val(grdManualPost.TextMatrix(llRow, WEEK5SOLDINDEX))
            End Select
            ilCount = 0
            llNext = Val(grdManualPost.TextMatrix(llRow, TIMESINDEX))
            'If ilAired < ilSold Then
            '    ilTimes = ilAired
            'Else
                ilTimes = ilSold
            'End If
            Do While llNext <> -1
                If tmSdfDateTimeInfo(llNext).iWeek = ilWeek Then
                    ilUpper = UBound(tgPostManualTimesInfo)
                    tgPostManualTimesInfo(ilUpper).lRow = llRow
                    tgPostManualTimesInfo(ilUpper).sSchStatus = tmSdfDateTimeInfo(llNext).sTimeSchStatus
                    tgPostManualTimesInfo(ilUpper).sDate = tmSdfDateTimeInfo(llNext).sAirDate
                    tgPostManualTimesInfo(ilUpper).sTime = tmSdfDateTimeInfo(llNext).sAirTime
                    tgPostManualTimesInfo(ilUpper).sISCI = tmSdfDateTimeInfo(llNext).sISCI
                    tgPostManualTimesInfo(ilUpper).lCifCode = tmSdfDateTimeInfo(llNext).lCifCode
                    tgPostManualTimesInfo(ilUpper).iWeek = ilWeek
                    tgPostManualTimesInfo(ilUpper).sWkDate = DateAdd("d", 7 * (ilWeek - 1), sgManualPostDate)
                    tgPostManualTimesInfo(ilUpper).lIndex = llNext
                    ReDim Preserve tgPostManualTimesInfo(0 To ilUpper + 1) As POSTMANUALTIMESINFO
                    ilCount = ilCount + 1
                    If ilCount = ilTimes Then
                        Exit Do
                    End If
                End If
                llNext = tmSdfDateTimeInfo(llNext).lNext
            Loop
        Next ilWeek
        ReDim Preserve tgPostManualCntrInfo(0 To UBound(tgPostManualCntrInfo) + 1) As POSTMANUALCNTRINFO
        'Find Next Row
        blRowFound = False
        For llCntr = llRow + 1 To grdManualPost.Rows - 1 Step 1
            If Val(grdManualPost.TextMatrix(llCntr, CNTRNOINDEX)) = Val(grdManualPost.TextMatrix(llCntrRow, CNTRNOINDEX)) Then
                llRow = llCntr
                blRowFound = True
                Exit For
            End If
        Next llCntr
    Loop While blRowFound
    llRow = llCntrRow
    igPostManualAdfCode = 0
    tmChfSrchKey0.lCode = Val(grdManualPost.TextMatrix(llRow, CHFCODEINDEX))
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        igPostManualAdfCode = tmChf.iAdfCode
    End If
    sgCntrLineInfo = grdManualPost.TextMatrix(llRow, ADVTNAMEINDEX) & " # " & grdManualPost.TextMatrix(llRow, CNTRNOINDEX) & " Est# " & grdManualPost.TextMatrix(llRow, AGYESTNOINDEX)
    PostManualTimes.Show vbModal
    If igPostManualTimes = 1 Then
        For llCntr = 0 To UBound(tgPostManualCntrInfo) - 1 Step 1
            llRow = tgPostManualCntrInfo(llCntr).lRow
            imChg = True
            grdManualPost.TextMatrix(llRow, CHGINDEX) = "Y"
            'grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX) = "C"
            grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX) = "T"
            For ilLoop = 0 To UBound(tgPostManualTimesInfo) - 1 Step 1
                If llRow = tgPostManualTimesInfo(ilLoop).lRow Then
                    llIndex = tgPostManualTimesInfo(ilLoop).lIndex
                    If (Trim$(tgPostManualTimesInfo(ilLoop).sDate) = "") Or (Trim$(tgPostManualTimesInfo(ilLoop).sTime) = "") Then
                        tmSdfDateTimeInfo(llIndex).sTimeSchStatus = "M"
                    Else
                        tmSdfDateTimeInfo(llIndex).sTimeSchStatus = "S"
                        tmSdfDateTimeInfo(llIndex).sAirDate = tgPostManualTimesInfo(ilLoop).sDate
                        tmSdfDateTimeInfo(llIndex).sAirTime = tgPostManualTimesInfo(ilLoop).sTime
                        tmSdfDateTimeInfo(llIndex).sISCI = tgPostManualTimesInfo(ilLoop).sISCI
                        tmSdfDateTimeInfo(llIndex).lCifCode = tgPostManualTimesInfo(ilLoop).lCifCode
                        'grdManualPost.TextMatrix(llRow, SOURCEFORMINDEX) = "T"
                    End If
                End If
            Next ilLoop
            grdManualPost.Row = llRow
            ilTotal = 0
            For ilWeek = 1 To 5 Step 1
                ilCount = 0
                llNext = Val(grdManualPost.TextMatrix(llRow, TIMESINDEX))
                Do While llNext <> -1
                    If (tmSdfDateTimeInfo(llNext).sTimeSchStatus <> "M") And (tmSdfDateTimeInfo(llNext).iWeek = ilWeek) Then
                        ilCount = ilCount + 1
                        ilTotal = ilTotal + 1
                    End If
                    llNext = tmSdfDateTimeInfo(llNext).lNext
                Loop
                Select Case ilWeek
                    Case 1
                        grdManualPost.Col = WEEK1AIREDINDEX
                        grdManualPost.CellForeColor = vbBlack
                        If (Val(grdManualPost.TextMatrix(llRow, WEEK1SOLDINDEX)) <> 0) Then 'Or (ilCount > 0) Then
                            grdManualPost.TextMatrix(llRow, WEEK1AIREDINDEX) = ilCount
                        End If
                    Case 2
                        grdManualPost.Col = WEEK2AIREDINDEX
                        grdManualPost.CellForeColor = vbBlack
                        If (Val(grdManualPost.TextMatrix(llRow, WEEK2SOLDINDEX)) <> 0) Then 'Or (ilCount > 0) Then
                            grdManualPost.TextMatrix(llRow, WEEK2AIREDINDEX) = ilCount
                        End If
                    Case 3
                        grdManualPost.Col = WEEK3AIREDINDEX
                        grdManualPost.CellForeColor = vbBlack
                        If (Val(grdManualPost.TextMatrix(llRow, WEEK3SOLDINDEX)) <> 0) Then 'Or (ilCount > 0) Then
                            grdManualPost.TextMatrix(llRow, WEEK3AIREDINDEX) = ilCount
                        End If
                    Case 4
                        grdManualPost.Col = WEEK4AIREDINDEX
                        grdManualPost.CellForeColor = vbBlack
                        If (Val(grdManualPost.TextMatrix(llRow, WEEK4SOLDINDEX)) <> 0) Then 'Or (ilCount > 0) Then
                            grdManualPost.TextMatrix(llRow, WEEK4AIREDINDEX) = ilCount
                        End If
                    Case 5
                        grdManualPost.Col = WEEK5AIREDINDEX
                        grdManualPost.CellForeColor = vbBlack
                        If (Val(grdManualPost.TextMatrix(llRow, WEEK5SOLDINDEX)) <> 0) Then 'Or (ilCount > 0) Then
                            grdManualPost.TextMatrix(llRow, WEEK5AIREDINDEX) = ilCount
                        End If
                End Select
            Next ilWeek
            grdManualPost.TextMatrix(llRow, TOTALINDEX) = ilTotal
            grdManualPost.Col = TOTALINDEX
            grdManualPost.CellForeColor = vbBlack
            'If ilTotal > 0 Then
                grdManualPost.TextMatrix(llRow, AFFIDAVITINDEX) = "Date/Time"
            'Else
            '    grdManualPost.TextMatrix(llRow, AFFIDAVITINDEX) = "Count"
            'End If
        Next llCntr
        mSetCommands
    End If
    pbcArrow.Visible = False
End Sub

Private Function mAddOrChgIIHF(ilVefCode As Integer, llChfCode As Long, slSourceForm As String, slStationEstNo As String, slStationInvNo As String, slStationCntrNo As String) As Integer
    Dim ilRet As Integer
    Dim ilVff As Integer
    Dim llInvStartdate As Long
    
    llInvStartdate = gDateValue(sgManualPostDate)
    tmIihfSrchKey2.lChfCode = llChfCode
    tmIihfSrchKey2.iVefCode = ilVefCode
    gPackDateLong llInvStartdate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        If ((Trim$(tmIihf.sSourceForm) <> slSourceForm)) Or ((Trim$(tmIihf.sStnEstimateNo) <> slStationEstNo)) Or ((Trim$(tmIihf.sStnInvoiceNo) <> slStationInvNo)) Or ((Trim$(tmIihf.sStnContractNo) <> slStationCntrNo)) Then
            tmIihf.sSourceForm = slSourceForm
            tmIihf.sStnEstimateNo = slStationEstNo
            tmIihf.sStnInvoiceNo = slStationInvNo
            tmIihf.sStnContractNo = slStationCntrNo
            ilRet = btrUpdate(hmIihf, tmIihf, imIihfRecLen)
        End If
        mAddOrChgIIHF = True
        Exit Function
    End If
    tmIihf.lCode = 0
    tmIihf.iVefCode = ilVefCode
    tmIihf.lChfCode = llChfCode
    gPackDateLong llInvStartdate, tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1)
    tmIihf.sFileName = "Manual Post"
    tmIihf.sStnEstimateNo = slStationEstNo
    tmIihf.sStnInvoiceNo = slStationInvNo
    tmIihf.sStnContractNo = slStationCntrNo
    tmIihf.lAmfCode = 0
    tmIihf.sSourceForm = slSourceForm
    tmIihf.sUnused = ""
    ilRet = btrInsert(hmIihf, tmIihf, imIihfRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        mAddOrChgIIHF = True
    Else
        mAddOrChgIIHF = False
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindAvail                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Private Function mFindAvail(ilVefCode As Integer, slSchDate As String, slFindTime As String, ilGameNo As Integer, ilFindAdjAvail As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mFindAvail(slSchDate, slSchTime, ilAvailIndex)
'   Where:
'       slSchDate(I)- Scheduled Date
'       slSchTime(I)- Time that avail is to be found at
'       ilFindAdjAvail(I)- Find closest avail to specified time
'       llSsfRecPos(O)- Ssf record position
'       ilAvailIndex(O)- Index into Ssf where avail is located
'       ilRet(O)- True=Avail found; False=Avail not found
'       lmSsfRecPos(O)- Ssf record position
'
    Dim ilRet As Integer
    Dim llSchDate As Long
    Dim llTime As Long
    Dim llTstTime As Long
    Dim llFndAdjTime As Long
    Dim ilLoop As Integer
    llTime = CLng(gTimeToCurrency(slFindTime, False))
    llSchDate = gDateValue(slSchDate)
    imSelectedDay = gWeekDayStr(slSchDate)
    lmSsfDate(imSelectedDay) = 0
    ilRet = gObtainSsfForDateOrGame(ilVefCode, llSchDate, slFindTime, ilGameNo, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
    llFndAdjTime = -1
    For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
       LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTstTime
            If llTime = llTstTime Then 'Replace
                ilAvailIndex = ilLoop
                mFindAvail = True
                Exit Function
            ElseIf (llTstTime < llTime) And (ilFindAdjAvail) Then
                ilAvailIndex = ilLoop
                llFndAdjTime = llTstTime
            ElseIf (llTime < llTstTime) And (ilFindAdjAvail) Then
                If llFndAdjTime = -1 Then
                    ilAvailIndex = ilLoop
                    mFindAvail = True
                    Exit Function
                Else
                    If (llTime - llFndAdjTime) < (llTstTime - llTime) Then
                        mFindAvail = True
                        Exit Function
                    Else
                        ilAvailIndex = ilLoop
                        mFindAvail = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ilLoop
    If (llFndAdjTime <> -1) And (ilFindAdjAvail) Then
        mFindAvail = True
        Exit Function
    End If
    mFindAvail = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailRoom                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if room exist for    *
'*                      spot within avail              *
'*                                                     *
'*******************************************************
Private Function mAvailRoom(ilVefCode As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mAvailRoom(ilAvailIndex)
'   where:
'       ilAvailIndex(I)- location of avail within Ssf (use mFindAvail)
'       ilRet(O)- True=Avail has room; False=insufficient room within avail
'
'       tmSdf(I)- spot records
'
'       Code later: ask if avail should be overbooked
'                   If so, create a version zero (0) of the library with the new
'                   units/seconds
'
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer
    Dim ilUnitsSold As Integer
    Dim ilSecSold As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotIndex As Integer
    Dim ilNewUnit As Integer
    Dim ilNewSec As Integer
    Dim ilRet As Integer
    Dim ilVpfIndex As Integer
    
    ilVpfIndex = gBinarySearchVpfPlus(ilVefCode)    'gVpfFind(PostLog, imVefCode)
    If ilVpfIndex = -1 Then
        mAvailRoom = False
        Exit Function
    End If
   LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    '10/27/11: Disallow more then 31 spots in any avail
    If tmAvail.iNoSpotsThis >= 31 Then
        'ilRet = MsgBox("Move not allowed because Avail contains the maximum number of spots (31).", vbOkOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
       LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilSpotIndex)
        If tmSpot.lSdfCode = tmSdf.lCode Then
            mAvailRoom = True
            Exit Function
        End If
        If (tmSpot.iRecType And &HF) >= 10 Then
            ilSpotLen = tmSpot.iPosLen And &HFFF
            If (tgVpf(ilVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
                If ilSpotUnits <= 0 Then
                    ilSpotUnits = 1
                End If
                ilSpotLen = 0
            Else
                ilSpotUnits = 1
                'If (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                '    ilSpotLen = 0
                'End If
            End If
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilUnitsSold = ilUnitsSold + ilSpotUnits
                ilSecSold = ilSecSold + ilSpotLen
            End If
        End If
    Next ilSpotIndex
    ilSpotLen = tmSdf.iLen
    If (tgVpf(ilVpfIndex).sSSellOut = "T") Then
        ilSpotUnits = ilSpotLen \ 30
        If ilSpotUnits <= 0 Then
            ilSpotUnits = 1
        End If
        ilSpotLen = 0
    Else
        ilSpotUnits = 1
        'If (tgVpf(ilVpfIndex).sSSellOut = "U") Then
        '    ilSpotLen = 0
        'End If
    End If
    ilNewUnit = 0
    ilNewSec = 0
    If (tgVpf(ilVpfIndex).sSSellOut = "M") Then
        If (ilSpotLen + ilSecSold <> ilAvailSec) Or (ilSpotUnits + ilUnitsSold <> ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    Else
        If (ilSpotLen + ilSecSold > ilAvailSec) Or (ilSpotUnits + ilUnitsSold > ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    End If
    If (tgVpf(ilVpfIndex).sSOverBook <> "Y") Then
        'ilRet = MsgBox("Move not allowed because Avail would be Overbooked.", vbOkOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    Do
        imSsfRecLen = Len(tmSsf(imSelectedDay))
        ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        '5/20/11
        If (tmAvail.iOrigUnit = 0) And (tmAvail.iOrigLen = 0) Then
            tmAvail.iOrigUnit = tmAvail.iAvInfo And &H1F
            tmAvail.iOrigLen = tmAvail.iLen
        End If
        tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilNewUnit
        tmAvail.iLen = ilNewSec
        tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        imSsfRecLen = igSSFBaseLen + tmSsf(imSelectedDay).iCount * Len(tmProg)
        ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAvailRoom = False
        Exit Function
    End If
    mAvailRoom = True
    Exit Function
End Function
Private Function mBookSpot(tlSdfDateTimeInfo As SDFDATETIMEINFO, tlSdf As SDF, llCifCode As Long) As Boolean
    Dim slAirDate As String
    Dim slAirTime As String
    Dim ilAvailIndex As Integer
    Dim ilBkQH As Integer
    Dim ilRet As Integer
    Dim llSdfRecPos As Long
    Dim slRet As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilPriceLevel As Integer
    Dim ilRdf As Integer
    Dim slAgyCompliant As String
    
    mBookSpot = False
    
    ilVefCode = tlSdf.iVefCode
    ilVpfIndex = gBinarySearchVpfPlus(ilVefCode)    'gVpfFind(PostLog, imVefCode)
    If ilVpfIndex = -1 Then
        Exit Function
    End If
    
    slAirDate = tlSdfDateTimeInfo.sAirDate
    slAirTime = tlSdfDateTimeInfo.sAirTime
    slAgyCompliant = mGetCompliantStatus(gDateValue(slAirDate), gTimeToLong(slAirTime, False))

    tmChfSrchKey0.lCode = tlSdf.lChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmClfSrchKey0.lChfCode = tlSdf.lChfCode
        tmClfSrchKey0.iLine = tlSdf.iLineNo
        tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
        tmClfSrchKey0.iPropVer = tmChf.iPropVer
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
        If ilRdf <> -1 Then
            tmRdf = tgMRdf(ilRdf)
        End If
    Else
        Exit Function
    End If
    If Not mFindAvail(ilVefCode, slAirDate, slAirTime, 0, True, ilAvailIndex) Then
        Exit Function
    End If
    If Not mAvailRoom(ilVefCode, ilAvailIndex) Then
        Exit Function
    End If
    If Not mFindAvail(ilVefCode, slAirDate, slAirTime, 0, True, ilAvailIndex) Then
        Exit Function
    End If
    tmSdfSrchKey3.lCode = tlSdf.lCode
    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If
    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
    If slAgyCompliant = "O" Then
        slRet = "O"
    Else
        slRet = "S"
    End If
    'Test if time within daypart, if not set to Outside
    ilBkQH = IMPORTINVOICESPOT
    ilPriceLevel = 0
    ilRet = gBookSpot(slRet, hmSdf, tlSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmRdf, ilVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, ilPriceLevel, False, hmSxf, hmGsf)
    mBookSpot = ilRet
    If ilRet Then
    Else
    End If
    tmSdfSrchKey3.lCode = tlSdf.lCode
    ilRet = btrGetEqual(hmSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        If gDateValue(slAirDate) <= lmLastStdMnthBilled Then
            tlSdf.sBill = "Y"
        End If
        tlSdf.sAffChg = "Y"
        If llCifCode > 0 Then
            If tlSdf.lCopyCode <= 0 Then
                tlSdf.sPtType = "1"
                tlSdf.lCopyCode = llCifCode
            End If
        ElseIf llCifCode <= 0 Then
            If tlSdf.lCopyCode > 0 Then
                tlSdf.sPtType = ""
                tlSdf.lCopyCode = 0
            End If
        End If
        gPackDate slAirDate, tlSdf.iDate(0), tlSdf.iDate(1)
        gPackTime slAirTime, tlSdf.iTime(0), tlSdf.iTime(1)
        ilRet = btrUpdate(hmSdf, tlSdf, imSdfRecLen)
    End If
    
End Function
Private Sub mPopListKey()
    Dim llMaxWidth As Long
    lbcKey.Clear
    lbcKey.AddItem "Spot Color Code:"
    lbcKey.AddItem "     Cyan: Sold Count"
    lbcKey.AddItem "     Black: Count changed or Time specified"
    lbcKey.AddItem "When Save-All pressed:"
    lbcKey.AddItem "     All rows will be saved"
    lbcKey.AddItem "     as if all rows changed"
    lbcKey.AddItem "When Save-Changes pressed:"
    lbcKey.AddItem "     All Contract rows will be saved when"
    lbcKey.AddItem "     any value changed for the Contract"
    
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontName = "Arial"
    Traffic.pbcArial.FontBold = False
    Traffic.pbcArial.FontSize = 8
    llMaxWidth = (Traffic.pbcArial.TextWidth("     Black: Count changed or Time specified")) + 180
    lbcKey.Width = llMaxWidth
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.Height = (lbcKey.ListCount) * 225
    lbcKey.Height = gListBoxHeight(lbcKey.ListCount, 9)
    lbcKey.Move imcKey.Left, imcKey.Top + imcKey.Height
End Sub


Private Function mUpdateApf(llChfCode As Long, ilVefCode As Integer, slStnInvNo As String, slStnCntrNo As String) As Integer
    Dim llCntrNo As Long
    Dim llInvStartdate As Long
    Dim llInvEndDate As Long
    Dim llSdfDate As Long
    Dim llApfInvDate As Long
    Dim ilRet As Integer
    Dim ilApf As Integer
    ReDim llApfCode(0 To 0) As Long
    
    If (Asc(tgSaf(0).sFeatures3) And REQSTATIONPOSTING) = REQSTATIONPOSTING Then 'Require to Post spot prior to invoicing
        mUpdateApf = True
        Exit Function
    End If
    If (Asc(tgSaf(0).sFeatures2) And PAYMENTONCOLLECTION) <> PAYMENTONCOLLECTION Then 'Payment on Collection
        mUpdateApf = True
        Exit Function
    End If
    tmChfSrchKey0.lCode = llChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mUpdateApf = False
        Exit Function
    End If
    llCntrNo = tmChf.lCntrNo
    llInvStartdate = gDateValue(sgManualPostDate)
    llInvEndDate = gDateValue(gObtainEndStd(sgManualPostDate))
    tmApfSrchKey4.lCntrNo = llCntrNo
    gPackDate "1/1/1970", tmApfSrchKey4.iFullyPaidDate(0), tmApfSrchKey4.iFullyPaidDate(1)
    'ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORWRITE)
    ilRet = btrGetGreaterOrEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmApf.lCntrNo = llCntrNo)
        gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llApfInvDate
        '4/20/20: use date range
        'If (llInvEndDate = llApfInvDate) And (tmApf.iVefCode = ilVefCode) And (tmApf.lSbfCode = 0) Then
        If (llApfInvDate >= llInvStartdate) And (llApfInvDate <= llInvEndDate) And (tmApf.iVefCode = ilVefCode) And (tmApf.lSbfCode = 0) Then
            llApfCode(UBound(llApfCode)) = tmApf.lCode
            ReDim Preserve llApfCode(0 To UBound(llApfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hmApf, tmApf, imApfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    For ilApf = 0 To UBound(llApfCode) - 1 Step 1
        tmApfSrchKey0.lCode = llApfCode(ilApf)
        ilRet = btrGetEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            'Find lines
            tmApf.iAiredSpotCount = 0
            tmClfSrchKey1.lChfCode = llChfCode
            tmClfSrchKey1.iVefCode = ilVefCode
            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iVefCode = ilVefCode)
                If tmApf.lAcquisitionCost = tmClf.lAcquisitionCost Then
                    'Update aired count
                    tmSdfSrchKey0.iVefCode = ilVefCode
                    tmSdfSrchKey0.lChfCode = llChfCode
                    tmSdfSrchKey0.iLineNo = tmClf.iLine
                    tmSdfSrchKey0.lFsfCode = 0
                    gPackDateLong llInvStartdate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
                    tmSdfSrchKey0.sSchStatus = ""
                    gPackTime "12AM", tmSdfSrchKey0.iTime(0), tmSdfSrchKey0.iTime(1)
                    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.lChfCode = llChfCode) And (tmSdf.iLineNo = tmClf.iLine)
                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                        If llSdfDate > llInvEndDate Then
                            Exit Do
                        End If
                        If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                            tmApf.iAiredSpotCount = tmApf.iAiredSpotCount + 1
                        End If
                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            tmApf.sStationCntrNo = slStnCntrNo
            tmApf.sStationInvNo = slStnInvNo
            ilRet = btrUpdate(hmApf, tmApf, imApfRecLen)
        End If
    Next ilApf
    mUpdateApf = True
End Function
Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    If Not imUpdateAllowed Then
        mColOk = False
    Else
        mColOk = True
        If grdManualPost.CellBackColor = LIGHTYELLOW Then
            mColOk = False
            Exit Function
        End If
    End If
End Function

Private Sub mDetermineDP(slDays As String, slTimeRange As String, slDateRange As String, slEDIDays As String)
    Dim ilRet As Integer
    Dim ilRdf As Integer
    Dim llSdfDate As Long
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilDay As Integer
    Dim ilDP As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim llLnStartDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    
    slDays = ""
    slTimeRange = ""
    slDateRange = ""
    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
    tmChfSrchKey0.lCode = tmSdf.lChfCode
    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmClfSrchKey0.lChfCode = tmSdf.lChfCode
        tmClfSrchKey0.iLine = tmSdf.iLineNo
        tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
        tmClfSrchKey0.iPropVer = tmChf.iPropVer
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLnStartDate
        ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
        If ilRdf <> -1 Then
            tmRdf = tgMRdf(ilRdf)
        End If
        If (tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0) Then
            ReDim llStartTime(0 To 1) As Long
            ReDim llEndTime(0 To 1) As Long
            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "4", slStartTime
            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "4", slEndTime
            slTimeRange = slStartTime & "-" & slEndTime
        Else
            ReDim llStartTime(0 To 0) As Long
            ReDim llEndTime(0 To 0) As Long
            If ilRdf <> -1 Then
                For ilDP = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                    If (tmRdf.iStartTime(0, ilDP) <> 1) Or (tmRdf.iStartTime(1, ilDP) <> 0) Then
                        gUnpackTime tmRdf.iStartTime(0, ilDP), tmRdf.iStartTime(1, ilDP), "A", "4", slStartTime
                        gUnpackTime tmRdf.iEndTime(0, ilDP), tmRdf.iEndTime(1, ilDP), "A", "4", slEndTime
                        gUnpackTimeLong tmRdf.iStartTime(0, ilDP), tmRdf.iStartTime(1, ilDP), False, llRdfStartTime
                        gUnpackTimeLong tmRdf.iEndTime(0, ilDP), tmRdf.iEndTime(1, ilDP), True, llRdfEndTime
                        If llRdfStartTime < llRdfEndTime Then
                            If slTimeRange = "" Then
                                slTimeRange = slStartTime & "-" & slEndTime
                            Else
                                slTimeRange = slTimeRange & "; " & slStartTime & "-" & slEndTime
                            End If
                        Else
                            If slTimeRange = "" Then
                                slTimeRange = slStartTime & "-" & "12AM"
                            Else
                                slTimeRange = slTimeRange & "; " & slStartTime & "-" & "12AM"
                            End If
                            slTimeRange = slTimeRange & "; " & "12AM" & "-" & slEndTime
                        End If
                    End If
                Next ilDP
            End If
        End If
        tmCffSrchKey0.lChfCode = tmClf.lChfCode
        tmCffSrchKey0.iClfLine = tmClf.iLine
        tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
        tmCffSrchKey0.iPropVer = tmClf.iPropVer
        gPackDateLong llLnStartDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
        ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmCff.iClfLine = tmClf.iLine) And (tmCff.lChfCode = tmClf.lChfCode)
            gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llCffStartDate
            gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llCffEndDate
            If (llCffStartDate <= llSdfDate) And (llCffEndDate >= llSdfDate) Then
                'slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                If tmCff.sDyWk <> "W" Then
                    For ilDay = 0 To 6 Step 1
                        If ilDay <> gWeekDayLong(llSdfDate) Then
                            tmCff.iDay(ilDay) = 0
                        End If
                    Next ilDay
                    slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                    slStartDate = Format(llSdfDate, "m/d/yy")
                    slEndDate = slStartDate
                Else
                    slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                    For ilDay = 0 To 6 Step 1
                        If tmCff.iDay(ilDay) > 0 Then
                            slStartDate = Format(llSdfDate - (gWeekDayLong(llSdfDate) - ilDay), "m/d/yy")
                            Exit For
                        End If
                    Next ilDay
                    For ilDay = 6 To 0 Step -1
                        If tmCff.iDay(ilDay) > 0 Then
                            slEndDate = Format(llSdfDate + (ilDay - gWeekDayLong(llSdfDate)), "m/d/yy")
                            Exit For
                        End If
                    Next ilDay
                End If
                If gDateValue(slStartDate) = gDateValue(slEndDate) Then
                    slDateRange = slStartDate
                Else
                    slDateRange = slStartDate & "-" & slEndDate
                End If
                Exit Sub
            End If
            ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
End Sub

Private Function mGetCompliantStatus(llAirDate As Long, llAirTime As Long) As String
    Dim slDays As String
    Dim slTimeRange As String
    Dim slDateRange As String
    Dim slEDIDays As String
    Dim ilPos As Integer
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDPTime As Integer
    Dim slDPTimes() As String
    
    mGetCompliantStatus = "O"
    mDetermineDP slDays, slTimeRange, slDateRange, slEDIDays
    ilPos = InStr(1, slDateRange, "-", vbTextCompare)
    If ilPos > 0 Then
        llStartDate = gDateValue(Left$(slDateRange, ilPos - 1))
        llEndDate = gDateValue(Mid$(slDateRange, ilPos + 1))
    Else
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llStartDate
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llEndDate
    End If
    If (llAirDate >= llStartDate) And (llAirDate <= llEndDate) Then
        slDPTimes = Split(slTimeRange, ";")
        If IsArray(slDPTimes) Then
            For ilDPTime = 0 To UBound(slDPTimes) Step 1
                ilPos = InStr(1, slDPTimes(ilDPTime), "-", vbTextCompare)
                If ilPos > 0 Then
                    llStartTime = gTimeToLong(Left$(slDPTimes(ilDPTime), ilPos - 1), False)
                    llEndTime = gTimeToLong(Mid$(slDPTimes(ilDPTime), ilPos + 1), True)
                    If (llAirTime >= llStartTime) And (llAirTime <= llEndTime) Then
                        ilDay = gWeekDayLong(llAirDate)
                        If Mid$(slEDIDays, ilDay + 1, 1) = "Y" Then
                            mGetCompliantStatus = "A"
                            Exit For
                        End If
                    End If
                End If
            Next ilDPTime
        End If
    End If
    
End Function

Private Function mGetAcqCost() As String
    Dim llCntrNo As Long
    Dim llInvStartdate As Long
    Dim llInvEndDate As Long
    Dim llApfInvDate As Long
    Dim ilRet As Integer

    mGetAcqCost = gLongToStrDec(tmClf.lAcquisitionCost, 2)
    llCntrNo = tmChf.lCntrNo
    llInvStartdate = gDateValue(sgManualPostDate)
    llInvEndDate = gDateValue(gObtainEndStd(sgManualPostDate))
    tmApfSrchKey4.lCntrNo = llCntrNo
    gPackDate "1/1/1970", tmApfSrchKey4.iFullyPaidDate(0), tmApfSrchKey4.iFullyPaidDate(1)
    ilRet = btrGetGreaterOrEqual(hmApf, tmApf, imApfRecLen, tmApfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmApf.lCntrNo = llCntrNo)
        gUnpackDateLong tmApf.iInvDate(0), tmApf.iInvDate(1), llApfInvDate
        If (llInvEndDate = llApfInvDate) And (tmApf.iVefCode = igManualPostVefCode) And (tmApf.lSbfCode = 0) Then
        End If
        ilRet = btrGetNext(hmApf, tmApf, imApfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
End Function
