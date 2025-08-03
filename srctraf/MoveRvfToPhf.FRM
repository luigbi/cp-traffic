VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MoveRvfToPhf 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move Receivables To History"
   ClientHeight    =   2670
   ClientLeft      =   465
   ClientTop       =   2340
   ClientWidth     =   7095
   ClipControls    =   0   'False
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
   Icon            =   "MoveRvfToPhf.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2670
   ScaleWidth      =   7095
   Begin vb6MoveRvfToPhf.CSI_Calendar edcDate 
      Height          =   330
      Left            =   2730
      TabIndex        =   5
      Top             =   135
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      Text            =   "7/25/2019"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   0   'False
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   4
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Left            =   2085
      TabIndex        =   4
      Top             =   1455
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6345
      Top             =   165
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Move"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   1785
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   1785
      Width           =   1050
   End
   Begin VB.Label lacDate 
      Caption         =   "Move all transaction prior to"
      Height          =   315
      Left            =   135
      TabIndex        =   6
      Top             =   165
      Width           =   2625
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   2190
      Width           =   6585
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1875
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "MoveRvfToPhf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: MoveRvfToPhf.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
Dim hmRvf As Integer    'Avail Summary file handle
Dim tmRvf As RVF
Dim imRvfRecLen As Integer        'Asf record length
'Spot file
Dim hmPhf As Integer
Dim imUserInput As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imMovingRecs As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmStartDate As Long
Dim lmEndDate As Long
Dim rvf_rst As ADODB.Recordset
Private Sub cmcCancel_Click()
    If imMovingRecs Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'
'           1-29-01 Remove everything in file before genning the weeks summary
'                   previously only removed the vehicle it was processing.  Dormant ones remained on file.
Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim ilRif As Integer
    Dim ilWk As Integer
    Dim ilWkMinus1 As Integer
    Dim llPercent As Long
    Dim slDate As String
    Dim slSQLQuery As String
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim llRet As Long
    
    lacInfo.Visible = False
    If imMovingRecs Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If slDate = "" Then
        MsgBox "Please enter the transaction date"
        Exit Sub
    End If
    If Not gIsDate(slDate) Then
        MsgBox "Transaction date in error"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    slSQLQuery = "Select Count(1) as RvfCount From rvf_Receivables "
    slSQLQuery = slSQLQuery & " Where rvfTranDate < '" & Format(slDate, sgSQLDateForm) & "'"
    Set rvf_rst = gSQLSelectCall(slSQLQuery)
    If rvf_rst.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No Transactions to move"
        Exit Sub
    End If
    If rvf_rst!RvfCount <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "No Transactions to move"
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    ilRet = MsgBox(rvf_rst!RvfCount & " Receivable records will be transferred to History, Ok to Continue", vbOKCancel + vbQuestion, "Question")
    If ilRet = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmcCancel.Caption = "&Cancel"
    imMovingRecs = True
    lacInfo.Visible = True
    lacInfo.Caption = ""
    lmTotalNoRecs = rvf_rst!RvfCount
    '2-02-01
    hmRvf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", MoveRvfToPhf
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf)
    
    hmPhf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", MoveRvfToPhf
    On Error GoTo 0
    '4/12/18: required to load tgSaf(0) which is required for testing options
    gObtainSAF
    '5/7/18
    lmProcessedNoRecs = 0
    plcGauge.Visible = True
    plcGauge.Value = 0
    
        
    
    ilExtLen = Len(tmRvf)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlItf) 'Obtain number of records
    btrExtClear hmRvf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to access records, error # " & ilRet
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        Exit Sub
    End If
    Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "Rvf", "") 'Set extract limits (all records)
    gPackDate gDecOneDay(slDate), tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
    ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    ilOffSet = 0
    ilRet = btrExtAddField(hmRvf, ilOffSet, ilExtLen)  'Extract iCode field
    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tmRvf)  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            
            tmRvf.lCode = 0
            ilRet = btrInsert(hmPhf, tmRvf, imRvfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                MsgBox "Unable to insert History record, error # " & ilRet & ". Restore database"
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                ilRet = btrClose(hmPhf)
                btrDestroy hmPhf
                Exit Sub
            End If
            
            lmProcessedNoRecs = lmProcessedNoRecs + 1
            llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            plcGauge.Value = llPercent
            
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    If lmProcessedNoRecs <> lmTotalNoRecs Then
        Screen.MousePointer = vbDefault
        MsgBox "Not all Receivable records trasferred to History" & ". Restore database"
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        Exit Sub
    End If
    slSQLQuery = "Delete From rvf_Receivables "
    slSQLQuery = slSQLQuery & " Where rvfTranDate < '" & Format(slDate, sgSQLDateForm) & "'"
    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    If llRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to complete removal task, error # " & ilRet & ". Restore database"
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        Exit Sub
    End If
    lacInfo.Visible = True
    imMovingRecs = False
    cmcCancel.Caption = "&Done"
    MsgBox "Transafer of the transaction records completed "
    '9/27/18: Getting Invalid procedure call
    'cmcCancel.SetFocus
    cmcGen.Enabled = False
    ilRet = btrClose(hmPhf)
    ilRet = btrClose(hmRvf)
    btrDestroy hmPhf
    btrDestroy hmRvf
    Screen.MousePointer = vbDefault
    If imTerminate Then
        'tmcTimer.Enabled = False
        cmcCancel_Click
    End If
    Exit Sub
cmcExportErr:
    imTerminate = True
    ilRet = Err
    Resume Next
End Sub

Private Sub Form_Load()
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String
    
    mInit
    If imTerminate Then
        tmcTimer.Enabled = True
    Else
            slDate = Format$(Now(), "m/d/yy")
    slMonth = Month(slDate)
    slYear = Year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(Str$(ilValue))
    Do While Len(slStr) < 4
        slStr = "0" & slStr
    Loop
    sgSpecialPassword = slStr
    'smSpecial = "Login" & slStr
    'smGGBGSpecial = "8373191"

        CSPWord.Show vbModal
        If Not igPasswordOk Then
            imTerminate = True
            tmcTimer.Enabled = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    
    ilRet = btrClose(hmPhf)
    btrDestroy hmPhf
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    btrStopAppl
    
    Set MoveRvfToPhf = Nothing   'Remove data segment

    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'
'      2-02-01 Move Opens to gcmcGen; right before the
'           generation of summary records.  The timer
'           wasnt activated until after opening; therefore
'           files could remain open and the backup
'           aborted
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    imTerminate = False
    mParseCmmdLine
    
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
'Rm**    btrDestroy hgHlf
    Screen.MousePointer = vbHourglass
    imMovingRecs = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    lmStartDate = lmNowDate
    Do While gWeekDayLong(lmStartDate) <> 0
        lmStartDate = lmStartDate - 1
    Loop
    lmEndDate = lmStartDate + 7 * 26 - 1
    lacTo.Caption = sgDBPath
    gCenterStdAlone MoveRvfToPhf
    Screen.MousePointer = vbDefault
    'imMinElapsed = imTimeoutInterval - 2
    'If imTimeoutInterval > 0 Then
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim slChar As String
    slCommand = Command$
    igDirectCall = 0
    sgIniPath = ""
    sgCommandStr = "debug"
    igStdAloneMode = True 'Switch from/to stand alone mode
    sgCallAppName = ""
    ilTestSystem = False
    ilPos = InStr(1, slCommand, "/S:Test", 1)
    If ilPos > 0 Then
        ilTestSystem = True
    End If
    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    If ilPos > 0 Then
        ilTestSystem = False
    End If
    imUserInput = True
    igBkgdProg = 0
    'If imTimeoutInterval > 0 Then
    '    igBkgdProg = 5
    'End If
    slStr = "Guide"
    gInitStdAlone MoveRvfToPhf, slStr, ilTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "MoveRvfToPhf.Frm"
    gLogMsg "Current Mode: UserInput.", "MoveRvfToPhf.txt", False
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload MoveRvfToPhf
    igManUnload = NO
End Sub



Private Sub tmcTimer_Timer()
    tmcTimer.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
End Sub


Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Rvf.Btr", Len(tmRvf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Phf.Btr", Len(tmRvf)) Then
        mCheckRecLength = False
        Exit Function
    End If

    mCheckRecLength = True
End Function


