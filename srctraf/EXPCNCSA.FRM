VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ExpCnCSA 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2355
   ClientLeft      =   465
   ClientTop       =   2340
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   7095
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6060
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   0
      Width           =   2910
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5745
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6015
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcTo 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
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
      Left            =   5415
      TabIndex        =   4
      Top             =   330
      Width           =   1485
   End
   Begin VB.PictureBox plcTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1035
      ScaleHeight     =   345
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   300
      Width           =   4305
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   45
         TabIndex        =   3
         Top             =   105
         Width           =   4185
      End
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   1695
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
      Left            =   3735
      TabIndex        =   6
      Top             =   1695
      Width           =   1050
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   375
      Left            =   2295
      TabIndex        =   11
      Top             =   1215
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   1455
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1890
      TabIndex        =   10
      Top             =   900
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      Caption         =   "To File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   375
      Width           =   810
   End
End
Attribute VB_Name = "ExpCnCSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPCNCSA.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ExpCnCSA.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
Dim hmTo As Integer   'From file hanle
'Vehicle Link
Dim hmVlf As Integer    'Vehicle Link file handle
Dim tmVlf As VLF
Dim imVlfRecLen As Integer        'Vlf record length
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim imVefRecLen As Integer     'VEF record length
Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim imFirstActivate As Integer

'' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    lacInfo.Visible = False
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    slToFile = Trim$(edcTo.Text)
    If Len(slToFile) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Sub
    End If
    'If InStr(slToFile, ":") = 0 Then
    If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
        slToFile = sgExportPath & slToFile
    End If
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        edcTo.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    sgMessageFile = sgDBPath & "Messages\" & "ExpCncSAL.Txt"
    
    
    gAutomationAlertAndLogHandler "** Export CNC Selling to Airing Links **"
    gAutomationAlertAndLogHandler "* Storing Output into " & slToFile
    
    
    imExporting = True
    lmTotalNoRecs = btrRecords(hmVlf)
    plcGauge.Visible = True
    'plcGauge.FloodPercent = 0
    plcGauge.Visible = True
    lacInfo.Visible = True
    lacInfo.Caption = ""
    ilRet = mExpRecs()
    If ilRet Then
        'plcGauge.FloodPercent = 100
        plcGauge.Value = 100
        lacInfo.Caption = "Export Successfully Completed"
        gAutomationAlertAndLogHandler "Export Successfully Completed"
    Else
        lacInfo.Caption = "Export Failed or was Terminated"
        gAutomationAlertAndLogHandler "Export Failed or was Terminated"
    End If
    lacInfo.Visible = True
    Close hmTo
    imExporting = False
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Csv"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo.Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub
Private Sub cmcTo_GotFocus()
    lacInfo.Visible = False
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcTo_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
    lacInfo.Visible = False
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
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
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
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    
    Set ExpCnCSA = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mExpRecs                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Export Recs                    *
'*                                                     *
'*******************************************************
Private Function mExpRecs() As Integer
'
'   iRet = mExpRecs()
'   Where:
'       iRet (O)- True if record exported,
'                 False if error
'
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim slTime As String
    Dim slRecord As String
    Dim llPercent As Long
    Dim llTermDate As Long
    ilExtLen = Len(tmVlf)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
    btrExtClear hmVlf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmVlf, tmVlf, imVlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mExpRecs = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mExpRecs = False
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hmVlf, llNoRec, -1, "UC", "VLF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hmVlf, ilOffSet, ilExtLen)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        mExpRecs = False
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlVef)    'Extract record
    ilRet = btrExtGetNext(hmVlf, tmVlf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mExpRecs = False
            Exit Function
        End If
        ilExtLen = Len(tmVlf)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlVef, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmVlf, tmVlf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            'DoEvents
            igDOE = igDOE + 1
            If igDOE > 1000 Then
                DoEvents
                igDOE = 0
            End If
            If imTerminate Then
                mExpRecs = False
                Exit Function
            End If
            If tmVlf.sStatus <> "P" Then
                'Build Export Record
                slRecord = Trim$(str$(tmVlf.lCode)) & sgCnCDelimiter
                'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilLoop).iCode = tmVlf.iSellCode Then
                    ilLoop = gBinarySearchVef(tmVlf.iSellCode)
                    If ilLoop <> -1 Then
                        slRecord = slRecord & Trim$(tgMVef(ilLoop).sCodeStn) & sgCnCDelimiter & Trim$(str$(tmVlf.iSellCode)) & sgCnCDelimiter & Trim$(tgMVef(ilLoop).sName) & sgCnCDelimiter
                '        Exit For
                    End If
                'Next ilLoop
                slRecord = slRecord & Trim$(str$(tmVlf.iSellDay)) & sgCnCDelimiter
                gUnpackTime tmVlf.iSellTime(0), tmVlf.iSellTime(1), "A", "1", slTime
                slTime = Format$(gConvertTime(slTime), "hhmmss")
                slRecord = slRecord & slTime & sgCnCDelimiter
                slRecord = slRecord & Trim$(str$(tmVlf.iSellPosNo)) & sgCnCDelimiter
                slRecord = slRecord & Trim$(str$(tmVlf.iSellSeq)) & sgCnCDelimiter
                slRecord = slRecord & Trim$(tmVlf.sStatus) & sgCnCDelimiter
                'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilLoop).iCode = tmVlf.iAirCode Then
                    ilLoop = gBinarySearchVef(tmVlf.iAirCode)
                    If ilLoop <> -1 Then
                        slRecord = slRecord & Trim$(tgMVef(ilLoop).sCodeStn) & sgCnCDelimiter & Trim$(str$(tmVlf.iAirCode)) & sgCnCDelimiter & Trim$(tgMVef(ilLoop).sName) & sgCnCDelimiter
                '        Exit For
                    End If
                'Next ilLoop
                slRecord = slRecord & Trim$(str$(tmVlf.iAirDay)) & sgCnCDelimiter
                gUnpackTime tmVlf.iAirTime(0), tmVlf.iAirTime(1), "A", "1", slTime
                slTime = Format$(gConvertTime(slTime), "hhmmss")
                slRecord = slRecord & slTime & sgCnCDelimiter
                slRecord = slRecord & Trim$(str$(tmVlf.iAirPosNo)) & sgCnCDelimiter
                slRecord = slRecord & Trim$(str$(tmVlf.iAirSeq)) & sgCnCDelimiter
                gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slStr
                slStr = Format$(gDateValue(slStr), "mmddyyyy")
                slRecord = slRecord & slStr & sgCnCDelimiter
                gUnpackDate tmVlf.iTermDate(0), tmVlf.iTermDate(1), slStr
                If slStr <> "" Then
                    llTermDate = gDateValue(slStr)
                    slStr = Format$(llTermDate, "mmddyyyy")
                Else
                    llTermDate = 0
                End If
                slRecord = slRecord & slStr
                If (llTermDate = 0) Or (llTermDate >= lmNowDate) Then
                    Print #hmTo, slRecord
                End If
            End If
            lmProcessedNoRecs = lmProcessedNoRecs + 1
            lacInfo.Caption = Trim$(str$(lmProcessedNoRecs)) & " of" & str$(lmTotalNoRecs)
            llPercent = (lmProcessedNoRecs * CSng(100)) / lmTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            'plcGauge.FloodPercent = llPercent
            plcGauge.Value = llPercent
            ilRet = btrExtGetNext(hmVlf, tmVlf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmVlf, tmVlf, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mExpRecs = True
    Exit Function

    On Error GoTo 0
    mExpRecs = False
    Exit Function
End Function
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
    Dim slDate As String

    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSA
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCSA
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    ilRet = gObtainVef()
    gCenterStdAlone ExpCnCSA
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr  'Command$
    ilPos = InStr(1, slCommand, "/CS", 1)
'    If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
'        igStdAloneMode = True 'False  'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = False
'        ilPos = InStr(1, slCommand, "/S:Test", 1)
'        If ilPos > 0 Then
'            ilTestSystem = True
'        End If
'        ilPos = InStr(1, slCommand, "/S:Prod", 1)
'        If ilPos > 0 Then
'            ilTestSystem = False
'        End If
'        slCommand = "Traffic\Guide"
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            ''MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            gAutomationAlertAndLogHandler "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        'imShowHelpMsg = True
        'ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        'If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
        '    imShowHelpMsg = False
        'End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ExpCnCSA, slStr, ilTestSystem
    igCmmCallSource = Val(slStr)
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

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpCnCSA
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Export Selling To Airing  Links"
End Sub
