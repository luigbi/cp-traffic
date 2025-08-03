VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ExpCnCAP 
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
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   375
      Left            =   2295
      TabIndex        =   10
      Top             =   1185
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
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
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5745
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6015
      TabIndex        =   7
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   300
      Width           =   4305
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   45
         TabIndex        =   2
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
      TabIndex        =   4
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
      TabIndex        =   5
      Top             =   1695
      Width           =   1050
   End
   Begin VB.Label plcScreen 
      Caption         =   "Export Orders by Advertisers/Product"
      Height          =   210
      Left            =   45
      TabIndex        =   11
      Top             =   30
      Width           =   5160
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1890
      TabIndex        =   9
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
      TabIndex        =   0
      Top             =   375
      Width           =   810
   End
End
Attribute VB_Name = "ExpCnCAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPCNCAP.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ExpCncAP.Frm
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
'Contract Header
Dim hmCHF As Integer
Dim tmChf As CHF
Dim imCHFRecLen As Integer  'CHF record length
'Contract Line
Dim hmClf As Integer
Dim tmClf As CLF
Dim imClfRecLen As Integer  'CLF record length
'Contract Flight
Dim hmCff As Integer
Dim tmCff As CFF
Dim imCffRecLen As Integer  'CHF record length
'Product name
Dim hmPrf As Integer
Dim tmPrf As PRF
Dim tmPrfSrchKey As PRFKEY1 'PRF key record image
Dim imPrfRecLen As Integer  'PRF record length
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim imAdfRecLen As Integer  'ANF record length
'Short Title
Dim hmSif As Integer
Dim tmSif As SIF
Dim tmSifSrchKey As LONGKEY0 'SIF key record image
Dim imSifRecLen As Integer  'SIF record length
'Vehicle/Short Title
Dim hmVsf As Integer
Dim tmVsf As VSF
Dim tmVsfSrchKey As LONGKEY0
Dim imVsfRecLen As Integer  'VSF record length
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
    sgMessageFile = sgDBPath & "Messages\" & "ExpCncAP.Txt"
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
    
    gAutomationAlertAndLogHandler "** Export CNC Orders by Advertisers/Product **"

    
    imExporting = True
    lacInfo.Visible = True
    lacInfo.Caption = "Obtain Contracts to be Exported"
    
    gAutomationAlertAndLogHandler "Obtaining Contracts to be Exported.."
    
    ilRet = mObtainCntrCodes()
    lacInfo.Caption = ""
    lmTotalNoRecs = UBound(lgCntrCode)
    plcGauge.Visible = True
    
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    
    gAutomationAlertAndLogHandler "Exporting..."
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
    
    Erase lgCntrCode
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    
    Set ExpCnCAP = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
'    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
'    Traffic!cdcSetup.Action = 6
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
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return status
    Dim slRecord As String
    Dim llPercent As Long
    Dim ilCntr As Integer
    Dim slAdvtName As String
    Dim llPrfCode As Long
    Dim llSifCode As Long
    Dim slCompName As String
    Dim slDemoName As String
    Dim llDate As Long
    For ilCntr = LBound(lgCntrCode) To UBound(lgCntrCode) - 1 Step 1
        'ilRet = gObtainCntr(hmChf, hmClf, hmCff, lgCntrCode(ilCntr), False, tgChf, tgClf(), tgCff())
        ilRet = gObtainChfClf(hmCHF, hmClf, lgCntrCode(ilCntr), False, tgChf, tgClf())
        If Not ilRet Then
            mExpRecs = False
            Exit Function
        End If
        DoEvents
        If imTerminate Then
            mExpRecs = False
            Exit Function
        End If
        slAdvtName = "Advt Missing"
        'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
        '    If tgCommAdf(ilLoop).iCode = tgChf.iAdfCode Then
            ilLoop = gBinarySearchAdf(tgChf.iAdfCode)
            If ilLoop <> -1 Then
                slAdvtName = sgCnCNameBracket & Trim$(tgCommAdf(ilLoop).sName) & sgCnCNameBracket
        '        Exit For
            End If
        'Next ilLoop
        llPrfCode = 0
        tmPrfSrchKey.iAdfCode = tgChf.iAdfCode
        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = tgChf.iAdfCode)
            If StrComp(Trim$(tgChf.sProduct), Trim$(tmPrf.sName), 1) = 0 Then
                llPrfCode = tmPrf.lCode
                Exit Do
            End If
            ilRet = btrGetNext(hmPrf, tmPrf, imPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
        If tgChf.lVefCode < 0 Then
            tmVsfSrchKey.lCode = -tgChf.lVefCode
            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                mExpRecs = False
                Exit Function
            End If
        End If
        If tgChf.iMnfComp(0) > 0 Then
            slCompName = "Competitive Missing"
            For ilLoop = LBound(tgExclMnf) To UBound(tgExclMnf) - 1 Step 1
                If tgExclMnf(ilLoop).iCode = tgChf.iMnfComp(0) Then
                    slCompName = Trim$(tgExclMnf(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
        Else
            slCompName = ""
        End If
        If tgChf.iMnfDemo(0) > 0 Then
            slDemoName = "Demo Missing"
            For ilLoop = LBound(tgDemoMnf) To UBound(tgDemoMnf) - 1 Step 1
                If tgDemoMnf(ilLoop).iCode = tgChf.iMnfDemo(0) Then
                    slDemoName = Trim$(tgDemoMnf(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
        Else
            slDemoName = ""
        End If
        'For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            slRecord = Trim$(str$(tgChf.iAdfCode)) & sgCnCDelimiter
            slRecord = slRecord & slAdvtName & sgCnCDelimiter
            slRecord = slRecord & Trim$(str$(llPrfCode)) & sgCnCDelimiter
            slRecord = slRecord & sgCnCNameBracket & Trim$(tgChf.sProduct) & sgCnCNameBracket & sgCnCDelimiter
            llSifCode = 0
            'If tgChf.lVefCode < 0 Then
            '    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
            '        If tmVsf.iFSCode(ilVsf) > 0 Then
            '            If (tgClf(ilClf).ClfRec.iVefCode = tmVsf.iFSCode(ilVsf)) Then
            '                If tmVsf.lFSComm(ilVsf) > 0 Then
            '                    llSifCode = tmVsf.lFSComm(ilVsf)
            '                End If
            '                Exit For
            '            End If
            '        End If
            '    Next ilVsf
            'End If
            If llSifCode <= 0 Then
                If tgChf.lSifCode <= 0 Then
                    slRecord = slRecord & Trim$(str$(0)) & sgCnCDelimiter
                    slRecord = slRecord & Trim$("No Short Title") & sgCnCDelimiter
                Else
                    tmSifSrchKey.lCode = tgChf.lSifCode
                    ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slRecord = slRecord & Trim$(str$(tgChf.lSifCode)) & sgCnCDelimiter
                        slRecord = slRecord & sgCnCNameBracket & Trim$(tmSif.sName) & sgCnCNameBracket & sgCnCDelimiter
                    Else
                        slRecord = slRecord & Trim$(str$(tgChf.lSifCode)) & sgCnCDelimiter
                        slRecord = slRecord & Trim$("No Short Title") & sgCnCDelimiter
                    End If
                End If
            Else
                tmSifSrchKey.lCode = llSifCode
                ilRet = btrGetEqual(hmSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slRecord = slRecord & Trim$(str$(llSifCode)) & sgCnCDelimiter
                    slRecord = slRecord & sgCnCNameBracket & Trim$(tmSif.sName) & sgCnCNameBracket & sgCnCDelimiter
                Else
                    slRecord = slRecord & Trim$(str$(llSifCode)) & sgCnCDelimiter
                    slRecord = slRecord & Trim$("No Short Title") & sgCnCDelimiter
                End If
            End If
            slRecord = slRecord & Trim$(str$(tgChf.iMnfComp(0))) & sgCnCDelimiter
            slRecord = slRecord & sgCnCNameBracket & slCompName & sgCnCNameBracket & sgCnCDelimiter
            gUnpackDateLong tgChf.iStartDate(0), tgChf.iStartDate(1), llDate
            slRecord = slRecord & Format$(llDate, "MMDDYYYY") & sgCnCDelimiter
            gUnpackDateLong tgChf.iEndDate(0), tgChf.iEndDate(1), llDate
            slRecord = slRecord & Format$(llDate, "MMDDYYYY") & sgCnCDelimiter
            'slRecord = slRecord & Trim$(Str$(tgClf(ilClf).ClfRec.iLen)) & sgCnCDelimiter
            'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If tgMVef(ilLoop).iCode = tgClf(ilClf).ClfRec.iVefCode Then
            '        slRecord = slRecord & Trim$(Str$(tgClf(ilClf).ClfRec.iVefCode)) & sgCnCDelimiter & Trim$(tgMVef(ilLoop).sName) & sgCnCDelimiter
            '        Exit For
            '    End If
            'Next ilLoop
            slRecord = slRecord & Trim$(str$(tgChf.lInputGross)) & sgCnCDelimiter
            slRecord = slRecord & Trim$(str$(tgChf.lCntrNo)) & sgCnCDelimiter
            slRecord = slRecord & slDemoName & sgCnCDelimiter
            slRecord = slRecord & Trim$(str$(tgChf.iPctTrade))
            Print #hmTo, slRecord
        'Next ilClf
        lmProcessedNoRecs = lmProcessedNoRecs + 1
        lacInfo.Caption = Trim$(str$(lmProcessedNoRecs)) & " of" & str$(lmTotalNoRecs)
        llPercent = (lmProcessedNoRecs * CSng(100)) / lmTotalNoRecs
        If llPercent >= 100 Then
            llPercent = 100
        End If
        'plcGauge.FloodPercent = llPercent
        plcGauge.Value = llPercent
    Next ilCntr
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
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCnCAP
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    ilRet = gObtainAdvt()
    ilRet = gObtainVef()
    sgExclMnfStamp = ""
    ilRet = gObtainMnfForType("C", sgExclMnfStamp, tgExclMnf())
    sgDemoMnfStamp = ""
    ilRet = gObtainMnfForType("D", sgDemoMnfStamp, tgDemoMnf())
    gCenterStdAlone ExpCnCAP
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
'*      Procedure Name:mObtainCntrCodes                *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain contracts to be exported*
'*                                                     *
'*******************************************************
Private Function mObtainCntrCodes() As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer    'Return status
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    ReDim lgCntrCode(0 To 0) As Long
    ilExtLen = 4  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrGetFirst(hmCHF, tmChf, imCHFRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainCntrCodes = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainCntrCodes = False
            Exit Function
        End If
    End If
    'Call btrExtSetBounds(hmChf, llNoRec, -1, "UC") 'Set extract limits (all records)
    Call btrExtSetBounds(hmCHF, llNoRec, -1, "UC", "CHFCODEPK", "L") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "Y"
    ilOffSet = gFieldOffset("Chf", "ChfDelete")
    ilRet = btrExtAddLogicConst(hmCHF, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
    tlCharTypeBuff.sType = "O"
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddLogicConst(hmCHF, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
    tlCharTypeBuff.sType = "H"
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddLogicConst(hmCHF, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)

    ilOffSet = gFieldOffset("Chf", "ChfCode")
    ilRet = btrExtAddField(hmCHF, ilOffSet, 4)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainCntrCodes = False
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlVef)    'Extract record
    ilRet = btrExtGetNext(hmCHF, lgCntrCode(UBound(lgCntrCode)), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainCntrCodes = False
            Exit Function
        End If
        ilExtLen = 4  'Extract operation record size
        'ilRet = btrExtGetFirst(hlVef, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmCHF, lgCntrCode(UBound(lgCntrCode)), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            DoEvents
            If imTerminate Then
                mObtainCntrCodes = False
                Exit Function
            End If
            ReDim Preserve lgCntrCode(0 To UBound(lgCntrCode) + 1) As Long
            ilRet = btrExtGetNext(hmCHF, lgCntrCode(UBound(lgCntrCode)), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmCHF, lgCntrCode(UBound(lgCntrCode)), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mObtainCntrCodes = True
    Exit Function

    On Error GoTo 0
    mObtainCntrCodes = False
    Exit Function
End Function
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
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
    '    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    ilPos = InStr(1, slCommand, "/S:Test", 1)
    '    If ilPos > 0 Then
    '        ilTestSystem = True
    '    End If
    '    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    '    If ilPos > 0 Then
    '        ilTestSystem = False
    '    End If
    '    slCommand = "Traffic\Guide"
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
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
        igCmmCallSource = Val(slStr)
    'End If
    'gInitStdAlone ExpCnCAP, slStr, ilTestSystem
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

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpCnCAP
    igManUnload = NO
End Sub
