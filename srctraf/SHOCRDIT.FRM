VERSION 5.00
Begin VB.Form ShoCrdit 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6150
   ClientLeft      =   210
   ClientTop       =   1530
   ClientWidth     =   9225
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   9225
   Begin VB.CommandButton cmcCheck 
      Appearance      =   0  'Flat
      Caption         =   "&Check"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5430
      TabIndex        =   10
      Top             =   5685
      Width           =   945
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1260
      Top             =   5655
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   4140
      TabIndex        =   4
      Top             =   5685
      Width           =   945
   End
   Begin VB.PictureBox plcCredit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5205
      Left            =   60
      ScaleHeight     =   5145
      ScaleWidth      =   8985
      TabIndex        =   3
      Top             =   240
      Width           =   9045
      Begin VB.PictureBox pbcLbcCredit 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4830
         Left            =   75
         ScaleHeight     =   4830
         ScaleWidth      =   8610
         TabIndex        =   5
         Top             =   255
         Width           =   8610
      End
      Begin VB.ListBox lbcCredit 
         Appearance      =   0  'Flat
         Height          =   4860
         Left            =   60
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   8925
      End
      Begin VB.Label lacDollars 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDFFD7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dollars"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   6630
         TabIndex        =   9
         Top             =   45
         Width           =   2355
      End
      Begin VB.Label lacCreditStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDFFD7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Credit Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   3510
         TabIndex        =   8
         Top             =   45
         Width           =   3135
      End
      Begin VB.Label lacAdvt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDFFD7&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Advertiser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   60
         TabIndex        =   7
         Top             =   45
         Width           =   3465
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   1950
      Width           =   15
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2805
      TabIndex        =   1
      Top             =   5685
      Width           =   945
   End
   Begin VB.Label lacScreen 
      BackColor       =   &H000000FF&
      Caption         =   "Credit Alert"
      Height          =   225
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   5895
   End
End
Attribute VB_Name = "ShoCrdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SHOCRDIT.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmAgfSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ShoCrdit.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Show Advertiser/Agency Credit limit input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmAgf As AGF        'Agf record image
Dim hmAgf As Integer    'Agency file handle
Dim imAgfRecLen As Integer        'Agf record length
Dim tmAgyCredit As AGYCREDIT
Dim tmAdf As ADF        'Adf record image
Dim hmAdf As Integer    'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAdfSrchKey As INTKEY0 'ADF key record image
Dim tmAdvtCredit As ADVTCREDIT
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imListFieldCredit(0 To 4) As Integer    'Index zero ignored
Dim imFirstActivate As Integer
Dim imLastSortCol As Integer
Dim imSortDirection As Integer
Dim tmShoCreditSort() As SHOCREDITSORT


Private Sub cmcCheck_Click()
    mBranch
End Sub

Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    If imFirstFocus Then
        If lbcCredit.ListCount <= 0 Then
            'mTerminate
            'Exit Sub
        End If
        imFirstFocus = False
        'mInitDDE
        If igUpdateAllowed = False Then
            'mSendHelpMsg "BF"
        Else
            'mSendHelpMsg "BT"
        End If
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String        'General string

    igRptCallType = COLLECTIONSJOB
    igRptType = 1
        If igTestSystem Then
            slStr = "ShoCrdit^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "ShoCrdit^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    sgCommandStr = slStr & "\||9\Advertiser and Agency Credit Status"       '2-24-05 automatically run the agy&adv credit status report
    RptSelCreditStatus.Show vbModal         '7-6-15 changed to remove modules from traffic due to out of memory error
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
'    gShowBranner
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        'Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    'Me.KeyPreview = True
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    'Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
    '    'plcCalendar.Visible = False
    '    gFunctionKeyBranch KeyCode
    'End If

End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        ShoCrdit.Height = 0
        ShoCrdit.Width = 0
        tmcStart.Enabled = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    Erase tmShoCreditSort
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmAgf   'Clear any previous extend operation
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf

    Set ShoCrdit = Nothing   'Remove data segment

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser/Agency     *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slCredit                                                                              *
'******************************************************************************************

    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slCreditLimit As String
    Dim slCurrAR As String
    Dim slUnbilled As String
    Dim slOwned As String
    Dim slStr As String
    ilExtLen = Len(tmAdvtCredit)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmAdf, tmAdf, imAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'Call btrExtSetBounds(hmAdf, llNoRec, -1, "UC") 'Set extract limits (all records)
    Call btrExtSetBounds(hmAdf, llNoRec, -1, "UC", "ADVTCREDITEXTPK", ADVTCREDITEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Adf", "AdfCode")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfName")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 30)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfAddrID")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 9) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfCreditRestr")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfCreditLimit")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 4) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfCurrAR")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 6) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfUnbilled")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 6) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfCrdApp")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Adf", "AdfBilAgyDir")
    ilRet = btrExtAddField(hmAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hmAdf)    'Extract record
    ilRet = btrExtGetNext(hmAdf, tmAdvtCredit, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
        ilExtLen = Len(tmAdvtCredit)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmAdf, tmAdvtCredit, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            tmShoCreditSort(UBound(tmShoCreditSort)).iType = 0
            tmShoCreditSort(UBound(tmShoCreditSort)).sKey = ""
            tmShoCreditSort(UBound(tmShoCreditSort)).sName = tmAdvtCredit.sName
            tmShoCreditSort(UBound(tmShoCreditSort)).sAddrID = tmAdvtCredit.sAddrID
            tmShoCreditSort(UBound(tmShoCreditSort)).sCityID = ""
            tmShoCreditSort(UBound(tmShoCreditSort)).sCreditRestr = tmAdvtCredit.sCreditRestr
            tmShoCreditSort(UBound(tmShoCreditSort)).lCreditLimit = tmAdvtCredit.lCreditLimit
            tmShoCreditSort(UBound(tmShoCreditSort)).lDollars = 0
            tmShoCreditSort(UBound(tmShoCreditSort)).sCrdApp = tmAdvtCredit.sCrdApp
            tmShoCreditSort(UBound(tmShoCreditSort)).sBillAgyDir = tmAdvtCredit.sBillAgyDir
            tmShoCreditSort(UBound(tmShoCreditSort)).iCode = tmAdvtCredit.iCode
            If (tmAdvtCredit.sCrdApp = "R") Then
                ReDim Preserve tmShoCreditSort(0 To UBound(tmShoCreditSort) + 1) As SHOCREDITSORT
            Else
                If tmAdvtCredit.sCreditRestr = "L" Then  'Credit Limit
                    slCreditLimit = gLongToStrDec(tmAdvtCredit.lCreditLimit, 2)
                    gPDNToStr tmAdvtCredit.sCurrAR, 2, slCurrAR
                    gPDNToStr tmAdvtCredit.sUnbilled, 2, slUnbilled
                    slOwned = gAddStr(slCurrAR, slUnbilled)
                    If gCompNumberStr(slOwned, slCreditLimit) > 0 Then
                        slStr = gSubStr(slOwned, slCreditLimit)
                        tmShoCreditSort(UBound(tmShoCreditSort)).lDollars = gStrDecToLong(slStr, 2)
                        ReDim Preserve tmShoCreditSort(0 To UBound(tmShoCreditSort) + 1) As SHOCREDITSORT
                    End If
                ElseIf (tmAdvtCredit.sCreditRestr = "W") Or (tmAdvtCredit.sCreditRestr = "M") Or (tmAdvtCredit.sCreditRestr = "T") Then
                    gPDNToStr tmAdvtCredit.sCurrAR, 2, slCurrAR
                    gPDNToStr tmAdvtCredit.sUnbilled, 2, slUnbilled
                    slOwned = gAddStr(slCurrAR, slUnbilled)
                    If gCompNumberStr(slOwned, ".00") > 0 Then
                        tmShoCreditSort(UBound(tmShoCreditSort)).lDollars = gStrDecToLong(slOwned, 2)
                        ReDim Preserve tmShoCreditSort(0 To UBound(tmShoCreditSort) + 1) As SHOCREDITSORT
                    End If
                End If
            End If
            'If tmAdvtCredit.sCrdApp = "R" Then
            '    slStr = Trim$(tmAdvtCredit.sName) & "|Requires Credit Check" & "| "
            '    lbcCredit.AddItem slStr
            '    pbcLbcCredit_Paint
            'Else
            '    'Check Credit
            '    If tmAdvtCredit.sCreditRestr = "L" Then  'Credit Limit
            '        'If A/R + UnBilled > Credit Limit- Show Advertiser
            '        'gPDNToStr tmAdvtCredit.sCreditLimit, 2, slCreditLimit
            '        slCreditLimit = gLongToStrDec(tmAdvtCredit.lCreditLimit, 2)
            '        gPDNToStr tmAdvtCredit.sCurrAR, 2, slCurrAR
            '        gPDNToStr tmAdvtCredit.sUnbilled, 2, slUnbilled
            '        slOwned = gAddStr(slCurrAR, slUnbilled)
            '        If gCompNumberStr(slOwned, slCreditLimit) > 0 Then
            '            slStr = gSubStr(slOwned, slCreditLimit)
            '            gFormatStr slCreditLimit, FMTCOMMA + FMTDOLLARSIGN, 2, slCreditLimit
            '            gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
            '            slStr = Trim$(tmAdvtCredit.sName) & "|Limit: " & slCreditLimit & "|Exceeded: " & slCredit
            '            lbcCredit.AddItem slStr
            '            pbcLbcCredit_Paint
            '        End If
            '    ElseIf (tmAdvtCredit.sCreditRestr = "W") Or (tmAdvtCredit.sCreditRestr = "M") Or (tmAdvtCredit.sCreditRestr = "T") Then
            '        'If A/R + Unbilled > 0- Show Advertiser
            '        gPDNToStr tmAdvtCredit.sCurrAR, 2, slCurrAR
            '        gPDNToStr tmAdvtCredit.sUnbilled, 2, slUnbilled
            '        slOwned = gAddStr(slCurrAR, slUnbilled)
            '        If gCompNumberStr(slOwned, ".00") > 0 Then
            '            gFormatStr slOwned, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
            '            slStr = Trim$(tmAdvtCredit.sName) & "|Cash in Advance" & "|Short: " & slCredit
            '            lbcCredit.AddItem slStr
            '            pbcLbcCredit_Paint
            '        End If
            '    End If
            'End If
            ilExtLen = Len(tmAdvtCredit)
            ilRet = btrExtGetNext(hmAdf, tmAdvtCredit, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmAdf, tmAdvtCredit, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    btrExtClear hmAdf   'Clear any previous extend operation
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAgyPop                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Agency                *
'*                                                     *
'*******************************************************
Private Sub mAgyPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slCredit                                                                              *
'******************************************************************************************

    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slCreditLimit As String
    Dim slCurrAR As String
    Dim slUnbilled As String
    Dim slOwned As String
    Dim slStr As String
    ilExtLen = Len(tmAgyCredit)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAgf) 'Obtain number of records
    btrExtClear hmAgf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmAgf, tmAgf, imAgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'Call btrExtSetBounds(hmAgf, llNoRec, -1, "UC") 'Set extract limits (all records)
    Call btrExtSetBounds(hmAgf, llNoRec, -1, "UC", "AGYCREDITEXTPK", AGYCREDITEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Agf", "AgfCode")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfName")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 40)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCity")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 5)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCreditRestr")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCreditLimit")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 4) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCurrAR")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 6) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfUnbilled")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 6) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCrdApp")
    ilRet = btrExtAddField(hmAgf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hmAgf)    'Extract record
    ilRet = btrExtGetNext(hmAgf, tmAgyCredit, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
        ilExtLen = Len(tmAgyCredit)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlAgf, tgCommAgf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmAgf, tmAgyCredit, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            'If tmAgyCredit.sCrdApp = "R" Then
            '    slStr = Trim$(tmAgyCredit.sName) & "|Requires Credit Check" & "| "
            '    lbcCredit.AddItem slStr
            '    pbcLbcCredit_Paint
            'Else
            '    'Check Credit
            '    If tmAgyCredit.sCreditRestr = "L" Then  'Credit Limit
            '        'If A/R + UnBilled > Credit Limit- Show Advertiser
            '        'gPDNToStr tmAgyCredit.sCreditLimit, 2, slCreditLimit
            '        slCreditLimit = gLongToStrDec(tmAgyCredit.lCreditLimit, 2)
            '        gPDNToStr tmAgyCredit.sCurrAR, 2, slCurrAR
            '        gPDNToStr tmAgyCredit.sUnbilled, 2, slUnbilled
            '        slOwned = gAddStr(slCurrAR, slUnbilled)
            '        If gCompNumberStr(slOwned, slCreditLimit) > 0 Then
            '            slStr = gSubStr(slOwned, slCreditLimit)
            '            gFormatStr slCreditLimit, FMTCOMMA + FMTDOLLARSIGN, 2, slCreditLimit
            '            gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
            '            slStr = Trim$(tmAgyCredit.sName) & ", " & Trim$(tmAgyCredit.sCityID) & "|Limit: " & slCreditLimit & "|Exceeded: " & slCredit
            '            lbcCredit.AddItem slStr
            '            pbcLbcCredit_Paint
            '        End If
            '    ElseIf (tmAgyCredit.sCreditRestr = "W") Or (tmAgyCredit.sCreditRestr = "M") Or (tmAgyCredit.sCreditRestr = "T") Then
            '        'If A/R + Unbilled > 0- Show Advertiser
            '        gPDNToStr tmAgyCredit.sCurrAR, 2, slCurrAR
            '        gPDNToStr tmAgyCredit.sUnbilled, 2, slUnbilled
            '        slOwned = gAddStr(slCurrAR, slUnbilled)
            '        If gCompNumberStr(slOwned, ".00") > 0 Then
            '            gFormatStr slOwned, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
            '            slStr = Trim$(tmAgyCredit.sName) & ", " & Trim$(tmAgyCredit.sCityID) & "|Cash in Advance" & "|Short: " & slCredit
            '            lbcCredit.AddItem slStr
            '            pbcLbcCredit_Paint
            '        End If
            '    End If
            'End If
            tmShoCreditSort(UBound(tmShoCreditSort)).iType = 1
            tmShoCreditSort(UBound(tmShoCreditSort)).sKey = ""
            tmShoCreditSort(UBound(tmShoCreditSort)).sName = tmAgyCredit.sName
            tmShoCreditSort(UBound(tmShoCreditSort)).sCityID = tmAgyCredit.sCityID
            tmShoCreditSort(UBound(tmShoCreditSort)).sAddrID = ""
            tmShoCreditSort(UBound(tmShoCreditSort)).sCreditRestr = tmAgyCredit.sCreditRestr
            tmShoCreditSort(UBound(tmShoCreditSort)).lCreditLimit = tmAgyCredit.lCreditLimit
            tmShoCreditSort(UBound(tmShoCreditSort)).lDollars = 0
            tmShoCreditSort(UBound(tmShoCreditSort)).sCrdApp = tmAgyCredit.sCrdApp
            tmShoCreditSort(UBound(tmShoCreditSort)).sBillAgyDir = ""
            tmShoCreditSort(UBound(tmShoCreditSort)).iCode = tmAgyCredit.iCode
            If (tmAgyCredit.sCrdApp = "R") Then
                ReDim Preserve tmShoCreditSort(0 To UBound(tmShoCreditSort) + 1) As SHOCREDITSORT
            Else
                If tmAgyCredit.sCreditRestr = "L" Then  'Credit Limit
                    slCreditLimit = gLongToStrDec(tmAgyCredit.lCreditLimit, 2)
                    gPDNToStr tmAgyCredit.sCurrAR, 2, slCurrAR
                    gPDNToStr tmAgyCredit.sUnbilled, 2, slUnbilled
                    slOwned = gAddStr(slCurrAR, slUnbilled)
                    If gCompNumberStr(slOwned, slCreditLimit) > 0 Then
                        slStr = gSubStr(slOwned, slCreditLimit)
                        tmShoCreditSort(UBound(tmShoCreditSort)).lDollars = gStrDecToLong(slStr, 2)
                        ReDim Preserve tmShoCreditSort(0 To UBound(tmShoCreditSort) + 1) As SHOCREDITSORT
                    End If
                ElseIf (tmAgyCredit.sCreditRestr = "W") Or (tmAgyCredit.sCreditRestr = "M") Or (tmAgyCredit.sCreditRestr = "T") Then
                    gPDNToStr tmAgyCredit.sCurrAR, 2, slCurrAR
                    gPDNToStr tmAgyCredit.sUnbilled, 2, slUnbilled
                    slOwned = gAddStr(slCurrAR, slUnbilled)
                    If gCompNumberStr(slOwned, ".00") > 0 Then
                        tmShoCreditSort(UBound(tmShoCreditSort)).lDollars = gStrDecToLong(slOwned, 2)
                        ReDim Preserve tmShoCreditSort(0 To UBound(tmShoCreditSort) + 1) As SHOCREDITSORT
                    End If
                End If
            End If
            ilExtLen = Len(tmAgyCredit)
            ilRet = btrExtGetNext(hmAgf, tmAgyCredit, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmAgf, tmAgyCredit, ilExtLen, llRecPos)
            Loop
        Loop
    End If
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slLimit                       slCredit                  *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slSpfDate As String

    imTerminate = False

    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    'If test here- then must remove DDE communication to traffic
    'If igWinStatus(COLLECTIONSJOB) <> 2 Then
    '    imTerminate = True
    '    mTerminate
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
'    mInitDDE

    imListFieldCredit(1) = 15
    imListFieldCredit(2) = 50 * igAlignCharWidth
    imListFieldCredit(3) = 77 * igAlignCharWidth
    imListFieldCredit(4) = 100 * igAlignCharWidth
    lacAdvt.Width = imListFieldCredit(2) - imListFieldCredit(1) + 15
    lacAdvt.Left = lbcCredit.Left
    lacAdvt.Top = lbcCredit.Top - lacAdvt.Height + 15
    lacCreditStatus.Width = imListFieldCredit(3) - imListFieldCredit(2) + 15
    lacCreditStatus.Left = lacAdvt.Left + lacAdvt.Width - 15
    lacCreditStatus.Top = lacAdvt.Top
    lacDollars.Width = lbcCredit.Width - lacAdvt.Width - lacCreditStatus.Width + 15
    lacDollars.Left = lbcCredit.Left + lbcCredit.Width - lacDollars.Width
    lacDollars.Top = lacAdvt.Top
    lacCreditStatus.Width = lacCreditStatus.Width + 15
    igShowCredit = True
    imFirstFocus = True
    hmAgf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", ShoCrdit
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", ShoCrdit
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    ReDim tmShoCreditSort(0 To 0) As SHOCREDITSORT
    tmAgf.iCode = 0
    tmAdf.iCode = 0
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    mAgyPop
    If imTerminate Then
        Exit Sub
    End If
    imLastSortCol = -1
    mSort 1
    If lbcCredit.ListCount <= 0 Then
        ''mTerminate
        'imTerminate = True
        'Exit Sub
        lbcCredit.AddItem " No Credit Alerts"
    End If
    'Don't show unless something to show
    gUnpackDate tgSpf.iRCreditDate(0), tgSpf.iRCreditDate(1), slSpfDate
    lacScreen.Caption = "Credit Alerts, Last date Set Credit Ran: " & slSpfDate
    ShoCrdit.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterStdAlone ShoCrdit
    gCenterForm ShoCrdit
    'ShoCrdit.Show
'    gCenterModalForm ShoCrdit
    Screen.MousePointer = vbDefault
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
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    'slCommand = Command$
    slCommand = sgCommandStr  'Command$
'    If StrComp(slCommand, "Debug", 1) = 0 Then
'        igStdAloneMode = True 'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = True 'False
'        imShowHelpMsg = False
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        slStr = "ShoCrdit"
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
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
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        igCmmCallSource = Val(slStr)
    'End If
    'gInitStdAlone ShoCrdit, slStr, ilTestSystem
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

    igCollCallFrom = 0
    sgCollCallName = ""
    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    igShowCredit = False
    'Unload Traffic
    Unload ShoCrdit
    igManUnload = NO
End Sub

Private Sub lacAdvt_Click()
    mSort 1
End Sub

Private Sub lacCreditStatus_Click()
    mSort 2
End Sub

Private Sub lacDollars_Click()
    mSort 3
End Sub

Private Sub lbcCredit_Click()
    pbcLbcCredit_Paint
    mSetCommands
End Sub

Private Sub lbcCredit_DblClick()
    mBranch
End Sub

Private Sub lbcCredit_Scroll()
    pbcLbcCredit_Paint
End Sub


Private Sub pbcClickFocus_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
        'mInitDDE
        If igUpdateAllowed = False Then
            'mSendHelpMsg "BF"
        Else
            'mSendHelpMsg "BT"
        End If
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcCredit_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilCreditEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 2) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilFieldIndex As Integer

    ilCreditEnd = lbcCredit.TopIndex + lbcCredit.Height \ fgListHtArial825
    If ilCreditEnd > lbcCredit.ListCount Then
        ilCreditEnd = lbcCredit.ListCount
    End If
    If lbcCredit.ListCount <= lbcCredit.Height \ fgListHtArial825 Then
        llWidth = lbcCredit.Width - 30
    Else
        llWidth = lbcCredit.Width - igScrollBarWidth - 30
    End If
    pbcLbcCredit.Width = llWidth
    pbcLbcCredit.Cls
    llFgColor = pbcLbcCredit.ForeColor
    For ilLoop = lbcCredit.TopIndex To ilCreditEnd - 1 Step 1
        pbcLbcCredit.ForeColor = llFgColor
        If lbcCredit.MultiSelect = 0 Then
            If lbcCredit.ListIndex = ilLoop Then
                gPaintArea pbcLbcCredit, CSng(0), CSng((ilLoop - lbcCredit.TopIndex) * fgListHtArial825), CSng(pbcLbcCredit.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcCredit.ForeColor = vbWhite
            End If
        Else
            If lbcCredit.Selected(ilLoop) Then
                gPaintArea pbcLbcCredit, CSng(0), CSng((ilLoop - lbcCredit.TopIndex) * fgListHtArial825), CSng(pbcLbcCredit.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcCredit.ForeColor = vbWhite
            End If
        End If
        slStr = lbcCredit.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = 1 To 3 Step 1
            pbcLbcCredit.CurrentX = imListFieldCredit(ilField)
            pbcLbcCredit.CurrentY = (ilLoop - lbcCredit.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcCredit, slStr, imListFieldCredit(ilField + 1) - imListFieldCredit(ilField)
            pbcLbcCredit.Print slStr
        Next ilField
        pbcLbcCredit.ForeColor = llFgColor
    Next ilLoop

End Sub


Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    If Not igJobShowing(COLLECTIONSJOB) Then
        Me.ZOrder 0
    Else
        DoEvents
        tmcStart.Enabled = True
    End If
End Sub

Private Sub mSort(ilSortCol As Integer)
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slLimit As String
    Dim slCredit As String

    lbcCredit.Clear
    pbcLbcCredit.Cls
    If imLastSortCol <> ilSortCol Then
        For ilLoop = 0 To UBound(tmShoCreditSort) - 1 Step 1
            If ilSortCol = 1 Then
                If tmShoCreditSort(ilLoop).iType = 0 Then
                    If Trim$(tmShoCreditSort(ilLoop).sBillAgyDir) = "D" Then
                        If Trim$(tmShoCreditSort(ilLoop).sAddrID) <> "" Then
                            tmShoCreditSort(ilLoop).sKey = tmShoCreditSort(ilLoop).sName & ", " & tmShoCreditSort(ilLoop).sAddrID & "/Direct"
                        Else
                            tmShoCreditSort(ilLoop).sKey = tmShoCreditSort(ilLoop).sName & "/Direct"
                        End If
                    Else
                        tmShoCreditSort(ilLoop).sKey = tmShoCreditSort(ilLoop).sName
                    End If
                Else
                    tmShoCreditSort(ilLoop).sKey = tmShoCreditSort(ilLoop).sName & tmShoCreditSort(ilLoop).sCityID
                End If
            ElseIf ilSortCol = 2 Then
                If tmShoCreditSort(ilLoop).sCrdApp = "R" Then
                    slStr = "A-Requires Credit Check"
                Else
                    If tmShoCreditSort(ilLoop).sCreditRestr = "L" Then
                        slLimit = Trim$(str$(tmShoCreditSort(ilLoop).lCreditLimit))
                        Do While Len(slLimit) < 10
                            slLimit = "0" & slLimit
                        Loop
                        slStr = "B-" & slLimit
                    ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "W") Then
                        slStr = "D-Cash in Advance"
                    ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "M") Then
                        slStr = "E-Cash in Advance"
                    ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "T") Then
                        slStr = "F-Cash in Advance"
                    ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "P") Then
                        slStr = "C-Prohibit New Orders"
                    End If
                End If
                tmShoCreditSort(ilLoop).sKey = slStr & tmShoCreditSort(ilLoop).sName
            ElseIf ilSortCol = 3 Then
                If tmShoCreditSort(ilLoop).sCrdApp = "R" Then
                    slStr = "0"
                    Do While Len(slStr) < 10
                        slStr = "0" & slStr
                    Loop
                Else
                    If tmShoCreditSort(ilLoop).sCreditRestr = "L" Then
                        slStr = Trim$(str$(tmShoCreditSort(ilLoop).lDollars))
                        Do While Len(slStr) < 10
                            slStr = "0" & slStr
                        Loop
                        'slStr = slStr
                    ElseIf tmShoCreditSort(ilLoop).sCreditRestr = "P" Then
                        slStr = Trim$(str$(tmShoCreditSort(ilLoop).lDollars))
                        Do While Len(slStr) < 10
                            slStr = "0" & slStr
                        Loop
                        'slStr = slStr
                    ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "W") Or (tmShoCreditSort(ilLoop).sCreditRestr = "M") Or (tmShoCreditSort(ilLoop).sCreditRestr = "T") Then
                        slStr = Trim$(str$(tmShoCreditSort(ilLoop).lDollars))
                        Do While Len(slStr) < 10
                            slStr = "0" & slStr
                        Loop
                        'If (tmShoCreditSort(ilLoop).sCreditRestr = "W") Then
                        '    slStr = slStr
                        'ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "M") Then
                        '    slStr = slStr
                        'Else
                        '    slStr = slStr
                        'End If
                    End If
                End If
                tmShoCreditSort(ilLoop).sKey = slStr & tmShoCreditSort(ilLoop).sName
            End If
        Next ilLoop
        imSortDirection = 0
        If UBound(tmShoCreditSort) - 1 > 0 Then
            ArraySortTyp fnAV(tmShoCreditSort(), 0), UBound(tmShoCreditSort), imSortDirection, LenB(tmShoCreditSort(0)), 0, LenB(tmShoCreditSort(0).sKey), 0
        End If
    Else
        If UBound(tmShoCreditSort) - 1 > 0 Then
            If imSortDirection = 0 Then
                imSortDirection = 1
            Else
                imSortDirection = 0
            End If
            ArraySortTyp fnAV(tmShoCreditSort(), 0), UBound(tmShoCreditSort), imSortDirection, LenB(tmShoCreditSort(0)), 0, LenB(tmShoCreditSort(0).sKey), 0
        End If
    End If
    For ilLoop = 0 To UBound(tmShoCreditSort) - 1 Step 1
        If tmShoCreditSort(ilLoop).iType = 0 Then
            If Trim$(tmShoCreditSort(ilLoop).sBillAgyDir) = "D" Then
                If Trim$(tmShoCreditSort(ilLoop).sAddrID) <> "" Then
                    slStr = Trim$(tmShoCreditSort(ilLoop).sName) & ", " & Trim$(tmShoCreditSort(ilLoop).sAddrID) & "/Direct"
                Else
                    slStr = Trim$(tmShoCreditSort(ilLoop).sName) & "/Direct"
                End If
            Else
                slStr = Trim$(tmShoCreditSort(ilLoop).sName)
            End If
        Else
            slStr = Trim$(tmShoCreditSort(ilLoop).sName) & ", " & Trim$(tmShoCreditSort(ilLoop).sCityID)
        End If
        If tmShoCreditSort(ilLoop).sCrdApp = "R" Then
            slStr = slStr & "|Requires Credit Check" & "| "
        Else
            If tmShoCreditSort(ilLoop).sCreditRestr = "L" Then
                slLimit = gLongToStrDec(tmShoCreditSort(ilLoop).lCreditLimit, 2)
                gFormatStr slLimit, FMTCOMMA + FMTDOLLARSIGN, 2, slLimit
                slCredit = gLongToStrDec(tmShoCreditSort(ilLoop).lDollars, 2)
                gFormatStr slCredit, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
                slStr = slStr & "|Limit: " & slLimit & "|Exceeded: " & slCredit
            ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "P") Then
                slCredit = gLongToStrDec(tmShoCreditSort(ilLoop).lDollars, 2)
                gFormatStr slCredit, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
                slStr = slStr & "|Prohibit New Orders" & "|" & slCredit
            ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "W") Or (tmShoCreditSort(ilLoop).sCreditRestr = "M") Or (tmShoCreditSort(ilLoop).sCreditRestr = "T") Then
                slCredit = gLongToStrDec(tmShoCreditSort(ilLoop).lDollars, 2)
                gFormatStr slCredit, FMTCOMMA + FMTDOLLARSIGN, 2, slCredit
                If (tmShoCreditSort(ilLoop).sCreditRestr = "W") Then
                    slStr = slStr & "|Cash in Advance-Weekly" & "|Short: " & slCredit
                ElseIf (tmShoCreditSort(ilLoop).sCreditRestr = "M") Then
                    slStr = slStr & "|Cash in Advance-Monthly" & "|Short: " & slCredit
                Else
                    slStr = slStr & "|Cash in Advance-Contract" & "|Short: " & slCredit
                End If
            End If
        End If
        lbcCredit.AddItem slStr
    Next ilLoop
    pbcLbcCredit_Paint
    imLastSortCol = ilSortCol

End Sub

Private Sub mSetCommands()
    Dim ilRow As Integer

    cmcCheck.Enabled = False
    ilRow = lbcCredit.ListIndex
    If (ilRow >= 0) And (ilRow <= UBound(tmShoCreditSort)) Then
        If (tmShoCreditSort(ilRow).sCreditRestr <> "P") Then
            cmcCheck.Enabled = True
        End If
    End If
End Sub

Private Sub mBranch()
    Dim ilRow As Integer
    Dim slName As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilAgf As Integer

    ilRow = lbcCredit.ListIndex
    If (ilRow >= 0) And (ilRow <= UBound(tmShoCreditSort)) Then
        slName = lbcCredit.List(ilRow)
        ilRet = gParseItem(slName, 1, "|", slName)    'Get application name
        If tmShoCreditSort(ilRow).iType = 0 Then
            If (tmShoCreditSort(ilRow).sCrdApp = "R") Or (tmShoCreditSort(ilRow).sCreditRestr = "L") Then
                '3/13/08:  Force repopulation of advertiser names
                sgAdvertiserTag = ""
                igAdvtCallSource = CALLSOURCEALERTS
                If igTestSystem Then
                    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource)) & "\" & slName
                Else
                    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource)) & "\" & slName
                End If
                sgCommandStr = slStr
                Advt.Show vbModal
            Else
                If igWinStatus(COLLECTIONSJOB) = 0 Then
                    Exit Sub
                End If
                If tmShoCreditSort(ilRow).sBillAgyDir <> "D" Then
                    tmAdfSrchKey.iCode = tmShoCreditSort(ilRow).iCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilAgf = gBinarySearchAgf(tmAdf.iAgfCode)
                        If ilAgf <> -1 Then
                            slName = Trim$(tgCommAgf(ilAgf).sName) & ", " & Trim$(tgCommAgf(ilAgf).sCityID)
                        End If
                    End If
                Else
                    slName = Trim$(tmShoCreditSort(ilRow).sName)
                End If
                igCollCallFrom = 1
                sgCollCallName = slName
                On Error Resume Next
                If Not igJobShowing(COLLECTIONSJOB) Then
                    Collect.Show vbModeless  'This is required to activate form if previousely activated
                Else
                    Collect.SetFocus
                End If
                tmcStart.Enabled = True
            End If
        Else
            If (tmShoCreditSort(ilRow).sCrdApp = "R") Or (tmShoCreditSort(ilRow).sCreditRestr = "L") Then
                '3/13/08:  Force repopulation of agency names
                sgAgencyTag = ""
                igAgyCallSource = CALLSOURCEALERTS
                If igTestSystem Then
                    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource)) & "\" & slName
                Else
                    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(igAdvtCallSource)) & "\" & slName
                End If
                sgCommandStr = slStr
                Agency.Show vbModal
            Else
                If igWinStatus(COLLECTIONSJOB) = 0 Then
                    Exit Sub
                End If
                igCollCallFrom = 1
                sgCollCallName = slName
                On Error Resume Next
                If Not igJobShowing(COLLECTIONSJOB) Then
                    Collect.Show vbModeless  'This is required to activate form if previousely activated
                Else
                    Collect.SetFocus
                End If
                tmcStart.Enabled = True
            End If
        End If
    End If
End Sub
