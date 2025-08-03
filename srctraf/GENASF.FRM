VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form GenASF 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate ASF"
   ClientHeight    =   2040
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
   Icon            =   "GENASF.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2040
   ScaleWidth      =   7095
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1005
      Top             =   1500
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Left            =   2085
      TabIndex        =   8
      Top             =   810
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6345
      Top             =   165
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
      Left            =   75
      ScaleHeight     =   270
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   30
      Width           =   2910
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6015
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
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
      TabIndex        =   2
      Top             =   1140
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
      TabIndex        =   3
      Top             =   1140
      Width           =   1050
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   1
      Top             =   1620
      Width           =   5820
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   900
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1875
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "GenASF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: GenASF.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
'Avail Summary File
Dim hmAsf As Integer    'Avail Summary file handle
Dim tmAsf As ASF
Dim imAsfReclen As Integer        'Asf record length
'Spot file
Dim hmSdf As Integer
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey0 As SDFKEY0 'SDF key record image (agency)
Dim tmSdfSrchKey1 As SDFKEY1 'SDF key record image (advt)
Dim tmSdfSrchKey2 As SDFKEY2 'SDF key record image (advt)
Dim tmSdfSrchKey3 As LONGKEY0 'SDF key record image (code)
Dim imSdfRecLen As Integer     'SDF record length
'Gather Info file (Used to communicate to stand alone program that contract should be sent)
Dim tmGif As GIF            'GIF record image
Dim tmGifSrchKey0 As LONGKEY0 'GIF key record image (agency)
Dim hmGif As Integer        'GIF Handle
Dim imGifRecLen As Integer     'GIF record length
'Library
Dim hmLcf As Integer 'Library file handle
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim tmLLC() As LLC
'Contract record information
Dim hmClf As Integer        'Contract line file handle
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Rate card
Dim tmRcf As RCF
Dim tmAvailInfo() As AVAILINFO
'Dim imTimeoutInterval As Integer    '0=No timeout interval- ask questions; > 0=Don't ask question, just run at each timeout interval
'Dim imMinElapsed As Integer
Dim imUserInput As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imGeneratingASF As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmStartDate As Long
Dim lmEndDate As Long
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
'Spot summary
Dim hmSsf As Integer        'Spot summary file handle
Dim tmSsf As SSF               'SSF record image
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer  'SSF record length
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmBBSpot As BBSPOTSS
Dim tmProgTest As PROGRAMSS
Dim tmAvailTest As AVAILSS
Dim tmSpotTest As CSPOTSS
Dim tmBBSpotTest As BBSPOTSS
Private Sub cmcCancel_Click()
    If imGeneratingASF Then
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
    lacInfo.Visible = False
    If imGeneratingASF Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmcCancel.Caption = "&Cancel"
    imGeneratingASF = True
    lacInfo.Visible = True
    lacInfo.Caption = ""
    lmTotalNoRecs = 0
    '2-02-01
    hmAsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmAsf, "", sgDBPath & "Asf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", GenASF
    On Error GoTo 0
    imAsfReclen = Len(tmAsf)
    
    hmGif = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmGif, "", sgDBPath & "Gif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", GenASF
    On Error GoTo 0
    imGifRecLen = Len(tmGif)
    
    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen: Sdf.Btr)", GenASF
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf) 'btrRecordLength(hmSdf)    'Get Cff size
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen: Clf.Btr)", GenASF
    On Error GoTo 0
    imClfRecLen = Len(tmClf) 'btrRecordLength(hmClf)    'Get Clf size
    
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "SSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen: SSF.Btr)", GenASF
    On Error GoTo 0
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "LCF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen: LCF.Btr)", GenASF
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf) 'btrRecordLength(hmLcf)    'Get Cff size
    ilRet = gObtainLatestRcf(tmRcf)
    If Not ilRet Then
        imTerminate = True
        Exit Sub
    End If
    gRCRead GenASF, tmRcf.iCode, tmRcf, tgMRif(), tgMRdf()
    ilRet = gObtainVef()
    If Not ilRet Then
        imTerminate = True
        Exit Sub
    End If
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S")) And (tgMVef(ilVef).sState = "A") Then
            lmTotalNoRecs = lmTotalNoRecs + 1
        End If
    Next ilVef
    '4/12/18: required to load tgSaf(0) which is required for testing options
    gObtainSAF
    '5/7/18
    ilRet = gObtainAvail()         'tgavailanf

    If tgSpf.sCIncludeMissDB <> "N" Then
        lmTotalNoRecs = (lmEndDate - lmStartDate + 1) * lmTotalNoRecs + 3 * lmTotalNoRecs
    Else
        lmTotalNoRecs = (lmEndDate - lmStartDate + 1) * lmTotalNoRecs
    End If
    lmProcessedNoRecs = 0
    plcGauge.Visible = True
    plcGauge.Value = 0
    
    
    '1-29-01 remove everything that exists
    ilRet = btrGetFirst(hmAsf, tmAsf, imAsfReclen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        ilRet = btrDelete(hmAsf)
        On Error GoTo cmcExportErr
        gBtrvErrorMsg ilRet, "cmcGen (btrDelete: Asf.Btr)", GenASF
        On Error GoTo 0
        ilRet = btrGetFirst(hmAsf, tmAsf, imAsfReclen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        On Error GoTo cmcExportErr
        gBtrvErrorMsg ilRet, "cmcGen (btrGetFirst: Asf.Btr)", GenASF
        On Error GoTo 0
    Loop
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S")) And (tgMVef(ilVef).sState = "A") Then
            lacInfo.Caption = "Processing: " & Trim$(tgMVef(ilVef).sName)
            DoEvents
            ReDim tmAvailInfo(0 To 0) As AVAILINFO
            ilVefCode = tgMVef(ilVef).iCode
            'Determine the number of dayparts to generate avails for
            For ilRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                If tgMRif(ilRif).iVefCode = ilVefCode Then
                    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        If tgMRdf(ilRdf).iCode = tgMRif(ilRif).iRdfCode Then
                            'If tgMRdf(ilRdf).sBase <> "N" Then
                            '1/22/09: Jim requested that the avail summary only be performed for dayparts that have report set to yes
                            If tgMRif(ilRif).sRpt = "Y" Then
                                tmAvailInfo(UBound(tmAvailInfo)).iRdfIndex = ilRdf
                                For ilLoop = 1 To 26
                                    tmAvailInfo(UBound(tmAvailInfo)).iAvail60(ilLoop) = -29999
                                    tmAvailInfo(UBound(tmAvailInfo)).iAvail30(ilLoop) = -29999
                                    tmAvailInfo(UBound(tmAvailInfo)).iAvail10(ilLoop) = -29999
                                Next ilLoop
                                ReDim Preserve tmAvailInfo(0 To UBound(tmAvailInfo) + 1) As AVAILINFO
                            End If
                            'End If
                            Exit For
                        End If
                    Next ilRdf
                End If
            Next ilRif
            If UBound(tmAvailInfo) > LBound(tmAvailInfo) Then
                mAvailCount ilVefCode, lmStartDate, lmEndDate
                If Not imTerminate Then
                    'Delete ASF records for the vehicle
                    '1-29-01 removed everything above, dont do each vehicle
                    'tmRUISrchKey.iRemoteID = tgMVef(ilVef).iRemoteID
                    'tmRUISrchKey.iAutoCode = tgMVef(ilVef).iAutoCode
                    'ilRet = btrGetEqual(hmAsf, tmAsf, imAsfReclen, tmRUISrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    'Do While ilRet = BTRV_ERR_NONE
                    '    DoEvents
                    '    ilRet = btrDelete(hmAsf)
                    '    tmRUISrchKey.iRemoteID = tgMVef(ilVef).iRemoteID
                    '    tmRUISrchKey.iAutoCode = tgMVef(ilVef).iAutoCode
                    '    ilRet = btrGetEqual(hmAsf, tmAsf, imAsfReclen, tmRUISrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    'Loop
                    'Update ASF
                    For ilLoop = LBound(tmAvailInfo) To UBound(tmAvailInfo) - 1 Step 1
                        DoEvents
                        gPackDateLong lmStartDate, tmAsf.iStartWkDate(0), tmAsf.iStartWkDate(1)
                        'tmAsf.iRdfRemoteID = tgMRdf(tmAvailInfo(ilLoop).iRdfIndex).iRemoteID
                        tmAsf.iUnused1 = 0
                        tmAsf.iRdfCode = tgMRdf(tmAvailInfo(ilLoop).iRdfIndex).iCode
                        'tmAsf.iVefRemoteID = tgMVef(ilVef).iRemoteID
                        tmAsf.iUnused2 = 0
                        tmAsf.iVefCode = tgMVef(ilVef).iCode
                        tmAsf.iRcfCode = tmRcf.iCode                '1-24-01
                        tmAsf.lCode = 0                             '1-24-01 auto increment
                        For ilWk = 1 To 26 Step 1
                            ilWkMinus1 = ilWk - 1
                            If tmAvailInfo(ilLoop).iAvail60(ilWk) = -29999 Then
                                tmAsf.sAvail60(ilWkMinus1) = Chr$(255)
                            ElseIf tmAvailInfo(ilLoop).iAvail60(ilWk) < 0 Then
                                tmAvailInfo(ilLoop).iAvail60(ilWk) = -1 * tmAvailInfo(ilLoop).iAvail60(ilWk)
                                If tmAvailInfo(ilLoop).iAvail60(ilWk) >= 50 Then
                                    tmAsf.sAvail60(ilWkMinus1) = Chr$(250)
                                Else
                                    tmAsf.sAvail60(ilWkMinus1) = Chr$(200 + tmAvailInfo(ilLoop).iAvail60(ilWk))
                                End If
                            Else
                                If tmAvailInfo(ilLoop).iAvail60(ilWk) >= 200 Then
                                    tmAsf.sAvail60(ilWkMinus1) = Chr$(200)
                                Else
                                    tmAsf.sAvail60(ilWkMinus1) = Chr$(tmAvailInfo(ilLoop).iAvail60(ilWk))
                                End If
                            End If
                            If tmAvailInfo(ilLoop).iAvail30(ilWk) = -29999 Then
                                tmAsf.sAvail30(ilWkMinus1) = Chr$(255)
                            ElseIf tmAvailInfo(ilLoop).iAvail30(ilWk) < 0 Then
                                tmAvailInfo(ilLoop).iAvail30(ilWk) = -1 * tmAvailInfo(ilLoop).iAvail30(ilWk)
                                If tmAvailInfo(ilLoop).iAvail30(ilWk) >= 50 Then
                                    tmAsf.sAvail30(ilWkMinus1) = Chr$(250)
                                Else
                                    tmAsf.sAvail30(ilWkMinus1) = Chr$(200 + tmAvailInfo(ilLoop).iAvail30(ilWk))
                                End If
                            Else
                                If tmAvailInfo(ilLoop).iAvail30(ilWk) >= 200 Then
                                    tmAsf.sAvail30(ilWkMinus1) = Chr$(200)
                                Else
                                    tmAsf.sAvail30(ilWkMinus1) = Chr$(tmAvailInfo(ilLoop).iAvail30(ilWk))
                                End If
                            End If
                            If tmAvailInfo(ilLoop).iAvail10(ilWk) = -29999 Then
                                tmAsf.sAvail10(ilWkMinus1) = Chr$(255)
                            ElseIf tmAvailInfo(ilLoop).iAvail10(ilWk) < 0 Then
                                tmAvailInfo(ilLoop).iAvail10(ilWk) = -1 * tmAvailInfo(ilLoop).iAvail10(ilWk)
                                If tmAvailInfo(ilLoop).iAvail10(ilWk) >= 50 Then
                                    tmAsf.sAvail10(ilWkMinus1) = Chr$(250)
                                Else
                                    tmAsf.sAvail10(ilWkMinus1) = Chr$(200 + tmAvailInfo(ilLoop).iAvail10(ilWk))
                                End If
                            Else
                                If tmAvailInfo(ilLoop).iAvail10(ilWk) >= 200 Then
                                    tmAsf.sAvail10(ilWkMinus1) = Chr$(200)
                                Else
                                    tmAsf.sAvail10(ilWkMinus1) = Chr$(tmAvailInfo(ilLoop).iAvail10(ilWk))
                                End If
                            End If
                        Next ilWk
                        gPackDate Format(gNow(), "m/d/yy"), tmAsf.iCreatedDate(0), tmAsf.iCreatedDate(1)
                        gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tmAsf.iCreatedTime(0), tmAsf.iCreatedTime(1)
                        tmAsf.sUnused = ""
                        ilRet = btrInsert(hmAsf, tmAsf, imAsfReclen, INDEXKEY0)
                        
                    Next ilLoop
                    ilRet = mCreateGIF()
                End If
            Else
                If tgSpf.sCIncludeMissDB <> "N" Then
                    lmProcessedNoRecs = lmProcessedNoRecs + 182 + 3
                Else
                    lmProcessedNoRecs = lmProcessedNoRecs + 182
                End If
                'lacInfo.Caption = Trim$(Str$(lmProcessedNoRecs)) & " of" & Str$(lmTotalNoRecs)
                llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
                If llPercent >= 100 Then
                    llPercent = 100
                End If
                plcGauge.Value = llPercent
            End If
            If imTerminate Then
                lacInfo.Caption = "Processing Terminated"
                Exit For
            End If
        End If
    Next ilVef
    lacInfo.Visible = True
    imGeneratingASF = False
    cmcCancel.Caption = "&Done"
    '9/27/18: Getting Invalid procedure call
    'cmcCancel.SetFocus
    cmcGen.Enabled = False
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmAsf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmGif)
    btrDestroy hmGif
    btrDestroy hmSdf
    btrDestroy hmAsf
    btrDestroy hmClf
    btrDestroy hmSsf
    btrDestroy hmLcf
    Screen.MousePointer = vbDefault
    If imTerminate Then
        'tmcTimer.Enabled = False
        cmcCancel_Click
    End If
    If Not imUserInput Then
        gUpdateTaskMonitor 2, "ASG"
        cmcCancel_Click
    End If
    Exit Sub
cmcExportErr:
    imTerminate = True
    ilRet = Err
    Resume Next
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    If Not imUserInput Then
        tmcSetTime.Enabled = False
        gCloseTmf
    End If
    
    Erase tmAvailInfo
    Erase tmLLC
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmGif)
    btrDestroy hmGif
    ilRet = btrClose(hmAsf)
    btrDestroy hmAsf
    btrStopAppl
    
    Set GenASF = Nothing   'Remove data segment

    End
End Sub
Private Sub imcHelp_Click()
    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailCount                     *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine avail count          *
'*                                                     *
'*******************************************************
Private Sub mAvailCount(ilVefCode As Integer, llStartDate As Long, llEndDate As Long)
'
'
'
    
    Dim ilVpfIndex As Integer
    Dim llLoopDate As Long
    Dim ilType As Integer
    Dim slType As String
    Dim ilWkIndex As Integer
    Dim slDate As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llLatestDate As Long
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilNo10 As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilPass As Integer
    Dim ilWeekDay As Integer
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim llTime As Long
    Dim ilRdf As Integer
    Dim ilRdfIndex As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDay As Integer
    Dim ilSpot As Integer
    Dim ilSpotOK As Integer
    Dim ilAI As Integer
    Dim ilAvail60 As Integer
    Dim llPercent As Long
    Dim ilVefIndex As Integer
    Dim slCntrType As String
    Dim ilPctTrade As Integer
    Dim ilAnf As Integer

    ReDim ilEvtType(0 To 14) As Integer
    imSdfRecLen = Len(tmSdf)
    ilVpfIndex = gVpfFind(GenASF, ilVefCode)
    If ilVpfIndex < 0 Then
        Exit Sub
    End If
    ilType = 0
    slType = "O"
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True
    
    ilVefIndex = gBinarySearchVef(ilVefCode)
    For llLoopDate = llStartDate To llEndDate Step 1
        DoEvents
        If imTerminate Then
            Exit Sub
        End If
        slDate = Format$(llLoopDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        ilWkIndex = (llLoopDate - llStartDate) \ 7 + 1
        gPackDate slDate, ilDate0, ilDate1
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        If tgMVef(ilVefIndex).sType <> "G" Then
            tmSsfSrchKey.iType = ilType
            tmSsfSrchKey.iVefCode = ilVefCode
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Else
            tmSsfSrchKey2.iVefCode = ilVefCode
            tmSsfSrchKey2.iDate(0) = ilDate0
            tmSsfSrchKey2.iDate(1) = ilDate1
            ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            ilType = tmSsf.iType
        End If

        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType And tgMVef(ilVefIndex).sType <> "G") Or (tmSsf.iVefCode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
            If (llLoopDate > llLatestDate) Then
                ReDim tmLLC(0 To 0) As LLC  'Merged library names
                ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                If (tgSsf(ilWeekDay - 1).iVefCode = ilVefCode) Then
                    tmSsf = tgSsf(ilWeekDay - 1)
                    tmSsf.iDate(0) = ilDate0
                    tmSsf.iDate(1) = ilDate1
                Else
                    If ilWeekDay = 1 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tmLLC())
                    ElseIf ilWeekDay = 2 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tmLLC())
                    ElseIf ilWeekDay = 3 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tmLLC())
                    ElseIf ilWeekDay = 4 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tmLLC())
                    ElseIf ilWeekDay = 5 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tmLLC())
                    ElseIf ilWeekDay = 6 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tmLLC())
                    ElseIf ilWeekDay = 7 Then
                            ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tmLLC())
                    End If
                    DoEvents
                    If imTerminate Then
                        Exit Sub
                    End If
                    tmSsf.iType = 0
                    tmSsf.iVefCode = ilVefCode
                    tmSsf.iDate(0) = ilDate0
                    tmSsf.iDate(1) = ilDate1
                    gPackTime tmLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                    tmSsf.iCount = 0
                    'tmSsf.iNextTime(0) = 1  'Time not defined
                    'tmSsf.iNextTime(1) = 0
                
                    For ilIndex = LBound(tmLLC) To UBound(tmLLC) - 1 Step 1
                        DoEvents
                        If imTerminate Then
                            Exit Sub
                        End If
                        tmAvail.iRecType = Val(tmLLC(ilIndex).sType)
                        gPackTime tmLLC(ilIndex).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                        tmAvail.iLtfCode = tmLLC(ilIndex).iLtfCode
                        tmAvail.iAvInfo = tmLLC(ilIndex).iAvailInfo Or tmLLC(ilIndex).iUnits
                        tmAvail.iLen = CInt(gLengthToCurrency(tmLLC(ilIndex).sLength))
                        tmAvail.ianfCode = Val(tmLLC(ilIndex).sName)
                        tmAvail.iNoSpotsThis = 0
                        tmAvail.iOrigUnit = 0
                        tmAvail.iOrigLen = 0
                        tmSsf.iCount = tmSsf.iCount + 1
                        tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                    Next ilIndex
                    tgSsf(ilWeekDay - 1) = tmSsf
                End If
                ilRet = BTRV_ERR_NONE
            End If
        End If
       ' Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
       Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            For ilAI = LBound(tmAvailInfo) To UBound(tmAvailInfo) - 1 Step 1
                If tmAvailInfo(ilAI).iAvail60(ilWkIndex) = -29999 Then
                    tmAvailInfo(ilAI).iAvail60(ilWkIndex) = 0
                End If
                If tmAvailInfo(ilAI).iAvail30(ilWkIndex) = -29999 Then
                    tmAvailInfo(ilAI).iAvail30(ilWkIndex) = 0
                End If
                If tmAvailInfo(ilAI).iAvail10(ilWkIndex) = -29999 Then
                    tmAvailInfo(ilAI).iAvail10(ilWkIndex) = 0
                End If
            Next ilAI
            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
            ilEvt = 1
            Do While ilEvt <= tmSsf.iCount
                DoEvents
                If imTerminate Then
                    Exit Sub
                End If
               LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                    ilLtfCode = tmProg.iLtfCode
                ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                    'Determine which rate card program this is associated with
                    For ilAI = LBound(tmAvailInfo) To UBound(tmAvailInfo) - 1 Step 1
                        ilAvailOk = False
                        ilRdfIndex = tmAvailInfo(ilAI).iRdfIndex
                        If (tgMRdf(ilRdfIndex).iLtfCode(0) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(1) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(2) <> 0) Then
                            If (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(0)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Then
                                ilAvailOk = False    'True- code later
                            End If
                        Else
                            For ilLoop = LBound(tgMRdf(ilRdfIndex).iStartTime, 2) To UBound(tgMRdf(ilRdfIndex).iStartTime, 2) Step 1 'Row
                                If (tgMRdf(ilRdfIndex).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdfIndex).iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTimeLong tgMRdf(ilRdfIndex).iStartTime(0, ilLoop), tgMRdf(ilRdfIndex).iStartTime(1, ilLoop), False, llStartTime
                                    gUnpackTimeLong tgMRdf(ilRdfIndex).iEndTime(0, ilLoop), tgMRdf(ilRdfIndex).iEndTime(1, ilLoop), True, llEndTime
                                    If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay) = "Y") Then
                                        ilAvailOk = True
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                        End If
                        If ilAvailOk Then
                            If tgMRdf(ilRdfIndex).sInOut = "I" Then   'Book into
                                If tmAvail.ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
                                    ilAvailOk = False
                                End If
                            ElseIf tgMRdf(ilRdfIndex).sInOut = "O" Then   'Exclude
                                If tmAvail.ianfCode = tgMRdf(ilRdfIndex).ianfCode Then
                                    ilAvailOk = False
                                End If
                            End If
                        End If
                        If ilAvailOk Then
                            DoEvents
                            If imTerminate Then
                                Exit Sub
                            End If
                            ilLen = tmAvail.iLen
                            ilUnits = tmAvail.iAvInfo And &H1F
                            ilNo30 = 0
                            ilNo60 = 0
                            ilNo10 = 0
                            If ilLen = 10 Then
                                ilNo10 = 1
                            End If
                            If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                'Convert inventory to number of 30's and 60's
                                Do While (ilLen >= 30) And (ilUnits > 0)
                                    ilNo30 = ilNo30 + 1
                                    ilLen = ilLen - 30
                                    ilUnits = ilUnits - 1
                                Loop
                                'il30InvCount = il30InvCount + ilNo30
                                ilNo30 = 0
                                ilNo60 = 0
                                ilLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                Do While (ilLen >= 60) And (ilUnits > 0)
                                    ilNo60 = ilNo60 + 1
                                    ilLen = ilLen - 60
                                    ilUnits = ilUnits - 1
                                Loop
                                Do While (ilLen >= 30) And (ilUnits > 0)
                                    ilNo30 = ilNo30 + 1
                                    ilLen = ilLen - 30
                                    ilUnits = ilUnits - 1
                                Loop
                                'il60InvCount = il60InvCount + ilNo60
                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                'Count 30 or 60 and set flag if neither
                                Do While (ilLen >= 30) And (ilUnits > 0)
                                    ilNo30 = ilNo30 + 1
                                    ilLen = ilLen - 30
                                    ilUnits = ilUnits - 1
                                Loop
                                'il30InvCount = il30InvCount + ilNo30
                                ilNo30 = 0
                                ilNo60 = 0
                                ilLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                Do While (ilLen >= 60) And (ilUnits > 0)
                                    ilNo60 = ilNo60 + 1
                                    ilLen = ilLen - 60
                                    ilUnits = ilUnits - 1
                                Loop
                                Do While (ilLen >= 30) And (ilUnits > 0)
                                    ilNo30 = ilNo30 + 1
                                    ilLen = ilLen - 30
                                    ilUnits = ilUnits - 1
                                Loop
                                'il60InvCount = il60InvCount + ilNo60
                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                'Count 30 or 60 and set flag if neither
                                If ilLen = 60 Then
                                    ilNo60 = 1
                                ElseIf ilLen = 30 Then
                                    ilNo30 = 1
                                Else
                                End If
                                'il30InvCount = il30InvCount + ilNo30
                                'il60InvCount = il60InvCount + ilNo60
                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                            End If
                            tmAvailInfo(ilAI).iAvail60(ilWkIndex) = tmAvailInfo(ilAI).iAvail60(ilWkIndex) + ilNo60
                            tmAvailInfo(ilAI).iAvail30(ilWkIndex) = tmAvailInfo(ilAI).iAvail30(ilWkIndex) + ilNo30
                            tmAvailInfo(ilAI).iAvail10(ilWkIndex) = tmAvailInfo(ilAI).iAvail10(ilWkIndex) + ilNo10
                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                DoEvents
                                If imTerminate Then
                                    Exit Sub
                                End If
                                LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                ilSpotOK = True                             'assume spot is OK to include
                                '4/12/18
                                'If tmSpot.iRank = 1010 Then     'DR
                                '    ilSpotOK = False
                                'End If
                                'If tmSpot.iRank = 1020 Then     'Remnant
                                '    ilSpotOK = False
                                'End If
                                'If tmSpot.iRank = 1030 Then     'PI
                                '    ilSpotOK = False
                                'End If
                                'If tmSpot.iRank = 1045 Then     'Extra
                                '    ilSpotOK = False
                                'End If
                                'If tmSpot.iRank = 1050 Then     'Promo
                                '    ilSpotOK = False
                                'End If
                                'If tmSpot.iRank = 1060 Then     'PSA
                                '    ilSpotOK = False
                                'End If
                                If (tmSpot.iRank And RANKMASK) = DIRECTRESPONSERANK Then     'DR
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSpot.iRank And RANKMASK) = REMNANTRANK Then     'Remnant
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSpot.iRank And RANKMASK) = PERINQUIRYRANK Then     'PI
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSpot.iRank And RANKMASK) = EXTRARANK Then     'Extra
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRank And RANKMASK) = PROMORANK Then     'Promo
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRank And RANKMASK) = PSARANK Then     'PSA
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRank And RANKMASK) = RESERVATIONRANK Then     'Reservation
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                    ilSpotOK = False
                                End If
                                
                                If ilSpotOK Then
                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                    '4/12/18
                                    gGetContractParameters tmSdf.lChfCode, slCntrType, ilPctTrade
                                    If (slCntrType = "R") Then     'DR
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If (slCntrType = "T") Then     'Remnant
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If (slCntrType = "Q") Then     'PI
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If (tmSdf.sSpotType = "X") Then     'Extra
                                        ilSpotOK = False
                                    End If
                                    If (slCntrType = "M") Then     'Promo
                                        ilSpotOK = False
                                    End If
                                    If (slCntrType = "S") Then     'PSA
                                        ilSpotOK = False
                                    End If
                                    If (slCntrType = "V") Then     'Reservation
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                End If
                                
                                
                                ilLen = tmSpot.iPosLen And &HFFF
                                If ilSpotOK Then 'continue testing other filters
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    ilNo10 = 0
                                    '5/7/18: added avail length test
                                    If ilLen = 10 And (tmAvail.iLen = 10) Then
                                        ilNo10 = 1
                                        ilLen = 0
                                    End If
                                    'ilLen = tmSpot.iPosLen And &HFFF
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then                   'both units and seconds
                                    'Convert inventory to number of 30's and 60's
                                        Do While ilLen >= 60
                                            ilNo60 = ilNo60 + 1
                                            ilLen = ilLen - 60
                                        Loop
                                        Do While ilLen >= 30
                                            ilNo30 = ilNo30 + 1
                                            ilLen = ilLen - 30
                                        Loop
                                        '5/17/18: Handle case when spot less the 30 sec booked into an avail non-10 sec length
                                        If (ilLen > 0) And (ilLen < 30) And (tmAvail.iLen <> 10) Then
                                            ilNo30 = ilNo30 + 1
                                            ilLen = 0
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then               'units sold
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            If ilLen <= 30 Then
                                                ilNo30 = 1
                                            Else
                                                ilNo60 = 1
                                            End If
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then               'matching units
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                    End If
                                    tmAvailInfo(ilAI).iAvail60(ilWkIndex) = tmAvailInfo(ilAI).iAvail60(ilWkIndex) - ilNo60
                                    tmAvailInfo(ilAI).iAvail30(ilWkIndex) = tmAvailInfo(ilAI).iAvail30(ilWkIndex) - ilNo30
                                    tmAvailInfo(ilAI).iAvail10(ilWkIndex) = tmAvailInfo(ilAI).iAvail10(ilWkIndex) - ilNo10
                                End If                              'ilspotOK
                            Next ilSpot                             'loop from ssf file for # spots in avail
                        End If  'Avail OK
                    Next ilAI
                    ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                End If
                ilEvt = ilEvt + 1   'Increment to next event
            Loop                                                        'do while ilEvt <= tmSsf.iCount
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If tgMVef(ilVefIndex).sType = "G" Then
                ilType = tmSsf.iType
            End If
        Loop
        lmProcessedNoRecs = lmProcessedNoRecs + 1
        'lacInfo.Caption = Trim$(Str$(lmProcessedNoRecs)) & " of" & Str$(lmTotalNoRecs)
        llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
        If llPercent >= 100 Then
            llPercent = 100
        End If
        plcGauge.Value = llPercent
    Next llLoopDate
    If tgSpf.sCIncludeMissDB <> "N" Then
        'Get missed
        'Key 2: VefCode; SchStatus; AdfCode; Date, Time
        For ilPass = 0 To 2 Step 1
            DoEvents
            If imTerminate Then
                Exit Sub
            End If
            tmSdfSrchKey2.iVefCode = ilVefCode
            If ilPass = 0 Then
                slType = "M"
            ElseIf ilPass = 1 Then
                slType = "R"
            ElseIf ilPass = 2 Then
                slType = "U"
            End If
            slDate = Format$(llStartDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            tmSdfSrchKey2.sSchStatus = slType
            tmSdfSrchKey2.iAdfCode = 0
            tmSdfSrchKey2.iDate(0) = ilDate0
            tmSdfSrchKey2.iDate(1) = ilDate1
            tmSdfSrchKey2.iTime(0) = 0
            tmSdfSrchKey2.iTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            'This code added as replacement for Ext operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
                DoEvents
                If imTerminate Then
                    Exit Sub
                End If
                ilSpotOK = True
                '4/12/18
                gGetContractParameters tmSdf.lChfCode, slCntrType, ilPctTrade
                If (slCntrType = "R") Then     'DR
                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                        ilSpotOK = False
                    End If
                End If
                If (slCntrType = "T") Then     'Remnant
                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                        ilSpotOK = False
                    End If
                End If
                If (slCntrType = "Q") Then     'PI
                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                        ilSpotOK = False
                    End If
                End If
                If (tmSdf.sSpotType = "X") Then     'Extra
                    ilSpotOK = False
                End If
                If (slCntrType = "M") Then     'Promo
                    ilSpotOK = False
                End If
                If (slCntrType = "S") Then     'PSA
                    ilSpotOK = False
                End If
                If (slCntrType = "V") Then     'Reservation
                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                        ilSpotOK = False
                    End If
                End If

                If ilSpotOK Then
                'end add 4/12/18
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    'If llDate > llEndDate Then
                    '    Exit Do
                    'End If
                    If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                        ilDay = gWeekDayLong(llDate)
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                        For ilAI = LBound(tmAvailInfo) To UBound(tmAvailInfo) - 1 Step 1
                            DoEvents
                            If imTerminate Then
                                Exit Sub
                            End If
                            ilSpotOK = False
                            ilRdfIndex = tmAvailInfo(ilAI).iRdfIndex
                            If (tgMRdf(ilRdfIndex).iLtfCode(0) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(1) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(2) <> 0) Then
                                If (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(0)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Then
                                    ilSpotOK = False    'True- code later
                                End If
                            Else
                                For ilLoop = LBound(tgMRdf(ilRdfIndex).iStartTime, 2) To UBound(tgMRdf(ilRdfIndex).iStartTime, 2) Step 1 'Row
                                    If (tgMRdf(ilRdfIndex).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdfIndex).iStartTime(1, ilLoop) <> 0) Then
                                        gUnpackTimeLong tgMRdf(ilRdfIndex).iStartTime(0, ilLoop), tgMRdf(ilRdfIndex).iStartTime(1, ilLoop), False, llStartTime
                                        gUnpackTimeLong tgMRdf(ilRdfIndex).iEndTime(0, ilLoop), tgMRdf(ilRdfIndex).iEndTime(1, ilLoop), True, llEndTime
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay) = "Y") Then
                                            ilSpotOK = True
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                            If ilSpotOK Then
                                ilAnf = gBinarySearchAnf(tgMRdf(ilRdfIndex).ianfCode, tgAvailAnf())
                                DoEvents
                                If imTerminate Then
                                    Exit Sub
                                End If
                                tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                tmClfSrchKey.iLine = tmSdf.iLineNo
                                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
                                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                                    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                        If tgMRdf(ilRdf).iCode = tmClf.iRdfCode Then
                                            If (tgMRdf(ilRdf).iCode <> tgMRdf(ilRdfIndex).iCode) Then   'Book into
                                                If (tgMRdf(ilRdf).sInOut = "I") And (tgMRdf(ilRdfIndex).sInOut = "I") Then   'Book into
                                                    If tgMRdf(ilRdf).ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
                                                        ilSpotOK = False
                                                    End If
                                                ElseIf (tgMRdf(ilRdf).sInOut = "O") And (tgMRdf(ilRdfIndex).sInOut = "O") Then    'Exclude
                                                    If tgMRdf(ilRdf).ianfCode = tgMRdf(ilRdfIndex).ianfCode Then
                                                        ilSpotOK = False
                                                    End If
                                                Else
                                                    '5/7/18
                                                    If (tgMRdf(ilRdf).sInOut <> "I") And (tgMRdf(ilRdf).sInOut <> "O") Then
                                                        If ilAnf <> -1 Then
                                                            If tgAvailAnf(ilAnf).sSustain <> "Y" Then
                                                                ilSpotOK = False
                                                            End If
                                                        End If
                                                    Else
                                                        ilSpotOK = False
                                                    End If
                                                End If
                                            End If
                                            Exit For
                                        End If
                                    Next ilRdf
                                Else
                                    ilSpotOK = False
                                End If
                            End If
                            If ilSpotOK Then
                                '4/12/18
                                If (tmSdf.sSpotType = "R") Then     'DR
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSdf.sSpotType = "T") Then     'Remnant
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSdf.sSpotType = "Q") Then     'PI
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                                        ilSpotOK = False
                                    End If
                                End If
                                If (tmSdf.sSpotType = "X") Then     'Extra
                                    ilSpotOK = False
                                End If
                                If (tmSdf.sSpotType = "M") Then     'Promo
                                    ilSpotOK = False
                                End If
                                If (tmSdf.sSpotType = "S") Then     'PSA
                                    ilSpotOK = False
                                End If
                                If (tmSdf.sSpotType = "V") Then     'Reservation
                                    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                                        ilSpotOK = False
                                    End If
                                End If
                            End If
                            ilLen = tmSdf.iLen
                            If ilSpotOK Then
                                DoEvents
                                If imTerminate Then
                                    Exit Sub
                                End If
                                'Determine if Avr created
                                ilNo30 = 0
                                ilNo60 = 0
                                ilNo10 = 0
                                If ilLen = 10 Then
                                    ilNo10 = 1
                                End If
                                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                'Convert inventory to number of 30's and 60's
                                    Do While ilLen >= 60
                                        ilNo60 = ilNo60 + 1
                                        ilLen = ilLen - 60
                                    Loop
                                    Do While ilLen >= 30
                                        ilNo30 = ilNo30 + 1
                                        ilLen = ilLen - 30
                                    Loop
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                    'Count 30 or 60 and set flag if neither
                                    If ilLen = 60 Then
                                        ilNo60 = 1
                                    ElseIf ilLen = 30 Then
                                        ilNo30 = 1
                                    Else
                                        If ilLen <= 30 Then
                                            ilNo30 = 1
                                        Else
                                            ilNo60 = 1
                                        End If
                                    End If
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                    'Count 30 or 60 and set flag if neither
                                    If ilLen = 60 Then
                                        ilNo60 = 1
                                    ElseIf ilLen = 30 Then
                                        ilNo30 = 1
                                    End If
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                End If
                                ilWkIndex = (llDate - llStartDate) \ 7 + 1
                                If tmAvailInfo(ilAI).iAvail60(ilWkIndex) = -29999 Then
                                    tmAvailInfo(ilAI).iAvail60(ilWkIndex) = 0
                                End If
                                tmAvailInfo(ilAI).iAvail60(ilWkIndex) = tmAvailInfo(ilAI).iAvail60(ilWkIndex) - ilNo60
                                If tmAvailInfo(ilAI).iAvail30(ilWkIndex) = -29999 Then
                                    tmAvailInfo(ilAI).iAvail30(ilWkIndex) = 0
                                End If
                                tmAvailInfo(ilAI).iAvail30(ilWkIndex) = tmAvailInfo(ilAI).iAvail30(ilWkIndex) - ilNo30
                                If tmAvailInfo(ilAI).iAvail10(ilWkIndex) = -29999 Then
                                    tmAvailInfo(ilAI).iAvail10(ilWkIndex) = 0
                                End If
                                tmAvailInfo(ilAI).iAvail10(ilWkIndex) = tmAvailInfo(ilAI).iAvail10(ilWkIndex) - ilNo60
                            End If
                        Next ilAI
                    End If
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            lmProcessedNoRecs = lmProcessedNoRecs + 1
            'lacInfo.Caption = Trim$(Str$(lmProcessedNoRecs)) & " of" & Str$(lmTotalNoRecs)
            llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            plcGauge.Value = llPercent
        Next ilPass
    End If
    For ilWkIndex = 1 To 26 Step 1
        DoEvents
        If imTerminate Then
            Exit Sub
        End If
        For ilAI = LBound(tmAvailInfo) To UBound(tmAvailInfo) - 1 Step 1
            DoEvents
            If imTerminate Then
                Exit Sub
            End If
            If tmAvailInfo(ilAI).iAvail30(ilWkIndex) <> -29999 Then
                'Adjust counts
                If (tgVpf(ilVpfIndex).sSSellOut = "B") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
                    ilAvail60 = tmAvailInfo(ilAI).iAvail60(ilWkIndex)
                    If tmAvailInfo(ilAI).iAvail30(ilWkIndex) < 0 Then
                        Do While (ilAvail60 > 0) And (tmAvailInfo(ilAI).iAvail30(ilWkIndex) < 0)
                            ilAvail60 = ilAvail60 - 1
                            tmAvailInfo(ilAI).iAvail30(ilWkIndex) = tmAvailInfo(ilAI).iAvail30(ilWkIndex) + 2
                        Loop
                    End If
                    If ilAvail60 > 0 Then
                        tmAvailInfo(ilAI).iAvail30(ilWkIndex) = tmAvailInfo(ilAI).iAvail30(ilWkIndex) + 2 * ilAvail60
                    End If
                    tmAvailInfo(ilAI).iAvail60(ilWkIndex) = ilAvail60
                End If
                If (tgVpf(ilVpfIndex).sSSellOut = "U") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
                    ilAvail60 = tmAvailInfo(ilAI).iAvail60(ilWkIndex)
                    If tmAvailInfo(ilAI).iAvail30(ilWkIndex) < 0 Then
                        Do While (ilAvail60 > 0) And (tmAvailInfo(ilAI).iAvail30(ilWkIndex) < 0)
                            ilAvail60 = ilAvail60 - 1
                            tmAvailInfo(ilAI).iAvail30(ilWkIndex) = tmAvailInfo(ilAI).iAvail30(ilWkIndex) + 1
                        Loop
                    End If
                    tmAvailInfo(ilAI).iAvail30(ilWkIndex) = ilAvail60
                    tmAvailInfo(ilAI).iAvail60(ilWkIndex) = ilAvail60
                End If
                'If ilLnLen <= 30 Then
                '    ilSpotCount = il30Count + 2 * il60Count
                '    ilAvail = il30Avail
                '    ilInvCount = il30InvCount   'this includes 60 inventory
                'Else
                '    ilSpotCount = 2 * il60Count + il30Count
                '    ilAvail = il60Avail
                '    ilInvCount = 2 * il60InvCount
                'End If
            End If
        Next ilAI
    Next ilWkIndex
    Erase ilEvtType
    Erase tmLLC
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
    imGeneratingASF = False
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
    gCenterStdAlone GenASF
    Screen.MousePointer = vbDefault
    imcHelp.Picture = Traffic!imcHelp.Picture
    'imMinElapsed = imTimeoutInterval - 2
    'If imTimeoutInterval > 0 Then
    If Not imUserInput Then
        GenASF.WindowState = 1
        'tmcTimer.Interval = imTimeoutInterval
        tmcTimer.Enabled = True
        cmcGen.Enabled = False
        gOpenTmf
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
    End If
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    sgTmfStatus = "E"
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
'*      Procedure Name:mInitDDE                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize stand alone mode    *
'*                                                     *
'*******************************************************
Private Sub mInitDDE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    If (Not igStdAloneMode) And (imShowHelpMsg) Then
        'Set up DDE for Branner and help message area within traffic
        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
        edcLinkDestHelpMsg.LinkTopic = "Traffic|HelpMsg"
        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceHelpMsg"
        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    End If
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
'    imTimeoutInterval = 0
'    ilPos = InStr(1, slCommand, "/CS", 1)
'    If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
'        igStdAloneMode = True 'False  'Switch from/to stand alone mode
'        sgCallAppName = ""
'        ilTestSystem = False
'        ilPos = InStr(1, slCommand, "/S:Test", 1)
'        If ilPos > 0 Then
'            ilTestSystem = True
'        End If
'        ilPos = InStr(1, slCommand, "/S:Prod", 1)
'        If ilPos > 0 Then
'            ilTestSystem = False
'        End If
'        slStr = ""
'        ilPos = InStr(1, slCommand, "/T:", 1)
'        If ilPos > 0 Then
'            ilPos = ilPos + 3
'            Do While ilPos <= Len(slCommand)
'                slChar = Mid$(slCommand, ilPos, 1)
'                If Trim$(slChar) = "" Then
'                    Exit Do
'                End If
'                slStr = slStr & slChar
'                ilPos = ilPos + 1
'            Loop
'            imTimeoutInterval = CInt(slStr)
'        End If
'        slStr = "Guide"
'        slCommand = "Traffic\Guide"
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
'        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
'        If Trim$(slStr) = "" Then
'            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
'            End
'        End If
'        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
'        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
'        If StrComp(slTestSystem, "Test", 1) = 0 Then
'            ilTestSystem = True
'        Else
'            ilTestSystem = False
'        End If
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
'        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
'    End If
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
    imUserInput = False
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imUserInput = True
        igBkgdProg = 0
'        gLogMsg "Current Mode: UserInput.", "GenASF.txt", False
    Else
        igBkgdProg = 5
'        gLogMsg "Current Mode: Background.", "GenASF.txt", False
    End If
    'If imTimeoutInterval > 0 Then
    '    igBkgdProg = 5
    'End If
    slStr = "Guide"
    gInitStdAlone GenASF, slStr, ilTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "GenASF.Frm"
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "GenASF.txt", False
    Else
        gLogMsg "Current Mode: Background.", "GenASF.txt", False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSendHelpMsg                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Send message if not stand alone*
'*                                                     *
'*******************************************************
Private Sub mSendHelpMsg(slMsg As String)
    Dim ilLoop As Integer
    If (Not igStdAloneMode) And (imShowHelpMsg) Then
        edcLinkDestHelpMsg.LinkExecute slMsg
        For ilLoop = 1 To igDDEDelay Step 1
            DoEvents
        Next ilLoop
    End If
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
    igParentRestarted = False
    If Not igStdAloneMode Then
        If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
            edcLinkDestHelpMsg.LinkExecute "@" & "Done"
        Else
            edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
            edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
            edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
            edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
            edcLinkDestHelpMsg.LinkExecute "Done"
        End If
        Do While Not igParentRestarted
            DoEvents
        Loop
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload GenASF
    igManUnload = NO
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "ASG"
End Sub

Private Sub tmcTimer_Timer()
    If imGeneratingASF Then
        Exit Sub
    End If
    tmcTimer.Enabled = False
    If Not imUserInput Then
        gUpdateTaskMonitor 1, "ASG"
    'imMinElapsed = imMinElapsed + 1
    'If imMinElapsed >= imTimeoutInterval Then
        cmcGen_Click
        gUpdateTaskMonitor 2, "ASG"
    Else
        cmcGen_Click
    End If
    '    imMinElapsed = 0
    'End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Avail Generation"
End Sub

Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Asf.Btr", Len(tmAsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Sdf.Btr", Len(tmSdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Lcf.Btr", Len(tmLcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tmClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Rcf.Btr", Len(tmRcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        If Not gRecLengthOk("Gif.Btr", Len(tmGif)) Then
            mCheckRecLength = False
            Exit Function
        End If
    End If
    mCheckRecLength = True
End Function

Private Function mCreateGIF() As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    If ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        ilVef = gBinarySearchVef(tmAsf.iVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).iNrfCode > 0 Then
                tmGif.lCode = 0
                tmGif.sNRProcessed = "N"
                tmGif.sType = "A"
                tmGif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                tmGif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                Select Case Left$(gGetLocalTZName(), 1)
                    Case "E"
                    Case "C"
                        tmGif.lGenTime = tmGif.lGenTime + 3600
                    Case "M"
                        tmGif.lGenTime = tmGif.lGenTime + 7200
                    Case "P"
                        tmGif.lGenTime = tmGif.lGenTime + 10800
                End Select
                tmGif.lCntrNo = 0
                tmGif.iVefCode = tmAsf.iVefCode
                tmGif.iStartDate(0) = 0
                tmGif.iStartDate(1) = 0
                tmGif.iEndDate(0) = 0
                tmGif.iEndDate(1) = 0
                tmGif.sMessageRepDBID = ""
                tmGif.sMessageNetDBID = ""
                tmGif.lMessageCefCode = 0
                gPackDate "", tmGif.iProcessedDate(0), tmGif.iProcessedDate(1)
                gPackTime "", tmGif.iProcessedTime(0), tmGif.iProcessedTime(1)
                ilRet = btrInsert(hmGif, tmGif, imGifRecLen, INDEXKEY0)
            End If
        End If
    End If
    mCreateGIF = True
End Function
