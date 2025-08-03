VERSION 5.00
Begin VB.Form RepNetLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rep-Net Link"
   ClientHeight    =   3600
   ClientLeft      =   4245
   ClientTop       =   3630
   ClientWidth     =   6030
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
   Icon            =   "RepNetLink.frx":0000
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.PictureBox plcSignon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      Picture         =   "RepNetLink.frx":08CA
      ScaleHeight     =   3555
      ScaleWidth      =   5985
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6045
      Begin VB.PictureBox pbcAvails 
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1785
         ScaleHeight     =   285
         ScaleWidth      =   3540
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1575
         Width           =   3600
         Begin VB.Label lacAvailsStatus 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Idle"
            Height          =   300
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   3540
         End
      End
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   330
         Top             =   2535
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   45
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   930
         Width           =   105
      End
      Begin VB.PictureBox plcName 
         Height          =   345
         Left            =   1785
         ScaleHeight     =   285
         ScaleWidth      =   3540
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   825
         Width           =   3600
         Begin VB.Label lacContractStatus 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Idle"
            Height          =   300
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   3555
         End
      End
      Begin VB.PictureBox plcPassword 
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1785
         ScaleHeight     =   285
         ScaleWidth      =   3540
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3600
         Begin VB.Label lacPostedSpotsStatus 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Idle"
            Height          =   300
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   3540
         End
      End
      Begin VB.Label lacAvails 
         BackStyle       =   0  'Transparent
         Caption         =   "Avails"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   345
         TabIndex        =   13
         Top             =   1620
         Width           =   1350
      End
      Begin VB.Label lacStop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   8
         Top             =   2985
         Width           =   1335
      End
      Begin VB.Label lacMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2325
         TabIndex        =   7
         Top             =   2460
         Width           =   1350
      End
      Begin VB.Label lacPostedSpots 
         BackStyle       =   0  'Transparent
         Caption         =   "Posted Spots"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   345
         TabIndex        =   6
         Top             =   1230
         Width           =   1350
      End
      Begin VB.Label lacContract 
         BackStyle       =   0  'Transparent
         Caption         =   "Contract"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   345
         TabIndex        =   5
         Top             =   840
         Width           =   1170
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   3210
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   2955
         TabIndex        =   4
         Top             =   2415
         Visible         =   0   'False
         Width           =   1410
      End
   End
End
Attribute VB_Name = "RepNetLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RepNetLink.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text


Dim hmTextChf As Integer
Dim hmTextClf As Integer
Dim hmTextMsg As Integer
Dim hmTextSdf As Integer
Dim hmTextAsf As Integer
Dim smFileName1 As String
Dim smFileName2 As String
Dim smFileName3 As String

Dim lmSeqNo As Long
Dim imCancelled As Integer
Dim imClosed As Integer
Dim lmSleepTime As Long
Dim imTaskCount As Integer  'Task count in seconds (i.e. 60 = 1 minute; 300=5 minutes)
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim smNowDate As String
Dim lmNowDate As Long
Dim imOkToUnload As Integer
Dim tmNetChfSort() As NETCHFSORT



Private Sub Form_Activate()
'    gShowBranner
End Sub


Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Only one copy of RepNetLink can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of RepNetLink path: " & App.Path & " from " & Trim$(gGetComputerName()), "RepNetLink.txt", False
        End
    End If
    gSetBkgdMode
    igBkgdProg = 6
    mInit
    tmcStart.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    If imClosed = True Then
        Exit Sub
    End If
    Cancel = 0
    tmcStart.Enabled = False
    
    'Doug
    'See if web is locked.  If so transfer is occuring, do you still want stop or not?  I so then unlock and get out.
    'if not locked just show the message below.
    
    ilRet = MsgBox("Stop the Rep-Net Link", vbQuestion + vbYesNo, "Stop Service")
    If ilRet = vbNo Then
        Cancel = 1
        imCancelled = False
        Call gEndWebSession("RepNetLink.txt")
        'tmcRestartTask.Enabled = True
        Exit Sub
    End If
    imClosed = True
    imCancelled = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    btrDestroy hgRNNrf
    btrDestroy hgAuf
    btrDestroy hgRNAuf
    btrDestroy hgRNSof
    btrDestroy hgRNSlf
    btrDestroy hgRNSmf
    btrDestroy hgRNCef
    btrDestroy hgRNRdf
    btrDestroy hgRNRcf
    btrDestroy hgRNMif
    btrDestroy hgRNMnf
    btrDestroy hgRNCxf
    btrDestroy hgRNAgf
    btrDestroy hgRNAdf
    btrDestroy hgRNQif
    btrDestroy hgRNGif
    btrDestroy hgRNVef
    btrDestroy hgRNVsf
    btrDestroy hgRNCff
    btrDestroy hgRNClf
    btrExtClear hgRNChf   'Clear any previous extend operation
    btrDestroy hgRNChf
    btrExtClear hgRNSdf
    btrDestroy hgRNSdf
    btrDestroy hgRNSbf
    btrDestroy hgRNCif
    btrDestroy hgRNCpf
    btrDestroy hgRNAsf
    btrDestroy hgRNAnf
        
    btrStopAppl
    Set RepNetLink = Nothing   'Remove data segment
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
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llAdfNoRecs As Integer
    Dim llAgfNoRecs As Integer
    imOkToUnload = True
    imTerminate = False
    
    'Screen.MousePointer = vbHourGlass
    mParseCmmdLine
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
    lmSeqNo = 1
    imTerminate = False
    imFirstFocus = True
    lmSleepTime = 1000 ' 1 seconds
    If (sgNetDBPath <> "") Or (sgRepDBPath <> "") Then
        imTaskCount = 5 '60
    Else
        imTaskCount = 15    '300
    End If
    imCancelled = False
    imClosed = False
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    
    hgAuf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgAuf, "", sgDBPath & "Auf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Auf.Btr)", RepNetLink
    On Error GoTo 0
    
    hgRNChf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", RepNetLink
    On Error GoTo 0
    igRNChfRecLen = Len(tgRNChf)
    hgRNClf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", RepNetLink
    On Error GoTo 0
    ReDim tgRNClf(0 To 0) As CLF
    igRNClfRecLen = Len(tgRNClf(0))
    hgRNCff = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", RepNetLink
    On Error GoTo 0
    ReDim tgRNCff(0 To 0) As CFF
    igRNCffRecLen = Len(tgRNCff(0))
    hgRNVsf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", RepNetLink
    On Error GoTo 0
    ReDim tgRNVsf(0 To 0) As VSF
    igRNVsfRecLen = Len(tgRNVsf(0))
    hgRNVef = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", RepNetLink
    On Error GoTo 0
    igRNVefRecLen = Len(tgRNVef)
    hgRNGif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNGif, "", sgDBPath & "Gif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gif.Btr)", RepNetLink
    On Error GoTo 0
    igRNGifRecLen = Len(tgRNGif)
    hgRNQif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNQif, "", sgDBPath & "Qif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Qif.Btr)", RepNetLink
    On Error GoTo 0
    igRNQifRecLen = Len(tgRNQif)
    hgRNNrf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNNrf, "", sgDBPath & "Nrf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Nrf.Btr)", RepNetLink
    On Error GoTo 0
    igRNNrfRecLen = Len(tgNrf)
    
    hgRNAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", RepNetLink
    On Error GoTo 0
    igRNAdfRecLen = Len(tgRNAdf)
    hgRNAgf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", RepNetLink
    On Error GoTo 0
    igRNAgfRecLen = Len(tgRNAgf)
    hgRNCxf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cxf.Btr)", RepNetLink
    On Error GoTo 0
    igRNCxfRecLen = Len(tgRNCxf)
    hgRNMnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", RepNetLink
    On Error GoTo 0
    igRNMnfRecLen = Len(tgRNMnf)
    hgRNMif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNMif, "", sgDBPath & "Mif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mif.Btr)", RepNetLink
    On Error GoTo 0
    igRNMifRecLen = Len(tgRNMif)
    hgRNRcf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rcf.Btr)", RepNetLink
    On Error GoTo 0
    igRNRcfRecLen = Len(tgRNRcf)
    hgRNRdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", RepNetLink
    On Error GoTo 0
    igRNRdfRecLen = Len(tgRNRdf)
    hgRNCef = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.Btr)", RepNetLink
    On Error GoTo 0
    igRNCefRecLen = Len(tgRNCef)
    hgRNSmf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", RepNetLink
    On Error GoTo 0
    igRNSmfRecLen = Len(tgRNSmf)
    hgRNSlf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf.Btr)", RepNetLink
    On Error GoTo 0
    igRNSlfRecLen = Len(tgRNSlf)
    hgRNSof = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sof.Btr)", RepNetLink
    On Error GoTo 0
    igRNSofRecLen = Len(tgRNSof)
    hgRNAuf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNAuf, "", sgDBPath & "Auf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Auf.Btr)", RepNetLink
    On Error GoTo 0
    igRNAufRecLen = Len(tgRNAuf)
    hgRNSdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", RepNetLink
    On Error GoTo 0
    igRNSdfRecLen = Len(tgRNSdf)
    hgRNSbf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hgRNSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", RepNetLink
    On Error GoTo 0
    igRNSbfRecLen = Len(tgRNSbf)
    hgRNCif = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", RepNetLink
    On Error GoTo 0
    igRNCifRecLen = Len(tgRNCif)
    hgRNCpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", RepNetLink
    On Error GoTo 0
    igRNCpfRecLen = Len(tgRNCpf)
    hgRNAsf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNAsf, "", sgDBPath & "Asf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Asf.Btr)", RepNetLink
    On Error GoTo 0
    igRNAsfRecLen = Len(tgRNAsf)
    hgRNAnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hgRNAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Anf.Btr)", RepNetLink
    On Error GoTo 0
    igRNAnfRecLen = Len(tgRNAnf)

    'If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
    '    lacAvails.Visible = False
    '    pbcAvails.Visible = False
    'End If
    If Not mRNNrfPop() Then
        imTerminate = True
    End If
    If imTerminate Then
        Exit Sub
    End If
    ilRet = mCheckVersions()
    If Not ilRet Then
        imTerminate = True
    End If
    ilRet = gObtainVef()
    If Not ilRet Then
        imTerminate = True
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
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = slCommand
    igDirectCall = 0
    sgIniPath = ""
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 8))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    sgXFerDBID = ""
    ilPos = InStr(1, slCommand, "/XFerDBID:", vbTextCompare)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgXFerDBID = Trim$(Mid$(slCommand, ilPos + 10))
        Else
            sgXFerDBID = Trim$(Mid$(slCommand, ilPos + 10, ilSpace - ilPos - 10))
        End If
    End If
    sgNetDBPath = ""
    ilPos = InStr(1, slCommand, "/NetDB:", vbTextCompare)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgNetDBPath = Trim$(Mid$(slCommand, ilPos + 7))
        Else
            sgNetDBPath = Trim$(Mid$(slCommand, ilPos + 7, ilSpace - ilPos - 7))
        End If
        If right$(sgNetDBPath, 1) <> "\" Then
            sgNetDBPath = sgNetDBPath & "\"
        End If
    End If
    sgRepDBPath = ""
    ilPos = InStr(1, slCommand, "/RepDB:", vbTextCompare)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgRepDBPath = Trim$(Mid$(slCommand, ilPos + 7))
        Else
            sgRepDBPath = Trim$(Mid$(slCommand, ilPos + 7, ilSpace - ilPos - 7))
        End If
        If right$(sgRepDBPath, 1) <> "\" Then
            sgRepDBPath = sgRepDBPath & "\"
        End If
    End If
    
    'If the the program in the start-up menu then Pervasive may not have had time to launch it's engine
    'so wait for one minute so long as we are not in debug mode.
    If slCommand <> "debug" Then
        Sleep (ONE_MINUTE)
    End If
    
    gInitStdAlone RepNetLink, slStr, ilTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "RepNetLink.Frm"

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

    igParentRestarted = False
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload RepNetLink
    igManUnload = NO
End Sub


Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Chf.Btr", Len(tgRNChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tgRNClf(0))) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tgRNCff(0))) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Adf.Btr", Len(tgRNAdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Agf.Btr", Len(tgRNAgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Mnf.Btr", Len(tgRNMnf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Qif.Btr", Len(tgRNQif)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Gif.Btr", Len(tgRNGif)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Nrf.Btr", Len(tgNrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

Private Sub lacMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacMin.Left - 30, lacMin.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub lacStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacStop.Left - 30, lacStop.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub pbcClickFocus_Click()
    imcOutline.Visible = False
End Sub

Private Sub plcSignon_Click()
    imcOutline.Visible = False
End Sub

Private Sub tmcStart_Timer()
    Dim ilLoop As Integer
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Or ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        mPollTask
    Else
        mTerminate
        Exit Sub
    End If
End Sub
Private Sub lacMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lacStop_Click()
    imCancelled = True
End Sub


Private Sub mPollTask()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilTaskCount As Integer
    Dim ilIdx As Integer

    ilTaskCount = -1
    Do
        DoEvents
        Sleep lmSleepTime
        If imCancelled Then
            mTerminate
            Exit Sub
        End If
        
        lacContract.Caption = "Contract"
        lacPostedSpots.Caption = "Posted Spots"
        lacAvails.Caption = "Avail Counts"
        If (ilTaskCount = -1) Or (ilTaskCount = imTaskCount) Then
            'For ilIdx = 0 To UBound(tgServerNRF) - 1 Step 1
            '    If (sgNetDBPath = "") And (sgRepDBPath = "") Then 'SQL Server workaround
            '        lacContract.Caption = "Testing"
            '        lacPostedSpots.Caption = "Testing"
            '        lacContractStatus.Caption = "Testing Internet Access"
            '        lacPostedSpotsStatus.Caption = Trim(tgServerNRF(ilIdx).sIISRootURL)
            '        If tgServerNRF(ilIdx).sType = "N" Then
            '            Call gTestAccessToWebServer(ilIdx)
            '            If Not gWebAccessTestedOk Then
            '                gURLHasAccess(ilIdx) = False
            '                lacContractStatus.Caption = "No Internet Access Found"
            '                lacPostedSpotsStatus.Caption = Trim(tgServerNRF(ilIdx).sIISRootURL)
            '                gLogMsg "No Internet Access Found For: " & Trim(tgServerNRF(ilIdx).sIISRootURL), "RepNetLink.txt", True
            '            Else
            '                gURLHasAccess(ilIdx) = True
            '            End If
            '        Else
            '            gURLHasAccess(ilIdx) = False
            '        End If
            '    Else
            '        gURLHasAccess(ilIdx) = True
            '    End If
            'Next ilIdx
        
            If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                'Check for Contracts to be sent to server
                mTrafficToText
                mTrafficTextToServer
                'Check for Posted Spots to be installed onto Traffic
                mServerToText
                mServerTextToTraffic
            ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                'Check for Posted Spots to be sent to server
                mTrafficToText
                mTrafficTextToServer
                'Check for Contracts from server to be installed on Traffic
                mServerToText
                mServerTextToTraffic
            End If
            ilTaskCount = 1
        Else
            ilTaskCount = ilTaskCount + 1
        End If
   Loop
End Sub



Private Sub mTrafficToText()

    'Check for contracts that need to be sent to the web server
    Dim ilRet As Integer
    Dim slChfCodes As String
    Dim slChfCode As String
    Dim ilItem As Integer
    Dim slRepDBID As String
    Dim slStr As String
    
    lacContractStatus.Caption = "Checking"
    lacPostedSpotsStatus.Caption = "Checking"
    lacAvailsStatus.Caption = "Checking"
    DoEvents
    tgRNGifSrchKey1.sNRProcessed = "N"
    tgRNGifSrchKey1.lGenDate = 0
    tgRNGifSrchKey1.lGenTime = 0
    ilRet = btrGetGreaterOrEqual(hgRNGif, tgRNGif, igRNGifRecLen, tgRNGifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record

    Do While (ilRet = BTRV_ERR_NONE) And (tgRNGif.sNRProcessed = "N")
        'Get Contract records
        lmSeqNo = 1
        Select Case tgRNGif.sType
            Case "C"    'Get Contract record
                lacContractStatus.Caption = "Extracting"
                lacPostedSpotsStatus.Caption = "Idle"
                lacAvailsStatus.Caption = "Idle"
                DoEvents
                ilRet = mContractToText()
                If Not ilRet Then
                    Exit Sub
                End If
            Case "S"    'Get Posted Spots
                lacContractStatus.Caption = "Idle"
                lacPostedSpotsStatus.Caption = "Extracting"
                lacAvailsStatus.Caption = "Idle"
                DoEvents
                'Gather all contracts to be exported
                ReDim tmNetChfSort(0 To 0) As NETCHFSORT
                tgRNCefSrchKey.lCode = tgRNGif.lMessageCefCode
                If tgRNCefSrchKey.lCode > 0 Then
                    igRNCefRecLen = Len(tgRNCef) '5027
                    ilRet = btrGetEqual(hgRNCef, tgRNCef, igRNCefRecLen, tgRNCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                    'If tgRNCef.iStrLen > 0 Then
                        slStr = gStripChr0(tgRNCef.sComment)
                        If slStr <> "" Then
                            slRepDBID = Trim$(tgRNGif.sMessageRepDBID)
                            slChfCodes = slStr  'Left$(tgRNCef.sComment, tgRNCef.iStrLen)
                            ilItem = 1
                            Do
                                ilRet = gParseItem(slChfCodes, ilItem, ",", slChfCode)
                                If slChfCode <> "" Then
                                    tmNetChfSort(UBound(tmNetChfSort)).lChfCode = Val(slChfCode)
                                    ReDim Preserve tmNetChfSort(0 To UBound(tmNetChfSort) + 1) As NETCHFSORT
                                Else
                                    Exit Do
                                End If
                                ilItem = ilItem + 1
                            Loop
                        End If
                    End If
                End If
                If UBound(tmNetChfSort) > 0 Then
                    ArraySortTyp fnAV(tmNetChfSort(), 0), UBound(tmNetChfSort), 0, LenB(tmNetChfSort(0)), 0, -2, 0
                    ilRet = mPostedSpotsToText(slRepDBID)
                    If Not ilRet Then
                        Exit Sub
                    End If
                End If
            Case "M"    'Get Message
                ilRet = mMessageToText()
                If Not ilRet Then
                    Exit Sub
                End If
            Case "A"    '6 month avails
                lacContractStatus.Caption = "Idle"
                lacPostedSpotsStatus.Caption = "Idle"
                lacAvailsStatus.Caption = "Extracting"
                ilRet = mAvailSumToText()
                If Not ilRet Then
                    Exit Sub
                End If
        End Select
        
        'Clear GIF Processed flag
        Do
            tgRNGifSrchKey0.lCode = tgRNGif.lCode
            ilRet = btrGetEqual(hgRNGif, tgRNGif, igRNGifRecLen, tgRNGifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tgRNGif.sNRProcessed = "Y"
                gPackDate Format$(gNow(), "m/d/yy"), tgRNGif.iProcessedDate(0), tgRNGif.iProcessedDate(1)
                gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNGif.iProcessedTime(0), tgRNGif.iProcessedTime(1)
                ilRet = btrUpdate(hgRNGif, tgRNGif, igRNGifRecLen)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        
        'Get Next recorded to be processed
        tgRNGifSrchKey1.sNRProcessed = "N"
        tgRNGifSrchKey1.lGenDate = 0
        tgRNGifSrchKey1.lGenTime = 0
        ilRet = btrGetGreaterOrEqual(hgRNGif, tgRNGif, igRNGifRecLen, tgRNGifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record
    Loop
    lacContractStatus.Caption = "Idle"
    lacPostedSpotsStatus.Caption = "Idle"
    lacAvailsStatus.Caption = "Idle"
    DoEvents
End Sub

Private Sub mTrafficTextToServer()
    Dim ilRet As Integer
    Dim ilSQLLocked As Integer
    Dim ilNrf As Integer
    Dim hlQif As Integer
    Dim tlQif As QIF
    Dim ilOKToSend As Integer
    Dim slTempPath As String
    Dim llETime As Long
    Dim llDate As Long
    Dim SQLQuery As String
    Dim slDateTime As String
    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilFound As Integer
    Dim slToRepOrNet As String
    Dim slDataArray() As String
    Dim llRet As Long
    Dim ilSvNrfIndex As Integer
    Dim ilSentToServer As Integer
    Dim ilBypassSendingToServer As Integer
    ReDim ilNRFFailedCode(0 To 0) As Integer
    
    lacContractStatus.Caption = "Checking"
    lacPostedSpotsStatus.Caption = "Checking"
    lacAvailsStatus.Caption = "Checking"
    DoEvents
    ilSQLLocked = -1
    ilSentToServer = False
    tgRNQifSrchKey1.sType = "T"
    tgRNQifSrchKey1.sNRProcessed = "N"
    tgRNQifSrchKey1.lGenDate = 0
    tgRNQifSrchKey1.lGenTime = 0
    ilRet = btrGetGreaterOrEqual(hgRNQif, tgRNQif, igRNQifRecLen, tgRNQifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record
    Do While (ilRet = BTRV_ERR_NONE) And (tgRNQif.sType = "T") And (tgRNQif.sNRProcessed = "N")
        ilSentToServer = False
        For ilNrf = 0 To UBound(tgServerNRF) - 1 Step 1
            If tgServerNRF(ilNrf).iCode = tgRNQif.iNrfCode Then
                If (sgNetDBPath = "") And (sgRepDBPath = "") Then 'SQL Server workaround
                    ilBypassSendingToServer = False
                    For ilIdx = 0 To UBound(ilNRFFailedCode) - 1 Step 1
                        If ilNRFFailedCode(ilIdx) = tgServerNRF(ilNrf).iCode Then
                            ilBypassSendingToServer = True
                            Exit For
                        End If
                    Next ilIdx
                    If ilBypassSendingToServer Then
                        Exit For
                    End If
                    smFileName1 = Trim$(tgRNQif.sFileName1)
                    smFileName2 = Trim$(tgRNQif.sFileName2)
                    smFileName3 = Trim$(tgRNQif.sFileName3)
                    'igNrfIndex = ilNrf
                    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                        igNrfIndex = ilNrf
                        sgRepDBID = Trim$(tgNrf.sDBID)
                        sgNetDBID = Trim(tgServerNRF(igNrfIndex).sDBID)
                    Else
                        For ilLoop = 0 To UBound(tgServerNRF) - 1 Step 1
                            If tgServerNRF(ilLoop).iCode = tgNrf.iCode Then
                                igNrfIndex = ilLoop
                            End If
                        Next ilLoop
                        sgRepDBID = Trim(tgServerNRF(ilNrf).sDBID)
                        sgNetDBID = Trim$(tgNrf.sDBID)
                    End If
                    ilSvNrfIndex = igNrfIndex
                    Call gTestAccessToWebServer(igNrfIndex)
                    If Not gWebAccessTestedOk Then
                        ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(igNrfIndex).iCode
                        ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                        Exit For
                    End If
                    
                    If (ilSQLLocked <> igNrfIndex) Then
                        If (ilSQLLocked <> -1) Then
                            igNrfIndex = ilSQLLocked
                            ilRet = gEndWebSession("RepNetLink.txt")
                            igNrfIndex = ilSvNrfIndex
                        End If
                        ilRet = gStartWebSession("U", "C", "RepNetLink.txt")
                        If ilRet = 0 Then
                            'we have control of the web
                            ilOKToSend = True
                        Else
                            'we could not get control at the moment
                            ilOKToSend = False
                        End If
                        
                        If Not ilOKToSend Then
                            ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                            ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                            Exit For
                        End If
                    End If
                    
                    ilSQLLocked = ilSvNrfIndex
                    ilSentToServer = True
                    Select Case tgRNQif.sSubType
                        Case "C"    'Contract
                            lacContractStatus.Caption = "Transferring"
                            lacPostedSpotsStatus.Caption = "Idle"
                            lacAvailsStatus.Caption = "Idle"
                            DoEvents
                            slTempPath = sgExportPath & smFileName1
                            If smFileName1 <> "" Then
                                ilRet = gFTPFileToWebServer(slTempPath, smFileName1)
                                If Not ilRet Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Sending File to SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                                If Not gSendCmdToWebServer("ImportData?TN=CHT_Contract_Header", smFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                            End If
'Jeff What happens if we import the data again
                            slTempPath = sgExportPath & smFileName2
                            If smFileName2 <> "" Then
                                ilRet = gFTPFileToWebServer(slTempPath, smFileName2)
                                If Not ilRet Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Sending File to SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                                If Not gSendCmdToWebServer("ImportData?TN=CLT_Contract_Line", smFileName2) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                            End If
                            gCreateAlert "M||" & "|" & tgRNQif.lCntrNo & "|" & "" & "||" & "Forward to Network " & sgNetDBID, True
                        Case "S"    'Posted spots
                            lacContractStatus.Caption = "Idle"
                            lacPostedSpotsStatus.Caption = "Transferring"
                            lacAvailsStatus.Caption = "Idle"
                            DoEvents
                            slTempPath = sgExportPath & smFileName1
                            If smFileName1 <> "" Then
                                ilRet = gFTPFileToWebServer(slTempPath, smFileName1)
                                If Not ilRet Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Sending File to SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                                If Not gSendCmdToWebServer("ImportData?TN=PST_Posted_Spots", smFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                            End If
                        Case "M"    'Message
                            slTempPath = sgExportPath & smFileName1
                            If smFileName1 <> "" Then
                                ilRet = gFTPFileToWebServer(slTempPath, smFileName1)
                                If Not ilRet Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Sending File to SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                                If Not gSendCmdToWebServer("ImportData?TN=MIT_Message_Info", smFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                            End If
                        Case "A"    '6 month avails
                            lacContractStatus.Caption = "Idle"
                            lacPostedSpotsStatus.Caption = "Idle"
                            lacAvailsStatus.Caption = "Transferring"
                            slTempPath = sgExportPath & smFileName1
                            If smFileName1 <> "" Then
                                ilRet = gFTPFileToWebServer(slTempPath, smFileName1)
                                If Not ilRet Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Sending File to SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                                If Not gSendCmdToWebServer("ImportData?TN=AMT_Avails_By_Months", smFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    ilNRFFailedCode(UBound(ilNRFFailedCode)) = tgServerNRF(ilNrf).iCode
                                    ReDim Preserve ilNRFFailedCode(0 To UBound(ilNRFFailedCode) + 1) As Integer
                                    ilSentToServer = False
                                    Exit For
                                End If
                            End If
                    End Select
                Else
                    For ilLoop = 0 To UBound(tgServerNRF) - 1 Step 1
                        If (Trim$(sgXFerDBID) = "") Or (StrComp(Trim$(sgXFerDBID), Trim$(tgServerNRF(ilLoop).sDBID), vbTextCompare) = 0) Then
                            ilSentToServer = True
                            FileCopy sgExportPath & tgRNQif.sFileName1, sgImportPath & tgRNQif.sFileName1
                            If Trim$(tgRNQif.sFileName2) <> "" Then
                                FileCopy sgExportPath & tgRNQif.sFileName2, sgImportPath & tgRNQif.sFileName2
                            End If
                            hlQif = CBtrvTable(TWOHANDLES)    'CBtrvObj()
                            Select Case tgRNQif.sSubType
                                Case "C"    'Contract
                                    lacContractStatus.Caption = "Transferring to Network Database"
                                    ilRet = btrOpen(hlQif, "", sgNetDBPath & "Qif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                                Case "S"    'Posted spots
                                    lacPostedSpotsStatus.Caption = "Transferring to Rep Database"
                                    ilRet = btrOpen(hlQif, "", sgRepDBPath & "Qif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                                Case "M"    'Message
                                    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                                        ilRet = btrOpen(hlQif, "", sgNetDBPath & "Qif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                                    ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                                        ilRet = btrOpen(hlQif, "", sgRepDBPath & "Qif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                                    End If
                                Case "A"    '6 month avails
                                    lacAvailsStatus.Caption = "Transferring to Rep Database"
                                    ilRet = btrOpen(hlQif, "", sgRepDBPath & "Qif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                            End Select
                            If ilRet = BTRV_ERR_NONE Then
                                tlQif.lCode = 0
                                tlQif.sNRProcessed = "N"
                                tlQif.sType = "S"
                                tlQif.sSubType = tgRNQif.sSubType
                                tlQif.sFileName1 = tgRNQif.sFileName1
                                tlQif.sFileName2 = tgRNQif.sFileName2
                                tlQif.sFileName3 = ""
                                tlQif.lCntrNo = tgRNQif.lCntrNo
                                tlQif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                                tlQif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                                Select Case Left$(gGetLocalTZName(), 1)
                                    Case "E"
                                    Case "C"
                                        tlQif.lGenTime = tlQif.lGenTime + 3600
                                    Case "M"
                                        tlQif.lGenTime = tgRNQif.lGenTime + 7200
                                    Case "P"
                                        tlQif.lGenTime = tlQif.lGenTime + 10800
                                End Select
                                gPackDate "", tlQif.iProcessedDate(0), tlQif.iProcessedDate(1)
                                gPackTime "", tlQif.iProcessedTime(0), tlQif.iProcessedTime(1)
                                tlQif.iNrfCode = 0  'Since this is placed on the other database it is not required
                                tlQif.sUnused = ""
                                ilRet = btrInsert(hlQif, tlQif, igRNQifRecLen, INDEXKEY0)
                                Sleep (1500)
                            End If
                            btrDestroy hlQif
                            Exit For
                        End If
                    Next ilLoop
                End If
                Exit For
            End If
        Next ilNrf
        If ilSentToServer Then
            If (sgNetDBPath = "") And (sgRepDBPath = "") Then
                'Doug
                'Create SQL Server Poll Table Record
                'we have sent the records now, create the PDT record to tell the other side that we
                'have information that needs to come in.
    '            llDate = gDateValue(Format$(gNow(), "m/d/yy"))
    '            llETime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
    '
    '            Select Case Left$(gGetLocalTZName(), 1)
    '                Case "E"
    '                Case "C"
    '                    llETime = llETime + 3600
    '                Case "M"
    '                    llETime = llETime + 7200
    '                Case "P"
    '                    llETime = llETime + 10800
    '            End Select
                
                llDate = tgRNQif.lGenDate
                llETime = tgRNQif.lGenTime
                
                llRet = gExecWebSQL(slDataArray, "Select GetDate() as ServerDateTime", True)
                If llRet = -1 Then  'Error condition, unable to connect to the web
                    ilSentToServer = False
                End If
                'If llRet = 0 Then   'No Data
                '    ilSentToServer = False
                'End If

                If ilSentToServer Then
                    slDateTime = ""
                    On Error Resume Next
                    slDateTime = gGetDataNoQuotes(slDataArray(0))
                    If slDateTime <> "" Then
                        If tgNrf.sType = "R" Then
                            slToRepOrNet = "N"
                        Else
                            slToRepOrNet = "R"
                        End If
                        SQLQuery = "Insert into PDT_Poll_Data (pdtRepDBID, pdtNetDBID, pdtProcessed, pdtGenDate, pdtGenTime, pdtType, pdtToRepOrNet, pdtContractNo, pdtFileName1, pdtFileName2, pdtFileName3, pdtDateTimeEntered) "
                        SQLQuery = SQLQuery & "Values('" & sgRepDBID & "', '" & sgNetDBID & "', '" & "N" & "', " & llDate & ", " & llETime & ", '" & Trim(tgRNQif.sSubType) & "', '" & slToRepOrNet & "', " & CLng(Trim$(Str$(tgRNQif.lCntrNo))) & ", '" & smFileName1 & "', '" & smFileName2 & "', '" & smFileName3 & "', '" & slDateTime & "')"
                        llRet = gExecWebSQLWithRowsEffected(SQLQuery)
                        If llRet = -1 Then  'Error condition, unable to connect to the web
                            ilSentToServer = False
                        End If
                    Else
                        ilSentToServer = False
                    End If
                End If

            End If
            
            If ilSentToServer Then
                'Doug clear the QIF Record
                Do
                    tgRNQifSrchKey0.lCode = tgRNQif.lCode
                    ilRet = btrGetEqual(hgRNQif, tgRNQif, igRNQifRecLen, tgRNQifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        tgRNQif.sNRProcessed = "Y"
                        gPackDate Format$(gNow(), "m/d/yy"), tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
                        gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
                        ilRet = btrUpdate(hgRNQif, tgRNQif, igRNQifRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
            End If
            
            'Get Next recorded to be processed
            tgRNQifSrchKey1.sType = "T"
            tgRNQifSrchKey1.sNRProcessed = "N"
            tgRNQifSrchKey1.lGenDate = 0
            tgRNQifSrchKey1.lGenTime = 0
            ilRet = btrGetGreaterOrEqual(hgRNQif, tgRNQif, igRNQifRecLen, tgRNQifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record

        Else
            ilRet = btrGetNext(hgRNQif, tgRNQif, igRNQifRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        End If
    Loop
    If ilSQLLocked <> -1 Then
        igNrfIndex = ilSQLLocked
        ilRet = gEndWebSession("RepNetLink.txt")
    End If
    lacContractStatus.Caption = "Idle"
    lacPostedSpotsStatus.Caption = "Idle"
    lacAvailsStatus.Caption = "Idle"
    DoEvents
End Sub

Private Sub mServerToText()
    Dim ilNrf As Integer
    Dim ilSQLLocked As Integer
    Dim ilRet As Integer
    Dim slDataArray() As String
    Dim llRet As Long
    Dim llIdx As Long
    Dim slFileName1 As String
    Dim slFileName2 As String
    Dim slFileName3 As String
    Dim ilOKToSend As Integer
    Dim llRecCount As Long
    Dim llDate As Long
    Dim llTime As Long
    Dim slTempPath As String
    Dim ilLoop As Integer
    Dim ilSvNrfIndex As Integer
    Dim llCntrNo As Long
        
    If (sgNetDBPath <> "") Or (sgRepDBPath <> "") Then
        Exit Sub
    End If
    lacContractStatus.Caption = "Checking"
    lacPostedSpotsStatus.Caption = "Checking"
    lacAvailsStatus.Caption = "Checking"
    DoEvents
    ilSQLLocked = -1
    For ilNrf = LBound(tgServerNRF) To UBound(tgServerNRF) - 1 Step 1
        If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Or (((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) And (tgServerNRF(ilNrf).sType = "N")) Then
            Do
                'Are there any records that I need to retrieve.  Look at the PDF table on SQL Server.  It's
                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                    igNrfIndex = ilNrf
                Else
                    For ilLoop = 0 To UBound(tgServerNRF) - 1 Step 1
                        If tgServerNRF(ilLoop).iCode = tgNrf.iCode Then
                            igNrfIndex = ilLoop
                        End If
                    Next ilLoop
                End If
                ilSvNrfIndex = igNrfIndex
                Call gTestAccessToWebServer(igNrfIndex)
                If Not gWebAccessTestedOk Then
                    Exit Do
                End If
                
                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                    llRet = gExecWebSQL(slDataArray, "Select Count(pdtRepDBID) from pdt_poll_data where pdtRepDBID = " & "'" & tgNrf.sDBID & "'" & " And pdtProcessed = " & "'" & "N" & "'" & " And pdtToRepOrNet = " & "'" & "R" & "'", True)
                Else
                    llRet = gExecWebSQL(slDataArray, "Select Count(pdtNetDBID) from pdt_poll_data where pdtNetDBID = " & "'" & tgNrf.sDBID & "'" & " And pdtProcessed = " & "'" & "N" & "'" & " And pdtToRepOrNet = " & "'" & "N" & "'", True)
                End If
                
                If llRet = -1 Then  'Error condition, unable to connect to the web
                    Exit Do
                End If
                'If (llRet = 0) Then 'No data
                '    Exit Do
                'End If
                
                llRecCount = 0
                On Error Resume Next
                llRecCount = CLng(slDataArray(0))
                On Error GoTo 0
                If llRecCount = 0 Then
                    Exit Do  'Nothing to do
                End If
                
                'Doug
                'Check if allowed to retrieve record from server.  If not exit this procedure
                
                If (ilSQLLocked <> igNrfIndex) Then
                    If (ilSQLLocked <> -1) Then
                        igNrfIndex = ilSQLLocked
                        ilRet = gEndWebSession("RepNetLink.txt")
                        igNrfIndex = ilSvNrfIndex
                    End If
                    ilRet = gStartWebSession("U", "C", "RepNetLink.txt")
                    If ilRet = 0 Then
                        'we have control of the web
                        ilOKToSend = True
                    Else
                        'we could not get control at the moment
                        ilOKToSend = False
                    End If
                    
                    If Not ilOKToSend Then
                        Exit Do 'Could not obtain a lock
                    End If
                End If
                
                ilSQLLocked = ilSvNrfIndex
                For llIdx = 0 To llRecCount - 1 Step 1
                    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                        llRet = gExecWebSQL(slDataArray, "Select * from pdt_poll_data where pdtRepDBID = " & "'" & tgNrf.sDBID & "'" & " And pdtProcessed = " & "'" & "N" & "'" & " And pdtToRepOrNet = " & "'" & "R" & "'" & " Order by pdtGenDate, pdtGenTime", True)
                    Else
                        llRet = gExecWebSQL(slDataArray, "Select * from pdt_poll_data where pdtNetDBID = " & "'" & tgNrf.sDBID & "'" & " And pdtProcessed = " & "'" & "N" & "'" & " And pdtToRepOrNet = " & "'" & "N" & "'" & " Order by pdtGenDate, pdtGenTime", True)
                    End If
                    
                    If llRet = -1 Then  'Error condition, unable to connect to the web
                        Screen.MousePointer = vbDefault
                        Exit Do
                    End If
                    
                    llDate = CLng(slDataArray(4))
                    llTime = CLng(slDataArray(5))
                    slFileName1 = ""
                    slFileName2 = ""
                    slFileName3 = ""
                    llCntrNo = 0
                    'Doug
                    'Initiate Retrieval of record and transfer of text file from SQL Server
                    'Create QIF_Queue_Info record
                    Select Case Trim$(slDataArray(6)) 'look a the PDT type
                        Case "C"    'Contract
                            lacContractStatus.Caption = "Retrieving"
                            lacPostedSpotsStatus.Caption = "Idle"
                            lacAvailsStatus.Caption = "Idle"
                            slFileName1 = Trim$(slDataArray(9))
                            If slFileName1 <> "" Then
                                If Not gSendCmdToWebServer("ExportData?SQL=Select * from CHT_Contract_Header where chtContractNo = " & slDataArray(8) & " And chtGenDate = " & llDate & " And chtGenTime = " & llTime, slFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                                slTempPath = sgImportPath
                                ilRet = gFTPFileFromWebServer(slTempPath, slFileName1)
                                If Not ilRet Then
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Retrieving File from SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                            End If
                            slFileName2 = Trim$(slDataArray(10))
                            If slFileName2 <> "" Then
                                If Not gSendCmdToWebServer("ExportData?SQL=Select * from CLT_Contract_Line  Where cltContractNo = " & slDataArray(8) & " And cltGenDate = " & llDate & " And cltGenTime = " & llTime, slFileName2) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                                slTempPath = sgImportPath
                                ilRet = gFTPFileFromWebServer(slTempPath, slFileName2)
                                If Not ilRet Then
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Retrieving File from SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If

                            End If
                            llCntrNo = Val(slDataArray(8))
                            DoEvents
                        
                        Case "S"    'Posted spots
                            lacContractStatus.Caption = "Idle"
                            lacPostedSpotsStatus.Caption = "Retrieving"
                            lacAvailsStatus.Caption = "Idle"
                            slFileName1 = Trim$(slDataArray(9))
                            If slFileName1 <> "" Then
                                If Not gSendCmdToWebServer("ExportData?SQL=Select * from PST_Posted_Spots where pstRepDBID = " & "'" & tgNrf.sDBID & "'" & " And pstGenDate = " & llDate & " And pstGenTime = " & llTime, slFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                                slTempPath = sgImportPath
                                ilRet = gFTPFileFromWebServer(slTempPath, slFileName1)
                                If Not ilRet Then
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Retrieving File from SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                            End If
                            llCntrNo = Val(slDataArray(8))  'Invoice end date
                            DoEvents
                        Case "M"    'Message
                            slFileName1 = Trim$(slDataArray(9))
                            If slFileName1 <> "" Then
                                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                                    If Not gSendCmdToWebServer("ExportData?SQL=Select * from MIT_Message_Info where mitRepDBID = " & "'" & tgNrf.sDBID & "'" & " And mitGenDate = " & llDate & " And mitGenTime = " & llTime, slFileName1) Then
                                        gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                        Exit Do
                                    End If
                                Else
                                    If Not gSendCmdToWebServer("ExportData?SQL=Select * from MIT_Message_Info where mitNetDBID = " & "'" & tgNrf.sDBID & "'" & " And mitGenDate = " & llDate & " And mitGenTime = " & llTime, slFileName1) Then
                                        Screen.MousePointer = vbDefault
                                        gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                        Exit Do
                                    End If
                                End If
                                slTempPath = sgImportPath
                                ilRet = gFTPFileFromWebServer(slTempPath, slFileName1)
                                If Not ilRet Then
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Retrieving File from SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                            End If
                        Case "A"    '6 month avails
                            slFileName1 = Trim$(slDataArray(9))
                            lacContractStatus.Caption = "Idle"
                            lacPostedSpotsStatus.Caption = "Idle"
                            lacAvailsStatus.Caption = "Retrieving"
                            If slFileName1 <> "" Then
                                If Not gSendCmdToWebServer("ExportData?SQL=Select * from AMT_Avails_By_Months where amtRepDBID = " & "'" & tgNrf.sDBID & "'" & " And amtGenDate = " & llDate & " And amtGenTime = " & llTime, slFileName1) Then
                                    Screen.MousePointer = vbDefault
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Request to Execute Command on SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                                slTempPath = sgImportPath
                                ilRet = gFTPFileFromWebServer(slTempPath, slFileName1)
                                If Not ilRet Then
                                    gCreateAlert "M||" & "|" & "|" & "" & "||" & "Retrieving File from SQL Server Failed, Error Posted into RepNetLink.Txt", True
                                    Exit Do
                                End If
                            End If
                            DoEvents
                    End Select
                    tgRNQif.lCode = 0
                    tgRNQif.sNRProcessed = "N"
                    tgRNQif.sType = "S"
                    'Doug
                    'Need to get subtype from the Poll table -
                    tgRNQif.sSubType = Trim$(slDataArray(6)) 'look a the PDTtype record
                    'Doug
                    'Need to get file name from the Poll table
                    'we are getting the SQL database and send it down for us
                    tgRNQif.sFileName1 = slFileName1 'look a the PDT filename 1
                    tgRNQif.sFileName2 = slFileName2 'look a the PDT filename 2
                    tgRNQif.sFileName3 = slFileName3  'look a the PDT filename 3
                    tgRNQif.lCntrNo = llCntrNo
                    'tgRNQif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                    'tgRNQif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                    'Select Case Left$(gGetLocalTZName(), 1)
                    '    Case "E"
                    '    Case "C"
                    '        tgRNQif.lGenTime = tgRNQif.lGenTime + 3600
                    '    Case "M"
                    '        tgRNQif.lGenTime = tgRNQif.lGenTime + 7200
                    '    Case "P"
                    '        tgRNQif.lGenTime = tgRNQif.lGenTime + 10800
                    'End Select
                    tgRNQif.lGenDate = llDate
                    tgRNQif.lGenTime = llTime
                    tgRNQif.iNrfCode = 0
                    gPackDate "", tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
                    gPackTime "", tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
                    tgRNQif.sUnused = ""
                    ilRet = btrInsert(hgRNQif, tgRNQif, igRNQifRecLen, INDEXKEY0)
                    'Mark the Pdt record as processed
                    llRet = gExecWebSQL(slDataArray, "Update pdt_poll_data Set pdtProcessed = " & "'Y'" & " Where UKey = " & CLng(slDataArray(0)), True)
                    If llRet = -1 Then  'Error condition, unable to connect to the web
                        Screen.MousePointer = vbDefault
                        Exit Do
                    End If
                Next llIdx
            Loop
        End If
    Next ilNrf
    If ilSQLLocked <> -1 Then
        igNrfIndex = ilSQLLocked
        ilRet = gEndWebSession("RepNetLink.txt")
    End If
    lacContractStatus.Caption = "Idle"
    lacPostedSpotsStatus.Caption = "Idle"
    lacAvailsStatus.Caption = "Idle"
    DoEvents
End Sub

Private Sub mServerTextToTraffic()
    Dim ilRet As Integer
    
    lacContractStatus.Caption = "Checking"
    lacPostedSpotsStatus.Caption = "Checking"
    lacAvailsStatus.Caption = "Checking"
    DoEvents
    tgRNQifSrchKey1.sType = "S"
    tgRNQifSrchKey1.sNRProcessed = "N"
    tgRNQifSrchKey1.lGenDate = 0
    tgRNQifSrchKey1.lGenTime = 0
    ilRet = btrGetGreaterOrEqual(hgRNQif, tgRNQif, igRNQifRecLen, tgRNQifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record
    Do While (ilRet = BTRV_ERR_NONE) And (tgRNQif.sType = "S") And (tgRNQif.sNRProcessed = "N")
        Select Case tgRNQif.sSubType
            Case "C"    'Contract
                lacContractStatus.Caption = "Creating"
                lacPostedSpotsStatus.Caption = "Idle"
                lacAvailsStatus.Caption = "Idle"
                DoEvents
                ilRet = gTextToContract(tgRNQif)
            Case "S"    'Posted spots
                lacContractStatus.Caption = "Idle"
                lacPostedSpotsStatus.Caption = "Creating"
                lacAvailsStatus.Caption = "Idle"
                DoEvents
                ilRet = gTextToPostedSpots(tgRNQif)
            Case "M"    'Message
                ilRet = gTextToMessage(tgRNQif)
            Case "A"    '6 month avails
                lacContractStatus.Caption = "Idle"
                lacPostedSpotsStatus.Caption = "Idle"
                lacAvailsStatus.Caption = "Creating"
                DoEvents
                ilRet = gTextToAvailSum(tgRNQif)
        End Select
        
        'Clear QIF Processed flag
        Do
            tgRNQifSrchKey0.lCode = tgRNQif.lCode
            ilRet = btrGetEqual(hgRNQif, tgRNQif, igRNQifRecLen, tgRNQifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tgRNQif.sNRProcessed = "Y"
                gPackDate Format$(gNow(), "m/d/yy"), tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
                gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
                ilRet = btrUpdate(hgRNQif, tgRNQif, igRNQifRecLen)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        
        'Get Next recorded to be processed
        tgRNQifSrchKey1.sType = "S"
        tgRNQifSrchKey1.sNRProcessed = "N"
        tgRNQifSrchKey1.lGenDate = 0
        tgRNQifSrchKey1.lGenTime = 0
        ilRet = btrGetGreaterOrEqual(hgRNQif, tgRNQif, igRNQifRecLen, tgRNQifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record
        DoEvents
    Loop
    lacContractStatus.Caption = "Idle"
    lacPostedSpotsStatus.Caption = "Idle"
    lacAvailsStatus.Caption = "Idle"
    DoEvents
End Sub



Private Function mRNNrfPop() As Integer
    Dim ilRet As Integer
    Dim tlNrf As NRF
    
    
    mRNNrfPop = True
    ReDim tgServerNRF(0 To 0) As NRF
    ilRet = btrGetFirst(hgRNNrf, tlNrf, igRNNrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
            If tlNrf.sType = "R" Then
                tgNrf = tlNrf
            Else
                If tlNrf.sInsertionType = "C" Then
                    tgServerNRF(UBound(tgServerNRF)) = tlNrf
                    ReDim Preserve tgServerNRF(0 To UBound(tgServerNRF) + 1) As NRF
                End If
            End If
        ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
            If tlNrf.sType = "N" Then
                tgNrf = tlNrf
            End If
            tgServerNRF(UBound(tgServerNRF)) = tlNrf
            ReDim Preserve tgServerNRF(0 To UBound(tgServerNRF) + 1) As NRF
        End If
        ilRet = btrGetNext(hgRNNrf, tlNrf, igRNNrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If UBound(tgServerNRF) > 0 Then
        ReDim gURLHasAccess(0 To UBound(tgServerNRF) - 1)
    End If
End Function

Private Function mContractToText() As Integer
    '
    'tgRNGif(I):  Contract Gather Info record
    '
    Dim ilFound As Integer
    Dim ilLastNrfCode As Integer
    Dim ilRet As Integer
    Dim ilVsf As Integer
    Dim ilVef As Integer
    Dim ilNrf As Integer
    Dim ilNrfIndex As Integer
    Dim ilProc As Integer
    Dim slGenDate As String
    Dim slGenTime As String
    Dim llGenDate As Long
    Dim llGenTime As Long
    Dim slNetDBID As String
    Dim slNetSlspFirst As String
    Dim slNetSlspLast As String
    Dim slTransmitPrice As String
    Dim slFileName As String
    Dim slHeader As String
    Dim slLine As String
    Dim slLineComment As String
    Dim slLineShowOn As String
    Dim slFlight As String
    Dim slTitle As String
    Dim tlVef As VEF
    Dim ilCreateQif As Integer
    ReDim ilVefCode(0 To 0) As Integer
    
    'Get latest version
    ilFound = False
    tgRNChfSrchKey1.lCntrNo = tgRNGif.lCntrNo
    tgRNChfSrchKey1.iCntRevNo = 32000
    tgRNChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hgRNChf, tgRNChf, igRNChfRecLen, tgRNChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tgRNChf.lCntrNo = tgRNGif.lCntrNo)
        If (tgRNChf.sStatus = "O") Then
            ilFound = True
            Exit Do
        End If
        ilRet = btrGetNext(hgRNChf, tgRNChf, igRNChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If Not ilFound Then
        gCreateAlert "C|E|" & "|" & Trim$(Str$(tgRNGif.lCntrNo)) & "|" & "||" & "Contract Not Found", True
        mContractToText = False
        Exit Function
    End If
    If tgRNChf.lVefCode < 0 Then
        tgRNVsfSrchKey0.lCode = -tgRNChf.lVefCode
        ilRet = btrGetEqual(hgRNVsf, tgRNVsf(0), igRNVsfRecLen, tgRNVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            For ilVsf = LBound(tgRNVsf(0).iFSCode) To UBound(tgRNVsf(0).iFSCode) Step 1
                If tgRNVsf(0).iFSCode(ilVsf) > 0 Then
                    ilVefCode(UBound(ilVefCode)) = tgRNVsf(0).iFSCode(ilVsf)
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
            Next ilVsf
            tgRNVsfSrchKey0.lCode = tgRNVsf(0).lLkVsfCode
            ilRet = btrGetEqual(hgRNVsf, tgRNVsf(0), igRNVsfRecLen, tgRNVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        ilVefCode(0) = tgRNChf.lVefCode
        ReDim Preserve ilVefCode(0 To 1) As Integer
    End If
    ilLastNrfCode = -1
    For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
        If ilVefCode(ilVef) > 0 Then
            ilFound = False
            tgRNVefSrchKey0.iCode = ilVefCode(ilVef)
            ilRet = btrGetEqual(hgRNVef, tgRNVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                gCreateAlert "C|E|" & "|" & Trim$(Str$(tgRNChf.lCntrNo)) & "|" & "||" & "Accessing Vehicle(vefCode = " & Trim$(Str$(ilVefCode(ilVef))) & ", Pervasive Error # " & ilRet, True
                mContractToText = False
                Exit Function
            End If
            If tgRNVef.iNrfCode > 0 Then
                For ilNrf = 0 To UBound(tgServerNRF) - 1 Step 1
                    If tgRNVef.iNrfCode = tgServerNRF(ilNrf).iCode Then
                        ilFound = True
                        ilNrfIndex = ilNrf
                        Exit For
                    End If
                Next ilNrf
            End If
            If ilFound Then
                slNetDBID = Trim$(tgServerNRF(ilNrfIndex).sDBID)
                slNetSlspFirst = Trim$(tgServerNRF(ilNrfIndex).sSlspFirstName)
                slNetSlspLast = Trim$(tgServerNRF(ilNrfIndex).sSlspLastName)
                slTransmitPrice = tgServerNRF(ilNrfIndex).sTransmitPrice
                ilCreateQif = True
            'Else
            '    slNetDBID = "None"
            '    slNetSlspFirst = ""
            '    slNetSlspLast = ""
            '    ilCreateQif = False
            'End If
                slFileName = slNetDBID & "_" & Trim$(tgNrf.sDBID)
                llGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                llGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                Select Case Left$(gGetLocalTZName(), 1)
                    Case "E"
                    Case "C"
                        llGenTime = llGenTime + 3600
                    Case "M"
                        llGenTime = llGenTime + 7200
                    Case "P"
                        llGenTime = llGenTime + 10800
                End Select
                slFileName = slFileName & "_" & Trim$(Str$(tgRNChf.lCntrNo)) & "_Header_" & Format$(llGenDate, "YY-MM-DD") & "_" & Trim$(Str$(llGenTime)) & ".csv"
                smFileName1 = slFileName
                ilRet = mOpenTextForWrite(hmTextChf, slFileName)
                If Not ilRet Then
                    mContractToText = False
                    Exit Function
                End If
                slTitle = "chtRepDBID,chtNetDBID,chtGenDate,chtGenTime,chtSeqNo,chtContractType,chtContractNo,chtExtRevNo,chtIntVerNo,chtProposalVersion,chtAdvertiserCode,chtAdvertiserName,chtAdvertiserAbbr,chtAdvertiserType,chtAdvertiserAddrID,chtAdvertiserCntrAddr_1,chtAdvertiserCntrAddr_2,chtAdvertiserCntrAddr_3,chtAdvertiserBillAddr_1,chtAdvertiserBillAddr_2,chtAdvertiserBillAddr_3,chtProductName,chtAgencyCode,chtAgencyName,chtAgencyAbbr,chtAgencyCityID,chtAgencyCntrAddr_1,chtAgencyCntrAddr_2,chtAgencyCntrAddr_3,chtAgencyBillAddr_1,chtAgencyBillAddr_2,chtAgencyBillAddr_3,chtAgencyPctComm,chtProductProtection_1,chtProductProtection_2,chtPctTrade,chtDemoGraphicName_1,chtDemoGraphicName_2,chtDemoGraphicName_3,chtDemoGraphicName_4,chtSalespersonFirstName,chtSalespersonLastName,chtBuyerName,chtBuyerPhoneNo,chtInternalComment,chtShowInternalOn,chtChangeReasonComment,chtShowChangeOn,chtCancelComment,chtShowCancelOn,chtOtherComment,chtShowOtherOn"
                Print #hmTextChf, slTitle
                slFileName = slNetDBID & "_" & Trim$(tgNrf.sDBID)
                slFileName = slFileName & "_" & Trim$(Str$(tgRNChf.lCntrNo)) & "_Lines_" & Format$(llGenDate, "YY-MM-DD") & "_" & Trim$(Str$(llGenTime)) & ".csv"
                smFileName2 = slFileName
                ilRet = mOpenTextForWrite(hmTextClf, slFileName)
                If Not ilRet Then
                    Close hmTextChf
                    mContractToText = False
                    Exit Function
                End If
                slTitle = "cltRepDBID,cltNetDBID,cltGenDate,cltGenTime,cltContractNo,cltIntVerNo,cltProposalVersion,cltLineNo,cltVehicleName,cltNetVehicleCode,cltDaypartName,cltNetDaypartCode,cltDPInOut,cltDPAvailName,cltDPDay_1,cltDPStartTime_1,cltDPEndTime_1,cltDPSpotPct_1,cltDPDay_2,cltDPStartTime_2,cltDPEndTime_2,cltDPSpotPct_2,cltSpotLength,cltBBOpenLength,cltBBCloseLength,cltSoloAvail,cltOv2DefinedBits,clt1stPosition,cltPctAllocation,cltAcquisitionCost,cltLineStartDate,cltLineEndDate,cltOverrideStartTime,cltOverrideEndTime,cltPreferredStartTime,cltPreferredEndTime,cltPreferredDays,cltPreemptible,cltLiveCopy,cltAdvtSepFlag,cltPriority,cltHideCBS,cltSeqNo,cltFlightStartDate,cltFlightEndDate,cltDailyOrWeeklyBuy,cltSpotsPerWeek,cltSpots_1_Mo,cltSpots_2_Tu,cltSpots_3_We,cltSpots_4_Th,cltSpots_5_Fr,cltSpots_6_Sa,cltSpots_7_Su,cltPriceType,cltActualPrice,cltComment,cltShowCommentOn"
                Print #hmTextClf, slTitle
                'Create Contract header record
                sgRepDBID = Trim$(tgNrf.sDBID)
                sgNetDBID = Trim$(slNetDBID)
                slHeader = gCreateContractHeaderRec(Trim$(tgNrf.sDBID), Trim$(slNetDBID), llGenDate, llGenTime, lmSeqNo, tgRNChf, slNetSlspFirst, slNetSlspLast)
                If slHeader <> "" Then
                    Print #hmTextChf, slHeader
                    lmSeqNo = lmSeqNo + 1
                    ilProc = ilVef
                    Do
                        tgRNClfSrchKey1.lChfCode = tgRNChf.lCode
                        tgRNClfSrchKey1.iVefCode = ilVefCode(ilProc)
                        ilRet = btrGetEqual(hgRNClf, tgRNClf(0), igRNClfRecLen, tgRNClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                        Do While (ilRet = BTRV_ERR_NONE) And (tgRNClf(0).lChfCode = tgRNChf.lCode) And (tgRNClf(0).iVefCode = ilVefCode(ilProc))
                            'Create Contract line record (bypass package vehicles:  Send only standard and hidden lines)
                            If (tgRNClf(0).sType <> "O") And (tgRNClf(0).sType <> "A") And (tgRNClf(0).sType <> "E") Then
                                slLine = gCreateContractLineRec(Trim$(tgNrf.sDBID), Trim$(slNetDBID), llGenDate, llGenTime, tgRNChf, tgRNClf(0))
                                slLineComment = gGetComment(tgRNClf(0).lCxfCode, slLineShowOn)
                                tgRNCffSrchKey0.lChfCode = tgRNClf(0).lChfCode
                                tgRNCffSrchKey0.iClfLine = tgRNClf(0).iLine
                                tgRNCffSrchKey0.iCntRevNo = tgRNClf(0).iCntRevNo
                                tgRNCffSrchKey0.iPropVer = tgRNClf(0).iPropVer
                                tgRNCffSrchKey0.iStartDate(0) = 0
                                tgRNCffSrchKey0.iStartDate(1) = 0
                                ilRet = btrGetGreaterOrEqual(hgRNCff, tgRNCff(0), igRNCffRecLen, tgRNCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tgRNCff(0).lChfCode = tgRNClf(0).lChfCode) And (tgRNCff(0).iClfLine = tgRNClf(0).iLine) And (tgRNCff(0).iCntRevNo = tgRNClf(0).iCntRevNo) And (tgRNCff(0).iPropVer = tgRNClf(0).iPropVer)
                                    'Add flight to line record
                                    If slTransmitPrice = "N" Then
                                        tgRNCff(0).lActPrice = 0
                                        tgRNCff(0).lPropPrice = 0
                                    End If
                                    slFlight = gCreateContractFlightRec(tgRNCff(0))
                                    Print #hmTextClf, slLine & Trim$(Str$(lmSeqNo)) & ", " & slFlight & gAddQuotesRemoveChars(slLineComment, True) & ", " & gAddQuotesRemoveChars(slLineShowOn, True)
                                    slLineComment = ""
                                    slLineShowOn = ""
                                    lmSeqNo = lmSeqNo + 1
                                    ilRet = btrGetNext(hgRNCff, tgRNCff(0), igRNCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                            ilRet = btrGetNext(hgRNClf, tgRNClf(0), igRNClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        ilVefCode(ilProc) = -1
                        ilProc = ilProc + 1
                        Do While ilProc < UBound(ilVefCode)
                            If ilVefCode(ilProc) > 0 Then
                                tgRNVefSrchKey0.iCode = ilVefCode(ilProc)
                                ilRet = btrGetEqual(hgRNVef, tlVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    Close hmTextChf
                                    Close hmTextClf
                                    If ilRet >= 30000 Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    gCreateAlert "C|E|" & "|" & Trim$(Str$(tgRNChf.lCntrNo)) & "|" & "||" & "Accessing Vehicle(vefCode = " & Trim$(Str$(ilVefCode(ilProc))) & ", Pervasive Error # " & ilRet, True
                                    mContractToText = False
                                    Exit Function
                                End If
                                If tgRNVef.iNrfCode = tlVef.iNrfCode Then
                                    Exit Do
                                End If
                            End If
                            ilProc = ilProc + 1
                        Loop
                    Loop While ilProc < UBound(ilVefCode)
                Else
                    Close hmTextChf
                    Close hmTextClf
                    mContractToText = False
                    Exit Function
                End If
                Close hmTextChf
                Close hmTextClf
                If ilCreateQif Then
                    'Create QIF_Queue_Info record
                    tgRNQif.lCode = 0
                    tgRNQif.sNRProcessed = "N"
                    tgRNQif.sType = "T"
                    tgRNQif.sSubType = tgRNGif.sType
                    tgRNQif.sFileName1 = smFileName1
                    tgRNQif.sFileName2 = smFileName2
                    tgRNQif.sFileName3 = ""
                    tgRNQif.lCntrNo = tgRNChf.lCntrNo
                    tgRNQif.lGenDate = llGenDate      'gDateValue(Format$(llGenDate, "m/d/yy"))
                    tgRNQif.lGenTime = llGenTime      'gTimeToLong(Format$(llGenTime, "h:mm:ssAM/PM"), False)
                    'Select Case Left$(gGetLocalTZName(), 1)
                    '    Case "E"
                    '    Case "C"
                    '        tgRNQif.lGenTime = tgRNQif.lGenTime + 3600
                    '    Case "M"
                    '        tgRNQif.lGenTime = tgRNQif.lGenTime + 7200
                    '    Case "P"
                    '        tgRNQif.lGenTime = tgRNQif.lGenTime + 10800
                    'End Select
                    tgRNQif.iNrfCode = tgRNVef.iNrfCode
                    gPackDate "", tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
                    gPackTime "", tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
                    tgRNQif.sUnused = ""
                    ilRet = btrInsert(hgRNQif, tgRNQif, igRNQifRecLen, INDEXKEY0)
                    Sleep (1500)
                End If
            Else
                Sleep (1500)
            End If
        End If
    Next ilVef
    mContractToText = True
End Function

Private Function mOpenTextForWrite(hlText As Integer, slFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenTextForWriteErr:
    slToFile = sgExportPath & slFileName
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenTextForWriteErr:
        'hlText = FreeFile
        'Open slToFile For Output As hlText
        ilRet = gFileOpen(slToFile, "Output", hlText)
        If ilRet <> 0 Then
            gCreateAlert "O|E|" & "|" & "|" & "||" & "Open Text file " & slFileName & " for Write failed, error #: " & ilRet, True
            mOpenTextForWrite = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenTextForWriteErr:
        'hlText = FreeFile
        'Open slToFile For Output As hlText
        ilRet = gFileOpen(slToFile, "Output", hlText)
        If ilRet <> 0 Then
            gCreateAlert "O|E|" & "|" & "|" & "||" & "Open Text file " & slFileName & " for Write failed, error #: " & ilRet, True
            mOpenTextForWrite = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    mOpenTextForWrite = True
    Exit Function
'mOpenTextForWriteErr:
'    ilRet = Err.Number
'    Resume Next
End Function



Private Function mCheckVersions() As Integer
    mCheckVersions = True
    
    If (sgNetDBPath <> "") Or (sgRepDBPath <> "") Then
        Exit Function
    End If
    
    'Doug
    'Check if running SQL Server on the latest Data Definition
    'If not, then update the SQL Server definitions
    
    
End Function

Private Function mMessageToText() As Integer
    Dim ilRet As Integer
    Dim slFileName As String
    Dim slMessage As String
    Dim slTitle As String
    Dim llGenDate As Long
    Dim llGenTime As Long
    Dim ilNrf As Integer
    Dim slStr As String
    
    tgRNCefSrchKey.lCode = tgRNGif.lMessageCefCode
    If tgRNCefSrchKey.lCode > 0 Then
        igRNCefRecLen = Len(tgRNCef) '5027
        ilRet = btrGetEqual(hgRNCef, tgRNCef, igRNCefRecLen, tgRNCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
        'If tgRNCef.iStrLen > 0 Then
            slStr = gStripChr0(tgRNCef.sComment)
            If slStr <> "" Then
                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                    slFileName = Trim$(tgRNGif.sMessageNetDBID) & "_" & Trim$(tgRNGif.sMessageRepDBID)
                ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                    slFileName = Trim$(tgRNGif.sMessageRepDBID) & "_" & Trim$(tgRNGif.sMessageNetDBID)
                End If
                llGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                llGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                Select Case Left$(gGetLocalTZName(), 1)
                    Case "E"
                    Case "C"
                        llGenTime = llGenTime + 3600
                    Case "M"
                        llGenTime = llGenTime + 7200
                    Case "P"
                        llGenTime = llGenTime + 10800
                End Select
                slFileName = slFileName & "_Message_" & Format$(llGenDate, "YY-MM-DD") & "_" & Trim$(Str$(llGenTime)) & ".csv"
                ilRet = mOpenTextForWrite(hmTextMsg, slFileName)
                If Not ilRet Then
                    mMessageToText = False
                    Exit Function
                End If
                smFileName1 = slFileName
                smFileName2 = ""
                smFileName3 = ""
                slTitle = "mitRepDBID,mitNetDBID,mitGenDate,mitGenTime,mitSeqNo,mitMessageType,mitMessageSubtype,mitMessage"
                Print #hmTextMsg, slTitle
                slMessage = gCreateMessageRec(Trim$(tgRNGif.sMessageRepDBID), Trim$(tgRNGif.sMessageNetDBID), llGenDate, llGenTime, lmSeqNo, tgRNCef)
                Print #hmTextMsg, slMessage
                lmSeqNo = lmSeqNo + 1
                Close hmTextMsg
                'Create QIF_Queue_Info record
                tgRNQif.lCode = 0
                tgRNQif.sNRProcessed = "N"
                tgRNQif.sType = "T"
                tgRNQif.sSubType = tgRNGif.sType
                tgRNQif.sFileName1 = smFileName1
                tgRNQif.sFileName2 = ""
                tgRNQif.sFileName3 = ""
                tgRNQif.lCntrNo = 0
                'tgRNQif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                'tgRNQif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                'Select Case Left$(gGetLocalTZName(), 1)
                '    Case "E"
                '    Case "C"
                '        tgRNQif.lGenTime = tgRNQif.lGenTime + 3600
                '    Case "M"
                '        tgRNQif.lGenTime = tgRNQif.lGenTime + 7200
                '    Case "P"
                '        tgRNQif.lGenTime = tgRNQif.lGenTime + 10800
                'End Select
                tgRNQif.lGenDate = llGenDate
                tgRNQif.lGenTime = llGenTime
                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                    For ilNrf = 0 To UBound(tgServerNRF) - 1 Step 1
                        If StrComp(Trim$(tgServerNRF(ilNrf).sDBID), Trim$(tgRNGif.sMessageNetDBID), vbTextCompare) = 0 Then
                            tgRNQif.iNrfCode = tgServerNRF(ilNrf).iCode
                            Exit For
                        End If
                    Next ilNrf
                ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                    For ilNrf = 0 To UBound(tgServerNRF) - 1 Step 1
                        If StrComp(Trim$(tgServerNRF(ilNrf).sDBID), Trim$(tgRNGif.sMessageRepDBID), vbTextCompare) = 0 Then
                            tgRNQif.iNrfCode = tgServerNRF(ilNrf).iCode
                            Exit For
                        End If
                    Next ilNrf
                End If
                gPackDate "", tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
                gPackTime "", tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
                tgRNQif.sUnused = ""
                ilRet = btrInsert(hgRNQif, tgRNQif, igRNQifRecLen, INDEXKEY0)
                Sleep (1500)
            End If
        End If
        
    End If
    mMessageToText = True
End Function

Private Function mPostedSpotsToText(slRepDBID As String) As Integer
    Dim ilLoop As Integer
    Dim ilNrf As Integer
    Dim llGenDate As Long
    Dim llGenTime As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim slFileName As String
    Dim slTitle As String
    Dim slPostedSpot As String
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilFileCreated As Integer
    Dim ilChf As Integer
    Dim ilFound As Integer
    
    'Gather Post spots
    ilFileCreated = False
    llGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
    llGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
    Select Case Left$(gGetLocalTZName(), 1)
        Case "E"
        Case "C"
            llGenTime = llGenTime + 3600
        Case "M"
            llGenTime = llGenTime + 7200
        Case "P"
            llGenTime = llGenTime + 10800
    End Select
    ilExtLen = Len(tgRNSdf)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hgRNSdf   'Clear any previous extend operation
    tgRNSdfSrchKey4.iDate(0) = tgRNGif.iStartDate(0)
    tgRNSdfSrchKey4.iDate(1) = tgRNGif.iStartDate(1)
    tgRNSdfSrchKey4.lChfCode = 0
    ilRet = btrGetGreaterOrEqual(hgRNSdf, tgRNSdf, igRNSdfRecLen, tgRNSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point
    If ilRet <> BTRV_ERR_NONE Then
        mPostedSpotsToText = False
        Exit Function
    End If
    Call btrExtSetBounds(hgRNSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
    tlDateTypeBuff.iDate0 = tgRNGif.iStartDate(0)
    tlDateTypeBuff.iDate1 = tgRNGif.iStartDate(1)
    ilOffSet = gFieldOffset("Sdf", "SdfDate")
    ilRet = btrExtAddLogicConst(hgRNSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    tlDateTypeBuff.iDate0 = tgRNGif.iEndDate(0)
    tlDateTypeBuff.iDate1 = tgRNGif.iEndDate(1)
    ilOffSet = gFieldOffset("Sdf", "SdfDate")
    ilRet = btrExtAddLogicConst(hgRNSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
    ilRet = btrExtAddField(hgRNSdf, 0, ilExtLen)  'Extract Name
    ilRet = btrExtGetNext(hgRNSdf, tgRNSdf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mPostedSpotsToText = False
            Exit Function
        End If
        ilExtLen = Len(tgRNSdf)  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hgRNSdf, tgRNSdf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            For ilChf = 0 To UBound(tmNetChfSort) - 1 Step 1
                If tgRNSdf.lChfCode = tmNetChfSort(ilChf).lChfCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilChf
            If ilFound Then
                If tgRNSdf.lChfCode <> tgRNChf.lCode Then
                    tgRNChfSrchKey0.lCode = tgRNSdf.lChfCode
                    ilRet = btrGetEqual(hgRNChf, tgRNChf, igRNChfRecLen, tgRNChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
                If Trim$(tgRNChf.sRepDBID) = slRepDBID Then
                    'Check if file handle defined, if not create one and generate output header record
                    If Not ilFileCreated Then
                        slFileName = Trim$(tgRNChf.sRepDBID) & "_" & Trim$(tgNrf.sDBID)
                        slFileName = slFileName & "_PostedSpots_" & Format$(llGenDate, "YY-MM-DD") & "_" & Trim$(Str$(llGenTime)) & ".csv"
                        ilRet = mOpenTextForWrite(hmTextSdf, slFileName)
                        If Not ilRet Then
                            mPostedSpotsToText = False
                            Exit Function
                        End If
                        slTitle = "pstRepDBID,pstNetDBID,pstGenDate,pstGenTime,pstSeqNo,pstContractNo,pstLineNo,pstSpotID,pstSpotType,pstPriceType,pstSchStatus,pstSchdVehicleName,pstSchdVehicleCode,pstSchdDate,pstSchdTime,pstAirVehicleName,pstAirVehicleCode,pstAirDate,pstAirTime,pstXmidnight,pstISCI"
                        Print #hmTextSdf, slTitle
                        ilFileCreated = True
                    End If
                    slPostedSpot = gCreatePostedSpotRec(Trim$(tgRNChf.sRepDBID), Trim$(tgNrf.sDBID), llGenDate, llGenTime, lmSeqNo, tgRNSdf, tgRNChf)
                    If slPostedSpot <> "" Then
                        Print #hmTextSdf, slPostedSpot
                        lmSeqNo = lmSeqNo + 1
                    End If
                End If
            End If
            ilRet = btrExtGetNext(hgRNSdf, tgRNSdf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hgRNSdf, tgRNSdf, ilExtLen, llRecPos)
            Loop
        Loop
        If ilFileCreated Then
            Close hmTextSdf
            'Create QIF_Queue_Info record
            tgRNQif.lCode = 0
            tgRNQif.sNRProcessed = "N"
            tgRNQif.sType = "T"
            tgRNQif.sSubType = tgRNGif.sType
            tgRNQif.sFileName1 = slFileName
            tgRNQif.sFileName2 = ""
            tgRNQif.sFileName3 = ""
            tgRNQif.lCntrNo = tgRNGif.lCntrNo   'Pass Invoice End Date
            tgRNQif.lGenDate = llGenDate
            tgRNQif.lGenTime = llGenTime
            For ilNrf = 0 To UBound(tgServerNRF) - 1 Step 1
                If StrComp(Trim$(tgServerNRF(ilNrf).sDBID), Trim$(slRepDBID), vbTextCompare) = 0 Then
                    tgRNQif.iNrfCode = tgServerNRF(ilNrf).iCode
                    Exit For
                End If
            Next ilNrf
            gPackDate "", tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
            gPackTime "", tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
            tgRNQif.sUnused = ""
            ilRet = btrInsert(hgRNQif, tgRNQif, igRNQifRecLen, INDEXKEY0)
            Sleep (1500)
        End If
    End If
    mPostedSpotsToText = True
End Function

Private Function mAvailSumToText() As Integer
    Dim llGenDate As Long
    Dim llGenTime As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilNrf As Integer
    Dim slFileName As String
    Dim slTitle As String
    Dim slAvailSum As String
    Dim slRepDBID As String

    llGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
    llGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
    Select Case Left$(gGetLocalTZName(), 1)
        Case "E"
        Case "C"
            llGenTime = llGenTime + 3600
        Case "M"
            llGenTime = llGenTime + 7200
        Case "P"
            llGenTime = llGenTime + 10800
    End Select
    
    tgRNVefSrchKey0.iCode = tgRNGif.iVefCode
    ilRet = btrGetEqual(hgRNVef, tgRNVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        'gCreateAlert "C|E|" & "|" & Trim$(Str$(tgRNChf.lCntrNo)) & "|" & "||" & "Accessing Vehicle(vefCode = " & Trim$(Str$(ilVefCode(ilVef))) & ", Pervasive Error # " & ilRet, True
        mAvailSumToText = False
        Exit Function
    End If
    
    If tgRNVef.iNrfCode > 0 Then
        tgRNAsfSrchKey0.iCode = tgRNGif.iVefCode
        ilRet = btrGetEqual(hgRNAsf, tgRNAsf, igRNAsfRecLen, tgRNAsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            'gCreateAlert "C|E|" & "|" & Trim$(Str$(tgRNChf.lCntrNo)) & "|" & "||" & "Accessing Vehicle(vefCode = " & Trim$(Str$(ilVefCode(ilVef))) & ", Pervasive Error # " & ilRet, True
            mAvailSumToText = False
            Exit Function
        End If
        For ilNrf = 0 To UBound(tgServerNRF) - 1 Step 1
            If tgRNVef.iNrfCode = tgServerNRF(ilNrf).iCode Then
                slRepDBID = Trim$(tgServerNRF(ilNrf).sDBID)
                slFileName = slRepDBID & "_" & Trim$(tgNrf.sDBID) & "_" & tgRNGif.iVefCode
                slFileName = slFileName & "_AvailSum_" & Format$(llGenDate, "YY-MM-DD") & "_" & Trim$(Str$(llGenTime)) & ".csv"
                ilRet = mOpenTextForWrite(hmTextAsf, slFileName)
                If Not ilRet Then
                    mAvailSumToText = False
                    Exit Function
                End If
                slTitle = "amtRepDBID,amtNetDBID,amtVehicleName,amtNetVehicleCode,amtGenDate,amtGenTime,amtSeqNo,"
                slTitle = slTitle & "amtDaypartName,amtNetDaypartCode,amtDPInOut,amtDPAvailName,amtDPDay_1,amtDPStartTime_1,amtDPEndTime_1,amtDPSpotPct_1,amtDPDay_2,amtDPStartTime_2,amtDPEndTime_2,amtDPSpotPct_2,amtStartWeekDate,"
                slTitle = slTitle & "amtAvail_60_Wk01,amtAvail_60_Wk03,amtAvail_60_Wk05,amtAvail_60_Wk07,amtAvail_60_Wk09,amtAvail_60_Wk11,amtAvail_60_Wk13,amtAvail_60_Wk15,amtAvail_60_Wk17,amtAvail_60_Wk19,amtAvail_60_Wk21,amtAvail_60_Wk23,amtAvail_60_Wk25,"
                slTitle = slTitle & "amtAvail_30_Wk01,amtAvail_30_Wk03,amtAvail_30_Wk05,amtAvail_30_Wk07,amtAvail_30_Wk09,amtAvail_30_Wk11,amtAvail_30_Wk13,amtAvail_30_Wk15,amtAvail_30_Wk17,amtAvail_30_Wk19,amtAvail_30_Wk21,amtAvail_30_Wk23,amtAvail_30_Wk25,"
                slTitle = slTitle & "amtAvail_10_Wk01,amtAvail_10_Wk03,amtAvail_10_Wk05,amtAvail_10_Wk07,amtAvail_10_Wk09,amtAvail_10_Wk11,amtAvail_10_Wk13,amtAvail_10_Wk15,amtAvail_10_Wk17,amtAvail_10_Wk19,amtAvail_10_Wk21,amtAvail_10_Wk23,amtAvail_10_Wk25,"
                slTitle = slTitle & "amtCreatedDate,amtCreatedTime"
                Print #hmTextAsf, slTitle
                Do While tgRNAsf.iVefCode = tgRNGif.iVefCode
                    slAvailSum = gCreateAvailSumRec(slRepDBID, Trim$(tgNrf.sDBID), llGenDate, llGenTime, lmSeqNo, tgRNAsf)
                    If slAvailSum <> "" Then
                        Print #hmTextAsf, slAvailSum
                        lmSeqNo = lmSeqNo + 1
                    End If
                    ilRet = btrGetNext(hgRNAsf, tgRNAsf, igRNAsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                Loop
                Close hmTextAsf
                tgRNQif.lCode = 0
                tgRNQif.sNRProcessed = "N"
                tgRNQif.sType = "T"
                tgRNQif.sSubType = tgRNGif.sType
                tgRNQif.sFileName1 = slFileName
                tgRNQif.sFileName2 = ""
                tgRNQif.sFileName3 = ""
                tgRNQif.lCntrNo = 0
                tgRNQif.lGenDate = llGenDate
                tgRNQif.lGenTime = llGenTime
                tgRNQif.iNrfCode = tgServerNRF(ilNrf).iCode
                gPackDate "", tgRNQif.iProcessedDate(0), tgRNQif.iProcessedDate(1)
                gPackTime "", tgRNQif.iProcessedTime(0), tgRNQif.iProcessedTime(1)
                tgRNQif.sUnused = ""
                ilRet = btrInsert(hgRNQif, tgRNQif, igRNQifRecLen, INDEXKEY0)
                Sleep (1500)
                Exit For
            End If
        Next ilNrf
    End If
    mAvailSumToText = True
End Function
