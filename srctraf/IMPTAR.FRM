VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptAR 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   360
   ClientTop       =   3600
   ClientWidth     =   8535
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
   Icon            =   "IMPTAR.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5055
   ScaleWidth      =   8535
   Begin VB.ListBox lbcResult 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "IMPTAR.frx":08CA
      Left            =   420
      List            =   "IMPTAR.frx":08CC
      TabIndex        =   10
      Top             =   1890
      Width           =   7635
   End
   Begin VB.DirListBox lbcResultDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   5130
      TabIndex        =   9
      Top             =   600
      Width           =   2940
   End
   Begin VB.DriveListBox cbcResultDir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3030
      TabIndex        =   8
      Top             =   585
      Width           =   1965
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6435
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7380
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3810
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6885
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3810
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Import"
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
      Left            =   3225
      TabIndex        =   0
      Top             =   4215
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
      Left            =   4530
      TabIndex        =   1
      Top             =   4215
      Width           =   1050
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   225
      Left            =   2835
      TabIndex        =   12
      Top             =   3840
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label plcScreen 
      Caption         =   "Import A/R"
      Height          =   240
      Left            =   75
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label lacResult 
      Appearance      =   0  'Flat
      Caption         =   "Folder containing RVF, PHF and VEF from Central office"
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   480
      TabIndex        =   11
      Top             =   540
      Width           =   2310
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "All Users Must Be Off The System During The Running Of This Program"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   255
      TabIndex        =   7
      Top             =   300
      Width           =   7920
   End
   Begin VB.Label lacToDB 
      Appearance      =   0  'Flat
      Caption         =   "Data To:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   4680
      Width           =   7635
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   3465
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1545
      TabIndex        =   5
      Top             =   3495
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "ImptAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptAR.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Import Rate Card input screen code
Option Explicit
Option Compare Text
Dim lmTotalNoRecords As Long
Dim lmProcessedNoRecords As Long
Dim lmPercent As Long
Dim smDefaultInitDir As String
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imImporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smTodaysDate As String
Dim lmNowDate As Long
Dim imNowYear As Integer
Dim smFromFilePath As String
Dim smResultFilePath As String
Dim hmMsg As Integer

Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0  'VEF key record image
Dim hmVef As Integer        'VEF Handle
Dim imVefRecLen As Integer      'VEF record length

Dim tmRvf As RVF            'VEF record image
Dim hmFRvf As Integer        'VEF Handle
Dim hmFPhf As Integer
Dim hmTRvf As Integer        'VEF Handle
Dim hmTPhf As Integer        'VEF Handle
Dim tmRvfSrchKey2 As LONGKEY0
Dim tmRvfSrchKey4 As RVFKEY4            'RVF record image (Advertiser code)
Dim imRvfRecLen As Integer      'VEF record length

Dim tmChf As CHF
Dim hmCHF As Integer
Dim tmChfSrchKey1 As CHFKEY1    'CntrNo; CntRevNo; PropVer
Dim imCHFRecLen As Integer        'CHF record length

Dim hmPrf As Integer            'Product file handle
Dim tmPrfSrchKey As PRFKEY1            'PRF record image
Dim imPrfRecLen As Integer        'PRF record length
Dim tmPrf As PRF

Dim lmHIRvfCode() As Long
Dim tmVefMap() As VEFMAP
Dim tmMissedVehNames() As MISSEDVEHNAMES
Dim tmMissedCntr() As MISSEDCNTR
Dim lmPONet As Long
Dim lmPOGross As Long
Dim lmNewRvfCntr() As Long


'*******************************************************
'*                                                     *
'*      Procedure Name:mRvfExist                       *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if contract          *
'*                      previously invoiced            *
'*                                                     *
'*******************************************************
Private Function mRvfExist(hlFile As Integer) As Integer
'   tgChfInv(I)- CHF Record
'   llInvNo(O) Invoice Number
'   ilRet:  True if RVF or PHF matching record exist;  False if no matching record exist
'
    Dim ilFound As Integer
    Dim ilCntrFound As Integer
    Dim ilRet As Integer
    Dim llGross1 As Long
    Dim llGross2 As Long
    Dim llNet1 As Long
    Dim llNet2 As Long
    Dim ilLoop As Integer
    Dim tlRvf As RVF
    Dim ilTranDate(0 To 1) As Integer
    
    For ilLoop = LBound(lmNewRvfCntr) To UBound(lmNewRvfCntr) - 1 Step 1
        If tmChf.lCntrNo = lmNewRvfCntr(ilLoop) Then
            mRvfExist = False
            Exit Function
        End If
    Next ilLoop
    ilFound = False
    ilCntrFound = False
    tmRvfSrchKey4.lCntrNo = tmChf.lCntrNo
    tmRvfSrchKey4.iTranDate(0) = tmRvf.iTranDate(0)
    tmRvfSrchKey4.iTranDate(1) = tmRvf.iTranDate(1)
    ilRet = btrGetGreaterOrEqual(hlFile, tlRvf, imRvfRecLen, tmRvfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'Do While (ilRet = BTRV_ERR_NONE) And (tlRvf.iAdfCode = tgChfInv.iAdfCode)
    Do While (ilRet = BTRV_ERR_NONE) And (tlRvf.lCntrNo = tmChf.lCntrNo)
        ilCntrFound = True
        'If (tgChfInv.iAgfCode = tlRvf.iAgfCode) And (tgChfInv.lCntrNo = tlRvf.lCntrNo) Then
            If (tmRvf.iTranDate(0) = tlRvf.iTranDate(0)) And (tmRvf.iTranDate(1) = tlRvf.iTranDate(1)) And (tmRvf.sTranType = tlRvf.sTranType) And (tmRvf.sCashTrade = tlRvf.sCashTrade) Then
                If (tmRvf.iAirVefCode = tlRvf.iAirVefCode) And (tmRvf.iBillVefCode = tlRvf.iBillVefCode) And (tmRvf.iAgePeriod = tlRvf.iAgePeriod) And (tmRvf.iAgingYear = tlRvf.iAgingYear) Then
                    gPDNToLong tmRvf.sGross, llGross1
                    gPDNToLong tlRvf.sGross, llGross2
                    gPDNToLong tmRvf.sNet, llNet1
                    gPDNToLong tlRvf.sNet, llNet2
                    If (llGross1 = llGross2) And (llNet1 = llNet2) Then
                        ilFound = True
                        Exit Do
                    End If
                End If
            End If
        'End If
        ilRet = btrGetNext(hlFile, tlRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If Not ilCntrFound Then
        lmNewRvfCntr(UBound(lmNewRvfCntr)) = tmChf.lCntrNo
        ReDim Preserve lmNewRvfCntr(0 To UBound(lmNewRvfCntr) + 1) As Long
    End If
    mRvfExist = ilFound
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveBlanks                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove blanks from name        *
'*                                                     *
'*******************************************************
Private Function mRemoveBlanks(slInStr As String) As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slChar As String * 1
    
    slStr = ""
    For ilLoop = 1 To Len(slInStr) Step 1
        slChar = Mid$(slInStr, ilLoop, 1)
        If slChar <> " " Then
            slStr = slStr & slChar
        End If
    Next ilLoop
    mRemoveBlanks = slStr
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpStnFd.Txt"
    slToFile = smFromFilePath & "ImptAR.Txt"
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        'slFileDate = Format$(slDateTime, "m/d/yy")
        'If DateValue(slFileDate) = DateValue(smTodaysDate) Then  'Append
        '    On Error GoTo 0
        '    ilRet = 0
        '    On Error GoTo mOpenMsgFileErr:
        '    hmMsg = FreeFile
        '    Open slToFile For Append As hmMsg
        '    If ilRet <> 0 Then
        '        Screen.MousePointer = vbDefault
        '        MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        '        mOpenMsgFile = False
        '        Exit Function
        '    End If
        'Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmMsg = FreeFile
            Open slToFile For Output As hmMsg
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        'End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Sub cbcResultDir_Change()
    lbcResultDir.Path = cbcResultDir.Drive
End Sub

'Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
'' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcImport_Click()
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDateTime As String
    Dim slMsgFile As String
    Dim slGross As String
    Dim slNet As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slStr As String
    
    lacInfo.Visible = False
    If imImporting Then
        Exit Sub
    End If
    smFromFilePath = lbcResultDir.Path & "\"
    If StrComp(smFromFilePath, sgDBPath, vbTextCompare) = 0 Then
        MsgBox "The from folder can't be the same as the database folder " & smFromFilePath & " same as " & sgDBPath
        Exit Sub
    End If
    ilPos1 = InStr(1, smFromFilePath, "\csi", vbTextCompare)
    ilPos2 = InStr(1, sgDBPath, "\csi", vbTextCompare)
    If StrComp(Mid$(smFromFilePath, ilPos1), Mid$(sgDBPath, ilPos2), vbTextCompare) = 0 Then
        MsgBox "The from folder can't be the same as the database folder " & smFromFilePath & " same as " & sgDBPath
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcImportErr:
    slDateTime = FileDateTime(smFromFilePath & "RVF.BTR")
    If ilRet <> 0 Then
        MsgBox "RVF.BTR missing from " & smFromFilePath
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcImportErr:
    slDateTime = FileDateTime(smFromFilePath & "PHF.BTR")
    If ilRet <> 0 Then
        MsgBox "PHF.BTR missing from " & smFromFilePath
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcImportErr:
    slDateTime = FileDateTime(smFromFilePath & "VEF.BTR")
    If ilRet <> 0 Then
        MsgBox "VEF.BTR missing from " & smFromFilePath
        Exit Sub
    End If
    If Not mOpenMsgFile(slMsgFile) Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imImporting = True
    
    On Error GoTo 0
    Print #hmMsg, "Import Payment Files " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", smFromFilePath & "VEF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    lmTotalNoRecords = btrRecords(hmVef)
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    lmTotalNoRecords = lmTotalNoRecords + btrRecords(hmTPhf)
    hmFRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFRvf, "", smFromFilePath & "RVF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    lmTotalNoRecords = lmTotalNoRecords + btrRecords(hmFRvf)
    ilRet = btrClose(hmFRvf)
    btrDestroy hmFRvf
    hmFPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFPhf, "", smFromFilePath & "PHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    lmTotalNoRecords = lmTotalNoRecords + btrRecords(hmFPhf)
    ilRet = btrClose(hmFPhf)
    btrDestroy hmFPhf
    
    lacInfo.Visible = True
    plcGauge.Visible = True
    plcGauge.Value = 0
   
    Print #hmMsg, ""
    Print #hmMsg, "From folder: " & smFromFilePath
    Print #hmMsg, "To folder: " & sgDBPath
    Print #hmMsg, ""
   
    ilRet = mMakeVefMap()
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "mMakeVefMap I/O error " & Str$(ilRet)
        lbcResult.AddItem "mMakeVefMap I/O error " & Str$(ilRet)
        Print #hmMsg, "Import Failed" & " " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close #hmMsg
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = mRemoveHI()
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "mRemoveHI I/O error " & Str$(ilRet)
        lbcResult.AddItem "mRemoveHI I/O error " & Str$(ilRet)
        Print #hmMsg, "Import Failed" & " " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close #hmMsg
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ilRet = mMoveRec()
    If UBound(tmMissedCntr) > LBound(tmMissedCntr) Then
        If ilRet = 0 Then
            slStr = "Import Completed, but Contracts Missing"
        Else
            slStr = "Import Failed plus Contracts Missing"
        End If
        Print #hmMsg, ""
        Print #hmMsg, "Missing Contract Summary"
        For ilLoop = LBound(tmMissedCntr) To UBound(tmMissedCntr) - 1 Step 1
            slGross = gLongToStrDec(tmMissedCntr(ilLoop).lGross, 2)
            slNet = gLongToStrDec(tmMissedCntr(ilLoop).lNet, 2)
            Print #hmMsg, "Contract " & Str$(tmMissedCntr(ilLoop).lCntrNo) & " missing, Gross = " & slGross & " Net = " & slNet & " Count = " & Str(tmMissedCntr(ilLoop).iCount)
            lbcResult.AddItem "Contract " & Str$(tmMissedCntr(ilLoop).lCntrNo) & " missing, Gross = " & slGross & " Net = " & slNet & " Count = " & Str(tmMissedCntr(ilLoop).iCount)
        Next ilLoop
        ilRet = 1
    Else
        If ilRet = 0 Then
            slStr = "Import Successful Completed"
        Else
            slStr = "Import Failed"
        End If
    End If
    If (lmPOGross <> 0) Or (lmPONet <> 0) Then
        slGross = gLongToStrDec(lmPOGross, 2)
        slNet = gLongToStrDec(lmPONet, 2)
        Print #hmMsg, "Amount of POs " & ", Gross = " & slGross & " Net = " & slNet
    End If
    lacInfo.Caption = slStr
    lacInfo.Visible = True
    Print #hmMsg, ""
    Print #hmMsg, lacInfo.Caption & " " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close #hmMsg
    plcGauge.Value = 100
    Screen.MousePointer = vbDefault
    If ilRet = 0 Then
        MsgBox "See " & slMsgFile & " for extra information"
    Else
        MsgBox "See " & slMsgFile & " for error messages"
    End If
    cmcCancel.Caption = "Done"
    cmcCancel.SetFocus
    cmcImport.Enabled = False
    imImporting = False
    Exit Sub
cmcImportErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    'igParentRestarted = True
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    'sgDoneMsg = CmdStr
    'igChildDone = True
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
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    btrStopAppl
    End
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    Dim llNoRec As Long 'Number of records
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    imTerminate = False
    mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imImporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecords = 0
    lmProcessedNoRecords = 0
    smTodaysDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    lacToDB.Caption = "Data To: " & Left$(sgDBPath, Len(sgDBPath) - 1)
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gObtainVef
    hmTRvf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmTRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        btrDestroy hmTRvf
        Exit Sub
    End If
    imRvfRecLen = Len(tmRvf)
    hmTPhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmTPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        btrDestroy hmTRvf
        btrDestroy hmTPhf
        Exit Sub
    End If
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        btrDestroy hmTRvf
        btrDestroy hmTPhf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmPrf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        imTerminate = True
        btrDestroy hmTRvf
        btrDestroy hmTPhf
        btrDestroy hmCHF
        btrDestroy hmPrf
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)

    gCenterStdAlone ImptAR
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
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
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim ilPos As Integer
    Dim slChar As String
    slCommand = Command$
    smDefaultInitDir = ""
    ilPos = InStr(1, slCommand, "/CS", 1)
    If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
        'imShowHelpMsg = False
        ilPos = InStr(1, slCommand, "/I:", 1)
        If ilPos > 0 Then
            smDefaultInitDir = ""
            ilPos = ilPos + 3
            slChar = Trim$(Mid$(slCommand, ilPos, 1))
            Do While slChar = ""
                ilPos = ilPos + 1
                slChar = Mid$(slCommand, ilPos, 1)
            Loop
            Do While slChar <> ""
                smDefaultInitDir = smDefaultInitDir & slChar
                ilPos = ilPos + 1
                If ilPos > Len(slCommand) Then
                    Exit Do
                End If
                slChar = Mid$(slCommand, ilPos, 1)
            Loop
            If right$(smDefaultInitDir, 1) = "\" Then
                smDefaultInitDir = Left$(smDefaultInitDir, Len(smDefaultInitDir) - 1)
            End If
        End If
    Else
        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            End
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        'imShowHelpMsg = True
        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
            'imShowHelpMsg = False
        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    End If
    gInitStdAlone ImptAR, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    'igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    If Len(smDefaultInitDir) = 0 Then
        smDefaultInitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    End If
    If InStr(1, smDefaultInitDir, ":", vbTextCompare) = 2 Then
        cbcResultDir.Drive = Left$(smDefaultInitDir, 1)
        lbcResultDir.Path = smDefaultInitDir
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
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    
    Erase tmVefMap
    Erase tmMissedVehNames
    Erase lmHIRvfCode
    Erase tmMissedCntr
    Erase lmNewRvfCntr
    
    btrDestroy hmTRvf
    btrDestroy hmTPhf
    btrDestroy hmCHF
    btrDestroy hmPrf
    
    'igParentRestarted = False
    'If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute "Done"
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    'End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptAR
    Set ImptAR = Nothing   'Remove data segment
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeVefMap                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create conversion table from   *
'*                      import to current              *
'*                                                     *
'*******************************************************
Private Function mMakeVefMap() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilMatch As Integer
    Dim slVehicleName As String
    Dim slVefName As String
    Dim slImptName As String
    Dim ilPos As Integer
    Dim llPercent As Long
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    
    ReDim tmVefMap(0 To 0) As VEFMAP
    ReDim tmMissedVehNames(0 To 0) As MISSEDVEHNAMES
    
    lacInfo.Caption = "Gathering Vehicle Information"
    Print #hmMsg, "Gathering Vehicle Information"
    mMakeVefMap = BTRV_ERR_NONE
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", smFromFilePath & "VEF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeVefMap = ilRet
        btrDestroy hmVef
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        mMakeVefMap = ilRet
        btrDestroy hmVef
        Exit Function
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        slVehicleName = Trim$(tmVef.sName)
        ilMatch = False
        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(slVehicleName), 1) = 0 Then
                tmVefMap(UBound(tmVefMap)).iFromVefCode = tmVef.iCode
                tmVefMap(UBound(tmVefMap)).iToVefCode = tgMVef(ilLoop).iCode
                ilMatch = True
                Exit For
            End If
            If StrComp(mRemoveBlanks(tgMVef(ilLoop).sName), mRemoveBlanks(slVehicleName), 1) = 0 Then
                tmVefMap(UBound(tmVefMap)).iFromVefCode = tmVef.iCode
                tmVefMap(UBound(tmVefMap)).iToVefCode = tgMVef(ilLoop).iCode
                ilMatch = True
                Exit For
            End If
        Next ilLoop
        If Not ilMatch Then
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                slVefName = mRemoveBlanks(tgMVef(ilLoop).sName)
                slImptName = mRemoveBlanks(slVehicleName)
                ilPos = InStr(1, slImptName, slVefName, vbTextCompare)
                If ilPos > 0 Then
                    tmVefMap(UBound(tmVefMap)).iFromVefCode = tmVef.iCode
                    tmVefMap(UBound(tmVefMap)).iToVefCode = tgMVef(ilLoop).iCode
                    ilMatch = True
                    Exit For
                End If
            Next ilLoop
        End If
        If Not ilMatch Then
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                slVefName = mRemoveBlanks(tgMVef(ilLoop).sName)
                ilPos1 = InStr(1, slVefName, "-", vbTextCompare)
                If ilPos1 > 0 Then
                    ilPos2 = InStr(ilPos1 + 1, slVefName, "-", vbTextCompare)
                    If ilPos2 > 0 Then
                        slVefName = Mid(slVefName, ilPos1 + 1)
                        slImptName = mRemoveBlanks(slVehicleName)
                        ilPos = InStr(1, slImptName, slVefName, vbTextCompare)
                        If ilPos > 0 Then
                            tmVefMap(UBound(tmVefMap)).iFromVefCode = tmVef.iCode
                            tmVefMap(UBound(tmVefMap)).iToVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                    End If
                End If
            Next ilLoop
        End If
        If ilMatch Then
            ReDim Preserve tmVefMap(0 To UBound(tmVefMap) + 1) As VEFMAP
        Else
            tmMissedVehNames(UBound(tmMissedVehNames)).iVefCode = tmVef.iCode
            tmMissedVehNames(UBound(tmMissedVehNames)).sName = slVehicleName
            tmMissedVehNames(UBound(tmMissedVehNames)).iStatus = 0
            ReDim Preserve tmMissedVehNames(0 To UBound(tmMissedVehNames) + 1) As MISSEDVEHNAMES
        End If
        lmProcessedNoRecords = lmProcessedNoRecords + 1
        llPercent = (lmProcessedNoRecords * 100) \ lmTotalNoRecords
        If llPercent >= 100 Then
            If lmProcessedNoRecords + 3 < lmTotalNoRecords Then
                llPercent = 99
            Else
                llPercent = 100
            End If
        End If
        If lmPercent <> llPercent Then
            plcGauge.Value = llPercent
            lmPercent = llPercent
        End If
        DoEvents
        ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move RVF and PHF record        *
'*                                                     *
'*******************************************************
Private Function mMoveRec() As Integer
    Dim ilPass As Integer
    Dim ilVef As Integer
    Dim llRefInvNo As Long
    Dim llInvNo As Long
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llPercent As Long
    Dim ilTest As Integer
    Dim llValue As Long
    Dim llCount As Long
    Dim slNet As String
    
    mMoveRec = BTRV_ERR_NONE
    ReDim tmMissedCntr(0 To 0) As MISSEDCNTR
    ReDim lmNewRvfCntr(0 To 0) As Long
    lmPONet = 0
    lmPOGross = 0
    For ilPass = 1 To 2 Step 1
        llCount = 0
        If ilPass = 1 Then
            lacInfo.Caption = "Moving Receivable Transactions"
            Print #hmMsg, "Moving Receivable Transactions"
            hmFRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmFRvf, "", smFromFilePath & "RVF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        Else
            lacInfo.Caption = "Moving History Transactions"
            Print #hmMsg, "Moving History Transactions"
            hmFPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmFPhf, "", smFromFilePath & "PHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        End If
        If ilRet <> BTRV_ERR_NONE Then
            mMoveRec = ilRet
            If ilPass = 1 Then
                btrDestroy hmFRvf
            Else
                btrDestroy hmFPhf
            End If
            Exit Function
        End If

        If ilPass = 1 Then
            ilRet = btrGetFirst(hmFRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = btrGetFirst(hmFPhf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        If ilRet <> BTRV_ERR_NONE Then
            mMoveRec = ilRet
            If ilPass = 1 Then
                btrDestroy hmFRvf
            Else
                btrDestroy hmFPhf
            End If
            Exit Function
        End If
        Do While (ilRet = BTRV_ERR_NONE)
            'Only process cash transactions.  Also, ignore POs
            If (tmRvf.sTranType <> "PO") And (tmRvf.sCashTrade = "C") And (tmRvf.iAgingYear >= 2002) Then
                gPDNToStr tmRvf.sNet, 2, slNet
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                'Find latest contract because it contains all current line from different revisions
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) Then
                    tmRvf.lCode = 0
                    tmRvf.iAgfCode = tmChf.iAgfCode
                    tmRvf.iAdfCode = tmChf.iAdfCode
                    'tmRvf.lPrfCode = 0
                    ilRet = mAddProduct()
                    tmRvf.iSlfCode = tmChf.iSlfCode(0)
                    llRefInvNo = tmRvf.lRefInvNo
                    llInvNo = tmRvf.lInvNo
                    tmRvf.lInvNo = llRefInvNo
                    tmRvf.lRefInvNo = llInvNo
                    ilFound = False
                    If tmRvf.iAirVefCode = 0 Then
                        ilFound = True
                    Else
                        For ilVef = LBound(tmVefMap) To UBound(tmVefMap) - 1 Step 1
                            If tmRvf.iAirVefCode = tmVefMap(ilVef).iFromVefCode Then
                                ilFound = True
                                tmRvf.iAirVefCode = tmVefMap(ilVef).iToVefCode
                                Exit For
                            End If
                        Next ilVef
                    End If
                    If Not ilFound Then
                        'Add message with missing vehicle name
                        For ilVef = LBound(tmMissedVehNames) To UBound(tmMissedVehNames) - 1 Step 1
                            If tmRvf.iAirVefCode = tmMissedVehNames(ilVef).iVefCode Then
                                Print #hmMsg, "Contract" & Str$(tmChf.lCntrNo) & " referenced vehicle " & Trim$(tmMissedVehNames(ilVef).sName) & " which is not defined. Net = " & slNet
                                If tmMissedVehNames(ilVef).iStatus = 0 Then
                                    lbcResult.AddItem "Vehicle " & Trim$(tmMissedVehNames(ilVef).sName) & " missing"
                                    tmMissedVehNames(ilVef).iStatus = 1
                                End If
                                ilFound = True
                                Exit For
                            End If
                        Next ilVef
                        If Not ilFound Then
                            Print #hmMsg, "Contract" & Str$(tmChf.lCntrNo) & " referenced air vehicle code " & Str(tmRvf.iAirVefCode) & " which had no matching name. Net = " & slNet
                        Else
                            ilFound = False
                        End If
                    Else
                        ilFound = False
                        If tmRvf.iBillVefCode = 0 Then
                            ilFound = True
                        Else
                            For ilVef = LBound(tmVefMap) To UBound(tmVefMap) - 1 Step 1
                                If tmRvf.iBillVefCode = tmVefMap(ilVef).iFromVefCode Then
                                    ilFound = True
                                    tmRvf.iBillVefCode = tmVefMap(ilVef).iToVefCode
                                    Exit For
                                End If
                            Next ilVef
                            If Not ilFound Then
                                'Add message with missing vehicle name
                                For ilVef = LBound(tmMissedVehNames) To UBound(tmMissedVehNames) - 1 Step 1
                                    If tmRvf.iBillVefCode = tmMissedVehNames(ilVef).iVefCode Then
                                        Print #hmMsg, "Contract" & Str$(tmChf.lCntrNo) & " referenced vehicle " & Trim$(tmMissedVehNames(ilVef).sName) & " which is not defined. Net = " & slNet
                                        If tmMissedVehNames(ilVef).iStatus = 0 Then
                                            lbcResult.AddItem "Vehicle " & Trim$(tmMissedVehNames(ilVef).sName) & " missing"
                                            tmMissedVehNames(ilVef).iStatus = 1
                                        End If
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilVef
                                If Not ilFound Then
                                    Print #hmMsg, "Contract" & Str$(tmChf.lCntrNo) & " referenced air vehicle code " & Str(tmRvf.iBillVefCode) & " which had no matching name. Net = " & slNet
                                Else
                                    ilFound = False
                                End If
                            End If
                        End If
                    End If
                    If ilFound Then
                        tmRvf.iUrfCode = 2
                        tmRvf.iPkLineNo = 0
                        tmRvf.iMnfGroup = 0 'The report will split values for participants.
                        tmRvf.iMnfItem = 0
                        tmRvf.sSlsComm = "N"
                        tmRvf.lCefCode = 0
                        If ilPass = 1 Then
                            ilRet = mRvfExist(hmTRvf)
                        Else
                            ilRet = mRvfExist(hmTPhf)
                        End If
                        If Not ilRet Then
                            If ilPass = 1 Then
                                ilRet = btrInsert(hmTRvf, tmRvf, imRvfRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    Print #hmMsg, "Insert RVF error" & Str$(ilRet)
                                    lbcResult.AddItem "Insert RVF error" & Str$(ilRet)
                                Else
                                    llCount = llCount + 1
                                End If
                            Else
                                ilRet = btrInsert(hmTPhf, tmRvf, imRvfRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    Print #hmMsg, "Insert PHF error" & Str$(ilRet)
                                    lbcResult.AddItem "Insert PHF error" & Str$(ilRet)
                                Else
                                    llCount = llCount + 1
                                End If
                            End If
                        End If
                    End If
                Else
                    'Test if any vehicle is part of this group
                    For ilVef = LBound(tmVefMap) To UBound(tmVefMap) - 1 Step 1
                        If tmRvf.iAirVefCode = tmVefMap(ilVef).iFromVefCode Then
                            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                If tgMVef(ilLoop).iCode = tmVefMap(ilVef).iToVefCode Then
                                    ilFound = False
                                    For ilTest = LBound(tmMissedCntr) To UBound(tmMissedCntr) - 1 Step 1
                                        If tmMissedCntr(ilTest).lCntrNo = tmRvf.lCntrNo Then
                                            ilFound = True
                                            gPDNToLong tmRvf.sGross, llValue
                                            tmMissedCntr(ilTest).lGross = tmMissedCntr(ilTest).lGross + llValue
                                            gPDNToLong tmRvf.sNet, llValue
                                            tmMissedCntr(ilTest).lNet = tmMissedCntr(ilTest).lNet + llValue
                                            tmMissedCntr(ilTest).iCount = tmMissedCntr(ilTest).iCount + 1
                                            Exit For
                                        End If
                                    Next ilTest
                                    If Not ilFound Then
                                        ilTest = UBound(tmMissedCntr)
                                        tmMissedCntr(ilTest).lCntrNo = tmRvf.lCntrNo
                                        gPDNToLong tmRvf.sGross, llValue
                                        tmMissedCntr(ilTest).lGross = llValue
                                        gPDNToLong tmRvf.sNet, llValue
                                        tmMissedCntr(ilTest).lNet = llValue
                                        tmMissedCntr(ilTest).iCount = 1
                                        ReDim Preserve tmMissedCntr(0 To ilTest + 1) As MISSEDCNTR
                                    End If
                                    Print #hmMsg, "Contract" & Str$(tmRvf.lCntrNo) & " referenced vehicle " & Trim$(tgMVef(ilLoop).sName) & " but contract is missing. Net =" & slNet
                                    'lbcResult.AddItem "Contract" & Str$(tmRvf.lCntrNo) & " referenced vehicle " & Trim$(tgMVef(ilLoop).sName) & " but contract is missing"
                                    Exit For
                                End If
                            Next ilLoop
                            Exit For
                        End If
                    Next ilVef
                End If
            ElseIf (tmRvf.sTranType = "PO") And (tmRvf.sCashTrade = "C") And (tmRvf.iAgingYear >= 2002) Then
                gPDNToLong tmRvf.sGross, llValue
                lmPOGross = lmPOGross + llValue
                gPDNToLong tmRvf.sNet, llValue
                lmPONet = lmPONet + llValue
            End If
            lmProcessedNoRecords = lmProcessedNoRecords + 1
            llPercent = (lmProcessedNoRecords * 100) \ lmTotalNoRecords
            If llPercent >= 100 Then
                If lmProcessedNoRecords + 3 < lmTotalNoRecords Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            DoEvents
            If ilPass = 1 Then
                ilRet = btrGetNext(hmFRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Else
                ilRet = btrGetNext(hmFPhf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            End If
        Loop

        If ilPass = 1 Then
            btrDestroy hmFRvf
            lacInfo.Caption = "Moved " & Str(llCount) & " Receivable Transactions"
        Else
            btrDestroy hmFPhf
            lacInfo.Caption = "Moved " & Str(llCount) & " History Transactions"
        End If
    Next ilPass
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveHI                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove transaction type HI from*
'*                      History                        *
'*                                                     *
'*******************************************************
Private Function mRemoveHI() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llPercent As Long
    ReDim lmHIRvfCode(0 To 0) As Long
    
    lacInfo.Caption = "Gathering HI Transactions from History"
    Print #hmMsg, "Gathering HI Transactions from History"
    ilRet = btrGetFirst(hmTPhf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        mRemoveHI = ilRet
        Exit Function
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If tmRvf.sTranType = "HI" Then
            lmHIRvfCode(UBound(lmHIRvfCode)) = tmRvf.lCode
            ReDim Preserve lmHIRvfCode(0 To UBound(lmHIRvfCode) + 1) As Long
        End If
        lmProcessedNoRecords = lmProcessedNoRecords + 1
        llPercent = (lmProcessedNoRecords * 100) \ lmTotalNoRecords
        If llPercent >= 100 Then
            If lmProcessedNoRecords + 3 < lmTotalNoRecords Then
                llPercent = 99
            Else
                llPercent = 100
            End If
        End If
        If lmPercent <> llPercent Then
            plcGauge.Value = llPercent
            lmPercent = llPercent
        End If
        DoEvents
        ilRet = btrGetNext(hmTPhf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    lmTotalNoRecords = lmTotalNoRecords + UBound(lmHIRvfCode)
    If UBound(lmHIRvfCode) > LBound(lmHIRvfCode) Then
        lacInfo.Caption = "Deleting HI Transactions from History"
        Print #hmMsg, "Deleting " & Str(UBound(lmHIRvfCode)) & " HI Transactions from History"
    End If
    DoEvents
    For ilLoop = 0 To UBound(lmHIRvfCode) - 1 Step 1
        tmRvfSrchKey2.lCode = lmHIRvfCode(ilLoop)
        ilRet = btrGetEqual(hmTPhf, tmRvf, imRvfRecLen, tmRvfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmTPhf)
            If ilRet <> BTRV_ERR_NONE Then
                mRemoveHI = ilRet
                Exit Function
            End If
        Else
            mRemoveHI = ilRet
            Exit Function
        End If
        lmProcessedNoRecords = lmProcessedNoRecords + 1
        llPercent = (lmProcessedNoRecords * 100) \ lmTotalNoRecords
        If llPercent >= 100 Then
            If lmProcessedNoRecords + 3 < lmTotalNoRecords Then
                llPercent = 99
            Else
                llPercent = 100
            End If
        End If
        If lmPercent <> llPercent Then
            plcGauge.Value = llPercent
            lmPercent = llPercent
        End If
        DoEvents
    Next ilLoop
    mRemoveHI = BTRV_ERR_NONE
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mAddProduct                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add product if not defined     *
'*                                                     *
'*******************************************************
Private Function mAddProduct() As Integer
    Dim slProduct As String
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    'Product- update prf now instead of after all invoices generated
    tmRvf.lPrfCode = 0
    slProduct = tmChf.sProduct
    If (Trim$(slProduct) <> "") Then
        ilFound = False
        tmPrfSrchKey.iAdfCode = tmChf.iAdfCode
        ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = tmChf.iAdfCode) 'tmRvf.iAdfCode)
            If StrComp(Trim$(Trim$(slProduct)), Trim$(tmPrf.sName), 1) = 0 Then
                ilFound = True
                tmRvf.lPrfCode = tmPrf.lCode
                Exit Do
            End If
            ilRet = btrGetNext(hmPrf, tmPrf, imPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (Not ilFound) Then
            Do  'Loop until record updated or added
                tmPrf.lCode = 0
                tmPrf.iAdfCode = tmChf.iAdfCode
                tmPrf.sName = Trim$(slProduct)
                tmPrf.iMnfComp(0) = 0
                tmPrf.iMnfComp(1) = 0
                tmPrf.iMnfExcl(0) = 0
                tmPrf.iMnfExcl(1) = 0
                tmPrf.iUrfCode = tgUrf(0).iCode
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = tmPrf.lCode
                ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                tmRvf.lPrfCode = 0
                mAddProduct = ilRet
                Exit Function
            End If
        End If
    End If
    mAddProduct = BTRV_ERR_NONE
End Function
