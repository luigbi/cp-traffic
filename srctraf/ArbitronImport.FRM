VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ArbitronImport 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2580
   ClientLeft      =   435
   ClientTop       =   1920
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
   Icon            =   "ArbitronImport.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   7095
   Begin VB.FileListBox lbcFile 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   5985
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   765
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.DriveListBox cbcDrive 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5895
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.DirListBox lbcPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5850
      TabIndex        =   6
      Top             =   390
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Index           =   0
      Left            =   3465
      TabIndex        =   4
      Top             =   450
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
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
      Left            =   2385
      TabIndex        =   1
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1695
      Width           =   1050
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Index           =   1
      Left            =   3465
      TabIndex        =   9
      Top             =   945
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   945
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Label lacScreen 
      Caption         =   "Arbitron Import"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   60
      Width           =   6705
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   2175
      Width           =   6615
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Visible         =   0   'False
      Width           =   3150
   End
End
Attribute VB_Name = "ArbitronImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ArbitronImport.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long

'Contract record information
Dim hmArb As Integer        'Contract line file handle
Dim tmArb() As Arbitron           'CLF record image
Dim tmArbSrchKey1 As ARBKEY1
Dim imArbRecLen As Integer
Dim hmFrom As Integer
Dim smBookName As String
Dim smBookPeriod As String
Dim smReportType As String
Dim smTempFileName As String
Dim smFileName As String
Dim imPageSize(0 To 3) As Integer
Dim imImporting As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmStartDate As Long
Dim lmEndDate As Long
Dim lmFileLen As Long
Dim lmFileProcessedRecs As Long
Dim hmMsg As Integer
Dim tmArbAudChar() As ARBAUDCHAR
Dim tmArbCallLetters() As ARBCALLLETTERS
Dim tmArbDaypart() As ARBDAYPART
Dim smInputBuffer As String
Dim imUpdateFile As Integer







Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'
'           1-29-01 Remove everything in file before genning the weeks summary
'                   previously only removed the vehicle it was processing.  Dormant ones remained on file.
Private Sub cmcImport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llPercent As Long
    Dim slToFile As String
    Dim slNewFile As String
    Dim slDateTime As String
    Dim ilEof As Integer
    Dim slName As String
    Dim ilCDNumber As Integer
    Dim slReportType As String
    
    lacInfo(0).Visible = False
    If imImporting Then
        Exit Sub
    End If
    'Test if file exist
    Select Case smReportType
        Case "1"
            slReportType = "RMR"
        Case "2"
            slReportType = "EB"
        Case "3"
            slReportType = "EH"
        Case "4"
            slReportType = "N"
        Case "6"
            slReportType = "SM"
        Case Else
            slReportType = smReportType
    End Select
    ReDim tmArb(0 To 0) As Arbitron
        
    imUpdateFile = False
    smFileName = smBookName & "_" & smBookPeriod & "_" & slReportType & ".arb"
    smTempFileName = "Arb.Dat"
    On Error GoTo cmcImportFileErr:
    slToFile = sgDBPath & smFileName
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        'ilRet = MsgBox("File Exist- Recreate", vbQuestion + vbYesNo, "File Exist")
        'If ilRet = vbNo Then
        sgGenMsg = "File Exist: Press either Recreate, Update or Cancel"
        sgCMCTitle(0) = "&Recreate"
        sgCMCTitle(1) = "&Update"
        sgCMCTitle(2) = "&Cancel"
        sgCMCTitle(3) = ""
        igDefCMC = 0
        igEditBox = 0
        sgEditValue = ""
        GenMsg.Show vbModal
        If (igAnsCMC = 2) Then
            imImporting = False
            Exit Sub
        End If
        If igAnsCMC = 0 Then
            Kill slToFile
        End If
        If igAnsCMC = 1 Then
            Kill sgDBPath & smTempFileName
            Name sgDBPath & smFileName As sgDBPath & smTempFileName
            imUpdateFile = True
        End If
    End If
    Screen.MousePointer = vbHourglass
    cmcCancel.Caption = "&Cancel"
    If Not mOpenMsgFile() Then
        imImporting = False
        cmcCancel.SetFocus
        Exit Sub
    End If
    
    imImporting = True
    lacInfo(0).Visible = True
    lacInfo(0).Caption = ""
    lacInfo(1).Visible = True
    lacInfo(1).Caption = ""
    If Not imUpdateFile Then
        ilRet = mCreateFile()
    Else
        ilRet = True
    End If
    If Not ilRet Then
        imImporting = False
        Close hmMsg
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '2-02-01
    hmArb = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmArb, "", sgDBPath & smTempFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo cmcExportErr
    gBtrvErrorMsg ilRet, "cmcImport (btrOpen)", ArbitronImport
    On Error GoTo 0
    imArbRecLen = Len(tmArb(0))
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    plcGauge(0).Visible = True
    plcGauge(0).Value = 0
    plcGauge(1).Visible = True
    plcGauge(1).Value = 0
    
    ilEof = False
    Do
        lmTotalNoRecs = lbcFile.ListCount
        lmProcessedNoRecs = 0
        plcGauge(0).Value = 0
        DoEvents
        For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
            lacInfo(0).Caption = "File" & Str(ilLoop + 1) & " of" & Str(lmTotalNoRecs)
            ilRet = 0
            slStr = lbcFile.List(ilLoop)
            lacInfo(1).Caption = "Processing " & slStr
            hmFrom = FreeFile
            slName = lbcFile.Path & slStr
            lmFileLen = FileLen(slName) \ 100
            Open slName For Input Access Read As hmFrom
            If ilRet = 0 Then
                ilRet = mReadImport()
                Close hmFrom
                If ilRet = 1 Then   'Wrong CD
                    Exit For
                End If
                If ilRet < 0 Then
                    imTerminate = True
                    Exit Do
                End If
            End If
            lmProcessedNoRecs = lmProcessedNoRecs + 1
            'lacInfo(0).Caption = Trim$(Str$(lmProcessedNoRecs)) & " of" & Str$(lmTotalNoRecs)
            llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            plcGauge(0).Value = llPercent
            DoEvents
            If imTerminate Then
                lacInfo(0).Caption = "Processing Terminated"
                Print #hmMsg, "** Terminated ** " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Exit Do
            End If
        Next ilLoop
        'Have User insert next CD
        'ilRet = MsgBox("Insert Next CD, Press Ok when Inserted or Press Cancel if Completed (Ignore Windows Explore if Opened when CD Inserted)", vbInformation + vbOKCancel, "Next CD")
        'If ilRet = vbCancel Then
        sgGenMsg = "Insert Next CD, Press Ok when Inserted or Press Done if Completed (Ignore Windows Explore if Opened when CD Inserted)"
        sgCMCTitle(0) = "&Ok"
        sgCMCTitle(1) = "&Done"
        sgCMCTitle(2) = ""
        sgCMCTitle(3) = ""
        igDefCMC = 0
        igEditBox = 0
        sgEditValue = ""
        GenMsg.Show vbModal
        If (igAnsCMC = 1) Then
            ilEof = True
        Else
            'lbcFile.Path = "C:\"
            'lbcFile.Path = lbcPath.Path
            lbcFile.Refresh
        End If
    Loop While Not ilEof
    lacInfo(0).Caption = ""
    plcGauge(0).Value = 0
    lacInfo(1).Caption = ""
    plcGauge(1).Value = 0
    imImporting = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    cmcImport.Enabled = False
    ilRet = btrClose(hmArb)
    btrDestroy hmArb
    If Not imTerminate Then
        Print #hmMsg, "** Completed ** " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    Close #hmMsg
    If Not imTerminate Then
        Name sgDBPath & smTempFileName As sgDBPath & smFileName
    End If
    Screen.MousePointer = vbDefault
    If imTerminate Then
        cmcCancel_Click
    End If
    Exit Sub
cmcImportFileErr:
    ilRet = 1
    Resume Next
cmcExportErr:
    imTerminate = True
    ilRet = Err
    Resume Next
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
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    btrStopAppl
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
'      2-02-01 Move Opens to gcmcFix; right before the
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
    Dim ilTest As Integer
    Dim slDrive As String
    Dim slExt As String
    Dim ilPos As Integer
    Dim slName As String
    Dim slLine As String
    Dim slNum As String
    
    imTerminate = False
    mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imImporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    smBookName = ""
    smBookPeriod = ""
    imPageSize(0) = 512
    imPageSize(1) = 1024
    imPageSize(2) = 2048
    imPageSize(3) = 4096
    lacTo.Caption = sgDBPath
    For ilLoop = 0 To cbcDrive.ListCount - 1 Step 1
        ilRet = 0
        On Error GoTo mInitErr
        slDrive = UCase(Left(cbcDrive.List(ilLoop), 1))
        If (slDrive <> "A") And (slDrive <> "C") Then
            cbcDrive.Drive = cbcDrive.List(ilLoop)
            If ilRet = 0 Then
                lbcPath.Path = cbcDrive.Drive & "\"
                If ilRet = 0 Then
                    'lbcFile.Clear
                    lbcFile.Path = lbcPath.Path
                    For ilTest = 0 To lbcFile.ListCount - 1 Step 1
                        slStr = lbcFile.List(ilTest)
                        ilPos = InStr(1, slStr, ".", vbTextCompare)
                        If ilPos > 0 Then
                            slExt = Mid(slStr, ilPos + 1)
                            If Len(slExt) = 3 Then
                                slNum = Mid(slExt, 1, 1)
                                If (slNum >= "0") And (slNum <= "9") Then
                                    slNum = Mid(slExt, 2, 1)
                                    If (slNum >= "0") And (slNum <= "9") Then
                                        slNum = Mid(slExt, 3, 1)
                                        'Open file and test if Line one has Copyright and Arbitron in it
                                        hmFrom = FreeFile
                                        slName = lbcFile.Path & slStr
                                        Open slName For Input Access Read As hmFrom
                                        If ilRet = 0 Then
                                            slLine = Input(145, #hmFrom)    'Remove this line to read each character
                                            Close hmFrom
                                            If (InStr(1, slLine, "Copyright", vbTextCompare) = 110) And (InStr(1, slLine, "Arbitron", vbTextCompare) > 110) Then
                                                smReportType = Trim$(Mid(slLine, 2, 2))
                                                smBookName = Trim$(Mid(slLine, 19, 36))
                                                smBookPeriod = Trim$(Mid(slLine, 55, 11)) & "-" & Trim$(Mid(slLine, 66, 11))
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next ilTest
                    If smBookName <> "" Then
                        Exit For
                    End If
                End If
            End If
        End If
    Next ilLoop
    If smBookName <> "" Then
        lacScreen.Caption = "Arbitron Import: " & smBookName
    Else
        lacScreen.Caption = "Arbitron Import: " & "Book Not Found"
        cmcImport.Enabled = False
    End If
    gCenterStdAlone ArbitronImport
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    ilRet = 1
    Resume Next
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
    igDirectCall = 0
    sgIniPath = ""
    slCommand = Command$
    ilPos = InStr(1, slCommand, "/CS", 1)
        igStdAloneMode = True 'False  'Switch from/to stand alone mode
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
        slStr = "Guide"
        slCommand = "Traffic\Guide"
    gInitStdAlone ArbitronImport, slStr, ilTestSystem
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
    
    Erase tmArbAudChar
    Erase tmArbCallLetters
    Erase tmArbDaypart
    Erase tmArb
    
    ilRet = btrClose(hmArb)
    btrDestroy hmArb
    Screen.MousePointer = vbDefault
    igParentRestarted = False
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ArbitronImport
    Set ArbitronImport = Nothing   'Remove data segment
    igManUnload = NO
End Sub




Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpDall.Txt"
    slToFile = sgDBPath & "Messages\" & "ArbitronImport.Txt"
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        'slFileDate = Format$(slDateTime, "m/d/yy")
        'If gDateValue(slFileDate) = lmNowDate Then  'Append
        '    On Error GoTo 0
        '    ilRet = 0
        '    On Error GoTo mOpenMsgFileErr:
        '    hmMsg = FreeFile
        '    Open slToFile For Append As hmMsg
        '    If ilRet <> 0 Then
        '        Screen.MousePointer = vbDefault
        '        MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
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
                MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
            MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "Contract: Set Delete Flags: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = Err.Number
    Resume Next
End Function

Private Function mCreateFile() As Integer
    Dim slNewFile As String
    Dim llBaseRecLen As Long
    Dim ilVarLenRec As Integer
    Dim ilVarType As Integer
    Dim ilFileFlags As Integer
    Dim ilPages As Integer
    Dim ilPageSize As Integer
    Dim ilNoKeys As Integer
    Dim ilRet As Integer
    Dim ilKeyOffset As Integer
    Dim ilKeySize As Integer
    Dim ilKeyFlag As Integer
    Dim ilDataType As Integer
    
    ilVarLenRec = False
    ilVarType = 0
    llBaseRecLen = Len(tmArb(0))   '4 + 4 + 2 + 1 + 1 + 1 + 2 + 2 + 7 + 1 + 18 * 4
    ilPageSize = 4096
    
    'Flag bits
    '0x0001       File contains variable-length records
    '0x0002       Trailing blanks are truncated
    '0x0004       Pages were preallocated
    '0x0008       Data in file is compressed
    '0x0010       File contains keys only
    '0x0020       Index balancing is in effect
    '0x0040       Free space threshold is 10%
    '0x0080       Free space threshold is 20%
    '0x00C0       Free space threshold is 30%
    '0x0100       File has reserved space for duplicate pointers
    '0x0200       Reserved for future use
    '0x0400       Specific key number was assigned
    '0x0800       File contains VATs (version 6.1+ only)
    '
    ilFileFlags = 0
    On Error Resume Next
    Kill sgDBPath & smTempFileName
    On Error GoTo 0
    hmArb = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilNoKeys = 2
    ilRet = btrCreFile(hmArb, llBaseRecLen, ilPageSize, ilNoKeys, ilFileFlags, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create File Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    'Key0: AutoCode
    ilKeyOffset = 0
    ilKeySize = 4
    ilKeyFlag = &H100
    ilDataType = 15
    ilRet = btrCreKey(hmArb, ilKeyOffset + 1, ilKeySize, ilKeyFlag, ilDataType, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Key Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Key Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    'Key1, Seg 1: CallLetters
    ilKeyOffset = 4
    ilKeySize = 4
    ilKeyFlag = &H113
    ilDataType = 0
    ilRet = btrCreKey(hmArb, ilKeyOffset + 1, ilKeySize, ilKeyFlag, ilDataType, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Key Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Key Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    'Key1, Seg 2: Band
    ilKeyOffset = 8
    ilKeySize = 2
    ilKeyFlag = &H113
    ilDataType = 0
    ilRet = btrCreKey(hmArb, ilKeyOffset + 1, ilKeySize, ilKeyFlag, ilDataType, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Key Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Key Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    'Key1, Seg 3: DataType
    ilKeyOffset = 10
    ilKeySize = 1
    ilKeyFlag = &H113
    ilDataType = 0
    ilRet = btrCreKey(hmArb, ilKeyOffset + 1, ilKeySize, ilKeyFlag, ilDataType, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Key Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Key Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    'Key1, Seg 4: DayIndex
    ilKeyOffset = 17
    ilKeySize = 2
    ilKeyFlag = &H113
    ilDataType = 1
    ilRet = btrCreKey(hmArb, ilKeyOffset + 1, ilKeySize, ilKeyFlag, ilDataType, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Key Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Key Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    'Key1, Seg 5: StartHour
    ilKeyOffset = 13
    ilKeySize = 2
    ilKeyFlag = &H103
    ilDataType = 1
    ilRet = btrCreKey(hmArb, ilKeyOffset + 1, ilKeySize, ilKeyFlag, ilDataType, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Key Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Key Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    ilRet = btrCreCreate(hmArb, sgDBPath & smTempFileName, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmMsg, "Create Failed for " & sgDBPath & smTempFileName & " import aborted"
        Screen.MousePointer = vbDefault
        MsgBox "Create Failed for " & sgDBPath & smTempFileName & " error # " & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Error"
        mCreateFile = False
        Exit Function
    End If
    ilRet = btrClose(hmArb)
    btrDestroy hmArb
    mCreateFile = True

End Function

Private Function mReadImport() As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim slRecordID As String
    Dim slReportType As String
    Dim slBookName As String
    Dim slBookPeriod As String
    Dim ilFirstLine As Integer
    Dim ilFound As Integer
    Dim ilUpper As Integer
    Dim ilValue As Integer
    Dim ilPos As Integer
    Dim slDayNames As String
    Dim slHours As String
    Dim slStartDay As String
    Dim slEndDay As String
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim slStartHour As String
    Dim slEndHour As String
    Dim ilStartHour As Integer
    Dim ilEndHour As Integer
    Dim llDaypartID As Long
    Dim llStationID As Long
    Dim llAudID As Long
    Dim ilStation As Integer
    Dim ilDP As Integer
    Dim ilAud As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    'Dim iAs Integer
    Dim llPercent As Long
    Dim slEstType As String
    Dim ilHour As Integer
    Dim ilUpdate As Integer
    Dim tlArb As Arbitron
    
    ilEof = False
    ilFirstLine = True
    plcGauge(1).Value = 0
    lmFileProcessedRecs = 0
    ReDim tmArbAudChar(0 To 0) As ARBAUDCHAR
    ReDim tmArbCallLetters(0 To 0) As ARBCALLLETTERS
    ReDim tmArbDaypart(0 To 0) As ARBDAYPART
    ReDim tmArb(0 To 0) As Arbitron
    Do
        ilRet = 0
        'On Error GoTo mReadImportErr:
        If EOF(hmFrom) Then
            ilEof = True
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilEof = True
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                DoEvents
                If imTerminate Then
                    mReadImport = -1
                    Exit Function
                End If
                slRecordID = Trim$(Left$(slLine, 1))
                Select Case slRecordID
                    Case "A"    'Copy Right
                        'Test that Create CD
                        slReportType = Trim$(Mid(slLine, 2, 2))
                        slBookName = Trim$(Mid(slLine, 19, 36))
                        slBookPeriod = Trim$(Mid(slLine, 55, 11)) & "-" & Trim$(Mid(slLine, 66, 11))
                        If StrComp(slReportType, smReportType, vbTextCompare) <> 0 Then
                            MsgBox "Wrong CD inserted to drive", vbCritical + vbOKOnly, "Wrong CD"
                            mReadImport = 1
                            Exit Function
                        End If
                    Case "D"    'Audience Char
                        'Create Table
                        ilUpper = UBound(tmArbAudChar)
                        tmArbAudChar(ilUpper).lID = Val(Trim$(Mid$(slLine, 4, 6)))
                        tmArbAudChar(ilUpper).sName = Trim$(Mid$(slLine, 60, 10))
                        tmArbAudChar(ilUpper).iBucketID = 0
                        If Trim$(Mid$(slLine, 70, 10)) = "" Then    'Bypass Qualitative
                            ilFound = False
                            For ilLoop = 0 To ilUpper - 1 Step 1
                                If StrComp(tmArbAudChar(ilLoop).sName, tmArbAudChar(ilUpper).sName, vbTextCompare) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                Select Case UCase(Trim$(tmArbAudChar(ilUpper).sName))
                                    Case "BOYS 12-17"
                                        tmArbAudChar(ilUpper).iBucketID = 1
                                    Case "MEN 18-20"
                                        tmArbAudChar(ilUpper).iBucketID = 2
                                    Case "MEN 21-24"
                                        tmArbAudChar(ilUpper).iBucketID = 3
                                    Case "MEN 25-34"
                                        tmArbAudChar(ilUpper).iBucketID = 4
                                    Case "MEN 35-44"
                                        tmArbAudChar(ilUpper).iBucketID = 5
                                    Case "MEN 45-49"
                                        tmArbAudChar(ilUpper).iBucketID = 6
                                    Case "MEN 50-54"
                                        tmArbAudChar(ilUpper).iBucketID = 7
                                    Case "MEN 55-64"
                                        tmArbAudChar(ilUpper).iBucketID = 8
                                    Case "MEN 65+"
                                        tmArbAudChar(ilUpper).iBucketID = 9
                                    Case "GRLS 12-17"
                                        tmArbAudChar(ilUpper).iBucketID = 10
                                    Case "WMN 18-20"
                                        tmArbAudChar(ilUpper).iBucketID = 11
                                    Case "WMN 21-24"
                                        tmArbAudChar(ilUpper).iBucketID = 12
                                    Case "WMN 25-34"
                                        tmArbAudChar(ilUpper).iBucketID = 13
                                    Case "WMN 35-44"
                                        tmArbAudChar(ilUpper).iBucketID = 14
                                    Case "WMN 45-49"
                                        tmArbAudChar(ilUpper).iBucketID = 15
                                    Case "WMN 50-54"
                                        tmArbAudChar(ilUpper).iBucketID = 16
                                    Case "WMN 55-64"
                                        tmArbAudChar(ilUpper).iBucketID = 17
                                    Case "WMN 65+"
                                        tmArbAudChar(ilUpper).iBucketID = 18
                                End Select
                                If tmArbAudChar(ilUpper).iBucketID > 0 Then
                                    ReDim Preserve tmArbAudChar(0 To ilUpper + 1) As ARBAUDCHAR
                                End If
                            End If
                        End If
                    Case "G"    'InTab
                        'Ignore
                    Case "H"    'Populate by Market
                        'Ignore at this time
                    Case "J"    'Station Info
                        'Create Table
                        ilUpper = UBound(tmArbCallLetters)
                        tmArbCallLetters(ilUpper).lID = Val(Trim$(Mid$(slLine, 16, 6)))
                        tmArbCallLetters(ilUpper).sName = Trim$(Mid$(slLine, 52, 4))
                        tmArbCallLetters(ilUpper).sBand = Trim$(Mid$(slLine, 56, 2))
                        If (tmArbCallLetters(ilUpper).sBand = "AM") Or (tmArbCallLetters(ilUpper).sBand = "FM") Then
                            ilFound = False
                            For ilLoop = 0 To ilUpper - 1 Step 1
                                If (StrComp(tmArbCallLetters(ilLoop).sName, tmArbCallLetters(ilUpper).sName, vbTextCompare) = 0) And (StrComp(tmArbCallLetters(ilLoop).sBand, tmArbCallLetters(ilUpper).sBand, vbTextCompare) = 0) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                ReDim Preserve tmArbCallLetters(0 To ilUpper + 1) As ARBCALLLETTERS
                            End If
                        End If
                    Case "M"    'Station Within Combo
                        'Ignore
                    Case "P"    'Network Affiliation
                        'Ignore
                    Case "S"    'Daypart Record
                        'Create Table
                        ilUpper = UBound(tmArbDaypart)
                        tmArbDaypart(ilUpper).lID = Val(Trim$(Mid$(slLine, 4, 4)))
                        slStr = Trim$(Mid$(slLine, 8, 30))
                        ilValue = Val(Trim$(Mid$(slLine, 38, 3)))
                        If (ilValue = 4) Or (ilValue = 16) Or (ilValue = 20) Or (ilValue = 80) Then
                            ilPos = InStr(1, slStr, " ", vbTextCompare)
                            If ilPos > 0 Then
                                slDayNames = Trim$(Left$(slStr, ilPos - 1))
                                slHours = Trim$(Mid$(slStr, ilPos + 1))
                                ilPos = InStr(1, slDayNames, "-", vbTextCompare)
                                If ilPos > 0 Then
                                    slStartDay = Left$(slDayNames, ilPos - 1)
                                    slEndDay = Mid$(slDayNames, ilPos + 1)
                                Else
                                    slStartDay = slDayNames
                                    slEndDay = slDayNames
                                End If
                                slStartDay = UCase(slStartDay)
                                slEndDay = UCase(slEndDay)
                                ilStartDay = Switch(slStartDay = "MON", 1, slStartDay = "TUE", 1, slStartDay = "WED", 1, slStartDay = "THU", 1, slStartDay = "FRI", 1, slStartDay = "SAT", 6, slStartDay = "SUN", 7)
                                ilEndDay = Switch(slEndDay = "MON", 1, slEndDay = "TUE", 1, slEndDay = "WED", 1, slEndDay = "THU", 1, slEndDay = "FRI", 1, slEndDay = "SAT", 6, slEndDay = "SUN", 7)
                                ilStartHour = -1
                                ilEndHour = -1
                                ilPos = InStr(1, slHours, "-", vbTextCompare)
                                If ilPos > 0 Then
                                    slStartHour = Left$(slHours, ilPos - 1)
                                    slEndHour = Mid$(slHours, ilPos + 1)
                                    If StrComp(slStartHour, "MID", vbTextCompare) = 0 Then
                                        slStartHour = "12AM"
                                    End If
                                    If StrComp(slEndHour, "MID", vbTextCompare) = 0 Then
                                        slEndHour = "12AM"
                                    End If
                                    ilPos = InStr(1, slStartHour, "AM", vbTextCompare)
                                    If ilPos > 0 Then
                                        ilStartHour = Val(Left(slStartHour, ilPos - 1))
                                        If ilStartHour = 12 Then
                                            ilStartHour = 0
                                        End If
                                    Else
                                        ilPos = InStr(1, slStartHour, "PM", vbTextCompare)
                                        If ilPos > 0 Then
                                            ilStartHour = Val(Left(slStartHour, ilPos - 1))
                                            If ilStartHour <> 12 Then
                                                ilStartHour = ilStartHour + 12
                                            End If
                                        End If
                                    End If
                                    ilPos = InStr(1, slEndHour, "AM", vbTextCompare)
                                    If ilPos > 0 Then
                                        ilEndHour = Val(Left(slEndHour, ilPos - 1))
                                        If ilEndHour = 12 Then
                                            ilEndHour = 24
                                        End If
                                    Else
                                        ilPos = InStr(1, slEndHour, "PM", vbTextCompare)
                                        If ilPos > 0 Then
                                            ilEndHour = Val(Left(slEndHour, ilPos - 1))
                                            If ilEndHour <> 12 Then
                                                ilEndHour = ilEndHour + 12
                                            End If
                                        End If
                                    End If
                                    ilFound = False
                                    If (ilStartDay = 1) And (ilEndDay = 1) And (ilStartHour = ilEndHour - 1) Then
                                        ilFound = True
                                    End If
                                    If (ilStartDay = 6) And (ilEndDay = 6) And (ilStartHour = ilEndHour - 1) Then
                                        ilFound = True
                                    End If
                                    If (ilStartDay = 7) And (ilEndDay = 7) And (ilStartHour = ilEndHour - 1) Then
                                        ilFound = True
                                    End If
                                    If (ilStartDay = 1) And (ilEndDay = 1) And (ilStartHour = 1) And (ilEndHour = 5) Then
                                        ilFound = True
                                    End If
                                    If (ilStartDay = 6) And (ilEndDay = 6) And (ilStartHour = 1) And (ilEndHour = 5) Then
                                        ilFound = True
                                    End If
                                    If (ilStartDay = 7) And (ilEndDay = 7) And (ilStartHour = 1) And (ilEndHour = 5) Then
                                        ilFound = True
                                    End If
                                    If ilFound Then
                                        tmArbDaypart(ilUpper).iDayIndex = ilStartDay
                                        tmArbDaypart(ilUpper).iStartHour = ilStartHour
                                        tmArbDaypart(ilUpper).iEndHour = ilEndHour
                                        ilFound = False
                                        For ilLoop = 0 To ilUpper - 1 Step 1
                                            If (tmArbDaypart(ilLoop).iDayIndex = tmArbDaypart(ilUpper).iDayIndex) And (tmArbDaypart(ilLoop).iStartHour = tmArbDaypart(ilUpper).iStartHour) And (tmArbDaypart(ilLoop).iEndHour = tmArbDaypart(ilUpper).iEndHour) Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        Next ilLoop
                                        If Not ilFound Then
                                            ReDim Preserve tmArbDaypart(0 To ilUpper + 1) As ARBDAYPART
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Case "V"
                        'Create Record
'                        ilUpper = UBound(tmArb)
'                        slEstType = Mid$(slLine, 14, 1)
'                        If slEstType = "1" Then
'                            llStationID = Val(Trim$(Mid$(slLine, 28, 6)))
'                            For ilStation = 0 To UBound(tmArbCallLetters) - 1 Step 1
'                                If llStationID = tmArbCallLetters(ilStation).lID Then
'                                    tmArb(ilUpper).sCallLetters = tmArbCallLetters(ilStation).sName
'                                    tmArb(ilUpper).sBand = tmArbCallLetters(ilStation).sBand
'                                    llDaypartID = Val(Trim$(Mid$(slLine, 15, 4)))
'                                    For ilDP = 0 To UBound(tmArbDaypart) - 1 Step 1
'                                        If llDaypartID = tmArbDaypart(ilDP).lID Then
'                                            tmArb(ilUpper).iDayIndex = tmArbDaypart(ilDP).iDayIndex
'                                            tmArb(ilUpper).iStartHour = tmArbDaypart(ilDP).iStartHour
'                                            tmArb(ilUpper).iEndHour = tmArbDaypart(ilDP).iEndHour
'                                            llAudID = Val(Trim$(Mid$(slLine, 20, 6)))
'                                            For ilAud = 0 To UBound(tmArbAudChar) - 1 Step 1
'                                                If llAudID = tmArbAudChar(ilAud).lID Then
'                                                    ilFound = False
'                                                    For ilLoop = 0 To ilUpper - 1 Step 1
'                                                        If (tmArb(ilLoop).iDayIndex = tmArb(ilUpper).iDayIndex) And (tmArb(ilLoop).iStartHour = tmArb(ilUpper).iStartHour) And (tmArb(ilLoop).iEndHour = tmArb(ilUpper).iEndHour) Then
'                                                            If (StrComp(tmArb(ilLoop).sCallLetters, tmArb(ilUpper).sCallLetters, vbTextCompare) = 0) And (StrComp(tmArb(ilLoop).sBand, tmArb(ilUpper).sBand, vbTextCompare) = 0) Then
'                                                                ilFound = True
'                                                                slStr = Mid$(slLine, 13, 1)
'                                                                If tmArb(ilLoop).sAreaCovered = slStr Then
'                                                                    tmArb(ilLoop).lDemo(tmArbAudChar(ilAud).iBucketID) = Val(Trim$(Mid$(slLine, 35, 8)))
'                                                                Else
'                                                                    If slStr = "3" Then
'                                                                        tmArb(ilLoop).sAreaCovered = slStr
'                                                                        For i= 1 To 18 Step 1
'                                                                            tmArb(ilLoop).lDemo(ilSet) = 0
'                                                                        Next ilSet
'                                                                        tmArb(ilLoop).lDemo(tmArbAudChar(ilAud).iBucketID) = Val(Trim$(Mid$(slLine, 35, 8)))
'                                                                    End If
'                                                                End If
'                                                                Exit For
'                                                            End If
'                                                        End If
'                                                    Next ilLoop
'                                                    If Not ilFound Then
'                                                        slStr = Mid$(slLine, 13, 1)
'                                                        If (slStr = "1") Or (slStr = "3") Then
'                                                            tmArb(ilUpper).sAreaCovered = slStr
'                                                            slStr = Trim$(Mid$(slStr, 2, 2))
'                                                            If (slStr = "2") Or (slStr = "3") Then
'                                                                tmArb(ilUpper).sDataType = slStr
'                                                            Else
'                                                                tmArb(ilUpper).sDataType = "1"
'                                                            End If
'                                                            tmArb(ilUpper).sDemoType = "A"
'                                                            tmArb(ilUpper).sForm = "8"
'                                                            For i= 1 To 18 Step 1
'                                                                tmArb(ilUpper).lDemo(ilSet) = 0
'                                                            Next ilSet
'                                                            tmArb(ilUpper).lDemo(tmArbAudChar(ilAud).iBucketID) = Val(Trim$(Mid$(slLine, 35, 8)))
'                                                            ReDim Preserve tmArb(0 To ilUpper + 1) As Arbitron
'                                                        End If
'                                                    End If
'                                                    Exit For
'                                                End If
'                                            Next ilAud
'                                            Exit For
'                                        End If
'                                    Next ilDP
'                                    Exit For
'                                End If
'                            Next ilStation
'                        End If
                    Case Else
                        'Write Error message
                        If ilFirstLine Then
                            MsgBox "Wrong CD inserted to drive", vbCritical + vbOKOnly, "Wrong CD"
                            mReadImport = 1
                            Exit Function
                        Else
                            Print #hmMsg, "Unknown Record ID " & slLine
                        End If
                End Select
            End If
            ilFirstLine = False
        End If
        lmFileProcessedRecs = lmFileProcessedRecs + Len(slLine) + 2
        llPercent = (lmFileProcessedRecs) \ lmFileLen
        If llPercent >= 100 Then
            llPercent = 100
        End If
        plcGauge(1).Value = llPercent
        DoEvents
    Loop Until ilEof
    For ilLoop = 0 To UBound(tmArb) - 1 Step 1
        ilStartHour = tmArb(ilLoop).iStartHour
        ilEndHour = tmArb(ilLoop).iEndHour
        For ilHour = ilStartHour To ilEndHour - 1 Step 1
            tmArb(ilLoop).iStartHour = ilHour
            tmArb(ilLoop).iEndHour = ilHour + 1
            If imUpdateFile Then
                tmArbSrchKey1.sCallLetters = tmArb(ilLoop).sCallLetters
                tmArbSrchKey1.sBand = tmArb(ilLoop).sBand
                tmArbSrchKey1.sDataType = tmArb(ilLoop).sDataType
                tmArbSrchKey1.iDayIndex = tmArb(ilLoop).iDayIndex
                tmArbSrchKey1.iStartHour = tmArb(ilLoop).iStartHour
                ilRet = btrGetEqual(hmArb, tlArb, imArbRecLen, tmArbSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = 0 Then
                    If tlArb.sDemoType = tmArb(ilLoop).sDemoType Then
                        ilUpdate = True
                    Else
                        Do
                            ilRet = btrGetNext(hmArb, tlArb, imArbRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                            If (ilRet = 0) And (tlArb.sCallLetters = tmArb(ilLoop).sCallLetters) And (tlArb.sBand = tmArb(ilLoop).sBand) And (tlArb.sDataType = tmArb(ilLoop).sDataType) And (tlArb.iDayIndex = tmArb(ilLoop).iDayIndex) And (tlArb.iStartHour = tmArb(ilLoop).iStartHour) Then
                                If tlArb.sDemoType = tmArb(ilLoop).sDemoType Then
                                    ilUpdate = True
                                    Exit Do
                                End If
                            Else
                                ilUpdate = False
                                Exit Do
                            End If
                        Loop
                    End If
                Else
                    ilUpdate = False
                End If
            Else
                ilUpdate = False
            End If
            If ilUpdate Then
                tmArb(ilLoop).lCode = tlArb.lCode
                ilRet = btrUpdate(hmArb, tmArb(ilLoop), imArbRecLen)
            Else
                tmArb(ilLoop).lCode = 0
                ilRet = btrInsert(hmArb, tmArb(ilLoop), imArbRecLen, INDEXKEY0)
            End If
            If ilRet <> 0 Then
                Print #hmMsg, "Pervasive Error " & Str(ilRet) & " Import Terminated"
                mReadImport = -1
                Exit Function
            End If
        Next ilHour
    Next ilLoop
    mReadImport = 0
    Exit Function
mReadImportErr:
    ilRet = Err.Number
    Resume Next
End Function
