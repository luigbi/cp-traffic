VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form EngrFileCheck 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5310
   ClientLeft      =   315
   ClientTop       =   1710
   ClientWidth     =   9030
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
   ForeColor       =   &H00000000&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   9030
   Begin VB.PictureBox pbcArial 
      Height          =   180
      Left            =   6945
      ScaleHeight     =   120
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   4530
      Visible         =   0   'False
      Width           =   255
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   255
      Left            =   285
      TabIndex        =   20
      Top             =   4410
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CheckBox ckcTest 
      Caption         =   "Key L-F"
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
      Height          =   195
      Index           =   3
      Left            =   7905
      TabIndex        =   19
      Top             =   315
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox ckcTest 
      Caption         =   "Key F-L"
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
      Height          =   195
      Index           =   2
      Left            =   6915
      TabIndex        =   18
      Top             =   315
      Value           =   1  'Checked
      Width           =   1005
   End
   Begin VB.CheckBox ckcTest 
      Caption         =   "Step L-F"
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
      Height          =   195
      Index           =   1
      Left            =   5835
      TabIndex        =   17
      Top             =   315
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.CheckBox ckcTest 
      Caption         =   "Step F-L"
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
      Height          =   195
      Index           =   0
      Left            =   4770
      TabIndex        =   16
      Top             =   315
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.OptionButton rbcCF 
      Caption         =   "Check"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   7305
      TabIndex        =   14
      Top             =   90
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.OptionButton rbcCF 
      Caption         =   "Fix"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8235
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   675
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Files"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   285
      TabIndex        =   3
      Top             =   3945
      Width           =   1350
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8400
      Top             =   4635
   End
   Begin VB.PictureBox plcChkFix 
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
      Height          =   3210
      Left            =   150
      ScaleHeight     =   3150
      ScaleWidth      =   8625
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   8685
      Begin VB.ListBox lbcResult 
         Appearance      =   0  'Flat
         Height          =   2970
         Left            =   1470
         TabIndex        =   12
         Top             =   60
         Width           =   7155
      End
      Begin VB.FileListBox lbcFile 
         Appearance      =   0  'Flat
         Height          =   2970
         Left            =   45
         MultiSelect     =   2  'Extended
         Pattern         =   "*.eng"
         TabIndex        =   2
         Top             =   60
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmcChk 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   3075
      TabIndex        =   5
      Top             =   4830
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4935
      TabIndex        =   6
      Top             =   4830
      Width           =   1140
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1905
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1905
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6525
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5310
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lacDBPath 
      Appearance      =   0  'Flat
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
      Height          =   180
      Left            =   60
      TabIndex        =   15
      Top             =   5100
      Width           =   6060
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2850
      TabIndex        =   4
      Top             =   3990
      Width           =   5970
   End
   Begin VB.Label lacMissed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3120
      TabIndex        =   11
      Top             =   4470
      Width           =   5685
   End
   Begin VB.Label lacTest 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3090
      TabIndex        =   10
      Top             =   4230
      Width           =   5730
   End
End
Attribute VB_Name = "EngrFileCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Check/Fix MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmTo As Integer   'From file hanle
Dim hmFile As Integer           'file handle
Dim tmFile As FILECHECK              'Largest file
Dim imFileRecLen As Integer     'record length
Dim hmFileNew As Integer
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim lmNowDate As Long   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim lmPriRecPos() As Long
Dim lmSecRecPos() As Long
Dim lmPriEngrFileCheckCount As Long
Dim lmSecEngrFileCheckCount As Long
Dim imFirstActivate As Integer
Dim llMaxWidth As Long

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcFile.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcFile.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    If imProcessing Then
        imCancelled = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcChk_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llLoop As Long
    Dim llTotalFile As Long
    Dim llFileProc As Long
    Dim slName As String
    Dim ilNoKeys As Integer
    Dim ilKeys As Integer
    Dim llTotalRec As Long
    Dim llRecNo As Long
    Dim ilFileIndex As Integer
    Dim llPercent As Long
    Dim ilTestOnly As Integer
    Dim slToFile As String
    Dim slError As String
    Dim slNewName As String
    Dim slOldName As String
    Dim slNewFile As String
    Dim slStamp As String
    Dim ilPos As Integer
    Dim slMsg As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim llRecPos As Long
    'Dim ilNoRows As Integer
    'Dim ilRow As Integer
    'Dim ilCol As Integer
    Dim llFirstRecNo As Long
    Dim llRecAdded As Long
    Dim llRecProc As Long
    Dim ilRes As Integer
    Dim llValue As Long
    Dim llRg As Long
    Dim llRet As Long
    Dim slStr As String
    Dim slExtension As String
    Dim ilExtNumber As Integer
    Dim llMod As Long

    Screen.MousePointer = vbHourglass
    llMaxWidth = 0
    If rbcCF(0).Value Then
        ilTestOnly = True   'Check Only
    Else
        ilTestOnly = False  'Fix
    End If
    ilRet = 0
    On Error GoTo cmcChkErr:
    If ilTestOnly Then
        'D.S. 03/18/03 slToFile = sgExportPath & "EngrFileCheck.Txt"
        slToFile = sgDBPath & "Messages\" & "FileCheck.Txt"
    Else
        'D.S. 03/18/03 slToFile = sgExportPath & "BTRFix.Txt"
        slToFile = sgDBPath & "Messages\" & "FileFix.Txt"
    End If
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        slFileDate = Format$(slDateTime, "ddddd")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            On Error GoTo cmcChkErr:
            hmTo = FreeFile
            Open slToFile For Append As hmTo
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOK + vbCritical, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            On Error GoTo cmcChkErr:
            hmTo = FreeFile
            Open slToFile For Output As hmTo
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOK + vbCritical, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo cmcChkErr:
        hmTo = FreeFile
        Open slToFile For Output As hmTo
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOK + vbCritical, "Open Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo 0
    If ilTestOnly Then
        Print #hmTo, "File Check " & Format$(gNow(), "ddddd") & " at " & Format$(Now, "ttttt")
    Else
        Print #hmTo, "File Fix " & Format$(gNow(), "ddddd") & " at " & Format$(Now, "ttttt")
    End If
    Print #hmTo, ""

    llTotalFile = 0
    llFileProc = 0
    For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
        If lbcFile.Selected(ilLoop) Then
            llTotalFile = llTotalFile + 1
        End If
    Next ilLoop
    lbcResult.Clear
    imProcessing = True
    plcGauge.Value = 0
    ilFileIndex = -1
    For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
        If lbcFile.Selected(ilLoop) Then
            slName = lbcFile.List(ilLoop)
            lbcResult.AddItem slName
            ilFileIndex = ilFileIndex + 1
            If Not ilTestOnly Then
                hmFile = CBtrvTable(ONEHANDLE) 'CBtrvObj()
                ilRet = btrOpen(hmFile, "", sgDBPath & slName, BTRV_OPEN_VERIFY, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    MsgBox "Unable to open " & slName & "- Error code " & Trim$(Str$(ilRet))
                    Screen.MousePointer = vbDefault
                    imProcessing = False
                    'ilRet = btrClose(hmFile)
                    btrDestroy hmFile
                    Print #hmTo, ""
                    Print #hmTo, "Incompleted, Open Error"
                    Close #hmTo
                    Exit Sub
                End If
                llTotalRec = btrRecords(hmFile)
                'ReDim lmRecNoMoved(1 To llTotalRec)  As Long
                'ilNoRows = llTotalRec \ 32000 + 1
                ReDim lmPriRecPos(0 To llTotalRec) As Long
                lmPriEngrFileCheckCount = 0
                For llLoop = LBound(lmPriRecPos) To UBound(lmPriRecPos) Step 1
                    lmPriRecPos(llLoop) = 0
                Next llLoop
                ReDim lmSecRecPos(0 To llTotalRec) As Long
                lmSecEngrFileCheckCount = 0
                For llLoop = LBound(lmSecRecPos) To UBound(lmSecRecPos) Step 1
                    lmSecRecPos(llLoop) = 0
                Next llLoop
                ilPos = InStr(slName, ".")
                slNewName = Left$(slName, ilPos - 1) & "_.New"
                ilRet = btrClone(hmFile, sgDBPath & slNewName, 0)
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    Close #hmTo
                    MsgBox "Unable to clone " & slNewName & "- Error code " & Trim$(Str$(ilRet))
                    imProcessing = False
                    Print #hmTo, ""
                    Print #hmTo, "Incompleted, Clone Error"
                    Close #hmTo
                    Exit Sub
                End If
                'ilRet = btrClose(hmFile)
                btrDestroy hmFile
            End If
            hmFile = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmFile, "", sgDBPath & slName, BTRV_OPEN_VERIFY, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                MsgBox "Unable to open " & slName & "- Error code " & Trim$(Str$(ilRet))
                Screen.MousePointer = vbDefault
                imProcessing = False
                'ilRet = btrClose(hmFile)
                btrDestroy hmFile
                Print #hmTo, ""
                Print #hmTo, "Incompleted, Open Error"
                Close #hmTo
                Exit Sub
            End If
            If Not ilTestOnly Then
                hmFileNew = CBtrvTable(ONEHANDLE) 'CBtrvObj()
                ilRet = btrOpen(hmFileNew, "", sgDBPath & slNewName, BTRV_OPEN_VERIFY, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    MsgBox "Unable to open " & slNewName & "- Error code " & Trim$(Str$(ilRet))
                    Screen.MousePointer = vbDefault
                    imProcessing = False
                    'ilRet = btrClose(hmFile)
                    btrDestroy hmFile
                    'ilRet = btrClose(hmFileNew)
                    btrDestroy hmFileNew
                    Print #hmTo, ""
                    Print #hmTo, "Incompleted, Open Error"
                    Close #hmTo
                    Exit Sub
                End If
            End If
            lacProcessing.Caption = "Testing " & slName & " in First to Next order"
            DoEvents
            llTotalRec = btrRecords(hmFile)
            llMod = llTotalRec / 1000
            If llMod <= 0 Then
                llMod = 1
            End If
    llMod = 1000
            lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & " # Recs" & Str$(llTotalRec)
            slStr = lbcResult.List(ilFileIndex)
            If (pbcArial.TextWidth(slStr)) > llMaxWidth Then
                llMaxWidth = pbcArial.TextWidth(slStr)
                If llMaxWidth > lbcResult.Width Then
                    llValue = llMaxWidth / 15 + 120
                    llRg = 0
                    llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
                End If
            End If
            llRecNo = 0
            llRecAdded = 0
            llRecProc = 0
            'If ilTestOnly Then
                imFileRecLen = Len(tmFile)
            'Else
            '    imFileRecLen = btrRecordLength(hmFile)
            'End If
            If ckcTest(0).Value = vbChecked Then
                ilRet = btrStepFirst(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE)
                Do While ilRet = BTRV_ERR_NONE
                    llRecProc = llRecProc + 1
                    If Not ilTestOnly Then
                        ilRet = btrGetPosition(hmFile, llRecPos)
                        ilFound = False
                        'ilCol = 1
                        'ilRow = 1
                        'For llLoop = 1 To llRecNo Step 1
                        '    If lmRecNoMoved(ilRow, ilCol) = llRecPos Then
                        '        ilFound = True
                        '        Exit For
                        '    End If
                        '    ilCol = ilCol + 1
                        '    If ilCol > 32000 Then
                        '        ilCol = 1
                        '        ilRow = ilRow + 1
                        '    End If
                        'Next llLoop
                        If Not ilFound Then
                            llRecNo = llRecNo + 1
                            llRecAdded = llRecAdded + 1
                            'ilRow = (llRecNo - 1) \ 32000 + 1
                            'ilCol = llRecNo - 32000 * CLng((ilRow - 1))
                            'lmRecNoMoved(llRecNo) = llRecPos
                            ilRet = btrInsert(hmFileNew, tmFile, imFileRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                'lmRecNoMoved(llRecNo) = 0
                                llRecNo = llRecNo - 1
                                llRecAdded = llRecAdded - 1
                            Else
                                lmSecRecPos(lmSecEngrFileCheckCount) = llRecPos
                                lmSecEngrFileCheckCount = lmSecEngrFileCheckCount + 1
                            End If
                        End If
                    Else
                        llRecNo = llRecNo + 1
                    End If
                    If (llRecAdded Mod llMod) = 0 Then
                        lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                        DoEvents
                    End If
                    If imCancelled Then
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        'ilRet = btrClose(hmFile)
                        btrDestroy hmFile
                        If Not ilTestOnly Then
                            'ilRet = btrClose(hmFileNew)
                            btrDestroy hmFileNew
                        End If
                        Print #hmTo, ""
                        Print #hmTo, "Incompleted, Cancelled"
                        Close #hmTo
                        Exit Sub
                    End If
                    If llRecNo >= llTotalRec Then
                        Exit Do
                    End If
                    'If ilTestOnly Then
                        imFileRecLen = Len(tmFile)
                    'Else
                    '    imFileRecLen = btrRecordLength(hmFile)
                    'End If
                    ilRet = btrStepNext(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE)
                    Do While (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE)
                        If (llRecAdded Mod llMod) = 0 Then
                            lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                            DoEvents
                        End If
                        imFileRecLen = Len(tmFile)
                        ilRet = btrStepNext(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE)
                        llRecProc = llRecProc + 1
                        If llRecProc >= llTotalRec Then
                            Exit Do
                        End If
                    Loop
                Loop
                lacTest.Caption = "Added" & Str$(llRecAdded) & " of" & Str(llTotalRec)
                If llTotalRec = llRecNo Then
                    lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & " F-N: Ok"
                Else
                    If ilTestOnly Then
                        lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & " F-N: Error " & Trim$(Str$(ilRet)) & " #" & Str$(llRecNo)
                    Else
                        lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & " F-N: Error " & Trim$(Str$(ilRet)) & " # added" & Str$(llRecAdded)
                    End If
                End If
                slStr = lbcResult.List(ilFileIndex)
                If (pbcArial.TextWidth(slStr)) > llMaxWidth Then
                    llMaxWidth = pbcArial.TextWidth(slStr)
                    If llMaxWidth > lbcResult.Width Then
                        llValue = llMaxWidth / 15 + 120
                        llRg = 0
                        llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
                    End If
                End If
                If Not ilTestOnly Then
                    For llLoop = LBound(lmSecRecPos) To lmSecEngrFileCheckCount - 1 Step 1
                        lmPriRecPos(lmPriEngrFileCheckCount) = lmSecRecPos(llLoop)
                        lmPriEngrFileCheckCount = lmPriEngrFileCheckCount + 1
                    Next llLoop
                    If lmSecEngrFileCheckCount > 0 Then
                        ArraySortTyp fnAV(lmPriRecPos(), 0), lmPriEngrFileCheckCount, 0, LenB(lmPriRecPos(0)), 0, -2, 0
                    End If
                    lmSecEngrFileCheckCount = 0
                    For llLoop = LBound(lmSecRecPos) To UBound(lmSecRecPos) Step 1
                        lmSecRecPos(llLoop) = 0
                    Next llLoop
                End If
            End If
            lacProcessing.Caption = "Testing " & slName & " in Last to Previous order"
            DoEvents
            If imCancelled Then
                imProcessing = False
                Screen.MousePointer = vbDefault
                'ilRet = btrClose(hmFile)
                btrDestroy hmFile
                If Not ilTestOnly Then
                    'ilRet = btrClose(hmFileNew)
                    btrDestroy hmFileNew
                End If
                slError = lbcResult.List(ilFileIndex)
                Print #hmTo, slError
                Print #hmTo, ""
                Print #hmTo, "Incompleted, Cancelled"
                Close #hmTo
                Exit Sub
            End If
            'try creating once again, this time last, prev
            'If ilTestOnly Then
                imFileRecLen = Len(tmFile)
            'Else
            '    imFileRecLen = btrRecordLength(hmFile)
            'End If
            If ckcTest(1).Value = vbChecked Then
                If ilTestOnly Or (btrRecords(hmFile) <> btrRecords(hmFileNew)) Then
                    If ilTestOnly Then
                        llRecNo = 0
                    Else
                        llFirstRecNo = llRecNo
                    End If
                    llRecProc = 0
                    ilRet = btrStepLast(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE)
                    Do While ilRet = BTRV_ERR_NONE
                        llRecProc = llRecProc + 1
                        If Not ilTestOnly Then
                            ilRet = btrGetPosition(hmFile, llRecPos)
                            ilFound = False
                            'ilCol = 1
                            'ilRow = 1
                            'For llLoop = 1 To llRecNo  Step 1
                            '    If lmRecNoMoved(llLoop) = llRecPos Then
                            '        ilFound = True
                            '        Exit For
                            '    End If
                            '    'ilCol = ilCol + 1
                            '    'If ilCol > 32000 Then
                            '    '    ilCol = 1
                            '    '    ilRow = ilRow + 1
                            '    'End If
                            'Next llLoop
                            If ckcTest(0).Value = vbChecked Then
                                'If llRecNo > 0 Then
                                '    'ArraySortTyp fnAV(lmRecNoMoved(), 0), llRecNo, 0, LenB(lmRecNoMoved(0)), 0, -2, 0
                                '    If mBinarySearch(llRecPos, llRecNo) <> -1 Then
                                '        ilFound = True
                                '    End If
                                'End If
                                ilFound = mCheckRecPos(llRecPos)
                            End If
                            If Not ilFound Then
                                llRecNo = llRecNo + 1
                                llRecAdded = llRecAdded + 1
                                'ilRow = (llRecNo - 1) \ 32000 + 1
                                'ilCol = llRecNo - 32000 * CLng((ilRow - 1))
                                'lmRecNoMoved(llRecNo) = llRecPos
                                ilRet = btrInsert(hmFileNew, tmFile, imFileRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    llRecNo = llRecNo - 1
                                    llRecAdded = llRecAdded - 1
                                Else
                                    lmSecRecPos(lmSecEngrFileCheckCount) = llRecPos
                                    lmSecEngrFileCheckCount = lmSecEngrFileCheckCount + 1
                                End If
                            End If
                        Else
                            llRecNo = llRecNo + 1
                        End If
                        If (llRecAdded Mod llMod) = 0 Then
                            lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                            DoEvents
                        End If
                        If imCancelled Then
                            imProcessing = False
                            Screen.MousePointer = vbDefault
                            'ilRet = btrClose(hmFile)
                            btrDestroy hmFile
                            If Not ilTestOnly Then
                                'ilRet = btrClose(hmFileNew)
                                btrDestroy hmFileNew
                            End If
                            slError = lbcResult.List(ilFileIndex)
                            Print #hmTo, slError
                            Print #hmTo, ""
                            Print #hmTo, "Incompleted, Cancelled"
                            Close #hmTo
                            Exit Sub
                        End If
                        If llRecNo >= llTotalRec Then
                            Exit Do
                        End If
                        'If ilTestOnly Then
                            imFileRecLen = Len(tmFile)
                        'Else
                        '    imFileRecLen = btrRecordLength(hmFile)
                        'End If
                        ilRet = btrStepPrevious(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE)
                        Do While (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE)
                            If (llRecAdded Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            If (llRecProc Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            imFileRecLen = Len(tmFile)
                            ilRet = btrStepPrevious(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE)
                            llRecProc = llRecProc + 1
                            DoEvents
                            If llRecProc >= llTotalRec Then
                                Exit Do
                            End If
                        Loop
                    Loop
                    lacTest.Caption = "Added" & Str$(llRecAdded) & " of" & Str(llTotalRec)
                    If llTotalRec = llRecNo Then
                        lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; L-P: Ok"
                    Else
                        If ilTestOnly Then
                            lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; L-P: Error " & Trim$(Str$(ilRet)) & " #" & Str$(llRecNo)
                        Else
                            lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; L-P: Error " & Trim$(Str$(ilRet)) & " # added" & Str$(llRecAdded)
                        End If
                    End If
                    slStr = lbcResult.List(ilFileIndex)
                    If (pbcArial.TextWidth(slStr)) > llMaxWidth Then
                        llMaxWidth = pbcArial.TextWidth(slStr)
                        If llMaxWidth > lbcResult.Width Then
                            llValue = llMaxWidth / 15 + 120
                            llRg = 0
                            llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
                        End If
                    End If
                    If (Not ilTestOnly) Then
                        For llLoop = LBound(lmSecRecPos) To lmSecEngrFileCheckCount - 1 Step 1
                            lmPriRecPos(lmPriEngrFileCheckCount) = lmSecRecPos(llLoop)
                            lmPriEngrFileCheckCount = lmPriEngrFileCheckCount + 1
                        Next llLoop
                        If lmSecEngrFileCheckCount > 0 Then
                            ArraySortTyp fnAV(lmPriRecPos(), 0), lmPriEngrFileCheckCount, 0, LenB(lmPriRecPos(0)), 0, -2, 0
                        End If
                        lmSecEngrFileCheckCount = 0
                        For llLoop = LBound(lmSecRecPos) To UBound(lmSecRecPos) Step 1
                            lmSecRecPos(llLoop) = 0
                        Next llLoop
                    End If
                End If
            End If
            ilNoKeys = btrIndexes(hmFile)
            If ckcTest(2).Value = vbChecked Then
                If ilTestOnly Or (btrRecords(hmFile) <> btrRecords(hmFileNew)) Then
                    For ilKeys = 0 To ilNoKeys - 1 Step 1
                        lacProcessing.Caption = "Testing First to Next " & slName & " Index number" & Str$(ilKeys)
                        DoEvents
                        If imCancelled Then
                            imProcessing = False
                            Screen.MousePointer = vbDefault
                            'ilRet = btrClose(hmFile)
                            btrDestroy hmFile
                            If Not ilTestOnly Then
                                'ilRet = btrClose(hmFileNew)
                                btrDestroy hmFileNew
                            End If
                            slError = lbcResult.List(ilFileIndex)
                            Print #hmTo, slError
                            Print #hmTo, ""
                            Print #hmTo, "Incompleted, Cancelled"
                            Close #hmTo
                            Exit Sub
                        End If
                        imFileRecLen = Len(tmFile)
                        If ilTestOnly Then
                            llRecNo = 0
                        End If
                        'llRecAdded = 0
                        llRecProc = 0
                        ilRet = btrGetFirst(hmFile, tmFile, imFileRecLen, ilKeys, BTRV_LOCK_NONE, SETFORREADONLY)
                        Do While ilRet = BTRV_ERR_NONE
                            DoEvents
                            llRecProc = llRecProc + 1
                            If Not ilTestOnly Then
                                ilRet = btrGetPosition(hmFile, llRecPos)
                                ilFound = False
                                ''ilCol = 1
                                ''ilRow = 1
                                'For llLoop = 1 To llRecNo Step 1
                                '    If lmRecNoMoved(llLoop) = llRecPos Then
                                '        ilFound = True
                                '        Exit For
                                '    End If
                                '    'ilCol = ilCol + 1
                                '    'If ilCol > 32000 Then
                                '    '    ilCol = 1
                                '    '    ilRow = ilRow + 1
                                '    'End If
                                'Next llLoop
                                If (ckcTest(0).Value = vbChecked) Or (ckcTest(1).Value = vbChecked) Or (ilKeys > 0) Then
                                    'If llRecNo > 0 Then
                                    '    'ArraySortTyp fnAV(lmRecNoMoved(), 0), llRecNo, 0, LenB(lmRecNoMoved(0)), 0, -2, 0
                                    '    If mBinarySearch(llRecPos, llRecNo) <> -1 Then
                                    '        ilFound = True
                                    '    End If
                                    'End If
                                    ilFound = mCheckRecPos(llRecPos)
                                End If
                                If Not ilFound Then
                                    llRecNo = llRecNo + 1
                                    llRecAdded = llRecAdded + 1
                                    'ilRow = (llRecNo - 1) \ 32000 + 1
                                    'ilCol = llRecNo - 32000 * CLng((ilRow - 1))
                                    'lmRecNoMoved(llRecNo) = llRecPos
                                    ilRet = btrInsert(hmFileNew, tmFile, imFileRecLen, INDEXKEY0)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        'lmRecNoMoved(llRecNo) = 0
                                        llRecNo = llRecNo - 1
                                        llRecAdded = llRecAdded - 1
                                    Else
                                        lmSecRecPos(lmSecEngrFileCheckCount) = llRecPos
                                        lmSecEngrFileCheckCount = lmSecEngrFileCheckCount + 1
                                    End If
                                End If
                            Else
                                llRecNo = llRecNo + 1
                            End If
                            If (llRecAdded Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            If (llRecProc Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            
                            If imCancelled Then
                                imProcessing = False
                                Screen.MousePointer = vbDefault
                                'ilRet = btrClose(hmFile)
                                btrDestroy hmFile
                                If Not ilTestOnly Then
                                    'ilRet = btrClose(hmFileNew)
                                    btrDestroy hmFileNew
                                End If
                                slError = lbcResult.List(ilFileIndex)
                                Print #hmTo, slError
                                Print #hmTo, ""
                                Print #hmTo, "Incompleted, Cancelled"
                                Close #hmTo
                                Exit Sub
                            End If
                            If llRecNo >= llTotalRec Then
                                Exit Do
                            End If
                            imFileRecLen = Len(tmFile)
                            ilRet = btrGetNext(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Do While (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE)
                                If (llRecAdded Mod llMod) = 0 Then
                                    lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                    DoEvents
                                End If
                                If (llRecProc Mod llMod) = 0 Then
                                    lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                    DoEvents
                                End If
                                
                                
                                imFileRecLen = Len(tmFile)
                                ilRet = btrGetNext(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                llRecProc = llRecProc + 1
                                If llRecProc >= llTotalRec Then
                                    Exit Do
                                End If
                                DoEvents
                            Loop
                        Loop
                        lacTest.Caption = "Added" & Str$(llRecAdded) & " of" & Str(llTotalRec)
                        If llTotalRec = llRecNo Then
                            lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; Key " & Str$(ilKeys) & " F-N: Ok"
                        Else
                            If ilTestOnly Then
                                lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; Key " & Str$(ilKeys) & " F-N: Error" & Trim$(Str$(ilRet)) & " #" & Str$(llRecNo)
                            Else
                                lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; Key " & Str$(ilKeys) & " F-N: Error" & Trim$(Str$(ilRet)) & " # added" & Str$(llRecAdded)
                            End If
                        End If
                        slStr = lbcResult.List(ilFileIndex)
                        If (pbcArial.TextWidth(slStr)) > llMaxWidth Then
                            llMaxWidth = pbcArial.TextWidth(slStr)
                            If llMaxWidth > lbcResult.Width Then
                                llValue = llMaxWidth / 15 + 120
                                llRg = 0
                                llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
                            End If
                        End If
                        If (Not ilTestOnly) Then
                            For llLoop = LBound(lmSecRecPos) To lmSecEngrFileCheckCount - 1 Step 1
                                lmPriRecPos(lmPriEngrFileCheckCount) = lmSecRecPos(llLoop)
                                lmPriEngrFileCheckCount = lmPriEngrFileCheckCount + 1
                            Next llLoop
                            If lmSecEngrFileCheckCount > 0 Then
                                ArraySortTyp fnAV(lmPriRecPos(), 0), lmPriEngrFileCheckCount, 0, LenB(lmPriRecPos(0)), 0, -2, 0
                            End If
                            lmSecEngrFileCheckCount = 0
                            For llLoop = LBound(lmSecRecPos) To UBound(lmSecRecPos) Step 1
                                lmSecRecPos(llLoop) = 0
                            Next llLoop
                        End If
                    Next ilKeys
                End If
            End If
            If ckcTest(3).Value = vbChecked Then
                If ilTestOnly Or (btrRecords(hmFile) <> btrRecords(hmFileNew)) Then
                    For ilKeys = 0 To ilNoKeys - 1 Step 1
                        lacProcessing.Caption = "Testing Last to Previous " & slName & " Index number" & Str$(ilKeys)
                        DoEvents
                        If imCancelled Then
                            imProcessing = False
                            Screen.MousePointer = vbDefault
                            'ilRet = btrClose(hmFile)
                            btrDestroy hmFile
                            If Not ilTestOnly Then
                                'ilRet = btrClose(hmFileNew)
                                btrDestroy hmFileNew
                            End If
                            slError = lbcResult.List(ilFileIndex)
                            Print #hmTo, slError
                            Print #hmTo, ""
                            Print #hmTo, "Incompleted, Cancelled"
                            Close #hmTo
                            Exit Sub
                        End If
                        imFileRecLen = Len(tmFile)
                        If ilTestOnly Then
                            llRecNo = 0
                        End If
                        'llRecAdded = 0
                        llRecProc = 0
                        ilRet = btrGetLast(hmFile, tmFile, imFileRecLen, ilKeys, BTRV_LOCK_NONE, SETFORREADONLY)
                        Do While ilRet = BTRV_ERR_NONE
                            llRecProc = llRecProc + 1
                            If Not ilTestOnly Then
                                ilRet = btrGetPosition(hmFile, llRecPos)
                                ilFound = False
                                ''ilCol = 1
                                ''ilRow = 1
                                'For llLoop = 1 To llRecNo Step 1
                                '    If lmRecNoMoved(llLoop) = llRecPos Then
                                '        ilFound = True
                                '        Exit For
                                '    End If
                                '    'ilCol = ilCol + 1
                                '    'If ilCol > 32000 Then
                                '    '    ilCol = 1
                                '    '    ilRow = ilRow + 1
                                '    'End If
                                'Next llLoop
                                If (ckcTest(0).Value = vbChecked) Or (ckcTest(1).Value = vbChecked) Or (ckcTest(2).Value = vbChecked) Or (ilKeys > 0) Then
                                    'If llRecNo > 0 Then
                                    '    'ArraySortTyp fnAV(lmRecNoMoved(), 0), llRecNo, 0, LenB(lmRecNoMoved(0)), 0, -2, 0
                                    '    If mBinarySearch(llRecPos, llRecNo) <> -1 Then
                                    '        ilFound = True
                                    '    End If
                                    'End If
                                    ilFound = mCheckRecPos(llRecPos)
                                End If
                                If Not ilFound Then
                                    llRecNo = llRecNo + 1
                                    llRecAdded = llRecAdded + 1
                                    'ilRow = (llRecNo - 1) \ 32000 + 1
                                    'ilCol = llRecNo - 32000 * CLng((ilRow - 1))
                                    'lmRecNoMoved(llRecNo) = llRecPos
                                    ilRet = btrInsert(hmFileNew, tmFile, imFileRecLen, INDEXKEY0)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        'lmRecNoMoved(llRecNo) = 0
                                        llRecNo = llRecNo - 1
                                        llRecAdded = llRecAdded - 1
                                    Else
                                        lmSecRecPos(lmSecEngrFileCheckCount) = llRecPos
                                        lmSecEngrFileCheckCount = lmSecEngrFileCheckCount + 1
                                    End If
                                End If
                            Else
                                llRecNo = llRecNo + 1
                            End If
                            If (llRecAdded Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            If (llRecProc Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            
                            If (llRecAdded Mod llMod) = 0 Then
                                lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                DoEvents
                            End If
                            If imCancelled Then
                                imProcessing = False
                                Screen.MousePointer = vbDefault
                                'ilRet = btrClose(hmFile)
                                btrDestroy hmFile
                                If Not ilTestOnly Then
                                    'ilRet = btrClose(hmFileNew)
                                    btrDestroy hmFileNew
                                End If
                                slError = lbcResult.List(ilFileIndex)
                                Print #hmTo, slError
                                Print #hmTo, ""
                                Print #hmTo, "Incompleted, Cancelled"
                                Close #hmTo
                                Exit Sub
                            End If
                            If llRecNo >= llTotalRec Then
                                Exit Do
                            End If
                            imFileRecLen = Len(tmFile)
                            ilRet = btrGetPrevious(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Do While (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE)
                                If (llRecAdded Mod llMod) = 0 Then
                                    lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                End If
                                If (llRecProc Mod llMod) = 0 Then
                                    lacTest.Caption = "Added" & Str$(llRecAdded) & " Processed" & Str$(llRecProc) & " of" & Str(llTotalRec)
                                End If
                                DoEvents
                                imFileRecLen = Len(tmFile)
                                ilRet = btrGetPrevious(hmFile, tmFile, imFileRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                llRecProc = llRecProc + 1
                                If llRecProc >= llTotalRec Then
                                    Exit Do
                                End If
                            Loop
                        Loop
                        lacTest.Caption = "Added" & Str$(llRecAdded) & " of" & Str(llTotalRec)
                        If llTotalRec = llRecNo Then
                            lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; Key " & Str$(ilKeys) & " L-P: Ok"
                        Else
                            If ilTestOnly Then
                                lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; Key " & Str$(ilKeys) & " L-P: Error" & Trim$(Str$(ilRet)) & " #" & Str$(llRecNo)
                            Else
                                lbcResult.List(ilFileIndex) = lbcResult.List(ilFileIndex) & "; Key " & Str$(ilKeys) & " L-P: Error" & Trim$(Str$(ilRet)) & " # added" & Str$(llRecAdded)
                            End If
                        End If
                        slStr = lbcResult.List(ilFileIndex)
                        If (pbcArial.TextWidth(slStr)) > llMaxWidth Then
                            llMaxWidth = pbcArial.TextWidth(slStr)
                            If llMaxWidth > lbcResult.Width Then
                                llValue = llMaxWidth / 15 + 120
                                llRg = 0
                                llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
                            End If
                        End If
                        If (Not ilTestOnly) Then
                            For llLoop = LBound(lmSecRecPos) To lmSecEngrFileCheckCount - 1 Step 1
                                lmPriRecPos(lmPriEngrFileCheckCount) = lmSecRecPos(llLoop)
                                lmPriEngrFileCheckCount = lmPriEngrFileCheckCount + 1
                            Next llLoop
                            If lmSecEngrFileCheckCount > 0 Then
                                ArraySortTyp fnAV(lmPriRecPos(), 0), lmPriEngrFileCheckCount, 0, LenB(lmPriRecPos(0)), 0, -2, 0
                            End If
                            lmSecEngrFileCheckCount = 0
                            For llLoop = LBound(lmSecRecPos) To UBound(lmSecRecPos) Step 1
                                lmSecRecPos(llLoop) = 0
                            Next llLoop
                        End If
                    Next ilKeys
                End If
            End If
            slError = lbcResult.List(ilFileIndex)
            Print #hmTo, slError
            If ilTestOnly Then
                'ilRet = btrClose(hmFile)
                btrDestroy hmFile
            Else
                slMsg = Str$(btrRecords(hmFileNew)) & " of" & Str$(btrRecords(hmFile)) & " copied"
                If (btrRecords(hmFile) = btrRecords(hmFileNew)) Then
                    'ilRet = btrClose(hmFileNew)
                    btrDestroy hmFileNew
                    'ilRet = btrClose(hmFile)
                    btrDestroy hmFile
                    ilPos = InStr(slName, ".")
                    slNewFile = Left$(slName, ilPos - 1)
                    slOldName = Left$(slName, ilPos) & "Old"
                    On Error Resume Next
                    Kill sgDBPath & slOldName
                    On Error GoTo 0
                    Name sgDBPath & slName As sgDBPath & slOldName
                    Name sgDBPath & slNewName As sgDBPath & slName
                    'Process extensions
                    slExtension = ".^01"
                    ilExtNumber = 1
                    Do
                        On Error Resume Next
                        Kill sgDBPath & slNewFile & "_old" & slExtension
                        On Error GoTo cmcChkErr:
                        ilRet = 0
                        slStamp = FileDateTime(sgDBPath & slNewFile & slExtension)
                        If ilRet = 0 Then
                            ilRet = 0
                            On Error GoTo cmcChkErr
                            Name sgDBPath & slNewFile & slExtension As sgDBPath & slNewFile & "_old" & slExtension
                            On Error GoTo 0
                            ilExtNumber = ilExtNumber + 1
                            If ilExtNumber <= 9 Then
                                slExtension = ".^0" & Trim$(Str$(ilExtNumber))
                            Else
                                slExtension = ".^" & Trim$(Str$(ilExtNumber))
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    slExtension = ".^01"
                    ilExtNumber = 1
                    Do
                        On Error GoTo cmcChkErr:
                        ilRet = 0
                        slStamp = FileDateTime(sgDBPath & slNewFile & "_" & slExtension)
                        If ilRet = 0 Then
                            ilRet = 0
                            On Error GoTo cmcChkErr
                            Name sgDBPath & slNewFile & "_" & slExtension As sgDBPath & slNewFile & slExtension
                            On Error GoTo 0
                            ilExtNumber = ilExtNumber + 1
                            If ilExtNumber <= 9 Then
                                slExtension = ".^0" & Trim$(Str$(ilExtNumber))
                            Else
                                slExtension = ".^" & Trim$(Str$(ilExtNumber))
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    

                    Print #hmTo, "File Completely Copied, File Renamed" & slMsg
                Else
                    'ilRet = btrClose(hmFileNew)
                    btrDestroy hmFileNew
                    'ilRet = btrClose(hmFile)
                    btrDestroy hmFile
                    Print #hmTo, "File Not Completely Copied, Names unchanged" & slMsg
                End If
            End If
        End If
        'ilRet = btrClose(hmFile)
        'btrDestroy hmFile
        llFileProc = llFileProc + 1
        llPercent = (llFileProc * CSng(100)) / llTotalFile
        If llPercent >= 100 Then
            llPercent = 100
        End If
        plcGauge.Value = llPercent

    Next ilLoop
    Print #hmTo, ""
    Print #hmTo, "Completed " & Format$(gNow(), "ddddd") & " at " & Format$(Now, "ttttt")
    Close #hmTo
    lacProcessing.Caption = "Processing- Done "
    imProcessing = False
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
    Exit Sub
cmcChkErr:
    ilRet = Err.Number
    Resume Next
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    'igParentRestarted = True
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    'End
End Sub
Private Sub lbcFile_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked 'False
        imSetAll = True
        mSetCommands
    End If
End Sub
Private Sub lbcFile_GotFocus()
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        tmcStart.Enabled = True
        Screen.MousePointer = vbDefault
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()

    Dim slDate As String

    imTerminate = False
    imFirstActivate = True
    imSetAll = True
    imAllClicked = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass

    'mParseCmmdLine
    'If imTerminate Then
    '    Exit Sub
    'End If
    lacDBPath.Caption = ""  '"D.B. Path=> " & sgDBPath
    EngrFileCheck.Height = cmcChk.Top + 5 * cmcChk.Height / 2
    gCenterFormModal EngrFileCheck
    'EngrFileCheck.Show

'    mInitDDE
    'gGetSchParameters
    Screen.MousePointer = vbHourglass
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    slDate = Format$(gNow(), "ddddd")
    lmNowDate = gDateValue(slDate)

    'Moved to lbcVehs GotFocus so hourglass will show up will building list box
'    Screen.MousePointer = vbHourGlass
'    lbcVehs.Clear
'    mPopulate

'    mSetCommands      'Disable LinksDef if No Vehicles are selected
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
         
    On Error GoTo 0
    imTerminate = True
    mTerminate
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
'    Dim slCommand As String
'    Dim slStr As String
'    Dim ilRet As Integer
'    Dim slTestSystem As String
'    Dim ilTestSystem As Integer
'    slCommand = sgCommandStr    'Command$
'    'ilPos = InStr(1, slCommand, "/CS", 1)
'    'If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
'    '    igStdAloneMode = True 'False  'Switch from/to stand alone mode
'    '    sgCallAppName = ""
'    '    slStr = "Guide"
'    '    ilTestSystem = False
'    '    ilPos = InStr(1, slCommand, "/S:Test", 1)
'    '    If ilPos > 0 Then
'    '        ilTestSystem = True
'    '    End If
'    '    ilPos = InStr(1, slCommand, "/S:Prod", 1)
'    '    If ilPos > 0 Then
'    '        ilTestSystem = False
'    '    End If
'    '    slCommand = "Traffic\Guide"
'    'Else
'    '    igStdAloneMode = False  'Switch from/to stand alone mode
'        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get user name
'        If Trim$(slStr) = "" Then
'            MsgBox "Application must be run from the Engineering application", MB_ICONSTOP, "Program Schedule"
'            'End
'            imTerminate = True
'            Exit Sub
'        End If
'        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
'        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
'        If StrComp(slTestSystem, "Test", 1) = 0 Then
'            ilTestSystem = True
'        Else
'            ilTestSystem = False
'        End If
'        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
'    'End If
'    'gInitStdAlone EngrFileCheck, slStr, ilTestSystem
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/25/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box              *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
    lbcFile.Path = Left$(sgDBPath, Len(sgDBPath) - 1)
End Sub
'************************************************************
'          Procedure Name : mSetCommands
'
'    Created : 4/17/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments:  Set Control properties
'
'
'************************************************************
'
Private Sub mSetCommands()
    Dim ilLoop As Integer

    If lbcFile.ListCount <= 0 Then
        cmcChk.Enabled = False
    Else
        cmcChk.Enabled = False
        For ilLoop = 0 To lbcFile.ListCount - 1 Step 1
            If lbcFile.Selected(ilLoop) Then
                cmcChk.Enabled = True
                Exit For
            End If
        Next ilLoop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate Links                *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
    'Erase lmRecNoMoved
    Erase lmPriRecPos
    Erase lmSecRecPos
    'igParentRestarted = False
    'If Not igStdAloneMode Then
    '    'If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '    '    edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    'Else
    '    '    edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '    '    edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '    '    edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '    '    edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '    '    edcLinkDestHelpMsg.LinkExecute "Done"
    '    'End If
    '    'Do While Not igParentRestarted
    '    '    DoEvents
    '    'Loop
    'End If
    Screen.MousePointer = vbDefault
    igManUnload = vbYes
    'Unload Traffic
    Unload EngrFileCheck
    Set EngrFileCheck = Nothing
    'End
    igManUnload = vbNo
End Sub
Private Sub rbcCF_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcCF(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            cmcChk.Caption = "C&heck"
        Else
            cmcChk.Caption = "&Fix"
        End If
    End If
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imFirstTime Then
        'mInitDDE
        imFirstTime = False
    End If
    Screen.MousePointer = vbHourglass
    lacDBPath.Caption = "D.B. Path=> " & sgDBPath
    If lbcFile.Path <> Left$(sgDBPath, Len(sgDBPath) - 1) Then
        'lbcFile.
        mPopulate
    End If
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "File Check/Fix"
End Sub

Private Function mCheckRecPos(llRecPos As Long) As Long

    
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = 0
    llMax = lmPriEngrFileCheckCount - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llRecPos = lmPriRecPos(llMiddle) Then
            'found the match
            mCheckRecPos = True
            Exit Function
        ElseIf llRecPos < lmPriRecPos(llMiddle) Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
'    llMin = 0
'    llMax = lmSecEngrFileCheckCount - 1
'    Do While llMin <= llMax
'        llMiddle = (llMin + llMax) \ 2
'        If llRecPos = lmSecRecPos(llMiddle) Then
'            'found the match
'            mCheckRecPos = True
'            Exit Function
'        ElseIf llRecPos < lmSecRecPos(llMiddle) Then
'            llMax = llMiddle - 1
'        Else
'            'search the right half
'            llMin = llMiddle + 1
'        End If
'    Loop
    mCheckRecPos = False
    Exit Function
End Function

