VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form EngrImport 
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "EngrImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton rbcImportType 
      Caption         =   "Relays"
      Height          =   195
      Index           =   3
      Left            =   3495
      TabIndex        =   11
      Top             =   435
      Width           =   1410
   End
   Begin VB.OptionButton rbcImportType 
      Caption         =   "Netcues"
      Height          =   195
      Index           =   2
      Left            =   2460
      TabIndex        =   10
      Top             =   435
      Width           =   1095
   End
   Begin VB.OptionButton rbcImportType 
      Caption         =   "Buses"
      Height          =   195
      Index           =   1
      Left            =   1545
      TabIndex        =   9
      Top             =   435
      Width           =   855
   End
   Begin VB.OptionButton rbcImportType 
      Caption         =   "Audio Names"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   435
      Width           =   1410
   End
   Begin VB.ListBox lbcError 
      Height          =   2010
      ItemData        =   "EngrImport.frx":08CA
      Left            =   240
      List            =   "EngrImport.frx":08CC
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
      Width           =   7245
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6225
      TabIndex        =   5
      Top             =   3960
      Width           =   1245
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   330
      Left            =   6075
      TabIndex        =   3
      Top             =   825
      Width           =   1395
   End
   Begin VB.CommandButton cmcImport 
      Caption         =   "Import"
      Height          =   315
      Left            =   4650
      TabIndex        =   4
      Top             =   3960
      Width           =   1245
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   945
      TabIndex        =   2
      Top             =   840
      Width           =   4770
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7530
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7665
      Top             =   2355
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4440
      FormDesignWidth =   7815
   End
   Begin VB.Label lacScreen 
      Caption         =   "Import Engineering Files"
      Height          =   270
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2460
   End
   Begin VB.Label lacPercent 
      Height          =   210
      Left            =   225
      TabIndex        =   8
      Top             =   4005
      Width           =   2955
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   1065
      TabIndex        =   7
      Top             =   3675
      Width           =   5625
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   855
      Width           =   780
   End
End
Attribute VB_Name = "EngrImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrImport - displays import csv information
'*
'*  Created Aug,1998 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imTerminate As Integer
Private imImporting As Integer
Private hmFrom As Integer
Private hmMsg As Integer
Private smMsgFile As String
Private smNowDate As String
Private smNowTime As String
Private lmTotalNoBytes As Long
Private lmProcessedNoBytes As Long
Private smCurDir As String
Private lmFloodPercent As Long
Private imImportSelection As Integer
Private smFields(1 To 3) As String
Private smBothANEStamp As String
Private tmBothANE() As ANE
Private tmANE As ANE
Private smBothNNEStamp As String
Private tmBothNNE() As NNE
Private tmNNE As NNE
Private smBothRNEStamp As String
Private tmBothRNE() As RNE
Private tmRNE As RNE
Private smBothBDEStamp As String
Private tmBothBDE() As BDE
Private tmBDE As BDE
Private smBothASEStamp As String
Private tmBothASE() As ASE
Private tmASE As ASE












'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileCP                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mImportFile(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim slStr As String
    Dim slChar As String
    Dim ilLinesToSkip As Integer
        
    ilRet = 0
    On Error GoTo mImportFileErr:
    hmFrom = FreeFile
    Open slFromFile For Input Access Read As hmFrom
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mImportFile = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    ilLinesToSkip = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mImportFileErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            lmProcessedNoBytes = lmProcessedNoBytes + 1
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf (slChar <> sgCR) And (slChar <> sgTB) Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mImportFile = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If (Len(slLine) > 0) Then
            If Left$(slLine, 1) <> "[" Then
                If ilLinesToSkip < 5 Then
                    ilLinesToSkip = ilLinesToSkip + 1
                Else
                    If InStr(1, slLine, "Printed:", vbTextCompare) = 1 Then
                        ilLinesToSkip = 1
                    Else
                        'Process Input
                        If imImportSelection <> 2 Then
                            smFields(1) = Trim$(Left$(slLine, 5))
                            smFields(2) = Trim$(Mid$(slLine, 6))
                            smFields(3) = ""
                        Else
                            smFields(1) = Trim$(Left$(slLine, 4))
                            smFields(2) = Trim$(Mid$(slLine, 5))
                            smFields(3) = ""
                        End If
                        If imImportSelection = 0 Then
                            ilFound = False
                            For ilLoop = LBound(tmBothANE) To UBound(tmBothANE) - 1 Step 1
                                If StrComp(Trim$(tmBothANE(ilLoop).sName), smFields(1), vbTextCompare) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                mMoveAudioToRec
                                ilRet = gPutInsert_ANE_AudioName(0, tmANE, "Audio Name-mImportFile: Insert ANE")
                                LSet tmBothANE(UBound(tmBothANE)) = tmANE
                                ReDim Preserve tmBothANE(LBound(tmBothANE) To UBound(tmBothANE) + 1) As ANE
                                mMoveASEToRec
                                ilRet = gPutInsert_ASE_AudioSource(0, tmASE, "Audio Source-ImportFile: Insert ASE")
                            Else
                                Print #hmMsg, "Name Previously Defined: " & smFields(1)
                            End If
                        ElseIf imImportSelection = 1 Then
                            smFields(3) = Trim$(Mid$(smFields(2), 35))
                            smFields(2) = Trim$(Left$(smFields(2), 34))
                            ilFound = False
                            For ilLoop = LBound(tmBothBDE) To UBound(tmBothBDE) - 1 Step 1
                                If StrComp(Trim$(tmBothBDE(ilLoop).sName), smFields(1), vbTextCompare) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                mMoveBusToRec
                                ilRet = gPutInsert_BDE_BusDefinition(0, tmBDE, "Bus-mImportFile: Insert BDE")
                                LSet tmBothBDE(UBound(tmBothBDE)) = tmBDE
                                ReDim Preserve tmBothBDE(LBound(tmBothBDE) To UBound(tmBothBDE) + 1) As BDE
                                If (smFields(3) <> "") And (tmBDE.iAseCode = 0) Then
                                    Print #hmMsg, "Audio Source " & smFields(3) & " not set for " & smFields(1)
                                End If
                            Else
                                Print #hmMsg, "Name Previously Defined: " & smFields(1)
                            End If
                        ElseIf imImportSelection = 2 Then
                            ilFound = -1
                            For ilLoop = LBound(tmBothNNE) To UBound(tmBothNNE) - 1 Step 1
                                If StrComp(Trim$(tmBothNNE(ilLoop).sName), smFields(1), vbTextCompare) = 0 Then
                                    ilFound = ilLoop
                                    Exit For
                                End If
                            Next ilLoop
                            If ilFound = -1 Then
                                mMoveNetcueToRec
                                ilRet = gPutInsert_NNE_NetcueName(0, tmNNE, "Netcue-mImportFile: Insert NNE")
                                LSet tmBothNNE(UBound(tmBothNNE)) = tmNNE
                                ReDim Preserve tmBothNNE(LBound(tmBothNNE) To UBound(tmBothNNE) + 1) As NNE
                            Else
                                If StrComp(Trim$(tmBothNNE(ilLoop).sDescription), smFields(2), vbTextCompare) = 0 Then
                                    Print #hmMsg, "Name Previously Defined: " & smFields(1)
                                Else
                                    tmBothNNE(ilLoop).sDescription = smFields(2)
                                    ilRet = gPutUpdate_NNE_NetcueName(3, tmBothNNE(ilLoop), "Netcue-mImportFile: Update description")
                                End If
                            End If
                        ElseIf imImportSelection = 3 Then
                            ilFound = False
                            For ilLoop = LBound(tmBothRNE) To UBound(tmBothRNE) - 1 Step 1
                                If StrComp(Trim$(tmBothRNE(ilLoop).sName), smFields(1), vbTextCompare) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                mMoveRelayToRec
                                ilRet = gPutInsert_RNE_RelayName(0, tmRNE, "Relay-mImportFile: Insert RNE")
                                LSet tmBothRNE(UBound(tmBothRNE)) = tmRNE
                                ReDim Preserve tmBothRNE(LBound(tmBothRNE) To UBound(tmBothRNE) + 1) As RNE
                            Else
                                Print #hmMsg, "Name Previously Defined: " & smFields(1)
                            End If
                        End If
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lacPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mImportFile = False
    Else
        mImportFile = True
        lmFloodPercent = 100
        lacPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mImportFileErr:
    ilRet = Err.Number
    Resume Next
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
Private Function mOpenMsgFile() As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    On Error GoTo mOpenMsgFileErr:
    slToFile = sgMsgDirectory & smMsgFile
    slNowDate = Format$(gNow(), sgShowDateForm)
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    If imImportSelection = 0 Then
        Print #hmMsg, "** Import Audio Name: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf imImportSelection = 1 Then
        Print #hmMsg, "** Import Buses: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf imImportSelection = 2 Then
        Print #hmMsg, "** Import Netcues: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf imImportSelection = 3 Then
        Print #hmMsg, "** Import Relays: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    End If
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function





Private Sub cmcBrowse_Click()

    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    cmcCancel.Caption = "&Cancel"
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload EngrImport
End Sub

Private Sub cmcImport_Click()
    Dim ilRet As Integer
    Dim slFromFile As String
    Dim slDate As String
    Dim ilVefCode As Integer
    Dim ilSelected As Integer
    Dim ilLoop As Integer
    
    If txtFile.Text = "" Then
        MsgBox "Import File must be specified.", vbOKOnly
        txtFile.SetFocus
        Exit Sub
    End If
    lmProcessedNoBytes = 0
    lbcError.Clear
    lacMsg.Caption = ""
    lacPercent.Caption = ""
    imImporting = True
    slFromFile = txtFile.Text
    If rbcImportType(0).Value Then
        imImportSelection = 0
        smMsgFile = "ImportAudio.Txt"
    ElseIf rbcImportType(1).Value Then
        imImportSelection = 1
        smMsgFile = "ImportBus.Txt"
        ilRet = MsgBox("Has Audio Been Import, Press Ok if it has been Imported", vbOKCancel + vbQuestion)
        If ilRet = vbCancel Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf rbcImportType(2).Value Then
        imImportSelection = 2
        smMsgFile = "ImportNetcue.Txt"
    ElseIf rbcImportType(3).Value Then
        imImportSelection = 3
        smMsgFile = "ImportRelay.Txt"
    Else
        MsgBox "Import Type must be specified.", vbOKOnly
        rbcImportType(0).SetFocus
        Exit Sub
    End If
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    ilRet = mOpenMsgFile()
    Screen.MousePointer = vbHourglass
    Print #hmMsg, "Import File: " & slFromFile
    If imImportSelection = 0 Then
        ilRet = gGetTypeOfRecs_ANE_AudioName("B", smBothANEStamp, "EngrImport-cmcImport", tmBothANE())
        ilRet = mImportFile(slFromFile)
        ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrImport-cmcImport Audio Names", tgCurrANE())
        ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrImport-cmcImport Audio", tgCurrASE())
    ElseIf imImportSelection = 1 Then
        ilRet = gGetTypeOfRecs_ASE_AudioSource("B", smBothASEStamp, "EngrImport-cmcImport Audio Source", tmBothASE())
        ilRet = gGetTypeOfRecs_ANE_AudioName("B", smBothANEStamp, "EngrImport-cmcImport Audio Name", tmBothANE())
        ilRet = gGetTypeOfRecs_BDE_BusDefinition("B", smBothBDEStamp, "EngrImport-cmcImport", tmBothBDE())
        ilRet = mImportFile(slFromFile)
        ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrImport-cmcImport", tgCurrBDE())
    ElseIf imImportSelection = 2 Then
        ilRet = gGetTypeOfRecs_NNE_NetcueName("B", smBothNNEStamp, "EngrImport-cmcImport", tmBothNNE())
        ilRet = mImportFile(slFromFile)
        ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrImport-cmcImport", tgCurrNNE())
    ElseIf imImportSelection = 3 Then   'Import Affiliate Spots
        ilRet = gGetTypeOfRecs_RNE_RelayName("B", smBothRNEStamp, "EngrImport-cmcImport", tmBothRNE())
        ilRet = mImportFile(slFromFile)
        ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrImport-cmcImport", tgCurrRNE())
    End If
    Close hmMsg
    Screen.MousePointer = vbDefault
    lacMsg.Caption = "See " & smMsgFile & " for Messages"
    'cmcImport.Enabled = False
    imImporting = False
    If Not imTerminate Then
        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
    Else
        Unload EngrImport
    End If
    Exit Sub

End Sub


Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrImport
    gCenterFormModal EngrImport
End Sub

Private Sub Form_Load()
    Dim iUpper As Integer
    
    smCurDir = CurDir
    Screen.MousePointer = vbHourglass
    imImporting = False
    imTerminate = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase tmBothANE
    Erase tmBothASE
    Erase tmBothRNE
    Erase tmBothBDE
    Erase tmBothNNE
    
    If InStr(1, smCurDir, ":") > 0 Then
        ChDrive Left$(smCurDir, 1)
        ChDir smCurDir
    End If
    Set EngrImport = Nothing
End Sub

Private Sub mMoveAudioToRec()
    
    tmANE.iCode = 0
    tmANE.sName = smFields(1)
    tmANE.sDescription = smFields(2)
    tmANE.iCceCode = 0
    tmANE.iAteCode = 0
    tmANE.sState = "A"
    tmANE.sUsedFlag = "N"
    tmANE.iVersion = 0
    tmANE.iOrigAneCode = tmANE.iCode
    tmANE.sCurrent = "Y"
    'tmANE.sEnteredDate = smNowDate
    'tmANE.sEnteredTime = smNowTime
    tmANE.sEnteredDate = Format(Now, sgShowDateForm)
    tmANE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmANE.iUieCode = tgUIE.iCode
    tmANE.sUnused = ""
End Sub
Private Sub mMoveNetcueToRec()
    
    tmNNE.iCode = 0
    tmNNE.sName = smFields(1)
    tmNNE.sDescription = smFields(2)
    tmNNE.lDneCode = 0
    tmNNE.sState = "A"
    tmNNE.sUsedFlag = "N"
    tmNNE.iVersion = 0
    tmNNE.iOrigNneCode = tmNNE.iCode
    tmNNE.sCurrent = "Y"
    'tmNNE.sEnteredDate = smNowDate
    'tmNNE.sEnteredTime = smNowTime
    tmNNE.sEnteredDate = Format(Now, sgShowDateForm)
    tmNNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmNNE.iUieCode = tgUIE.iCode
    tmNNE.sUnused = ""
End Sub

Private Sub mMoveRelayToRec()
    Dim slStr As String
    
    tmRNE.iCode = 0
    tmRNE.sName = smFields(1)
    tmRNE.sDescription = smFields(2)
    tmRNE.sState = "A"
    tmRNE.sUsedFlag = "N"
    tmRNE.iVersion = 0
    tmRNE.iOrigRneCode = tmRNE.iCode
    tmRNE.sCurrent = "Y"
    'tmRNE.sEnteredDate = smNowDate
    'tmRNE.sEnteredTime = smNowTime
    tmRNE.sEnteredDate = Format(Now, sgShowDateForm)
    tmRNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmRNE.iUieCode = tgUIE.iCode
    tmRNE.sUnused = ""
End Sub

Private Sub mMoveBusToRec()
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim slStr As String
    
    tmBDE.iCode = 0
    tmBDE.sName = smFields(1)
    tmBDE.sDescription = smFields(2)
    tmBDE.iCceCode = 0
    tmBDE.sChannel = ""
    tmBDE.iAseCode = 0
    slStr = smFields(3)
    For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                    tmBDE.iAseCode = tgCurrASE(ilASE).iCode
                    Exit For
                End If
            End If
        Next ilANE
        If tmBDE.iAseCode <> 0 Then
            Exit For
        End If
    Next ilASE
    tmBDE.sState = "A"
    tmBDE.sUsedFlag = "N"
    tmBDE.iVersion = 0
    tmBDE.iOrigBdeCode = tmBDE.iCode
    tmBDE.sCurrent = "Y"
    'tmBDE.sEnteredDate = smNowDate
    'tmBDE.sEnteredTime = smNowTime
    tmBDE.sEnteredDate = Format(Now, sgShowDateForm)
    tmBDE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmBDE.iUieCode = tgUIE.iCode
    tmBDE.sUnused = ""
End Sub

Private Sub mMoveASEToRec()
    Dim ilCCE As Integer
    Dim ilANE As Integer
    Dim ilASE As Integer
    Dim slStr As String
        
    tmASE.iCode = 0
    tmASE.iPriAneCode = tmANE.iCode
    tmASE.iPriCceCode = 0
    tmASE.sDescription = smFields(2)
    tmASE.iBkupAneCode = 0
    tmASE.iBkupCceCode = 0
    tmASE.iProtAneCode = 0
    tmASE.iProtCceCode = 0
    tmASE.sState = "A"
    tmASE.sUsedFlag = "N"
    tmASE.iVersion = 0
    tmASE.iOrigAseCode = tmASE.iCode
    tmASE.sCurrent = "Y"
    'tmASE.sEnteredDate = smNowDate
    'tmASE.sEnteredTime = smNowTime
    tmASE.sEnteredDate = Format(Now, sgShowDateForm)
    tmASE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmASE.iUieCode = tgUIE.iCode
    tmASE.sUnused = ""
End Sub

Private Sub rbcImportType_Click(Index As Integer)
    cmcCancel.Caption = "&Cancel"
End Sub

Private Sub txtFile_Change()
    cmcCancel.Caption = "&Cancel"
End Sub
