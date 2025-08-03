VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportAE 
   Caption         =   "Import Affiliate A/E"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "AffImportAE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6195
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   990
      TabIndex        =   6
      Top             =   480
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   4845
      TabIndex        =   5
      Top             =   480
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      Height          =   2205
      ItemData        =   "AffImportAE.frx":08CA
      Left            =   120
      List            =   "AffImportAE.frx":08CC
      TabIndex        =   1
      Top             =   1410
      Width           =   5790
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5790
      Top             =   4305
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5160
      FormDesignWidth =   6195
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   4650
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3165
      TabIndex        =   3
      Top             =   4650
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5895
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   615
      TabIndex        =   8
      Top             =   1125
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacAgreementNoAssigned 
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   4215
      Width           =   5715
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   495
      Width           =   780
   End
   Begin VB.Label lacResult 
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   3765
      Width           =   5730
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   5790
   End
End
Attribute VB_Name = "frmImportAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private smDate As String     'Export Date
Private imMktCode As Integer
Private smErrMarketName() As String
Private imVefCode As Integer
Private smErrVefName() As String
Private imMnfVehGp2 As Integer
Private lmMntCode As Long
Private smErrTerritoryName() As String
Private lmArttCode As Long
Private smErrAEName() As String
Private imAllClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
'Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private lmTotalNoBytes As Long
Private lmProcessedNoBytes As Long
Private lmPercent As Long








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
'Private Function mOpenMsgFile(sMsgFileName As String) As Integer
'    Dim slToFile As String
'    Dim slDateTime As String
'    Dim slFileDate As String
'    Dim slNowDate As String
'    Dim ilRet As Integer
'
'    On Error GoTo mOpenMsgFileErr:
'    ilRet = 0
'    slNowDate = Format$(gNow(), sgShowDateForm)
'    slToFile = sgMsgDirectory & "ImptAiredSpots.Txt"
'    slDateTime = FileDateTime(slToFile)
'    If ilRet = 0 Then
'        slFileDate = Format$(slDateTime, sgShowDateForm)
'        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                Close hmMsg
'                hmMsg = -1
'                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
'            Kill slToFile
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Output As hmMsg
'            If ilRet <> 0 Then
'                Close hmMsg
'                hmMsg = -1
'                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        End If
'    Else
'        On Error GoTo 0
'        ilRet = 0
'        On Error GoTo mOpenMsgFileErr:
'        hmMsg = FreeFile
'        Open slToFile For Output As hmMsg
'        If ilRet <> 0 Then
'            Close hmMsg
'            hmMsg = -1
'            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
'            mOpenMsgFile = False
'            Exit Function
'        End If
'    End If
'    On Error GoTo 0
'    'Print #hmMsg, "** Import Aired Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    'Print #hmMsg, ""
'    sMsgFileName = slToFile
'    mOpenMsgFile = True
'    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
'End Function

Private Sub cmcBrowse_Click()

    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.InitDir = sgImportDirectory
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdImport_Click()
    Dim iLoop As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String

    On Error GoTo ErrHand
    
    lbcMsg.Clear
    Screen.MousePointer = vbHourglass
    If Not mCheckFile() Then
        Screen.MousePointer = vbDefault
        txtFile.SetFocus
        Exit Sub
    End If
    If lmTotalNoBytes > 0 Then
        plcGauge.Value = 0
        plcGauge.Visible = True
    End If
    lmPercent = 0
    lmProcessedNoBytes = 0
    gPopMarkets
    gPopMntInfo "T", tgTerritoryInfo()
    gPopAffAE
'    If Not mOpenMsgFile(sMsgFileName) Then
'        cmdCancel.SetFocus
'        Exit Sub
'    End If
    imExporting = True
    On Error GoTo 0
    lbcMsg.Enabled = True
    lacResult.Caption = ""
    iRet = mImportAE()
    If (iRet = False) Then
        gLogMsg "** Terminated - mImportAffiliateAE returned False**", "ImportAffiliateAELog.Txt", False
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        Close #hmTo
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        gLogMsg "** User Terminated **", "ImportAffiliateAELog.Txt", False
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        Close #hmTo
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    mCheckForAgreementsNotAssigned
    'Clear old aet records out
    On Error GoTo ErrHand:
    imExporting = False
    gLogMsg "** Completed Import Affiliate A/E Assignment" & " **", "ImportAffiliateAELog.Txt", False
    gLogMsg "", "ImportAffiliateAELog.Txt", False
    'Print #hmMsg, "** Completed Import Aired Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    'Close #hmMsg
    lacResult.Caption = "Results: " & "...data\messages\ImportAffiliateAELog.Txt"
    cmdImport.Enabled = False
    cmdCancel.Caption = "&Done"
    plcGauge.Visible = False
    Screen.MousePointer = vbDefault
    Exit Sub
cmdImportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportAE-cmdImport"
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportAE
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    frmImportAE.Caption = "Affiliate A/E Assignment- " & sgClientName
    imAllClick = False
    imTerminate = False
    imExporting = False
    ilRet = gPopSubTotalGroups()
    ilRet = gPopShttInfo()
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Erase smErrMarketName
    Erase smErrVefName
    Erase smErrTerritoryName
    Erase smErrAEName
    Set frmImportAE = Nothing
End Sub

Private Function mImportAE() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilRankFd As Integer
    Dim ilMkt As Integer
    Dim ilMnt As Integer
    Dim ilVef As Integer
    Dim llArtt As Long
    Dim ilShtt As Integer
    Dim ilAddMat As Integer
    Dim ilErr As Integer
    Dim ilPass As Integer
    Dim ilH2 As Integer
    Dim ilH2Fd As Integer
    Dim ilErrNameFd As Integer
    'Dim slFields(1 To 5) As String
    Dim slFields(0 To 4) As String
    
    slFromFile = txtFile.Text
    'Clear Affiliate A/E table
    SQLQuery = "DELETE FROM mat "
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
        mImportAE = False
        Exit Function
    End If
    'Clear Agreement assignments
    SQLQuery = "UPDATE att SET "
    SQLQuery = SQLQuery & "attArttCode = 0"
    'cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
        mImportAE = False
        Exit Function
    End If
    ReDim smErrMarketName(0 To 0) As String
    ReDim smErrVefName(0 To 0) As String
    ReDim smErrTerritoryName(0 To 0) As String
    ReDim smErrAEName(0 To 0) As String
    'cnn.CommitTrans
    For ilPass = 1 To 3 Step 1
        'Pass 1"  Assign those record without Vehicle or Territory
        'Pass 2:  Assign those record that have only Territory
        'Pass 3:  Assign those records with vehicle
        'ilRet = 0
        'On Error GoTo mImportAEErr:
        'hmFrom = FreeFile
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            Close hmFrom
            Exit Function
        End If
        ilRankFd = False
        Do While Not EOF(hmFrom)
            DoEvents
            ilRet = 0
            'On Error GoTo mImportAEErr:
            Line Input #hmFrom, slLine
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, False, slFields()
                    For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                        slFields(ilLoop) = Trim$(slFields(ilLoop))
                    Next ilLoop
                    If Not ilRankFd Then
                        'If StrComp(slFields(1), "Rank", vbTextCompare) = 0 Then
                        If StrComp(slFields(0), "Rank", vbTextCompare) = 0 Then
                            ilRankFd = True
                        End If
                    Else
                        'Find market Code from Market Name
                        ilAddMat = True
                        imMktCode = 0
                        For ilMkt = LBound(tgMarketInfo) To UBound(tgMarketInfo) - 1 Step 1
                            'If StrComp(Trim$(tgMarketInfo(ilMkt).sName), slFields(2), vbTextCompare) = 0 Then
                            If StrComp(Trim$(tgMarketInfo(ilMkt).sName), slFields(1), vbTextCompare) = 0 Then
                                imMktCode = tgMarketInfo(ilMkt).lCode
                                Exit For
                            End If
                        Next ilMkt
                        If imMktCode = 0 Then
                            ilAddMat = False
                            ilErrNameFd = False
                            For ilErr = 0 To UBound(smErrMarketName) - 1 Step 1
                                'If StrComp(smErrMarketName(ilErr), slFields(2), vbTextCompare) = 0 Then
                                If StrComp(smErrMarketName(ilErr), slFields(1), vbTextCompare) = 0 Then
                                    ilErrNameFd = True
                                    Exit For
                                End If
                            Next ilErr
                            If Not ilErrNameFd Then
                                'lbcMsg.AddItem "Market Not Found: " & slFields(2)
                                lbcMsg.AddItem "Market Not Found: " & slFields(1)
                                'gLogMsg "Market Not Found: " & slFields(2), "ImportAffiliateAELog.Txt", False
                                gLogMsg "Market Not Found: " & slFields(1), "ImportAffiliateAELog.Txt", False
                                'smErrMarketName(UBound(smErrMarketName)) = slFields(2)
                                smErrMarketName(UBound(smErrMarketName)) = slFields(1)
                                ReDim Preserve smErrMarketName(0 To UBound(smErrMarketName) + 1) As String
                            End If
                        End If
                        'Find vehicle code from vehicle name
                        imVefCode = 0
                        imMnfVehGp2 = 0
                        'If slFields(3) <> "" Then
                        If slFields(2) <> "" Then
                            ilH2Fd = False
                            For ilH2 = LBound(tgSubtotalGroupInfo) To UBound(tgSubtotalGroupInfo) - 1 Step 1
                                'If StrComp(Trim$(tgSubtotalGroupInfo(ilH2).sName), slFields(3), vbTextCompare) = 0 Then
                                If StrComp(Trim$(tgSubtotalGroupInfo(ilH2).sName), slFields(2), vbTextCompare) = 0 Then
                                    imMnfVehGp2 = tgSubtotalGroupInfo(ilH2).iCode
                                    ilH2Fd = True
                                    Exit For
                                End If
                            Next ilH2
                            If Not ilH2Fd Then
                                For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                                    'If StrComp(Trim$(tgVehicleInfo(ilVef).sVehicleName), slFields(3), vbTextCompare) = 0 Then
                                    If StrComp(Trim$(tgVehicleInfo(ilVef).sVehicleName), slFields(2), vbTextCompare) = 0 Then
                                        imVefCode = tgVehicleInfo(ilVef).iCode
                                        Exit For
                                    End If
                                Next ilVef
                            End If
                            If (imVefCode = 0) And (ilH2Fd = False) Then
                                ilAddMat = False
                                ilErrNameFd = False
                                For ilErr = 0 To UBound(smErrVefName) - 1 Step 1
                                    'If StrComp(smErrVefName(ilErr), slFields(3), vbTextCompare) = 0 Then
                                    If StrComp(smErrVefName(ilErr), slFields(2), vbTextCompare) = 0 Then
                                        ilErrNameFd = True
                                        Exit For
                                    End If
                                Next ilErr
                                If Not ilErrNameFd Then
                                    'lbcMsg.AddItem "Vehicle Not Found: " & slFields(3)
                                    lbcMsg.AddItem "Vehicle Not Found: " & slFields(2)
                                    'gLogMsg "Vehicle Not Found: " & slFields(3), "ImportAffiliateAELog.Txt", False
                                    gLogMsg "Vehicle Not Found: " & slFields(2), "ImportAffiliateAELog.Txt", False
                                    'smErrVefName(UBound(smErrVefName)) = slFields(3)
                                    smErrVefName(UBound(smErrVefName)) = slFields(2)
                                    ReDim Preserve smErrVefName(0 To UBound(smErrVefName) + 1) As String
                                End If
                            End If
                        End If
                        'Find Territory code from territory name
                        lmMntCode = 0
                        'If slFields(4) <> "" Then
                        If slFields(3) <> "" Then
                            For ilMnt = LBound(tgTerritoryInfo) To UBound(tgTerritoryInfo) - 1 Step 1
                                'If StrComp(Trim$(tgTerritoryInfo(ilMnt).sName), slFields(4), vbTextCompare) = 0 Then
                                If StrComp(Trim$(tgTerritoryInfo(ilMnt).sName), slFields(3), vbTextCompare) = 0 Then
                                    lmMntCode = tgTerritoryInfo(ilMnt).lCode
                                    Exit For
                                End If
                            Next ilMnt
                            If lmMntCode = 0 Then
                                ilAddMat = False
                                ilErrNameFd = False
                                For ilErr = 0 To UBound(smErrTerritoryName) - 1 Step 1
                                    'If StrComp(smErrTerritoryName(ilErr), slFields(4), vbTextCompare) = 0 Then
                                    If StrComp(smErrTerritoryName(ilErr), slFields(3), vbTextCompare) = 0 Then
                                        ilErrNameFd = True
                                        Exit For
                                    End If
                                Next ilErr
                                If Not ilErrNameFd Then
                                    'lbcMsg.AddItem "Territory Not Found: " & slFields(4)
                                    lbcMsg.AddItem "Territory Not Found: " & slFields(3)
                                    'gLogMsg "Territory Not Found: " & slFields(4), "ImportAffiliateAELog.Txt", False
                                    gLogMsg "Territory Not Found: " & slFields(3), "ImportAffiliateAELog.Txt", False
                                    'smErrTerritoryName(UBound(smErrTerritoryName)) = slFields(4)
                                    smErrTerritoryName(UBound(smErrTerritoryName)) = slFields(3)
                                    ReDim Preserve smErrTerritoryName(0 To UBound(smErrTerritoryName) + 1) As String
                                End If
                            End If
                        End If
                        'Find A/E code from A/E Name
                        lmArttCode = 0
                        For llArtt = LBound(tgAffAEInfo) To UBound(tgAffAEInfo) - 1 Step 1
                            'If StrComp(Trim$(tgAffAEInfo(llArtt).sName), slFields(5), vbTextCompare) = 0 Then
                            If StrComp(Trim$(tgAffAEInfo(llArtt).sName), slFields(4), vbTextCompare) = 0 Then
                                lmArttCode = tgAffAEInfo(llArtt).lCode
                                Exit For
                            End If
                        Next llArtt
                        If lmArttCode = 0 Then
                            ilAddMat = False
                            ilErrNameFd = False
                            For ilErr = 0 To UBound(smErrAEName) - 1 Step 1
                                'If StrComp(smErrAEName(ilErr), slFields(5), vbTextCompare) = 0 Then
                                If StrComp(smErrAEName(ilErr), slFields(4), vbTextCompare) = 0 Then
                                    ilErrNameFd = True
                                    Exit For
                                End If
                            Next ilErr
                            If Not ilErrNameFd Then
                                'lbcMsg.AddItem "Affiliate A/E Not Found: " & slFields(5)
                                lbcMsg.AddItem "Affiliate A/E Not Found: " & slFields(4)
                                'gLogMsg "Affiliate A/E Not Found: " & slFields(5), "ImportAffiliateAELog.Txt", False
                                gLogMsg "Affiliate A/E Not Found: " & slFields(4), "ImportAffiliateAELog.Txt", False
                                'smErrAEName(UBound(smErrAEName)) = slFields(5)
                                smErrAEName(UBound(smErrAEName)) = slFields(4)
                                ReDim Preserve smErrAEName(0 To UBound(smErrAEName) + 1) As String
                            End If
                        End If
                        If ilAddMat = True Then
                            'Save Affiliate A/E Assignment
                            ilAddMat = False
                            For ilShtt = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                                If tgStationInfo(ilShtt).iMktCode = imMktCode Then
                                    If (ilPass = 1) And (imVefCode = 0) And (lmMntCode = 0) And (imMnfVehGp2 = 0) Then
                                        ilAddMat = True
                                        Exit For
                                    End If
                                    If (ilPass = 2) And (tgStationInfo(ilShtt).lMntCode = lmMntCode) And (lmMntCode <> 0) And (imVefCode = 0) And (imMnfVehGp2 = 0) Then
                                        ilAddMat = True
                                        Exit For
                                    End If
                                    If (ilPass = 3) And ((tgStationInfo(ilShtt).lMntCode = lmMntCode) Or (lmMntCode = 0)) And ((imVefCode <> 0) Or (ilH2Fd)) Then
                                        ilAddMat = True
                                        Exit For
                                    End If
                                End If
                            Next ilShtt
                            If ilAddMat Then
                                SQLQuery = "Insert Into mat ( "
                                SQLQuery = SQLQuery & "matCode, "
                                SQLQuery = SQLQuery & "matMktCode, "
                                SQLQuery = SQLQuery & "matVefCode, "
                                SQLQuery = SQLQuery & "matMntCode, "
                                SQLQuery = SQLQuery & "matArttCode, "
                                SQLQuery = SQLQuery & "matMnfVehGp2, "
                                SQLQuery = SQLQuery & "matUnused "
                                SQLQuery = SQLQuery & ") "
                                SQLQuery = SQLQuery & "Values ( "
                                SQLQuery = SQLQuery & 0 & ", "
                                SQLQuery = SQLQuery & imMktCode & ", "
                                SQLQuery = SQLQuery & imVefCode & ", "
                                SQLQuery = SQLQuery & lmMntCode & ", "
                                SQLQuery = SQLQuery & lmArttCode & ", "
                                SQLQuery = SQLQuery & imMnfVehGp2 & ", "
                                SQLQuery = SQLQuery & "'" & "' "
                                SQLQuery = SQLQuery & ") "
                                'cnn.Execute SQLQuery, rdExecDirect
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
                                    mImportAE = False
                                    Exit Function
                                End If
                            End If
                            'Set Agreements
                            'Get all stations matching market
                            For ilShtt = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                                If tgStationInfo(ilShtt).iMktCode = imMktCode Then
                                
                                    If ilPass = 1 Then
                                        If (imVefCode = 0) And (lmMntCode = 0) And (imMnfVehGp2 = 0) Then
                                            SQLQuery = "UPDATE att SET "
                                            SQLQuery = SQLQuery & "attArttCode = " & lmArttCode
                                            SQLQuery = SQLQuery & " WHERE attShfCode = " & tgStationInfo(ilShtt).iCode & " AND attArttCode = 0"
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
                                                mImportAE = False
                                                Exit Function
                                            End If
                                        End If
                                    ElseIf ilPass = 2 Then
                                        If (tgStationInfo(ilShtt).lMntCode = lmMntCode) And (lmMntCode <> 0) And (imVefCode = 0) And (imMnfVehGp2 = 0) Then
                                            SQLQuery = "UPDATE att SET "
                                            SQLQuery = SQLQuery & "attArttCode = " & lmArttCode
                                            SQLQuery = SQLQuery & " WHERE attShfCode = " & tgStationInfo(ilShtt).iCode
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
                                                mImportAE = False
                                                Exit Function
                                            End If
                                        End If
                                    ElseIf ilPass = 3 Then
                                        If ((tgStationInfo(ilShtt).lMntCode = lmMntCode) Or (lmMntCode = 0)) And ((imVefCode <> 0) Or (ilH2Fd)) Then
                                            If imVefCode <> 0 Then
                                                SQLQuery = "UPDATE att SET "
                                                SQLQuery = SQLQuery & "attArttCode = " & lmArttCode
                                                SQLQuery = SQLQuery & " WHERE attShfCode = " & tgStationInfo(ilShtt).iCode
                                                SQLQuery = SQLQuery & " AND attVefCode = " & imVefCode
                                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                    '6/10/16: Replaced GoSub
                                                    'GoSub ErrHand:
                                                    Screen.MousePointer = vbDefault
                                                    gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
                                                    mImportAE = False
                                                    Exit Function
                                                End If
                                            Else
                                                For ilH2 = LBound(tgSubtotalGroupInfo) To UBound(tgSubtotalGroupInfo) - 1 Step 1
                                                    'If StrComp(Trim$(tgSubtotalGroupInfo(ilH2).sName), slFields(3), vbTextCompare) = 0 Then
                                                    If StrComp(Trim$(tgSubtotalGroupInfo(ilH2).sName), slFields(2), vbTextCompare) = 0 Then
                                                        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                                                            If tgVehicleInfo(ilVef).iMnfVehGp2 = tgSubtotalGroupInfo(ilH2).iCode Then
                                                                imVefCode = tgVehicleInfo(ilVef).iCode
                                                                SQLQuery = "UPDATE att SET "
                                                                SQLQuery = SQLQuery & "attArttCode = " & lmArttCode
                                                                SQLQuery = SQLQuery & " WHERE attShfCode = " & tgStationInfo(ilShtt).iCode
                                                                SQLQuery = SQLQuery & " AND attVefCode = " & imVefCode
                                                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                                    '6/10/16: Replaced GoSub
                                                                    'GoSub ErrHand:
                                                                    Screen.MousePointer = vbDefault
                                                                    gHandleError "ImportAffiliateAELog.Txt", "ImportAE-mImportAE"
                                                                    mImportAE = False
                                                                    Exit Function
                                                                End If
                                                            End If
                                                        Next ilVef
                                                        Exit For
                                                    End If
                                                Next ilH2
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilShtt
                        End If
                    End If
                End If
            End If
            If lmTotalNoBytes > 0 Then
                lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
                lmPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
                If lmPercent >= 100 Then
                    If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                        lmPercent = 99
                    Else
                        lmPercent = 100
                    End If
                End If
                If plcGauge.Value <> lmPercent Then
                    plcGauge.Value = lmPercent
                    DoEvents
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
    Next ilPass
    mImportAE = True
    Exit Function
'mImportAEErr:
'    ilRet = Err.Number
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportAE-mImportAE"
    mImportAE = False
    Exit Function
End Function


Private Function mCheckFile()
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilRankFd As Integer
    'Dim slFields(1 To 5) As String
    Dim slFields(0 To 4) As String
    
    slFromFile = txtFile.Text
    ilRet = 0
    'On Error GoTo mImportSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Beep
        gMsgBox "Unable to open the Import file error: " & Trim$(Str$(ilRet)), vbCritical
        mCheckFile = False
        Close hmFrom
        Exit Function
    End If
    mCheckFile = True
    ilRankFd = False
    Do While Not EOF(hmFrom)
        ilRet = 0
        'On Error GoTo mImportSpotsErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, False, slFields()
                For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                    slFields(ilLoop) = Trim$(slFields(ilLoop))
                Next ilLoop
                'If StrComp(slFields(1), "Rank", vbTextCompare) = 0 Then
                If StrComp(slFields(0), "Rank", vbTextCompare) = 0 Then
                    ilRankFd = True
                    'If StrComp(slFields(2), "Market", vbTextCompare) <> 0 Then
                    If StrComp(slFields(1), "Market", vbTextCompare) <> 0 Then
                        Beep
                        gMsgBox "Import file missing Market as 2st Column", vbCritical
                        Close hmFrom
                        mCheckFile = False
                        Exit Function
                    Else
                        'If InStr(1, slFields(3), "Vehicle", vbTextCompare) <= 0 Then
                        If InStr(1, slFields(2), "Vehicle", vbTextCompare) <= 0 Then
                            Beep
                            gMsgBox "Import file missing Vehicle as as 3rd Column", vbCritical
                            Close hmFrom
                            mCheckFile = False
                            Exit Function
                        Else
                            'If StrComp(slFields(4), "Territory", vbTextCompare) <> 0 Then
                            If StrComp(slFields(3), "Territory", vbTextCompare) <> 0 Then
                                Beep
                                gMsgBox "Import file missing Territory as 4th Column", vbCritical
                                Close hmFrom
                                mCheckFile = False
                                Exit Function
                            Else
                                'If StrComp(slFields(5), "A/E", vbTextCompare) <> 0 Then
                                If StrComp(slFields(4), "A/E", vbTextCompare) <> 0 Then
                                    Beep
                                    gMsgBox "Import file missing A/E as 5th Column", vbCritical
                                    Close hmFrom
                                    mCheckFile = False
                                    Exit Function
                                Else
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If Not ilRankFd Then
        Beep
        gMsgBox "Import file missing Rank as 1st Column", vbCritical
        mCheckFile = False
    Else
        'ilRet = 0
        'On Error GoTo mImportSpotsErr:
        'hmFrom = FreeFile
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet = 0 Then
            lmTotalNoBytes = 3 * LOF(hmFrom)
        Else
            lmTotalNoBytes = 0
        End If
        Close hmFrom
    End If
    Exit Function
'mImportSpotsErr:
'    ilRet = Err.Number
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportAE-mCheckFile"
End Function

Private Sub mCheckForAgreementsNotAssigned()
    Dim slDate As String
    Dim slVehicleName As String
    Dim llVefIndex As Long
    Dim ilShttIndex As Integer
    Dim slCallLetters As String
    Dim ilAnyAgreements As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    
    gLogMsg "Agreements Not Assigned ", "AgreementsAENotAssigned.Txt", True
    ilAnyAgreements = False
    slDate = gNow()
    slDate = gAdjYear(Format$(slDate, sgShowDateForm))
    slDate = Format$(slDate, sgSQLDateForm)
    SQLQuery = "SELECT * FROM att"
    SQLQuery = SQLQuery + " WHERE (attArttCode = 0" & " AND attDropDate >= '" & slDate & "' AND attOffAir >= '" & slDate & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        ilAnyAgreements = True
        llVefIndex = gBinarySearchVef(CLng(rst!attvefCode))
        If llVefIndex <> -1 Then
            slVehicleName = Trim$(tgVehicleInfo(llVefIndex).sVehicle)
        Else
            slVehicleName = "Vehicle Name missing " & rst!attvefCode
        End If
        ilShttIndex = gBinarySearchStationInfoByCode(rst!attshfCode)
        If ilShttIndex <> -1 Then
            slCallLetters = tgStationInfoByCode(ilShttIndex).sCallLetters
        Else
            slCallLetters = "Station Call Letters missing " & rst!attshfCode
        End If
        slStartDate = Format$(rst!attOnAir, sgShowDateForm)
        If DateValue(gAdjYear(rst!attDropDate)) < DateValue(gAdjYear(rst!attOffAir)) Then
            slEndDate = Format$(rst!attDropDate, sgShowDateForm)
        Else
            slEndDate = Format$(rst!attOffAir, sgShowDateForm)
        End If
        gLogMsg slCallLetters & " " & slVehicleName & " " & slStartDate & "-" & slEndDate, "AgreementsAENotAssigned.Txt", False
        rst.MoveNext
    Loop
    If ilAnyAgreements Then
        lacAgreementNoAssigned.Caption = "Agreements Not Assigned: " & "...data\messages\AgreementsAENotAssigned.Txt"
    Else
        lacAgreementNoAssigned.Caption = ""
    End If
End Sub
