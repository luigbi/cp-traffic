VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmImportWWOSpot 
   Caption         =   "Import WWO Traffic Spots"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "AffImportWWOSpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6150
   Begin V81Affiliate.CSI_Calendar txtLogWeekDate 
      Height          =   285
      Left            =   1785
      TabIndex        =   9
      Top             =   825
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "11/8/2010"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   0   'False
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   3
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5955
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   5070
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtImportPath 
      Height          =   300
      Left            =   990
      TabIndex        =   6
      Top             =   195
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   4845
      TabIndex        =   5
      Top             =   195
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   120
      TabIndex        =   1
      Top             =   1635
      Width           =   5790
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5790
      Top             =   4545
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5550
      FormDesignWidth =   6150
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   5025
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3150
      TabIndex        =   3
      Top             =   5025
      Width           =   1575
   End
   Begin VB.Label lacProgress 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   3930
      Width           =   5790
   End
   Begin VB.Label lacFileInfo 
      Caption         =   "(Import: H_Log_WSJ.txt; C_Log_WSJ.txt; Optional S_Log_WSJ.Txt)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   570
      Width           =   3915
   End
   Begin VB.Label lacLogWeekDate 
      Caption         =   "Log Week Start Date"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   870
      Width           =   1590
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import Path"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   210
      Width           =   840
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   150
      TabIndex        =   4
      Top             =   4410
      Width           =   5490
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5790
   End
End
Attribute VB_Name = "frmImportWWOSpot"
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

Private imImporting As Integer
Private imTerminate As Integer
Private smWeekStartDate As String
Private lmWeekStartDate As Long
Private smWeekEndDate As String
Private lmWeekEndDate As Long
Private lmMaxWidth As Long

Private smVehicle As String
Private imVefCode As Integer
Private smVehicleType As String
Private imAdfCode As Integer
Private smAdvertiser As String
Private smAdvAbbv As String
Private smProduct As String
Private lmOrderNumber As Long
Private smLogDate As String
Private smAvailTime As String
Private smLogTime As String
Private imSpotLen As Integer
Private imBreak As Integer
Private imPosition As Integer
Private smISCI As String
Private smCreative As String
Private lmCpfCode As Long
Private smRotStartDate As String
Private smRotEndDate As String
Private lmAttCode As Long
Private imShttCode As Integer
Private smMonDate As String
Private lmCifCode As Long
Private lmLstCode As Long
Private smCallLetters As String
Private imLSTVefCode As Integer
Private lmLstDate As Long

Private tmSALinkInfo() As SALINKINFO
Private tmLstSpot() As LSTSPOT
'Private hmMsg As Integer
Private hmFrom As Integer
Private tmClearImportInfo() As CLEASRIMPORTINFO
Private lst_rst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private vlf_rst As ADODB.Recordset
Private cpf_rst As ADODB.Recordset
Private cif_rst As ADODB.Recordset

Private Sub cmcBrowse_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    igPathType = 0
    sgGetPath = txtImportPath.Text
    frmGetPath.Show vbModal
    If igGetPath = 0 Then
        txtImportPath.Text = sgGetPath
    End If
    
    ChDir slCurDir
    
    Exit Sub
End Sub

Private Sub cmdImport_Click()
    Dim iRet As Integer
    Dim ilFile As Integer

    On Error GoTo ErrHand
    
    lbcMsg.Clear
    lbcMsg.Enabled = True
    Screen.MousePointer = vbHourglass
    If txtLogWeekDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        txtLogWeekDate.SetFocus
        Exit Sub
    End If
    If gIsDate(txtLogWeekDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtLogWeekDate.SetFocus
        Exit Sub
    Else
        smWeekStartDate = Format(txtLogWeekDate.Text, sgShowDateForm)
        smWeekEndDate = DateAdd("d", 6, smWeekStartDate)
        lmWeekStartDate = gDateValue(smWeekStartDate)
        lmWeekEndDate = gDateValue(smWeekEndDate)
    End If
    If Not mCheckFile() Then
        Beep
        lbcMsg.AddItem "Import Stopped"
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons import stopped"
        Screen.MousePointer = vbDefault
        txtImportPath.SetFocus
        Exit Sub
    End If
    lbcMsg.AddItem "Import File Structure is Ok"
    imImporting = True
    On Error GoTo 0
    lacResult.Caption = ""
    ReDim tmClearImportInfo(0 To 0) As CLEASRIMPORTINFO
    iRet = mImportCopy()
    If (iRet = False) Then
        gLogMsg "** Error during Import **", "WWOTrafficSpots.Txt", False
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons Import stopped"
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        'Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        gLogMsg "** User Terminated Import**", "WWOTrafficSpots.Txt", False
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons import stopped"
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        'Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    iRet = mImportSpots()
    If (iRet = False) Then
        gLogMsg "** Error during Import **", "WWOTrafficSpots.Txt", False
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons Import stopped"
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        'Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        gLogMsg "** User Terminated Import**", "WWOTrafficSpots.Txt", False
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons import stopped"
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        'Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    iRet = mImportSplits()
    If (iRet = False) Then
        gLogMsg "** Error during Import **", "WWOTrafficSpots.Txt", False
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons Import stopped"
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        'Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        gLogMsg "** User Terminated Import**", "WWOTrafficSpots.Txt", False
        lacResult.Caption = "see WWOTrafficSpots.Txt for list reasons import stopped"
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        'Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrHand:
    imImporting = False
    gLogMsg "** Completed Import Aired Station Spots" & " **", "WWOTrafficSpots.Txt", False
    gLogMsg "", "WWOTrafficSpots.Txt", False
    'Print #hmMsg, "** Completed Import Aired Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    'Close #hmMsg
    lacResult.Caption = "Completed Import Aired Spots"
    cmdImport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdImportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-cmdImport"
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportWWOSpot
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.5
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    imTerminate = False
    imImporting = False
    lmMaxWidth = 0
    If Len(sgImportDirectory) > 0 Then
        txtImportPath.Text = Left$(sgImportDirectory, Len(sgImportDirectory) - 1)
    Else
        txtImportPath.Text = ""
    End If
    ilRet = gPopAdvertisers()
    ilRet = gPopTeams()
    ilRet = gPopLangs()
    ilRet = gPopAvailNames()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    lacFileInfo.FontSize = 7
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmLstSpot
    Erase tmSALinkInfo
    Erase tmClearImportInfo
    lst_rst.Close
    att_rst.Close
    vlf_rst.Close
    cpf_rst.Close
    cif_rst.Close
    Set frmImportWWOSpot = Nothing
End Sub

Private Function mImportSpots() As Integer
    Dim slFromFile As String
    Dim slPath As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilAdf As Integer
    Dim slChar As String
    Dim ilPrevVefCode As Integer
    Dim llPrevLogDate As Long
    Dim ilLink As Integer
    Dim slTime As String
    Dim llAvailTime As Long
    Dim ilPrevWeekDay As Integer
    
    gLogMsg "Importing: " & "H_Log_WSJ.Txt", "WWOTrafficSpots.Txt", False
    On Error GoTo mImportSpotsErr:
    ilPrevVefCode = -1
    llPrevLogDate = -1
    slPath = txtImportPath.Text
    If right$(slPath, 1) <> "\" Then
        slPath = slPath & "\"
    End If
    'ilRet = 0
    slFromFile = slPath & "H_Log_WSJ.Txt"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read ", hmFrom)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        mImportSpots = False
        Exit Function
    End If
    Do While Not EOF(hmFrom)
        DoEvents
        ilRet = 0
        On Error GoTo mImportSpotsErr:
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
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
            mAddMsgToList "User Cancelled Import"
            mImportSpots = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                smVehicle = Trim$(Mid(slLine, 7, 4))
                imVefCode = -1
                For ilLoop = 0 To UBound(tgSellingVehicleInfo) - 1 Step 1
                    If UCase(Trim$(tgSellingVehicleInfo(ilLoop).sVehicle)) = UCase(smVehicle) Then
                        imVefCode = tgSellingVehicleInfo(ilLoop).iCode
                        smVehicleType = tgSellingVehicleInfo(ilLoop).sVehType
                        Exit For
                    End If
                Next ilLoop
                If imVefCode = -1 Then
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If UCase(Trim$(tgVehicleInfo(ilLoop).sVehicle)) = UCase(smVehicle) Then
                            imVefCode = tgVehicleInfo(ilLoop).iCode
                            smVehicleType = tgVehicleInfo(ilLoop).sVehType
                            Exit For
                        End If
                    Next ilLoop
                End If
                If imVefCode = -1 Then
                    lbcMsg.AddItem "Vehicle " & smVehicle & " Not Found in Import"
                    gLogMsg "Vehicle " & smVehicle & " Not Found in Stopped", "WWOTrafficSpots.Txt", False
                Else
                    smLogDate = Trim$(Mid(slLine, 135, 10))
                    If imVefCode <> ilPrevVefCode Then
                        mClearPrevImport smLogDate
                        imBreak = 0
                        ilPrevWeekDay = -1
                        ilPrevVefCode = imVefCode
                    End If
                    smAdvertiser = Trim$(Mid(slLine, 23, 30))
                    smAdvAbbv = Trim$(Mid(slLine, 15, 5))
                    imAdfCode = -1
                    For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
                        If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), smAdvertiser, vbTextCompare) = 0 Then
                            imAdfCode = tgAdvtInfo(ilAdf).iCode
                            Exit For
                        End If
                    Next ilAdf
                    If imAdfCode = -1 Then
                        'Add Advertiser
                        imAdfCode = mAddAdvt()
                        If imAdfCode = -1 Then
                            lbcMsg.AddItem "Unable to add Advertiser " & smAdvertiser & ", Spot Not Added"
                            gLogMsg "Unable to add Advertiser " & smAdvertiser, "WWOTrafficSpots.Txt", False
                        End If
                    End If
                    If imAdfCode <> -1 Then
                        smProduct = Trim$(Mid(slLine, 65, 35))
                        lmOrderNumber = Val(Trim$(Mid(slLine, 128, 6)))
                        smLogDate = Trim$(Mid(slLine, 135, 10))
                        smMonDate = gObtainPrevMonday(smLogDate)
                        smLogTime = Trim$(Mid(slLine, 146, 8))
                        imSpotLen = Val(Trim$(Mid(slLine, 179, 2)))
                        imPosition = Val(Trim$(Mid(slLine, 183, 2)))
                        If imPosition = 1 Then
                            smAvailTime = smLogTime
                            imBreak = imBreak + 1
                        End If
                        smISCI = Trim$(Mid(slLine, 232, 15))
                        smCreative = ""
                        smRotEndDate = ""
                        'Add CPF or Update CPF
                        ilRet = mGetCpfCode(True)
                        If Not ilRet Then
                            'Add Error message: Generic copy does not exist
                        End If
                        'Add call to gGetAvails
                        'Check Selling avail time against Program times
                        If smVehicleType = "S" Then
                            If ilPrevWeekDay <> Weekday(smLogDate, vbMonday) Then
                                If ilPrevWeekDay = -1 Then
                                    ilRet = mGetSALink()
                                Else
                                    If ilPrevWeekDay = vbSaturday Then
                                        ilRet = mGetSALink()
                                    ElseIf ilPrevWeekDay = vbSunday Then
                                        ilRet = mGetSALink()
                                    ElseIf Weekday(smLogDate, vbMonday) = vbSaturday Then
                                        ilRet = mGetSALink()
                                    ElseIf Weekday(smLogDate, vbMonday) = vbSunday Then
                                        ilRet = mGetSALink()
                                    End If
                                End If
                                ilPrevWeekDay = Weekday(smLogDate, vbMonday)
                            End If
                            llAvailTime = gTimeToLong(smAvailTime, False)
                            For ilLink = 0 To UBound(tmSALinkInfo) - 1 Step 1
                                If tmSALinkInfo(ilLink).lSellTime = llAvailTime Then
                                    slTime = Format(tmSALinkInfo(ilLink).lAirTime, sgShowTimeWSecForm)
                                    lmLstCode = mAddLst(tmSALinkInfo(ilLink).iAirCode, slTime, tmSALinkInfo(ilLink).iBreak, tmSALinkInfo(ilLink).iPosition + imPosition - 1)
                                End If
                            Next ilLink
                        Else
                            lmLstCode = mAddLst(imVefCode, smAvailTime, imBreak, imPosition)
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    On Error GoTo 0
    mImportSpots = True
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mImportSpots"
    Close hmFrom
    mImportSpots = False
    Exit Function
End Function
Private Function mImportCopy() As Integer
    Dim slFromFile As String
    Dim slPath As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilAdf As Integer
    Dim slChar As String
    Dim slTime As String
    
    gLogMsg "Importing: " & "C_Log_WSJ.Txt", "WWOTrafficSpots.Txt", False
    On Error GoTo mImportCopyErr:
    slPath = txtImportPath.Text
    If right$(slPath, 1) <> "\" Then
        slPath = slPath & "\"
    End If
    'ilRet = 0
    slFromFile = slPath & "C_Log_WSJ.Txt"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read ", hmFrom)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        mImportCopy = False
        Exit Function
    End If
    Do While Not EOF(hmFrom)
        DoEvents
        ilRet = 0
        On Error GoTo mImportCopyErr:
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
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
            mAddMsgToList "User Cancelled Import"
            mImportCopy = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                smAdvertiser = Trim$(Mid(slLine, 8, 30))
                smAdvAbbv = ""
                imAdfCode = -1
                For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
                    If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), smAdvertiser, vbTextCompare) = 0 Then
                        imAdfCode = tgAdvtInfo(ilAdf).iCode
                        Exit For
                    End If
                Next ilAdf
                If imAdfCode = -1 Then
                    'Add Advertiser
                    imAdfCode = mAddAdvt()
                    If imAdfCode = -1 Then
                        lbcMsg.AddItem "Unable to add Advertiser " & smAdvertiser & ", Spot Not Added"
                        gLogMsg "Unable to add Advertiser " & smAdvertiser, "WWOTrafficSpots.Txt", False
                    End If
                End If
                If imAdfCode <> -1 Then
                    smProduct = Trim$(Mid(slLine, 40, 35))
                    smISCI = Trim$(Mid(slLine, 102, 13))
                    smCreative = Trim$(Mid(slLine, 119, 30))
                    imSpotLen = Val(Trim$(Mid(slLine, 181, 2)))
                    smRotStartDate = Trim$(Mid(slLine, 208, 10))
                    smRotEndDate = Trim$(Mid(slLine, 220, 10))
                    smLogDate = ""
                    'Add CPF or Update CPF
                    ilRet = mGetCpfCode(False)
                    ilRet = mGetCifCode()
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    On Error GoTo 0
    mImportCopy = True
    Exit Function
mImportCopyErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mImportCopy"
    Close hmFrom
    mImportCopy = False
    Exit Function
End Function

Private Function mImportSplits() As Integer
    Dim slFromFile As String
    Dim slPath As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilAdf As Integer
    Dim slChar As String
    Dim ilPrevVefCode As Integer
    Dim llPrevLogDate As Long
    Dim ilLink As Integer
    Dim slTime As String
    Dim llAvailTime As Long
    Dim ilPrevWeekDay As Integer
    Dim slStr As String
    
    gLogMsg "Importing: " & "S_Log_WSJ.Txt", "WWOTrafficSpots.Txt", False
    On Error GoTo mImportSplitsErr:
    ilPrevVefCode = -1
    llPrevLogDate = -1
    slPath = txtImportPath.Text
    If right$(slPath, 1) <> "\" Then
        slPath = slPath & "\"
    End If
    'ilRet = 0
    slFromFile = slPath & "S_Log_WSJ.Txt"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read ", hmFrom)
    If ilRet <> 0 Then
        'mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        'mImportSplits = False
        mImportSplits = True
        Exit Function
    End If
    Do While Not EOF(hmFrom)
        DoEvents
        ilRet = 0
        On Error GoTo mImportSplitsErr:
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
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
            mAddMsgToList "User Cancelled Import"
            mImportSplits = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                smVehicle = Trim$(Mid(slLine, 7, 4))
                ilRet = gParseItem(slLine, 1, Chr(9), smVehicle)
                imVefCode = -1
                For ilLoop = 0 To UBound(tgSellingVehicleInfo) - 1 Step 1
                    If UCase(Trim$(tgSellingVehicleInfo(ilLoop).sVehicle)) = UCase(smVehicle) Then
                        imVefCode = tgSellingVehicleInfo(ilLoop).iCode
                        smVehicleType = tgSellingVehicleInfo(ilLoop).sVehType
                        Exit For
                    End If
                Next ilLoop
                If imVefCode = -1 Then
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If UCase(Trim$(tgVehicleInfo(ilLoop).sVehicle)) = UCase(smVehicle) Then
                            imVefCode = tgVehicleInfo(ilLoop).iCode
                            smVehicleType = tgVehicleInfo(ilLoop).sVehType
                            Exit For
                        End If
                    Next ilLoop
                End If
                If imVefCode = -1 Then
                    lbcMsg.AddItem "Vehicle " & smVehicle & " Not Found in Import"
                    gLogMsg "Vehicle " & smVehicle & " Not Found in Stopped", "WWOTrafficSpots.Txt", False
                Else
                    ilRet = gParseItem(slLine, 2, Chr(9), smLogDate)
                    If (imVefCode <> ilPrevVefCode) Or (llPrevLogDate <> gDateValue(smLogDate)) Then
                        ilPrevVefCode = imVefCode
                        llPrevLogDate = gDateValue(smLogDate)
                    End If
                    ilRet = gParseItem(slLine, 7, Chr(9), smAdvertiser)
                    smAdvAbbv = ""
                    imAdfCode = -1
                    For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
                        If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), smAdvertiser, vbTextCompare) = 0 Then
                            imAdfCode = tgAdvtInfo(ilAdf).iCode
                            Exit For
                        End If
                    Next ilAdf
                    If imAdfCode = -1 Then
                        'Add Advertiser
                        imAdfCode = mAddAdvt()
                        If imAdfCode = -1 Then
                            lbcMsg.AddItem "Unable to add Advertiser " & smAdvertiser & ", Spot Not Added"
                            gLogMsg "Unable to add Advertiser " & smAdvertiser, "WWOTrafficSpots.Txt", False
                        End If
                    End If
                    If imAdfCode <> -1 Then
                        ilRet = gParseItem(slLine, 10, Chr(9), smProduct)
                        ilRet = gParseItem(slLine, 2, Chr(9), smLogDate)
                        ilRet = gParseItem(slLine, 3, Chr(9), smLogTime)
                        ilRet = gParseItem(slLine, 4, Chr(9), slStr)
                        imPosition = Val(slStr)
                        ilRet = gParseItem(slLine, 5, Chr(9), slStr)
                        imSpotLen = Val(slStr)
                        If imPosition = 1 Then
                            smAvailTime = smLogTime
                        End If
                        ilRet = gParseItem(slLine, 6, Chr(9), smCallLetters)
                        ilRet = gParseItem(slLine, 7, Chr(9), smISCI)
                        ilRet = gParseItem(slLine, 11, Chr(9), smCreative)
                        ilRet = gParseItem(slLine, 12, Chr(9), smRotStartDate)
                        ilRet = gParseItem(slLine, 13, Chr(9), smRotEndDate)
                        If smVehicleType = "S" Then
                            If ilPrevWeekDay <> Weekday(smLogDate, vbMonday) Then
                                If ilPrevWeekDay = -1 Then
                                    ilRet = mGetSALink()
                                Else
                                    If ilPrevWeekDay = vbSaturday Then
                                        ilRet = mGetSALink()
                                    ElseIf ilPrevWeekDay = vbSunday Then
                                        ilRet = mGetSALink()
                                    ElseIf Weekday(smLogDate, vbMonday) = vbSaturday Then
                                        ilRet = mGetSALink()
                                    ElseIf Weekday(smLogDate, vbMonday) = vbSunday Then
                                        ilRet = mGetSALink()
                                    End If
                                End If
                                ilPrevWeekDay = Weekday(smLogDate, vbMonday)
                            End If
                            llAvailTime = gTimeToLong(smAvailTime, False)
                            For ilLink = 0 To UBound(tmSALinkInfo) - 1 Step 1
                                If ((tmSALinkInfo(ilLink).lSellTime = llAvailTime) And (imPosition = 1)) Or ((tmSALinkInfo(ilLink).lSellTime + 30 = llAvailTime) And (imPosition = 2)) Then
                                    slTime = Format(tmSALinkInfo(ilLink).lAirTime, sgShowTimeWSecForm)
                                    lmLstCode = mFindLST(tmSALinkInfo(ilLink).iAirCode, slTime, tmSALinkInfo(ilLink).iPosition + imPosition - 1)
                                    If lmLstCode <> -1 Then
                                        Exit For
                                    End If
                                End If
                            Next ilLink
                        Else
                            lmLstCode = mFindLST(imVefCode, smLogTime, imPosition)
                        End If
                        If lmLstCode <> -1 Then
                            'Add CPF or Update CPF
                            ilRet = mGetCpfCode(False)
                            ilRet = mGetCifCode()
                            'add new file
                        Else
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    On Error GoTo 0
    mImportSplits = True
    Exit Function
mImportSplitsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mImportSplits"
    Close hmFrom
    mImportSplits = False
    Exit Function
End Function
Private Function mCheckFile() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    'Dim slFields(1 To 13) As String
    Dim slFields(0 To 12) As String
    Dim ilFile As Integer
    Dim slPath As String
    Dim slChar As String
    Dim slDate As String
    Dim smVehicle As String
    Dim ilFound As Integer
    Dim ilCount As Integer
    
    mCheckFile = True
    lacProgress.Caption = "Checking Station Info in H_Log_WSJ.Txt..."
    DoEvents
    On Error GoTo mCheckFileErr:
    slPath = txtImportPath.Text
    If right$(slPath, 1) <> "\" Then
        slPath = slPath & "\"
    End If
    'ilRet = 0
    slFromFile = slPath & "H_Log_WSJ.Txt"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read ", hmFrom)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        mCheckFile = False
        Exit Function
    End If
    ilCount = 0
    Do While Not EOF(hmFrom)
        DoEvents
        ilRet = 0
        On Error GoTo mCheckFileErr:
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
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
            mAddMsgToList "User Cancelled Import"
            mCheckFile = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                slChar = Left$(slLine, 1)
                If slChar <> "T" Then
                    mAddMsgToList "Spot File Form in Error as 'T' missing from first column " & slLine
                    mCheckFile = False
                    Exit Function
                End If
                slChar = Mid(slLine, 4, 1)
                If (slChar <> "R") And (slChar <> "M") Then
                    mAddMsgToList "Spot File Form in Error as 'R' or 'M' missing from 4th column " & slLine
                    mCheckFile = False
                    Exit Function
                End If
                smVehicle = Mid(slLine, 7, 4)
                'Verify name
                ilFound = False
                For ilLoop = 0 To UBound(tgSellingVehicleInfo) - 1 Step 1
                    If UCase(Trim$(tgSellingVehicleInfo(ilLoop).sVehicle)) = UCase(smVehicle) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If UCase(Trim$(tgVehicleInfo(ilLoop).sVehicle)) = UCase(smVehicle) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If Not ilFound Then
                    mAddMsgToList "Vehicle Name not found within System " & slLine
                    mCheckFile = False
                    Exit Function
                End If
                slDate = Mid(slLine, 135, 10)
                'Check if date and within Import date range
                If Not gIsDate(slDate) Then
                    mAddMsgToList "Spot Log Date not valid form " & slLine
                    mCheckFile = False
                    Exit Function
                End If
                If (gDateValue(slDate) < lmWeekStartDate) Or (gDateValue(slDate) > lmWeekEndDate) Then
                    mAddMsgToList "Spot Log Date " & slDate & " not within Week Import Date " & smWeekStartDate & "-" & smWeekEndDate
                    mCheckFile = False
                    Exit Function
                End If
                ilCount = ilCount + 1
                If ilCount > 5 Then
                    Exit Do
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    
    On Error GoTo mCheckFileErr:
    'ilRet = 0
    slFromFile = slPath & "C_Log_WSJ.Txt"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read ", hmFrom)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        mCheckFile = False
        Exit Function
    End If
    ilCount = 0
    Do While Not EOF(hmFrom)
        DoEvents
        ilRet = 0
        On Error GoTo mCheckFileErr:
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
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
            mAddMsgToList "User Cancelled Import"
            mCheckFile = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                'slChar = Left$(slLine, 1)
                'If slChar <> "1" Then
                '    mAddMsgToList "Copy File Form in Error as '1' missing from first column " & slLine
                '    mCheckFile = False
                '    Exit Function
                'End If
                slDate = Mid(slLine, 208, 10)
                If Not gIsDate(slDate) Then
                    mAddMsgToList "Copy Earliest Date not valid form " & slLine
                    mCheckFile = False
                    Exit Function
                End If
                slDate = Mid(slLine, 220, 10)
                If Not gIsDate(slDate) Then
                    mAddMsgToList "Copy Latest Date not valid form " & slLine
                    mCheckFile = False
                    Exit Function
                End If
                ilCount = ilCount + 1
                If ilCount > 5 Then
                    Exit Do
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    
    On Error GoTo mCheckFileErr:
    'ilRet = 0
    slFromFile = slPath & "S_Log_WSJ.Txt"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read ", hmFrom)
    If ilRet = 0 Then
        ilCount = 0
        Do While Not EOF(hmFrom)
            DoEvents
            ilRet = 0
            On Error GoTo mCheckFileErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
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
                mAddMsgToList "User Cancelled Import"
                mCheckFile = False
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    ilRet = gParseItem(slLine, 2, Chr(9), slFields(0))
                    ilRet = gParseItem(slLine, 3, Chr(9), slFields(1))
                    ilRet = gParseItem(slLine, 12, Chr(9), slFields(2))
                    'Check if date and within Import date range
                    If Not gIsDate(slFields(0)) Then
                        mAddMsgToList "Split Copy Log Date not valid form " & slLine
                        mCheckFile = False
                        Exit Function
                    End If
                    If (gDateValue(slFields(0)) < lmWeekStartDate) Or (gDateValue(slFields(0)) > lmWeekEndDate) Then
                        mAddMsgToList "Split Copy Log Date " & slFields(0) & " not within Week Import Date " & smWeekStartDate & "-" & smWeekEndDate
                        mCheckFile = False
                        Exit Function
                    End If
                    'Check if time
                    If Not gIsTime(slFields(1)) Then
                        mAddMsgToList "Split Copy Log Time not valid form " & slLine
                        mCheckFile = False
                        Exit Function
                    End If
                    'Check if date
                    If Not gIsDate(slFields(2)) Then
                        mAddMsgToList "Split Copy Earliest Date not valid form " & slLine
                        mCheckFile = False
                        Exit Function
                    End If
                    ilCount = ilCount + 1
                    If ilCount > 5 Then
                        Exit Do
                    End If
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
    End If
    mCheckFile = True
    On Error GoTo 0
    Exit Function
mCheckFileErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mCheckFile"
    mCheckFile = False
End Function



Private Function mAddLst(ilVefCode As Integer, slAvailTime As String, ilBreak As Integer, ilPosition As Integer) As Long
    Dim tlLST As LST
    Dim ilRet As Integer

    On Error GoTo ErrHand
    

    tlLST.lCode = 0
    tlLST.iType = 0
    tlLST.lSdfCode = 0
    tlLST.lCntrNo = 0
    tlLST.iAdfCode = imAdfCode
    tlLST.iAgfCode = 0
    tlLST.sProd = smProduct
    tlLST.iLineNo = 0
    tlLST.iLnVefCode = 0
    tlLST.sStartDate = Format$("1/1/1970", sgShowDateForm)
    tlLST.sEndDate = Format$("1/1/1970", sgShowDateForm)
    tlLST.iMon = 0
    tlLST.iTue = 0
    tlLST.iWed = 0
    tlLST.iThu = 0
    tlLST.iFri = 0
    tlLST.iSat = 0
    tlLST.iSun = 0
    tlLST.iSpotsWk = 0
    tlLST.iPriceType = 1
    tlLST.lPrice = 0
    tlLST.iSpotType = 5
    tlLST.iLogVefCode = ilVefCode
    tlLST.sLogDate = Format$(smLogDate, sgShowDateForm)
    tlLST.sLogTime = Format$(slAvailTime, sgShowTimeWSecForm)
    tlLST.sDemo = ""
    tlLST.lAud = 0
    tlLST.sISCI = gFixQuote(smISCI)
    tlLST.iWkNo = 0
    tlLST.iBreakNo = ilBreak
    tlLST.iPositionNo = ilPosition
    tlLST.iSeqNo = 0
    tlLST.sZone = ""
    tlLST.sCart = ""
    tlLST.lCpfCode = lmCpfCode
    tlLST.lCrfCsfCode = 0
    tlLST.iStatus = 0
    tlLST.iLen = imSpotLen
    tlLST.iUnits = 0
    tlLST.lCifCode = lmCifCode
    tlLST.iAnfCode = 0
    tlLST.lEvtIDCefCode = 0
    tlLST.sSplitNetwork = "N"
    tlLST.lRafCode = 0
    tlLST.lFsfCode = 0
    tlLST.lgsfCode = 0
    tlLST.sImportedSpot = "Y"
    tlLST.lBkoutLstCode = 0
    tlLST.sLnStartTime = "12am"
    tlLST.sLnEndTime = "12am"
    tlLST.sUnused = ""
    
    SQLQuery = "Insert Into lst ( "
    SQLQuery = SQLQuery & "lstCode, "
    SQLQuery = SQLQuery & "lstType, "
    SQLQuery = SQLQuery & "lstSdfCode, "
    SQLQuery = SQLQuery & "lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, "
    SQLQuery = SQLQuery & "lstAgfCode, "
    SQLQuery = SQLQuery & "lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, "
    SQLQuery = SQLQuery & "lstLnVefCode, "
    SQLQuery = SQLQuery & "lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, "
    SQLQuery = SQLQuery & "lstMon, "
    SQLQuery = SQLQuery & "lstTue, "
    SQLQuery = SQLQuery & "lstWed, "
    SQLQuery = SQLQuery & "lstThu, "
    SQLQuery = SQLQuery & "lstFri, "
    SQLQuery = SQLQuery & "lstSat, "
    SQLQuery = SQLQuery & "lstSun, "
    SQLQuery = SQLQuery & "lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, "
    SQLQuery = SQLQuery & "lstPrice, "
    SQLQuery = SQLQuery & "lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, "
    SQLQuery = SQLQuery & "lstLogDate, "
    SQLQuery = SQLQuery & "lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, "
    SQLQuery = SQLQuery & "lstAud, "
    SQLQuery = SQLQuery & "lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, "
    SQLQuery = SQLQuery & "lstBreakNo, "
    SQLQuery = SQLQuery & "lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, "
    SQLQuery = SQLQuery & "lstZone, "
    SQLQuery = SQLQuery & "lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, "
    SQLQuery = SQLQuery & "lstCrfCsfCode, "
    SQLQuery = SQLQuery & "lstStatus, "
    SQLQuery = SQLQuery & "lstLen, "
    SQLQuery = SQLQuery & "lstUnits, "
    SQLQuery = SQLQuery & "lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, "
    SQLQuery = SQLQuery & "lstEvtIDCefCode, "
    SQLQuery = SQLQuery & "lstSplitNetwork, "
    SQLQuery = SQLQuery & "lstRafCode, "
    SQLQuery = SQLQuery & "lstFsfCode, "
    SQLQuery = SQLQuery & "lstGsfCode, "
    SQLQuery = SQLQuery & "lstImportedSpot, "
    SQLQuery = SQLQuery & "lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, "
    SQLQuery = SQLQuery & "lstLnEndTime, "
    SQLQuery = SQLQuery & "lstUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlLST.lCode & ", "
    SQLQuery = SQLQuery & tlLST.iType & ", "
    SQLQuery = SQLQuery & tlLST.lSdfCode & ", "
    SQLQuery = SQLQuery & tlLST.lCntrNo & ", "
    SQLQuery = SQLQuery & tlLST.iAdfCode & ", "
    SQLQuery = SQLQuery & tlLST.iAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sProd) & "', "
    SQLQuery = SQLQuery & tlLST.iLineNo & ", "
    SQLQuery = SQLQuery & tlLST.iLnVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sEndDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlLST.iMon & ", "
    SQLQuery = SQLQuery & tlLST.iTue & ", "
    SQLQuery = SQLQuery & tlLST.iWed & ", "
    SQLQuery = SQLQuery & tlLST.iThu & ", "
    SQLQuery = SQLQuery & tlLST.iFri & ", "
    SQLQuery = SQLQuery & tlLST.iSat & ", "
    SQLQuery = SQLQuery & tlLST.iSun & ", "
    SQLQuery = SQLQuery & tlLST.iSpotsWk & ", "
    SQLQuery = SQLQuery & tlLST.iPriceType & ", "
    SQLQuery = SQLQuery & tlLST.lPrice & ", "
    SQLQuery = SQLQuery & tlLST.iSpotType & ", "
    SQLQuery = SQLQuery & tlLST.iLogVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLogDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLogTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sDemo) & "', "
    SQLQuery = SQLQuery & tlLST.lAud & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sISCI) & "', "
    SQLQuery = SQLQuery & tlLST.iWkNo & ", "
    SQLQuery = SQLQuery & tlLST.iBreakNo & ", "
    SQLQuery = SQLQuery & tlLST.iPositionNo & ", "
    SQLQuery = SQLQuery & tlLST.iSeqNo & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sCart) & "', "
    SQLQuery = SQLQuery & tlLST.lCpfCode & ", "
    SQLQuery = SQLQuery & tlLST.lCrfCsfCode & ", "
    SQLQuery = SQLQuery & tlLST.iStatus & ", "
    SQLQuery = SQLQuery & tlLST.iLen & ", "
    SQLQuery = SQLQuery & tlLST.iUnits & ", "
    SQLQuery = SQLQuery & tlLST.lCifCode & ", "
    SQLQuery = SQLQuery & tlLST.iAnfCode & ", "
    SQLQuery = SQLQuery & tlLST.lEvtIDCefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sSplitNetwork) & "', "
    SQLQuery = SQLQuery & tlLST.lRafCode & ", "
    SQLQuery = SQLQuery & tlLST.lFsfCode & ", "
    SQLQuery = SQLQuery & tlLST.lgsfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sImportedSpot) & "', "
    SQLQuery = SQLQuery & tlLST.lBkoutLstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLnStartTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLnEndTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sUnused) & "' "
    SQLQuery = SQLQuery & ")"

    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.Txt", "Import WWO Spots-mAddLst"
        mAddLst = False
        Exit Function
    End If

    On Error GoTo 0
    mAddLst = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mAddLST"
    mAddLst = False
End Function

Private Function mAddAdvt() As Integer
    Dim tlAdf As ADF
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim imAdfCode As Integer
    Dim ilCode As Integer

    On Error GoTo ErrHand
    ilCode = -1
    tlAdf.iCode = 0
    tlAdf.sName = smAdvertiser
    tlAdf.sAbbr = smAdvAbbv
    tlAdf.sProd = ""
    tlAdf.iSlfCode = 0
    tlAdf.iAgfCode = 0
    tlAdf.sBuyer = ""
    tlAdf.sCodeRep = ""
    tlAdf.sCodeAgy = ""
    tlAdf.sCodeStn = ""
    tlAdf.imnfComp2 = 0
    tlAdf.imnfExcl1 = 0
    tlAdf.imnfExcl2 = 0
    tlAdf.sCppCpm = ""
    tlAdf.sDemo1 = ""
    tlAdf.sDemo2 = ""
    tlAdf.sDemo3 = ""
    tlAdf.sDemo4 = ""
    tlAdf.imnfDemo1 = 0
    tlAdf.imnfDemo2 = 0
    tlAdf.imnfDemo3 = 0
    tlAdf.imnfDemo4 = 0
    tlAdf.lTarget1 = 0
    tlAdf.lTarget2 = 0
    tlAdf.lTarget3 = 0
    tlAdf.lTarget4 = 0
    tlAdf.sCreditRestr = "N"
    tlAdf.lCreditLimit = 0
    tlAdf.sPaymRating = "1"
    tlAdf.sISCI = ""
    tlAdf.imnfSort = 0
    tlAdf.sBilAgyDir = ""
    tlAdf.sCntrAddr1 = ""
    tlAdf.sCntrAddr2 = ""
    tlAdf.sCntrAddr3 = ""
    tlAdf.sBillAddr1 = ""
    tlAdf.sBillAddr2 = ""
    tlAdf.sBillAddr3 = ""
    tlAdf.iarfLkCode = 0
    tlAdf.sPhone = ""
    tlAdf.sFax = ""
    tlAdf.iarfContrCode = 0
    tlAdf.iarfInvCode = 0
    tlAdf.sCntrPrtSz = ""
    tlAdf.iTrfCode = 0
    tlAdf.sCrdApp = "R"
    tlAdf.sCrdRtg = ""
    tlAdf.ipnfBuyer = 0
    tlAdf.ipnfPay = 0
    tlAdf.iPct90 = 0
    tlAdf.sCurrAR = ""  'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sUnbilled = ""    'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sHiCredit = ""    'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sTotalGross = ""  'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sDateEntrd = Format$(gNow(), sgShowDateForm)
    tlAdf.iNSFChks = 0
    tlAdf.sDateLstInv = Format$("1/1/1970", sgShowDateForm)
    tlAdf.sDateLstPaym = Format$("1/1/1970", sgShowDateForm)
    tlAdf.iAvgToPay = 0
    tlAdf.iLstToPay = 0
    tlAdf.iNoInvPd = 0
    tlAdf.sNewBus = ""
    tlAdf.sEndDate = Format$("1/1/1970", sgShowDateForm)
    tlAdf.iMerge = 0
    tlAdf.iurfCode = 0
    tlAdf.sState = "A"
    tlAdf.sCrdAppDate = Format$("1/1/1970", sgShowDateForm)
    tlAdf.sCrdAppTime = Format$("1/1/1970", sgShowTimeWSecForm)
    tlAdf.sPkInvShow = ""
    tlAdf.lGuar = 0
    tlAdf.sBkoutPoolStatus = "N"
    tlAdf.sUnused2 = ""
    tlAdf.sRateOnInv = ""
    tlAdf.iMnfBus = 0
    tlAdf.iUnused1 = 0
    tlAdf.sAllowRepMG = ""
    tlAdf.sBonusOnInv = ""
    tlAdf.sRepInvGen = ""
    tlAdf.iMnfInvTerms = 0
    tlAdf.sPolitical = "N"
    tlAdf.sAddrID = ""
    
    
    SQLQuery = "Insert Into ADF_Advertisers ( "
    SQLQuery = SQLQuery & "adfCode, "
    SQLQuery = SQLQuery & "adfName, "
    SQLQuery = SQLQuery & "adfAbbr, "
    SQLQuery = SQLQuery & "adfProd, "
    SQLQuery = SQLQuery & "adfslfCode, "
    SQLQuery = SQLQuery & "adfagfCode, "
    SQLQuery = SQLQuery & "adfBuyer, "
    SQLQuery = SQLQuery & "adfCodeRep, "
    SQLQuery = SQLQuery & "adfCodeAgy, "
    SQLQuery = SQLQuery & "adfCodeStn, "
    SQLQuery = SQLQuery & "adfmnfComp1, "
    SQLQuery = SQLQuery & "adfmnfComp2, "
    SQLQuery = SQLQuery & "adfmnfExcl1, "
    SQLQuery = SQLQuery & "adfmnfExcl2, "
    SQLQuery = SQLQuery & "adfCppCpm, "
    SQLQuery = SQLQuery & "adfDemo1, "
    SQLQuery = SQLQuery & "adfDemo2, "
    SQLQuery = SQLQuery & "adfDemo3, "
    SQLQuery = SQLQuery & "adfDemo4, "
    SQLQuery = SQLQuery & "adfmnfDemo1, "
    SQLQuery = SQLQuery & "adfmnfDemo2, "
    SQLQuery = SQLQuery & "adfmnfDemo3, "
    SQLQuery = SQLQuery & "adfmnfDemo4, "
    SQLQuery = SQLQuery & "adfTarget1, "
    SQLQuery = SQLQuery & "adfTarget2, "
    SQLQuery = SQLQuery & "adfTarget3, "
    SQLQuery = SQLQuery & "adfTarget4, "
    SQLQuery = SQLQuery & "adfCreditRestr, "
    SQLQuery = SQLQuery & "adfCreditLimit, "
    SQLQuery = SQLQuery & "adfPaymRating, "
    SQLQuery = SQLQuery & "adfISCI, "
    SQLQuery = SQLQuery & "adfmnfSort, "
    SQLQuery = SQLQuery & "adfBilAgyDir, "
    SQLQuery = SQLQuery & "adfCntrAddr1, "
    SQLQuery = SQLQuery & "adfCntrAddr2, "
    SQLQuery = SQLQuery & "adfCntrAddr3, "
    SQLQuery = SQLQuery & "adfBillAddr1, "
    SQLQuery = SQLQuery & "adfBillAddr2, "
    SQLQuery = SQLQuery & "adfBillAddr3, "
    SQLQuery = SQLQuery & "adfarfLkCode, "
    SQLQuery = SQLQuery & "adfPhone, "
    SQLQuery = SQLQuery & "adfFax, "
    SQLQuery = SQLQuery & "adfarfContrCode, "
    SQLQuery = SQLQuery & "adfarfInvCode, "
    SQLQuery = SQLQuery & "adfCntrPrtSz, "
    SQLQuery = SQLQuery & "adfTrfCode, "
    SQLQuery = SQLQuery & "adfCrdApp, "
    SQLQuery = SQLQuery & "adfCrdRtg, "
    SQLQuery = SQLQuery & "adfpnfBuyer, "
    SQLQuery = SQLQuery & "adfpnfPay, "
    SQLQuery = SQLQuery & "adfPct90, "
    SQLQuery = SQLQuery & "adfCurrAR, "
    SQLQuery = SQLQuery & "adfUnbilled, "
    SQLQuery = SQLQuery & "adfHiCredit, "
    SQLQuery = SQLQuery & "adfTotalGross, "
    SQLQuery = SQLQuery & "adfDateEntrd, "
    SQLQuery = SQLQuery & "adfNSFChks, "
    SQLQuery = SQLQuery & "adfDateLstInv, "
    SQLQuery = SQLQuery & "adfDateLstPaym, "
    SQLQuery = SQLQuery & "adfAvgToPay, "
    SQLQuery = SQLQuery & "adfLstToPay, "
    SQLQuery = SQLQuery & "adfNoInvPd, "
    SQLQuery = SQLQuery & "adfNewBus, "
    SQLQuery = SQLQuery & "adfEndDate, "
    SQLQuery = SQLQuery & "adfMerge, "
    SQLQuery = SQLQuery & "adfurfCode, "
    SQLQuery = SQLQuery & "adfState, "
    SQLQuery = SQLQuery & "adfCrdAppDate, "
    SQLQuery = SQLQuery & "adfCrdAppTime, "
    SQLQuery = SQLQuery & "adfPkInvShow, "
    SQLQuery = SQLQuery & "adfGuar, "
    SQLQuery = SQLQuery & "adfBkoutPoolStatus, "
    SQLQuery = SQLQuery & "adfUnused2, "
    SQLQuery = SQLQuery & "adfRateOnInv, "
    SQLQuery = SQLQuery & "adfMnfBus, "
    SQLQuery = SQLQuery & "adfUnused1, "
    SQLQuery = SQLQuery & "adfAllowRepMG, "
    SQLQuery = SQLQuery & "adfBonusOnInv, "
    SQLQuery = SQLQuery & "adfRepInvGen, "
    SQLQuery = SQLQuery & "adfMnfInvTerms, "
    SQLQuery = SQLQuery & "adfPolitical, "
    SQLQuery = SQLQuery & "adfAddrID "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sAbbr) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sProd) & "', "
    SQLQuery = SQLQuery & tlAdf.iSlfCode & ", "
    SQLQuery = SQLQuery & tlAdf.iAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBuyer) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCodeRep) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCodeAgy) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCodeStn) & "', "
    SQLQuery = SQLQuery & tlAdf.imnfComp1 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfComp2 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfExcl1 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfExcl2 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCppCpm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo4) & "', "
    SQLQuery = SQLQuery & tlAdf.imnfDemo1 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfDemo2 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfDemo3 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfDemo4 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget1 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget2 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget3 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget4 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCreditRestr) & "', "
    SQLQuery = SQLQuery & tlAdf.lCreditLimit & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPaymRating) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sISCI) & "', "
    SQLQuery = SQLQuery & tlAdf.imnfSort & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBilAgyDir) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrAddr1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrAddr2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrAddr3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBillAddr1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBillAddr2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBillAddr3) & "', "
    SQLQuery = SQLQuery & tlAdf.iarfLkCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPhone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sFax) & "', "
    SQLQuery = SQLQuery & tlAdf.iarfContrCode & ", "
    SQLQuery = SQLQuery & tlAdf.iarfInvCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrPrtSz) & "', "
    SQLQuery = SQLQuery & tlAdf.iTrfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCrdApp) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCrdRtg) & "', "
    SQLQuery = SQLQuery & tlAdf.ipnfBuyer & ", "
    SQLQuery = SQLQuery & tlAdf.ipnfPay & ", "
    SQLQuery = SQLQuery & tlAdf.iPct90 & ", "
    SQLQuery = SQLQuery & "'" & tlAdf.sCurrAR & "', "
    SQLQuery = SQLQuery & "'" & tlAdf.sUnbilled & "', "
    SQLQuery = SQLQuery & "'" & tlAdf.sHiCredit & "', "
    SQLQuery = SQLQuery & "'" & tlAdf.sTotalGross & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sDateEntrd, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlAdf.iNSFChks & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sDateLstInv, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sDateLstPaym, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlAdf.iAvgToPay & ", "
    SQLQuery = SQLQuery & tlAdf.iLstToPay & ", "
    SQLQuery = SQLQuery & tlAdf.iNoInvPd & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sNewBus) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sEndDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlAdf.iMerge & ", "
    SQLQuery = SQLQuery & tlAdf.iurfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sState) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sCrdAppDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sCrdAppTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPkInvShow) & "', "
    SQLQuery = SQLQuery & tlAdf.lGuar & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBkoutPoolStatus) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sUnused2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sRateOnInv) & "', "
    SQLQuery = SQLQuery & tlAdf.iMnfBus & ", "
    SQLQuery = SQLQuery & tlAdf.iUnused1 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sAllowRepMG) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBonusOnInv) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sRepInvGen) & "', "
    SQLQuery = SQLQuery & tlAdf.iMnfInvTerms & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPolitical) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sAddrID) & "' "
    SQLQuery = SQLQuery & ")"
    
    ilCode = gInsertAndReturnCode(SQLQuery, "ADF_Advertisers", "adfCode", "Replace")
    If ilCode > 0 Then
        ilRet = gPopAdvertisers()
        mAddAdvt = ilCode
    End If
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mAddAdvt"
    mAddAdvt = -1
End Function


Private Sub mClearPrevImport(slClearDate As String)
    Dim ilCheck As Integer
    Dim ilDateClearedPrev As Integer
    Dim llClearDate As Long
    Dim llLst As Long
    Dim slLstDate As String
    Dim slMonDate As String
    Dim slSunDate As String
    ReDim llLsfCode(0 To 0) As Long
    
    On Error GoTo ErrHand
    
    ilDateClearedPrev = False
    llClearDate = DateValue(slClearDate)
    slMonDate = gObtainPrevMonday(slClearDate)
    slSunDate = gObtainNextSunday(slMonDate)
    For ilCheck = 0 To UBound(tmClearImportInfo) - 1 Step 1
        If (tmClearImportInfo(ilCheck).iVefCode = imVefCode) Then
            If tmClearImportInfo(ilCheck).lClearDate = llClearDate Then
                ilDateClearedPrev = True
            End If
            If ilDateClearedPrev Then
                Exit Sub
            End If
        End If
    Next ilCheck
    
    If Not ilDateClearedPrev Then
        SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "')" & ")"
        Set att_rst = gSQLSelectCall(SQLQuery)
        Do While Not att_rst.EOF
            
            SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & att_rst!attCode
            SQLQuery = SQLQuery & " AND astFeedDate = '" & Format$(slClearDate, sgSQLDateForm) & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.Txt", "Import WWO Spots-mClearPrevImport"
                Exit Sub
            End If
            
            SQLQuery = "DELETE FROM Aet WHERE (aetAtfCode = " & att_rst!attCode
            SQLQuery = SQLQuery & " AND aetFeedDate = '" & Format$(slClearDate, sgSQLDateForm) & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.Txt", "Import WWO Spots-mClearPrevImport"
                Exit Sub
            End If
            
            'Doug-Remove spots from web for this att and date
            
            att_rst.MoveNext
        Loop
        
        SQLQuery = "DELETE FROM lst WHERE (lstLogVefCode = " & imVefCode & " AND (lstLogDate = '" & Format$(slClearDate, sgSQLDateForm) & "')" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "Import WWO Spots-mClearPrevImport"
            Exit Sub
        End If
    
        tmClearImportInfo(UBound(tmClearImportInfo)).iVefCode = imVefCode
        tmClearImportInfo(UBound(tmClearImportInfo)).lClearDate = llClearDate
        tmClearImportInfo(UBound(tmClearImportInfo)).lgsfCode = 0
        ReDim Preserve tmClearImportInfo(0 To UBound(tmClearImportInfo) + 1) As CLEASRIMPORTINFO
        
    End If
    'Check if CPTT should be removed.  If no lst exist, then remove CPTT
    'CPTT will be recreated if required in mCheckCPTT
    SQLQuery = "SELECT * FROM lst WHERE (lstLogVefCode = " & imVefCode & " AND (lstLogDate >= '" & Format$(slMonDate, sgSQLDateForm) & "')" & " AND (lstLogDate <= '" & Format$(slSunDate, sgSQLDateForm) & "')" & ")"
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If lst_rst.EOF Then
        SQLQuery = "DELETE FROM CPTT WHERE (cpttVefCode = " & imVefCode & " AND (cpttStartDate = '" & Format$(slMonDate, sgSQLDateForm) & "')" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "Import WWO Spots-mClearPrevImport"
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mClearPrevImport"
End Sub


Private Function mAddCPTT() As Long
    Dim tlCPTT As CPTT
    Dim ilRet As Integer

    On Error GoTo ErrHand
    tlCPTT.lCode = 0
    tlCPTT.lAtfCode = lmAttCode
    tlCPTT.iShfCode = imShttCode
    tlCPTT.iVefCode = imVefCode
    tlCPTT.sCreateDate = Format$(gNow(), sgShowDateForm)
    tlCPTT.sStartDate = Format$(smMonDate, sgShowDateForm)
    'tlCPTT.iCycle = 1
    tlCPTT.sReturnDate = Format$("1/1/1970", sgShowDateForm)
    'tlCPTT.sAirTime = Format$("12:00AM", sgShowTimeWSecForm)
    tlCPTT.iStatus = 0
    tlCPTT.iUsfCode = igUstCode
    'tlCPTT.iPrintStatus = 0
    tlCPTT.iNoSpotsGen = 0
    tlCPTT.iNoSpotsAired = 0
    tlCPTT.iPostingStatus = 0
    tlCPTT.sAstStatus = "N"
    tlCPTT.sUnused = ""
    
    
    SQLQuery = "Insert Into cptt ( "
    SQLQuery = SQLQuery & "cpttCode, "
    SQLQuery = SQLQuery & "cpttAtfCode, "
    SQLQuery = SQLQuery & "cpttShfCode, "
    SQLQuery = SQLQuery & "cpttVefCode, "
    SQLQuery = SQLQuery & "cpttCreateDate, "
    SQLQuery = SQLQuery & "cpttStartDate, "
    SQLQuery = SQLQuery & "cpttReturnDate, "
    SQLQuery = SQLQuery & "cpttStatus, "
    SQLQuery = SQLQuery & "cpttUsfCode, "
    SQLQuery = SQLQuery & "cpttNoSpotsGen, "
    SQLQuery = SQLQuery & "cpttNoSpotsAired, "
    SQLQuery = SQLQuery & "cpttPostingStatus, "
    SQLQuery = SQLQuery & "cpttAstStatus, "
    SQLQuery = SQLQuery & "cpttUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlCPTT.lCode & ", "
    SQLQuery = SQLQuery & tlCPTT.lAtfCode & ", "
    SQLQuery = SQLQuery & tlCPTT.iShfCode & ", "
    SQLQuery = SQLQuery & tlCPTT.iVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlCPTT.sCreateDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlCPTT.sStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlCPTT.sReturnDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlCPTT.iStatus & ", "
    SQLQuery = SQLQuery & tlCPTT.iUsfCode & ", "
    SQLQuery = SQLQuery & tlCPTT.iNoSpotsGen & ", "
    SQLQuery = SQLQuery & tlCPTT.iNoSpotsAired & ", "
    SQLQuery = SQLQuery & tlCPTT.iPostingStatus & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlCPTT.sAstStatus) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlCPTT.sUnused) & "' "
    SQLQuery = SQLQuery & ")"
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.Txt", "Import WWO Spots-mAddCPTT"
        mAddCPTT = -1
        Exit Function
    End If
    gFileChgdUpdate "cptt.mkd", True
    SQLQuery = "SELECT * FROM CPTT WHERE (cpttVefCode = " & imVefCode & " AND cpttAtfCode = " & lmAttCode & " AND (cpttStartDate = '" & Format$(smMonDate, sgSQLDateForm) & "')" & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mAddCPTT = -1
    Else
        mAddCPTT = rst!cpttCode
    End If
    
    On Error GoTo 0
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mAddCPTT"
    mAddCPTT = -1
End Function



Private Function mUpdateLLD() As Integer
    Dim slLLD As String

    On Error GoTo ErrHand
    
    SQLQuery = "SELECT vpfLLD"
    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & imVefCode & ")"
    
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If IsNull(rst!vpfLLD) Then
            slLLD = "1/1/1970"
        Else
            If Not gIsDate(rst!vpfLLD) Then
                slLLD = "1/1/1970"
            Else
                'set sLLD to last log date
                slLLD = Format$(rst!vpfLLD, sgShowDateForm)
            End If
        End If
    Else
        mUpdateLLD = False
        Exit Function
    End If
    If DateValue(smLogDate) > DateValue(slLLD) Then
        SQLQuery = "UPDATE VPF_Vehicle_Options SET "
        SQLQuery = SQLQuery & "vpfLLD = '" & Format$(smLogDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " WHERE vpfvefKCode =" & imVefCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "Import WWO Spots-mUpdateLLD"
            mUpdateLLD = False
            Exit Function
        End If
        '11/26/17
        gFileChgdUpdate "vpf.btr", True
    End If
    mUpdateLLD = True
    On Error GoTo 0
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mUpdateLLD"
    mUpdateLLD = False
End Function

Private Sub txtImportPath_Change()
    lacResult.Caption = ""
    cmdImport.Caption = "&Import"
    cmdCancel.Caption = "&Cancel"
    cmdImport.Enabled = True
End Sub

Private Sub mAddMsgToList(slMsg As String)
    'Add horizontal scroll if required and add message to list box
    'The control pbcArial is used to get the approximate width of the text as the list box does not has a TextWidth command
    Dim llValue As Long
    Dim llRg As Long
    Dim llMaxWidth
    Dim llRet As Long
    Dim llRow As Long
    
    llMaxWidth = (pbcArial.TextWidth(slMsg))
    If llMaxWidth > lmMaxWidth Then
        lmMaxWidth = llMaxWidth
    End If
    If lmMaxWidth > lbcMsg.Width Then
        'Divide by 15 to convert units and add 120 for little extra room
        'Scale Mode is in Twips
        llValue = lmMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcMsg.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llRow = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slMsg)
    If llRow < 0 Then
        lbcMsg.AddItem slMsg
        gLogMsg slMsg, "WWOTrafficSpots.Txt", False
    End If
End Sub


Private Function mAddCIF() As Long
    Dim tlCif As CIF
    Dim slSQLQuery As String
    Dim llCode As Long

    On Error GoTo ErrHand

    mAddCIF = -1
    tlCif.lCode = 0
    tlCif.imcfCode = 0
    tlCif.sName = ""
    tlCif.sCut = ""
    tlCif.sReel = ""
    tlCif.iLen = imSpotLen
    tlCif.ietfCode = 0
    tlCif.ienfCode = 0
    tlCif.iAdfCode = imAdfCode
    tlCif.lCpfCode = lmCpfCode
    tlCif.imnfComp1 = 0
    tlCif.imnfComp2 = 0
    tlCif.imnfAnn = 0
    tlCif.sHouse = "N"
    tlCif.sCleared = "N"
    tlCif.lCSFCode = 0
    tlCif.iTimes = 1
    tlCif.sCDisp = "N"
    tlCif.sTDisp = "N"
    tlCif.sPurged = "A"
    tlCif.sPurgeDate = Format$("12/31/2069", sgShowDateForm)
    tlCif.sEntryDate = Format$(Now, sgShowDateForm)
    If smLogDate <> "" Then
        tlCif.sUsedDate = Format$(smLogDate, sgShowDateForm)
    Else
        tlCif.sUsedDate = Format$(Now, sgShowDateForm)
    End If
    If smRotStartDate <> "" Then
        tlCif.sRotStartDate = Format$(smRotStartDate, sgShowDateForm)
    Else
        If smLogDate <> "" Then
            tlCif.sRotStartDate = Format$(smLogDate, sgShowDateForm)
        Else
            tlCif.sRotStartDate = Format$(Now, sgShowDateForm)
        End If
    End If
    If smRotEndDate <> "" Then
        tlCif.sRotEndDate = Format$(smRotEndDate, sgShowDateForm)
    Else
        If smLogDate <> "" Then
            tlCif.sRotEndDate = Format$(smLogDate, sgShowDateForm)
        Else
            tlCif.sRotEndDate = Format$(Now, sgShowDateForm)
        End If
    End If
    tlCif.iurfCode = 0
    tlCif.sPrint = "N"
    tlCif.iLangMnfCode = 0
    tlCif.sInvSentDate = Format$("1/1/1970", sgShowDateForm)
    tlCif.sUnused = ""
    
    
    slSQLQuery = "Insert Into CIF_Copy_Inventory ( "
    slSQLQuery = slSQLQuery & "cifCode, "
    slSQLQuery = slSQLQuery & "cifmcfCode, "
    slSQLQuery = slSQLQuery & "cifName, "
    slSQLQuery = slSQLQuery & "cifCut, "
    slSQLQuery = slSQLQuery & "cifReel, "
    slSQLQuery = slSQLQuery & "cifLen, "
    slSQLQuery = slSQLQuery & "cifetfCode, "
    slSQLQuery = slSQLQuery & "cifenfCode, "
    slSQLQuery = slSQLQuery & "cifadfCode, "
    slSQLQuery = slSQLQuery & "cifcpfCode, "
    slSQLQuery = slSQLQuery & "cifmnfComp1, "
    slSQLQuery = slSQLQuery & "cifmnfComp2, "
    slSQLQuery = slSQLQuery & "cifmnfAnn, "
    slSQLQuery = slSQLQuery & "cifHouse, "
    slSQLQuery = slSQLQuery & "cifCleared, "
    slSQLQuery = slSQLQuery & "cifcsfCode, "
    slSQLQuery = slSQLQuery & "cifTimes, "
    slSQLQuery = slSQLQuery & "cifCDisp, "
    slSQLQuery = slSQLQuery & "cifTDisp, "
    slSQLQuery = slSQLQuery & "cifPurged, "
    slSQLQuery = slSQLQuery & "cifPurgeDate, "
    slSQLQuery = slSQLQuery & "cifEntryDate, "
    slSQLQuery = slSQLQuery & "cifUsedDate, "
    slSQLQuery = slSQLQuery & "cifRotStartDate, "
    slSQLQuery = slSQLQuery & "cifRotEndDate, "
    slSQLQuery = slSQLQuery & "cifurfCode, "
    slSQLQuery = slSQLQuery & "cifPrint, "
    slSQLQuery = slSQLQuery & "cifLangMnfCode, "
    slSQLQuery = slSQLQuery & "cifInvSentDate, "
    slSQLQuery = slSQLQuery & "cifUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "
    slSQLQuery = slSQLQuery & tlCif.imcfCode & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sName) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sCut) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sReel) & "', "
    slSQLQuery = slSQLQuery & tlCif.iLen & ", "
    slSQLQuery = slSQLQuery & tlCif.ietfCode & ", "
    slSQLQuery = slSQLQuery & tlCif.ienfCode & ", "
    slSQLQuery = slSQLQuery & tlCif.iAdfCode & ", "
    slSQLQuery = slSQLQuery & tlCif.lCpfCode & ", "
    slSQLQuery = slSQLQuery & tlCif.imnfComp1 & ", "
    slSQLQuery = slSQLQuery & tlCif.imnfComp2 & ", "
    slSQLQuery = slSQLQuery & tlCif.imnfAnn & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sHouse) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sCleared) & "', "
    slSQLQuery = slSQLQuery & tlCif.lCSFCode & ", "
    slSQLQuery = slSQLQuery & tlCif.iTimes & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sCDisp) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sTDisp) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sPurged) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCif.sPurgeDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCif.sEntryDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCif.sUsedDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCif.sRotStartDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCif.sRotEndDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & tlCif.iurfCode & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sPrint) & "', "
    slSQLQuery = slSQLQuery & tlCif.iLangMnfCode & ", "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCif.sInvSentDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCif.sUnused) & "' "
    slSQLQuery = slSQLQuery & ") "
    
    llCode = gInsertAndReturnCode(slSQLQuery, "Cif_Copy_Inventory", "CifCode", "Replace")
    If llCode > 0 Then
        mAddCIF = llCode
    End If
    On Error GoTo 0
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mAddCIF"
    mAddCIF = -1
End Function

Private Function mAddCPF() As Long
    Dim tlCpf As CPF
    Dim slSQLQuery As String
    Dim llCode As Long

    On Error GoTo ErrHand

    mAddCPF = -1
    tlCpf.lCode = 0
    tlCpf.sName = smProduct
    tlCpf.sISCI = smISCI
    tlCpf.sCreative = smCreative
    If smRotEndDate <> "" Then
        tlCpf.sRotEndDate = Format$(smRotEndDate, sgShowDateForm)
    Else
        tlCpf.sRotEndDate = Format$(smLogDate, sgShowDateForm)
    End If
    tlCpf.lsifCode = 0
    
    
    slSQLQuery = "Insert Into CPF_Copy_Prodct_ISCI ( "
    slSQLQuery = slSQLQuery & "cpfCode, "
    slSQLQuery = slSQLQuery & "cpfName, "
    slSQLQuery = slSQLQuery & "cpfISCI, "
    slSQLQuery = slSQLQuery & "cpfCreative, "
    slSQLQuery = slSQLQuery & "cpfRotEndDate, "
    slSQLQuery = slSQLQuery & "cpfsifCode "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCpf.sName) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCpf.sISCI) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tlCpf.sCreative) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tlCpf.sRotEndDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & tlCpf.lsifCode
    slSQLQuery = slSQLQuery & ") "
    
    llCode = gInsertAndReturnCode(slSQLQuery, "CPF_Copy_Prodct_ISCI", "CpfCode", "Replace")
    If llCode > 0 Then
        mAddCPF = llCode
    End If
    On Error GoTo 0
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mAddCPF"
    mAddCPF = -1
End Function

Private Function mUpdateCPF() As Integer
    Dim slSQLQuery As String

    On Error GoTo ErrHand

    slSQLQuery = "Update CPF_Copy_Prodct_ISCI Set "
    slSQLQuery = slSQLQuery & "cpfName = '" & gFixQuote(smProduct) & "', "
    If smCreative <> "" Then
        slSQLQuery = slSQLQuery & "cpfCreative = '" & gFixQuote(smCreative) & "', "
    End If
    If smRotEndDate <> "" Then
        slSQLQuery = slSQLQuery & "cpfRotEndDate = '" & Format$(smRotEndDate, sgSQLDateForm) & "', "
    End If
    slSQLQuery = slSQLQuery & "cpfsifCode = " & 0
    slSQLQuery = slSQLQuery & " WHERE cpfCode = " & lmCpfCode
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.Txt", "Import WWO Spots-mUpdateCPF"
        mUpdateCPF = False
        Exit Function
    End If
    mUpdateCPF = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mUpdateCPF"
    mUpdateCPF = False
End Function

Private Function mUpdateCIF() As Integer
    Dim slSQLQuery As String

    On Error GoTo ErrHand

    slSQLQuery = "Update CIF_Copy_Inventory Set "
    If smLogDate <> "" Then
        slSQLQuery = slSQLQuery & "cifUsedDate = '" & Format$(smLogDate, sgSQLDateForm) & "', "
    End If
    If smRotStartDate <> "" Then
        slSQLQuery = slSQLQuery & "cifRotStartDate = '" & Format$(smRotStartDate, sgSQLDateForm) & "', "
    End If
    If smRotEndDate <> "" Then
        slSQLQuery = slSQLQuery & "cifRotEndDate = '" & Format$(smRotEndDate, sgSQLDateForm) & "', "
    End If
    slSQLQuery = slSQLQuery & "cifUnused = '" & gFixQuote("") & "' "


    slSQLQuery = slSQLQuery & " WHERE cifCode = " & lmCifCode
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.Txt", "Import WWO Spots-mUpdateCIF"
        mUpdateCIF = False
        Exit Function
    End If
    mUpdateCIF = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mUpdateCIF"
    mUpdateCIF = False
End Function


Private Function mGetSALink() As Integer
    Dim ilSellDay As Integer
    Dim slEndDate As String
    Dim ilLink As Integer
    Dim ilFound As Integer
    Dim ilVlf As Integer
    Dim ilBreak As Integer
    Dim llPrevAirTime As Long
    ReDim ilAirCode(0 To 0) As Integer
    ReDim tmSALinkInfo(0 To 0) As SALINKINFO
    
    Select Case Weekday(smLogDate)
        Case vbMonday
            ilSellDay = 0
        Case vbTuesday
            ilSellDay = 0
        Case vbWednesday
            ilSellDay = 0
        Case vbThursday
            ilSellDay = 0
        Case vbFriday
            ilSellDay = 0
        Case vbSaturday
            ilSellDay = 6
        Case vbSunday
            ilSellDay = 7
    End Select
    
    SQLQuery = "SELECT * FROM vlf_Vehicle_Linkages WHERE (vlfSellCode = " & imVefCode
    SQLQuery = SQLQuery & " AND " & "(vlfSellDay = " & ilSellDay & ")"
    SQLQuery = SQLQuery & " AND " & "(vlfEffDate >= '" & Format$(gAdjYear(smLogDate), sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " AND " & "(vlfStatus = 'C')" & ")"
    Set vlf_rst = gSQLSelectCall(SQLQuery)
    Do While Not vlf_rst.EOF
        If IsNull(vlf_rst!vlfTermDate) Then
            slEndDate = "12/31/2069"
        Else
            slEndDate = Format(vlf_rst!vlfTermDate, "m/d/yy")
        End If
        If gDateValue(smLogDate) <= gDateValue(slEndDate) Then
            ilFound = False
            For ilVlf = 0 To UBound(ilAirCode) - 1 Step 1
                If vlf_rst!vlfAirCode = ilAirCode(ilVlf) Then
                    ilFound = True
                    Exit For
                End If
            Next ilVlf
            If Not ilFound Then
                ilAirCode(UBound(ilAirCode)) = vlf_rst!vlfAirCode
                ReDim Preserve ilAirCode(0 To UBound(ilAirCode) + 1) As Integer
            End If
        End If
        vlf_rst.MoveNext
    Loop
    For ilVlf = 0 To UBound(ilAirCode) - 1 Step 1
        llPrevAirTime = -1
        ilBreak = 0
        SQLQuery = "SELECT * FROM vlf_Vehicle_Linkages WHERE (vlfAirCode = " & ilAirCode(ilVlf)
        SQLQuery = SQLQuery & " AND " & "(vlfAirDay = " & ilSellDay & ")"
        SQLQuery = SQLQuery & " AND " & "(vlfEffDate >= '" & Format$(gAdjYear(smLogDate), sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND " & "(vlfStatus = 'C')" & ")" & "ORDER BY vlfAirTime, vlfAirPosNo"
        Set vlf_rst = gSQLSelectCall(SQLQuery)
        Do While Not vlf_rst.EOF
            If IsNull(vlf_rst!vlfTermDate) Then
                slEndDate = "12/31/2069"
            Else
                slEndDate = Format(vlf_rst!vlfTermDate, "m/d/yy")
            End If
            If gDateValue(smLogDate) <= gDateValue(slEndDate) Then
                If llPrevAirTime <> gTimeToLong(Format(vlf_rst!vlfAirTime, sgShowTimeWSecForm), False) Then
                    ilBreak = ilBreak + 1
                    llPrevAirTime = gTimeToLong(Format(vlf_rst!vlfAirTime, sgShowTimeWSecForm), False)
                End If
                If vlf_rst!vlfSellCode = imVefCode Then
                    tmSALinkInfo(UBound(tmSALinkInfo)).iAirCode = vlf_rst!vlfAirCode
                    tmSALinkInfo(UBound(tmSALinkInfo)).iSellCode = vlf_rst!vlfSellCode
                    tmSALinkInfo(UBound(tmSALinkInfo)).lAirTime = gTimeToLong(Format(vlf_rst!vlfAirTime, sgShowTimeWSecForm), False)
                    tmSALinkInfo(UBound(tmSALinkInfo)).lSellTime = gTimeToLong(Format(vlf_rst!vlfSellTime, sgShowTimeWSecForm), False)
                    tmSALinkInfo(UBound(tmSALinkInfo)).iBreak = ilBreak
                    tmSALinkInfo(UBound(tmSALinkInfo)).iPosition = 5 * (vlf_rst!vlfAirSeq - 1) + 1
                    ReDim Preserve tmSALinkInfo(0 To UBound(tmSALinkInfo) + 1) As SALINKINFO
                End If
            End If
            vlf_rst.MoveNext
        Loop
    Next ilVlf
    mGetSALink = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mGetSALink"
    mGetSALink = False
End Function

Private Function mGetCpfCode(blTestOnly As Boolean) As Integer
    Dim ilRet As Integer
    Dim llCode As Long
    
    mGetCpfCode = False
    lmCpfCode = 0
    SQLQuery = "SELECT * FROM cpf_Copy_Prodct_ISCI WHERE (cpfISCI = '" & smISCI & "')"
    Set cpf_rst = gSQLSelectCall(SQLQuery)
    If Not cpf_rst.EOF Then
        lmCpfCode = cpf_rst!cpfCode
        If Not blTestOnly Then
            ilRet = mUpdateCPF()
        End If
        mGetCpfCode = True
    Else
        If Not blTestOnly Then
            llCode = mAddCPF()
            If llCode <> -1 Then
                lmCpfCode = llCode
                mGetCpfCode = True
            End If
        End If
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mGetCpfCode"
End Function

Private Function mGetCifCode() As Integer
    Dim ilRet As Integer
    Dim llCode As Long
    
    mGetCifCode = False
    lmCifCode = 0
    SQLQuery = "SELECT * FROM cif_Copy_Inventory WHERE (cifcpfCode = " & lmCpfCode & "AND cifPurged = 'A'" & ")"
    Set cif_rst = gSQLSelectCall(SQLQuery)
    If Not cif_rst.EOF Then
        lmCifCode = cif_rst!cifCode
        ilRet = mUpdateCIF()
        mGetCifCode = True
    Else
        llCode = mAddCIF()
        If llCode <> -1 Then
            lmCifCode = llCode
            mGetCifCode = True
        End If
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-"
End Function


Private Function mGetLSTSpots(ilVefCode As Integer, slLogDate As String) As Integer
    mGetLSTSpots = True
    If (imLSTVefCode = ilVefCode) And (gDateValue(slLogDate) = lmLstDate) Then
        Exit Function
    End If
    imLSTVefCode = ilVefCode
    lmLstDate = gDateValue(slLogDate)
    'Verify if Load required
    ReDim tmLstSpot(0 To 0) As LSTSPOT
    SQLQuery = "SELECT * FROM lst WHERE ((Mod(lstStatus, 100) < " & ASTEXTENDED_MG & ") AND (lstLogVefCode = " & imVefCode & ") AND (lstLogDate = '" & Format$(smLogDate, sgSQLDateForm) & "')" & ")"
    Set lst_rst = gSQLSelectCall(SQLQuery)
    Do While Not lst_rst.EOF
        tmLstSpot(UBound(tmLstSpot)).lLstCode = lst_rst!lstCode
        tmLstSpot(UBound(tmLstSpot)).iAdfCode = lst_rst!lastAdfCode
        tmLstSpot(UBound(tmLstSpot)).lAvailTime = gTimeToLong(Format(lst_rst!lstLogTime, sgShowTimeWSecForm), False)
        ReDim Preserve tmLstSpot(0 To UBound(tmLstSpot) + 1) As LSTSPOT
        lst_rst.MoveNext
    Loop
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWWOSpot-mFindLSTSpot"
    mGetLSTSpots = False
End Function

Private Function mFindLST(ilVefCode As Integer, slLogDate As String, ilPosition As Integer) As Long
    Dim ilLst As Integer
    Dim llTime As Long
    Dim ilRet As Integer
    
    ilRet = mGetLSTSpots(ilVefCode, slLogDate)

    llTime = gTimeToLong(smLogTime, False)
    For ilLst = 0 To UBound(tmLstSpot) - 1 Step 1
        If tmLstSpot(ilLst).iAdfCode = imAdfCode Then
            If tmLstSpot(ilLst).lAvailTime = llTime Then
                mFindLST = tmLstSpot(ilLst).lLstCode
                Exit Function
            ElseIf imPosition = 2 Then
                If tmLstSpot(ilLst).lAvailTime + 30 = llTime Then
                    mFindLST = tmLstSpot(ilLst).lLstCode
                    Exit Function
                End If
            End If
        End If
    Next ilLst
    mFindLST = -1
End Function
