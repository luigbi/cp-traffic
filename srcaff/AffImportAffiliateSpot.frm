VERSION 5.00
Begin VB.Form frmImportAffiliateSpot 
   Caption         =   "Export Scheduled Station Spots"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffImportAffiliateSpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9615
   Begin VB.ListBox lbcStation 
      Height          =   1425
      ItemData        =   "AffImportAffiliateSpot.frx":030A
      Left            =   4200
      List            =   "AffImportAffiliateSpot.frx":030C
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   1770
      Width           =   1695
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   990
      TabIndex        =   19
      Top             =   3735
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   4845
      TabIndex        =   18
      Top             =   3735
      Width           =   1065
   End
   Begin VB.Frame frmVeh 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   975
      Width           =   3780
      Begin VB.OptionButton rbcSpots 
         Caption         =   "All Spots"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   6
         Top             =   45
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton rbcSpots 
         Caption         =   "Spot Changes"
         Height          =   195
         Index           =   1
         Left            =   2025
         TabIndex        =   7
         Top             =   45
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Export"
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   45
         Width           =   660
      End
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   360
      Left            =   4605
      TabIndex        =   3
      Text            =   "7"
      Top             =   450
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   3255
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      Height          =   3375
      ItemData        =   "AffImportAffiliateSpot.frx":030E
      Left            =   6585
      List            =   "AffImportAffiliateSpot.frx":0310
      TabIndex        =   14
      Top             =   450
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   1425
      ItemData        =   "AffImportAffiliateSpot.frx":0312
      Left            =   120
      List            =   "AffImportAffiliateSpot.frx":0314
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   1770
      Width           =   3855
   End
   Begin VB.TextBox txtDate 
      Height          =   360
      Left            =   1530
      TabIndex        =   1
      Top             =   450
      Width           =   1320
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   3495
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   21
      Top             =   4305
      Width           =   1200
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   5820
      TabIndex        =   15
      Top             =   4290
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   16
      Top             =   4275
      Width           =   1575
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   5205
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   22
      Top             =   4245
      Width           =   1200
   End
   Begin VB.Label lacTitle3 
      Alignment       =   2  'Center
      Caption         =   "Stations"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1470
      Width           =   1740
   End
   Begin VB.Label lbcFile 
      Caption         =   "Export File"
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   3750
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Days"
      Height          =   255
      Left            =   3195
      TabIndex        =   2
      Top             =   495
      Width           =   1335
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   17
      Top             =   4230
      Width           =   5490
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   7035
      TabIndex        =   13
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Vehicles"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1455
      Width           =   3885
   End
   Begin VB.Label Label1 
      Caption         =   "Export Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   495
      Width           =   1395
   End
End
Attribute VB_Name = "frmImportAffiliateSpot"
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

Private smDate As String     'Export Date
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private cprst As rdoResultset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmAetInfo() As AETINFO
Private tmAet() As AETINFO






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
Private Function mOpenMsgFile(sMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgMsgDirectory & "ExptSchdSpots.Txt"
    slNowDate = Format$(Now, sgShowDateForm)
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(slFileDate) = DateValue(slNowDate) Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                MsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Export Scheduled Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
End Sub




Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hWnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub cmcBrowse_Click()
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtFile.Text = Trim$(CommonDialog1.FileName)
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdExport_Click()
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
    If lbcVehicles.ListIndex < 0 Then
        Exit Sub
    End If
    If txtDate.Text = "" Then
        MsgBox "Date must be specified.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If IsDate(txtDate.Text) = False Then
        Beep
        MsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtDate.SetFocus
        Exit Sub
    Else
        smDate = Format(txtDate.Text, sgShowDateForm)
    End If
    sMoDate = gObtainPrevMonday(smDate)
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        MsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    If Not mCheckSelection() Then
        If rbcSpots(0).Value Then
            rbcSpots(1).SetFocus
        Else
            rbcSpots(0).SetFocus
        End If
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    If Not mOpenMsgFile(sMsgFileName) Then
        cmdCancel.SetFocus
        Exit Sub
    End If
    imExporting = True
    iRet = 0
    On Error GoTo cmdExportErr:
    sToFile = txtFile.Text
    sDateTime = FileDateTime(sToFile)
    If iRet = 0 Then
        Kill sToFile
    End If
    On Error GoTo 0
    'iRet = 0
    'On Error GoTo cmdExportErr:
    'hmTo = FreeFile
    'Open sToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If iRet <> 0 Then
        Print #hmMsg, "** Terminated **"
        Close #hmMsg
        Close #hmTo
        imExporting = False
        Screen.MousePointer = vbDefault
        MsgBox "Open Error #" & Str$(Err.Numner) & sToFile, vbOKOnly, "Open Error"
        Exit Sub
    End If
    Print #hmMsg, "** Storing Output into " & sToFile & " **"
    On Error GoTo 0
    lacResult.Caption = ""
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            smVefName = Trim$(lbcVehicles.List(iLoop))
            iRet = mExportSpots()
            If (iRet = False) Then
                Print #hmMsg, "** Terminated **"
                Close #hmMsg
                Close #hmTo
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                Exit Sub
            End If
            If imTerminate Then
                Print #hmMsg, "** User Terminated **"
                Close #hmMsg
                Close #hmTo
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                Exit Sub
            End If
       End If
    Next iLoop
    Close #hmTo
    'Clear old aet records out
    On Error GoTo ErrHand:
    env.BeginTrans
    SQLQuery = "DELETE FROM Aet WHERE (aetFeedDate <= '" & Format$(DateAdd("d", -28, sMoDate), sgSQLDateForm) & "')"
    cnn.Execute SQLQuery, rdExecDirect
    env.CommitTrans
    imExporting = False
    Print #hmMsg, "** Completed Export Scheduled Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close #hmMsg
    lacResult.Caption = "Results: " & sMsgFileName
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdExportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then             'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Export Schd Spots-cmdExport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Export Schd Spots-cmdExport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    txtDate.Text = ""
    Unload frmExportSchdSpot
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    
    Screen.MousePointer = vbHourglass
    smDate = gObtainNextMonday(Format$(Now, sgShowDateForm))
    txtDate.Text = smDate
    imNumberDays = 7
    txtNumberDays.Text = Trim$(Str$(imNumberDays))
    imAllClick = False
    imTerminate = False
    imExporting = False
    
    mFillVehicle
    txtFile.Text = sgExportDirectory & "Spots.csv"
    chkAll.Value = vbChecked
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmAetInfo
    Set frmExportSchdSpot = Nothing
End Sub


Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    lbcStation.Clear
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
    For iLoop = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(iLoop) Then
            imVefCode = lbcVehicles.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    If iCount = 1 Then
        mFillStations
    End If
End Sub

Private Sub txtDate_Change()
    lbcMsg.Clear
End Sub

Private Sub txtDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Function mExportSpots()
    Dim sDate As String
    Dim iNoWeeks As Integer
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim sAdvtProd As String
    Dim sPledgeStartDate As String
    Dim sPledgeEndDate As String
    Dim iIndex As Integer
    ReDim iDays(0 To 6) As Integer
    Dim sPledgeStartTime As String
    Dim sPledgeEndTime As String
    Dim sLen As String
    Dim sCart As String
    Dim sISCI As String
    Dim sCreative As String
    Dim iDay As Integer
    Dim iAddDelete As Integer
    Dim iUpper As Integer
    Dim slStr As String
    Dim iAet As Integer
    Dim iFound As Integer
    Dim iExport As Integer  '0=Don't export as it did not change
                            '1=Export and create aet record
                            '2=Export and don't create aet reord (nothing changed but generating all spot export)
    
    On Error GoTo ErrHand
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays, smDate)
    
    Do
        'Get CPTT so that Stations requiring CP can be obtained
        SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, attPrintCP, attTimeType, attGenCP"
        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        Set cprst = cnn.OpenResultset(SQLQuery)
        While Not cprst.EOF
            ReDim tgCPPosting(0 To 1) As CPPOSTING
            tgCPPosting(0).lCpttCode = cprst!cpttCode
            tgCPPosting(0).iStatus = cprst!cpttStatus
            tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
            tgCPPosting(0).iAttCode = cprst!cpttatfCode
            tgCPPosting(0).iAttTimeType = cprst!attTimeType
            tgCPPosting(0).iVefCode = imVefCode
            tgCPPosting(0).iShttCode = cprst!shttcode
            tgCPPosting(0).sZone = cprst!shttTimeZone
            tgCPPosting(0).sDate = Format$(sMoDate, sgShowDateForm)
            Print #hmTo, "A," & """" & "Counterpoint Software" & """"       'Network Provider Name
            Print #hmTo, "B," & """" & "Marketron" & """"                   'Web Provider
            Print #hmTo, "C," & """" & "Marketron" & """"                   'Station Provider
            Print #hmTo, "D," & """" & "HBC" & """"                         'Station Provider
            Print #hmTo, "E," & """" & Trim$(smVefName) & """" 'Vehicle name
            Print #hmTo, "F," & """" & Trim$(cprst!shttCallLetters) & """"
            'Create AST records
            igTimes = 1 'By Week
            imAdfCode = -1
            iRet = gGetAstInfo(tmCPDat(), tmAstInfo(), imAdfCode, True, True)
            ReDim tmAet(0 To 0) As AETINFO
            ReDim tmAetInfo(0 To 0) As AETINFO
            'Obtain past image
            SQLQuery = "SELECT aetCode, aetSdfCode, aetFeedDate, aetFeedTime, aetPledgeStartDate, aetPledgeEndDate, aetPledgeStartTime, aetPledgeEndTime, aetAdvtProd, aetCart, aetISCI, aetCreative, aetAstCode, aetLen"
            SQLQuery = SQLQuery + " FROM aet"
            SQLQuery = SQLQuery + " WHERE (aetShfCode = " & cprst!shttcode
            SQLQuery = SQLQuery + " AND aetAtfCode = " & cprst!cpttatfCode
            SQLQuery = SQLQuery + " AND aetVefCode = " & imVefCode & ")"
            SQLQuery = SQLQuery + " AND (aetFeedDate BETWEEN '" & Format$(smDate, sgSQLDateForm) & "' AND '" & Format$(sEndDate, sgSQLDateForm) & "')"
            Set rst = cnn.OpenResultset(SQLQuery)
            While Not rst.EOF
                iUpper = UBound(tmAetInfo)
                tmAetInfo(iUpper).lCode = rst!aetCode
                tmAetInfo(iUpper).lSdfCode = rst!aetSdfCode
                tmAetInfo(iUpper).sFeedDate = Format$(rst!aetFeedDate, sgShowDateForm)
                If Second(rst!aetFeedTime) <> 0 Then
                    tmAetInfo(iUpper).sFeedTime = Format$(rst!aetFeedTime, sgShowTimeWSecForm)
                Else
                    tmAetInfo(iUpper).sFeedTime = Format$(rst!aetFeedTime, sgShowTimeWOSecForm)
                End If
                tmAetInfo(iUpper).sPledgeStartDate = Format$(rst!aetPledgeStartDate, sgShowDateForm)
                tmAetInfo(iUpper).sPledgeEndDate = Format$(rst!aetPledgeEndDate, sgShowDateForm)
                If Second(rst!aetPledgeStartTime) <> 0 Then
                    tmAetInfo(iUpper).sPledgeStartTime = Format$(rst!aetPledgeStartTime, sgShowTimeWSecForm)
                Else
                    tmAetInfo(iUpper).sPledgeStartTime = Format$(rst!aetPledgeStartTime, sgShowTimeWOSecForm)
                End If
                If Not IsNull(rst!aetPledgeEndTime) Then
                    If Second(rst!aetPledgeEndTime) <> 0 Then
                        tmAetInfo(iUpper).sPledgeEndTime = Format$(rst!aetPledgeEndTime, sgShowTimeWSecForm)
                    Else
                        tmAetInfo(iUpper).sPledgeEndTime = Format$(rst!aetPledgeEndTime, sgShowTimeWOSecForm)
                    End If
                Else
                    tmAetInfo(iUpper).sPledgeEndTime = ""
                End If
                tmAetInfo(iUpper).sAdvtProd = rst!aetAdvtProd
                tmAetInfo(iUpper).sCart = rst!aetCart
                tmAetInfo(iUpper).sISCI = rst!aetISCI
                tmAetInfo(iUpper).sCreative = rst!aetCreative
                tmAetInfo(iUpper).lAstCode = rst!aetAstCode
                tmAetInfo(iUpper).iLen = rst!aetLen
                tmAetInfo(iUpper).iProcessed = False
                ReDim Preserve tmAetInfo(0 To iUpper + 1) As AETINFO
                rst.MoveNext
            Wend
            'Output AST
            For iLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                If (DateValue(tmAstInfo(iLoop).sFeedDate) >= DateValue(smDate)) And (DateValue(tmAstInfo(iLoop).sFeedDate) <= DateValue(sEndDate)) And (tgStatusTypes(tmAstInfo(iLoop).iPledgeStatus).iPledged <> 2) Then
                    iAddDelete = 0
                    sAdvtProd = "Missing"
                    sCart = ""
                    sISCI = ""
                    sCreative = ""
                    SQLQuery = "SELECT lstProd, lstCart, lstISCI, adfName, cpfCreative"
                    SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
                    SQLQuery = SQLQuery + " WHERE lstCode =" & Str(tmAstInfo(iLoop).lLstCode)
                    Set rst = cnn.OpenResultset(SQLQuery)
                    If Not rst.EOF Then
                        If IsNull(rst!adfName) = True Then
                            sAdvtProd = "Missing" & "/" & Trim$(rst!lstProd)
                        Else
                            sAdvtProd = Trim$(rst!adfName) & "/" & Trim$(rst!lstProd)
                        End If
                        If IsNull(rst!lstCart) = True Then
                            sCart = ""
                        Else
                            sCart = Trim$(rst!lstCart)
                        End If
                        If IsNull(rst!lstISCI) = True Then
                            sISCI = ""
                        Else
                            sISCI = Trim$(rst!lstISCI)
                        End If
                        If IsNull(rst!cpfCreative) = True Then
                            sCreative = ""
                        Else
                            sCreative = Trim$(rst!cpfCreative)
                        End If
                    End If
                    sPledgeStartDate = Format$(tmAstInfo(iLoop).sPledgeDate, "m/d/yyyy")
                    If tgStatusTypes(tmAstInfo(iLoop).iPledgeStatus).iPledged = 0 Then
                        sPledgeEndDate = sPledgeStartDate
                    Else
                        gUnMapDays tmAstInfo(iLoop).sPdDays, iDays()
                        iDay = Weekday(sPledgeStartDate, vbMonday) - 1
                        sPledgeEndDate = sPledgeStartDate
                        For iIndex = iDay + 1 To 6 Step 1
                            If iDays(iIndex) Then
                                sPledgeEndDate = DateAdd("d", 1, sPledgeEndDate)
                            Else
                                Exit For
                            End If
                        Next iIndex
                    End If
                    If Second(tmAstInfo(iLoop).sPledgeStartTime) <> 0 Then
                        sPledgeStartTime = Format$(tmAstInfo(iLoop).sPledgeStartTime, "h:mm:ssa/p")
                    Else
                        sPledgeStartTime = Format$(tmAstInfo(iLoop).sPledgeStartTime, "h:mma/p")
                    End If
                    If Len(Trim$(tmAstInfo(iLoop).sPledgeEndTime)) <= 0 Then
                        sPledgeEndTime = sPledgeStartTime
                    Else
                        If Second(tmAstInfo(iLoop).sPledgeEndTime) <> 0 Then
                            sPledgeEndTime = Format$(tmAstInfo(iLoop).sPledgeEndTime, "h:mm:ssa/p")
                        Else
                            sPledgeEndTime = Format$(tmAstInfo(iLoop).sPledgeEndTime, "h:mma/p")
                        End If
                    End If
                    sLen = Trim$(Str$(tmAstInfo(iLoop).iLen))
                    iFound = False
                    iExport = 1
                    For iAet = 0 To UBound(tmAetInfo) - 1 Step 1
                        If tmAetInfo(iAet).lSdfCode = tmAstInfo(iLoop).lSdfCode Then
                            iFound = True
                            tmAetInfo(iAet).iProcessed = True
                            'Compare to see if different
                            iExport = 0
                            If DateValue(sPledgeStartDate) <> DateValue(tmAetInfo(iAet).sPledgeStartDate) Then
                                iExport = 1
                            End If
                            If DateValue(sPledgeEndDate) <> DateValue(tmAetInfo(iAet).sPledgeEndDate) Then
                                iExport = 1
                            End If
                            If TimeValue(sPledgeStartTime) <> TimeValue(tmAetInfo(iAet).sPledgeStartTime) Then
                                iExport = 1
                            End If
                            If TimeValue(sPledgeEndTime) <> TimeValue(tmAetInfo(iAet).sPledgeEndTime) Then
                                iExport = 1
                            End If
                            If StrComp(sAdvtProd, Trim$(tmAetInfo(iAet).sAdvtProd), vbBinaryCompare) <> 0 Then
                                iExport = 1
                            End If
                            If StrComp(sCart, Trim$(tmAetInfo(iAet).sCart), vbBinaryCompare) <> 0 Then
                                iExport = 1
                            End If
                            If StrComp(sISCI, Trim$(tmAetInfo(iAet).sISCI), vbBinaryCompare) <> 0 Then
                                iExport = 1
                            End If
                            If StrComp(sCreative, Trim$(tmAetInfo(iAet).sCreative), vbBinaryCompare) <> 0 Then
                                iExport = 1
                            End If
                            If tmAstInfo(iLoop).lCode <> tmAetInfo(iAet).lAstCode Then
                                iExport = 1
                            End If
                            If tmAstInfo(iLoop).iLen <> tmAetInfo(iAet).iLen Then
                                iExport = 1
                            End If
                            If rbcSpots(1).Value Then
                                If iExport = 1 Then
                                    'Print #hmTo, "S," & """" & sAdvtProd & """" & "," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & "," & """" & sCart & """" & "," & """" & sISCI & """" & "," & """" & sCreative & """" & "," & Trim$(Str$(tmAstInfo(iLoop).lCode))
                                    slStr = "S," & """" & Trim$(tmAetInfo(iAet).sAdvtProd) & """" & ",D," & Trim$(tmAetInfo(iAet).sPledgeStartDate) & "-" & Trim$(tmAetInfo(iAet).sPledgeEndDate) & "," & Trim$(tmAetInfo(iAet).sPledgeStartTime) & "," & Trim$(tmAetInfo(iAet).sPledgeEndTime) & "," & Trim$(Str$(tmAetInfo(iAet).iLen)) & ","
                                    If Trim$(tmAetInfo(iAet).sCart) <> "" Then
                                        slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCart) & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    If Trim$(tmAetInfo(iAet).sISCI) <> "" Then
                                        slStr = slStr & """" & tmAetInfo(iAet).sISCI & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    If Trim$(tmAetInfo(iAet).sCreative) <> "" Then
                                        slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCreative) & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    slStr = slStr & Trim$(Str$(tmAetInfo(iAet).lAstCode))
                                    Print #hmTo, slStr
                                Else
                                    tmAetInfo(iAet).lCode = 0
                                End If
                            ElseIf rbcSpots(0).Value Then
                                If iExport = 0 Then
                                    tmAetInfo(iAet).lCode = 0
                                    iExport = 2
                                End If
                            End If
                            Exit For
                        End If
                    Next iAet
                    If InStr(1, sAdvtProd, "Missing", vbTextCompare) = 1 Then
                        Print #hmMsg, Trim$(smVefName) & ": Advertiser Missing on " & Format$(tmAstInfo(iLoop).sAirDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sAirTime, "h:mm:ssAM/PM")
                        lbcMsg.AddItem Trim$(smVefName) & ": Advertiser Missing on " & Format$(tmAstInfo(iLoop).sAirDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sAirTime, "h:mm:ssAM/PM")
                    Else
                        If iExport <> 0 Then
                            'Print #hmTo, "S," & """" & sAdvtProd & """" & "," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & "," & """" & sCart & """" & "," & """" & sISCI & """" & "," & """" & sCreative & """" & "," & Trim$(Str$(tmAstInfo(iLoop).lCode))
                            slStr = "S," & """" & sAdvtProd & """" & ",A," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & ","
                            If sCart <> "" Then
                                slStr = slStr & """" & sCart & """" & ","
                            Else
                                slStr = slStr & ","
                            End If
                            If sISCI <> "" Then
                                slStr = slStr & """" & sISCI & """" & ","
                            Else
                                slStr = slStr & ","
                            End If
                            If sCreative <> "" Then
                                slStr = slStr & """" & sCreative & """" & ","
                            Else
                                slStr = slStr & ","
                            End If
                            slStr = slStr & Trim$(Str$(tmAstInfo(iLoop).lCode))
                            Print #hmTo, slStr
                            If iExport = 1 Then
                                iUpper = UBound(tmAet)
                                tmAet(iUpper).lCode = 0
                                tmAet(iUpper).iAtfCode = tmAstInfo(iLoop).iAttCode
                                tmAet(iUpper).iShfCode = tmAstInfo(iLoop).iShttCode
                                tmAet(iUpper).iVefCode = tmAstInfo(iLoop).iVefCode
                                tmAet(iUpper).lSdfCode = tmAstInfo(iLoop).lSdfCode
                                tmAet(iUpper).sFeedDate = tmAstInfo(iLoop).sFeedDate
                                tmAet(iUpper).sFeedTime = tmAstInfo(iLoop).sFeedTime
                                tmAet(iUpper).sPledgeStartDate = sPledgeStartDate
                                tmAet(iUpper).sPledgeEndDate = sPledgeEndDate
                                tmAet(iUpper).sPledgeStartTime = sPledgeStartTime
                                tmAet(iUpper).sPledgeEndTime = sPledgeEndTime
                                tmAet(iUpper).sAdvtProd = sAdvtProd
                                tmAet(iUpper).sCart = sCart
                                tmAet(iUpper).sISCI = sISCI
                                tmAet(iUpper).sCreative = sCreative
                                tmAet(iUpper).lAstCode = tmAstInfo(iLoop).lCode
                                tmAet(iUpper).iLen = Val(sLen)
                                ReDim Preserve tmAet(0 To iUpper + 1) As AETINFO
                            End If
                        End If
                    End If
                End If
            Next iLoop
            If rbcSpots(1).Value Then
                For iAet = 0 To UBound(tmAetInfo) - 1 Step 1
                    If tmAetInfo(iAet).iProcessed = False Then
                        tmAetInfo(iAet).iProcessed = True
                        'Print #hmTo, "S," & """" & sAdvtProd & """" & "," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & "," & """" & sCart & """" & "," & """" & sISCI & """" & "," & """" & sCreative & """" & "," & Trim$(Str$(tmAstInfo(iLoop).lCode))
                        slStr = "S," & """" & Trim$(tmAetInfo(iAet).sAdvtProd) & """" & ",D," & Trim$(tmAetInfo(iAet).sPledgeStartDate) & "," & Trim$(tmAetInfo(iAet).sPledgeEndDate) & "," & Trim$(tmAetInfo(iAet).sPledgeStartTime) & "," & Trim$(tmAetInfo(iAet).sPledgeEndTime) & "," & Trim$(Str$(tmAetInfo(iAet).iLen)) & ","
                        If Trim$(tmAetInfo(iAet).sCart) <> "" Then
                            slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCart) & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        If tmAetInfo(iAet).sISCI <> "" Then
                            slStr = slStr & """" & tmAetInfo(iAet).sISCI & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        If Trim$(tmAetInfo(iAet).sCreative) <> "" Then
                            slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCreative) & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        slStr = slStr & Trim$(Str$(tmAetInfo(iAet).lAstCode))
                        Print #hmTo, slStr
                    End If
                Next iAet
            End If
            'Remove previously created aet
            For iAet = 0 To UBound(tmAetInfo) - 1 Step 1
                If tmAetInfo(iAet).lCode > 0 Then
                    env.BeginTrans
                    SQLQuery = "DELETE FROM Aet WHERE (aetCode = " & tmAetInfo(iAet).lCode & ")"
                    cnn.Execute SQLQuery, rdExecDirect
                    env.CommitTrans
                End If
            Next iAet
            'Add ast as aet
            For iAet = 0 To UBound(tmAet) - 1 Step 1
                SQLQuery = "INSERT INTO aet"
                SQLQuery = SQLQuery + "(aetAtfCode, aetShfCode, aetVefCode, "
                SQLQuery = SQLQuery + "aetSdfCode, aetFeedDate, aetFeedTime, "
                SQLQuery = SQLQuery + " aetPledgeStartDate, aetPledgeEndDate, "
                SQLQuery = SQLQuery + "aetPledgeStartTime, aetPledgeEndTime, aetAdvtProd,"
                SQLQuery = SQLQuery + "aetCart, aetISCI, aetCreative,"
                SQLQuery = SQLQuery + "aetAstCode, aetLen)"
                SQLQuery = SQLQuery + " VALUES "
                SQLQuery = SQLQuery + "(" & tmAet(iAet).iAtfCode & ", " & tmAet(iAet).iShfCode & ", "
                SQLQuery = SQLQuery & tmAet(iAet).iVefCode & ", " & tmAet(iAet).lSdfCode & ", "
                SQLQuery = SQLQuery + "'" & Format$(tmAet(iAet).sFeedDate, sgSQLDateForm) & "', '" & Format$(tmAet(iAet).sFeedTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(tmAet(iAet).sPledgeStartDate, sgSQLDateForm) & "', '" & Format$(tmAet(iAet).sPledgeEndDate, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "'" & Format$(tmAet(iAet).sPledgeStartTime, sgSQLTimeForm) & "', '" & Format$(tmAet(iAet).sPledgeEndTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(tmAet(iAet).sAdvtProd) & "', '" & Trim$(tmAet(iAet).sCart) & "', '" & Trim$(tmAet(iAet).sISCI) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(tmAet(iAet).sCreative) & "', " & tmAet(iAet).lAstCode & ", " & tmAet(iAet).iLen & ")"
                env.BeginTrans
                cnn.Execute SQLQuery, rdExecDirect
                env.CommitTrans
                SQLQuery = "Select MAX(aetCode) from aet"
                Set rst = cnn.OpenResultset(SQLQuery)
                tmAet(iAet).lCode = rst(0).Value
            Next iAet
            cprst.MoveNext
        Wend
        sMoDate = DateAdd("d", 7, sMoDate)
    Loop While DateValue(sMoDate) < DateValue(sEndDate)

    mExportSpots = True
    Exit Function
mExportSpotsErr:
    iRet = Err
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then             'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Export Scheduled Spots-mExportSpots: "
            Print #hmMsg, gMsg & gErrSQL.Description & " Error #" & gErrSQL.Number
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Export Scheduled Spots-mExportSpots: "
        Print #hmMsg, gMsg & Err.Description & " Error #" & Err.Number
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mExportSpots = False
    Exit Function
    
End Function

Private Sub mFillStations()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery + " FROM shtt, att"
    SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery + " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
    Set rst = cnn.OpenResultset(SQLQuery)
    While Not rst.EOF
        lbcStation.AddItem Trim$(rst!shttCallLetters)
        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttcode
        rst.MoveNext
    Wend
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then             'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in mFillStation-lbcStation: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mFillStation-lbcStation: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub

Private Function mCheckLastExportDate() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim slDate As String
    Dim ilLoop As Integer
    'Dim slFields(1 To 15) As String
    Dim slFields(0 To 15) As String
    
    slFromFile = txtFile.Text
    ilRet = 0
    'On Error GoTo mCheckLastExportDateErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        If rbcSpots(0).Value Then
            mCheckLastExportDate = True
        Else
        End If
        Exit Function
    End If
    slDate = ""
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mCheckLastExportDateErr:
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
                'If slFields(1) = "S" Then
                If slFields(0) = "S" Then
                    On Error GoTo ErrHand
                    If slDate = "" Then
                        'slDate = slFields(5)
                        slDate = slFields(4)
                    Else
                        'If DateValue(slFields(5)) > DateValue(slDate) Then
                        If DateValue(slFields(4)) > DateValue(slDate) Then
                            'slDate = slFields(5)
                            slDate = slFields(4)
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If slDate = "" Then
    Else
    End If
    Exit Function
mCheckLastExportDateErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    slMsg = ""
    'Close hmFrom
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            slMsg = "A SQL error has occured in Export Scheduled Spots-mCheckLastExportDate: "
            MsgBox slMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    env.RollbackTrans
    If (Err.Number <> 0) And (slMsg = "") Then
        slMsg = "A general error has occured in Export Scheduled Spots-mCheckLastExportDate: "
        MsgBox slMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    ilRet = 1
    Resume Next

End Function

Private Function mCheckSelection() As Integer

    Dim ilRet As Integer
    Dim slMsg As String
    
    SQLQuery = "SELECT DISTINCT aetFeedDate"
    SQLQuery = SQLQuery + " FROM aet"
    SQLQuery = SQLQuery + " WHERE (aetFeedDate BETWEEN '" & Format$(smDate, sgSQLDateForm) & "' AND '" & Format$(DateAdd("d", imNumberDays, smDate), sgSQLDateForm) & "')"
    Set rst = cnn.OpenResultset(SQLQuery)
    If rst.EOF Then
        If rbcSpots(1).Value Then
            MsgBox "You must select 'all spots' for the specified dates before selecting 'spot changes'", vbOKOnly
            mCheckSelection = False
            Exit Function
        End If
    Else
        If rbcSpots(0).Value Then
            ilRet = MsgBox("Warning: You have already generated 'all spots' for this period.  do not proceed if stations have already received 'all spots' export" & Chr$(13) & Chr$(10) & "Continue with 'All Spot' Export?", vbYesNo)
            If ilRet = vbNo Then
                mCheckSelection = False
                Exit Function
            End If
        End If
    End If
    mCheckSelection = True
    Exit Function
mCheckSelectionErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    slMsg = ""
    'Close hmFrom
    For Each gErrSQL In rdoErrors
        If gErrSQL.Number <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            slMsg = "A SQL error has occured in Export Scheduled Spots-mCheckSelection: "
            MsgBox slMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (slMsg = "") Then
        slMsg = "A general error has occured in Export Scheduled Spots-mCheckSelection: "
        MsgBox slMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    ilRet = 1
    Resume Next

End Function
