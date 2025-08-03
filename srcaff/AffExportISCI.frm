VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmExportISCI 
   Caption         =   "Export ISCI"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportISCI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9615
   Begin VB.TextBox lacStationMsg 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   2010
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   $"AffExportISCI.frx":08CA
      Top             =   2715
      Width           =   1695
   End
   Begin VB.Timer tmcFilterDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9480
      Top             =   2520
   End
   Begin VB.Frame frFilter 
      Caption         =   "Filters"
      Height          =   1290
      Left            =   6555
      TabIndex        =   19
      Top             =   165
      Width           =   2745
      Begin VB.ListBox lbcFilter 
         Height          =   1230
         Left            =   1320
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "Station"
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "State"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "MSA"
         Height          =   200
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "Format"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton rbcFilter 
         Caption         =   "DMA"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9420
      Top             =   945
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1275
      TabIndex        =   1
      Top             =   165
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "2/14/24"
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   3
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4230
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   2415
      Width           =   1635
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   2415
      Width           =   3825
   End
   Begin VB.ListBox lbcPledgeDateTime 
      Height          =   255
      ItemData        =   "AffExportISCI.frx":0906
      Left            =   9195
      List            =   "AffExportISCI.frx":0908
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   510
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.ListBox lbcSort 
      Height          =   255
      ItemData        =   "AffExportISCI.frx":090A
      Left            =   8985
      List            =   "AffExportISCI.frx":090C
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   150
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All Stations"
      Height          =   195
      Left            =   4215
      TabIndex        =   10
      Top             =   4830
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.ListBox lbcStation 
      Height          =   2010
      ItemData        =   "AffExportISCI.frx":090E
      Left            =   4200
      List            =   "AffExportISCI.frx":0910
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2715
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   4605
      TabIndex        =   3
      Text            =   "1"
      Top             =   165
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4830
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2400
      ItemData        =   "AffExportISCI.frx":0912
      Left            =   6555
      List            =   "AffExportISCI.frx":0914
      TabIndex        =   14
      Top             =   2355
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2010
      ItemData        =   "AffExportISCI.frx":0916
      Left            =   120
      List            =   "AffExportISCI.frx":0918
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2715
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9390
      Top             =   3810
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5700
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5820
      TabIndex        =   11
      Top             =   5130
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   12
      Top             =   5130
      Width           =   1575
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   5850
      TabIndex        =   17
      Top             =   4845
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   135
      TabIndex        =   4
      Top             =   795
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label lacNoDays 
      Caption         =   "Number of Days"
      Height          =   255
      Left            =   3195
      TabIndex        =   2
      Top             =   210
      Width           =   1335
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   135
      TabIndex        =   15
      Top             =   5070
      Width           =   5490
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   7035
      TabIndex        =   13
      Top             =   1905
      Width           =   1965
   End
   Begin VB.Label lacExportDate 
      Caption         =   "Export Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1395
   End
End
Attribute VB_Name = "frmExportISCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmExport - Export ISCI
'*
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'
'The Compress logic might need to be rewritten before it is used.
'The tmISCISendInfo must be generated across all stations since output is by commercial Provider
'It would require different logic to make it faster
'

Private smDate As String     'Export Date
Private smNowDate As String
Private imNumberDays As Integer
Private imShttCode As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private smCommProvArfCode As String
Private imCommProvArfCode As Integer
Private imEmbedded As Integer
Private imEmbeddedAllowed As Integer
Private smProducerName As String
Private imProducerArfCode As Integer
Private smProducerArfCode As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private imNoDaysResendISCI As Integer
Private imNoAirDays As Integer
Private lmFeedDateLastA As Long
Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private hmAst As Integer
Private hmISCIbyBreak As Integer
Private cprst As ADODB.Recordset
Private chfrst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmISCIXRef() As ISCIXREF
Private tmISCISendInfo() As ISCISENDINFO
Private tmISCISendInfo_E() As ISCISENDINFO
Private lmXRefUpper As Long
Private lmSendInfoUpper As Long
Private lmSendInfoUpper_E As Long
Private tmAstInfo() As ASTINFO
Dim imFinalAlertVefCode() As Integer 'Vehicles that should be set to 'New'
Dim lmFound As Long
Dim lmNotFound As Long
Dim imPrevVefCode() As Integer
Dim imPrintB As Integer
Dim imPrintC As Integer
Dim smFileNamesCreated() As String
Dim smExportToPath As String
Private lmTotalNumber As Long
Private lmProcessedNumber As Long
Private lmPercent As Long
Private smSvNumberDays As String
Private lmEqtCode As Long
'11/16/17
Private DuplInfo_rst As ADODB.Recordset
Private smFilterType As String
Private imBypassAll As Integer
'10927
Dim myZoneAndDSTHelper As cDST


Private Function mExportISCI_AllFormat(llCountISCISendInfo As Long, tlISCISendInfo() As ISCISENDINFO) As Integer
    Dim iNewFile As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim sNowDate As String
    Dim sName As String
    Dim sToFile As String
    Dim sSDate As String
    Dim llSDate As Long
    Dim sEDate As String
    Dim llEDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim iPrintISCI As Integer
    Dim llISCIInfo As Long
    Dim sDateTime As String
    Dim sISCI As String
    Dim sCallLetters As String
    Dim lTestLLDSent As Long
    Dim ilFound As Integer
    Dim slNewRevise As String
    Dim ilLoop As Integer
    Dim slContactEMail As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim llRTime As Long
    Dim ilPrevVefCode As Integer
    Dim ilAppendToFile As Integer
    
    sNowDate = Format$(smNowDate, "mmddyy")
    sSDate = Format$(smDate, "mm/dd/yy")
    llSDate = DateValue(gAdjYear(sSDate))
    sEDate = Format$(DateAdd("d", imNumberDays - 1, smDate), "mm/dd/yy")
    llEDate = DateValue(gAdjYear(sEDate))
    iNewFile = True
    imPrintB = True
    imPrintC = True
    '5/8/14: Moved to outside of vehicle loop
    'ReDim imFinalAlertVefCode(0 To 0) As Integer
    llISCIInfo = LBound(tlISCISendInfo)
    tlISCISendInfo(llCountISCISendInfo).iCommProvArfCode = -1
    'Do While llISCIInfo < UBound(tlISCISendInfo)
    Do While llISCIInfo < llCountISCISendInfo
        If igExportSource = 2 Then DoEvents
        imPrintB = True
        imPrintC = True
        'Create file
        On Error GoTo ErrHand:
        SQLQuery = "SELECT arfID"
        SQLQuery = SQLQuery & " FROM ARF_Addresses"
        SQLQuery = SQLQuery + " WHERE arfCode = " & tlISCISendInfo(llISCIInfo).iCommProvArfCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            sName = Trim$(rst!arfID)
        Else
            sName = "Missed"
        End If
        ilAppendToFile = False
        ilPrevVefCode = -1
        For ilLoop = 0 To UBound(smFileNamesCreated) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            If StrComp(sName, Trim$(smFileNamesCreated(ilLoop)), vbTextCompare) = 0 Then
                ilAppendToFile = True
                ilPrevVefCode = imPrevVefCode(ilLoop)
                Exit For
            End If
        Next ilLoop
        If ilAppendToFile Then
            If tlISCISendInfo(llISCIInfo).iVefCode = ilPrevVefCode Then
                imPrintB = False
            Else
                For ilLoop = 0 To UBound(smFileNamesCreated) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If StrComp(sName, Trim$(smFileNamesCreated(ilLoop)), vbTextCompare) = 0 Then
                        imPrevVefCode(ilLoop) = tlISCISendInfo(llISCIInfo).iVefCode
                        Exit For
                    End If
                Next ilLoop
            End If
        Else
            smFileNamesCreated(UBound(smFileNamesCreated)) = sName
            ReDim Preserve smFileNamesCreated(0 To UBound(smFileNamesCreated) + 1) As String
            imPrevVefCode(UBound(imPrevVefCode)) = tlISCISendInfo(llISCIInfo).iVefCode
            ReDim Preserve imPrevVefCode(0 To UBound(imPrevVefCode) + 1) As Integer
        End If
        
        If tlISCISendInfo(llISCIInfo).iEmbedded Then
            sToFile = smExportToPath & sName & "_E_" & sNowDate & ".txt"
        Else
            sToFile = smExportToPath & sName & "_" & sNowDate & ".txt"
        End If
        'On Error GoTo mExportISCI_AllFormatErr:
        iRet = 0
        'sDateTime = FileDateTime(sToFile)
        iRet = gFileExist(sToFile)
        If (iRet = 0) And (Not ilAppendToFile) Then
            Kill sToFile
            If iRet <> 0 Then
                Close hmTo
                gMsgBox "Kill File " & sToFile & " error#" & Str$(Err.Number), vbOKOnly
                mExportISCI_AllFormat = False
                Exit Function
            End If
        ElseIf (iRet = 1) And (ilAppendToFile) Then
            ilAppendToFile = False
            imPrintB = True
        End If
        'iRet = 0
        'hmTo = FreeFile
        If ilAppendToFile Then
            'Open sToFile For Append As hmTo
            iRet = gFileOpen(sToFile, "Append", hmTo)
        Else
            'Open sToFile For Output As hmTo
            iRet = gFileOpen(sToFile, "Output", hmTo)
        End If
        If iRet <> 0 Then
            Close hmTo
            gMsgBox "Open File " & sToFile & " error#" & Str$(Err.Number), vbOKOnly
            mExportISCI_AllFormat = False
            Exit Function
        End If
        If Not ilAppendToFile Then
            Print #hmMsg, "** Storing Output into " & sToFile & " **"
            Print #hmTo, "A" & sSDate & "-" & sEDate
        End If
        lacResult.Caption = sName
        llSTime = timeGetTime
        llRTime = llSTime
        Do While (llISCIInfo < llCountISCISendInfo)
            If igExportSource = 2 Then DoEvents
            iPrintISCI = True
            If imPrintB Then
                ilFound = False
                For ilLoop = 0 To UBound(imFinalAlertVefCode) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If tlISCISendInfo(llISCIInfo).iVefCode = imFinalAlertVefCode(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    ilFound = gAlertFound("F", "I", tlISCISendInfo(llISCIInfo).iVefCode, sSDate)
                    If ilFound Then
                        slNewRevise = "New"
                        imFinalAlertVefCode(UBound(imFinalAlertVefCode)) = tlISCISendInfo(llISCIInfo).iVefCode
                        ReDim Preserve imFinalAlertVefCode(0 To UBound(imFinalAlertVefCode) + 1) As Integer
                    Else
                        slNewRevise = "Revised"
                    End If
                Else
                    slNewRevise = "New"
                End If
                Print #hmTo, "B" & tlISCISendInfo(llISCIInfo).sVehName & Chr$(9) & slNewRevise
                imPrintB = False
            End If
            If Not tlISCISendInfo(llISCIInfo).iEmbedded Then
                sCallLetters = tlISCISendInfo(llISCIInfo).sCallLetters
                If InStr(1, sCallLetters, "!", vbTextCompare) = 1 Then
                    sCallLetters = Mid(sCallLetters, 2)
                End If
                sCallLetters = Trim$(sCallLetters)
                If imPrintC Then
                    Print #hmTo, "C" & sCallLetters
                    imPrintC = False
                    If udcCriteria.IIncludeCommands(1) = vbChecked Then
                        '4/28/10:  Until E-Mail address added to agreement, use E-Mail address from Station Affiliate Contact
                        'If tlISCISendInfo(llISCIInfo).iACExistWithAtt = False Then
                            'SQLQuery = "SELECT arttEmail FROM artt WHERE (arttShttCode = '" & tlISCISendInfo(llISCIInfo).iShfCode & "' And arttAffContact = '1')"
                            'Set rst = gSQLSelectCall(SQLQuery)
                            'If Not rst.EOF Then
                            '    slContactEMail = Trim$(rst!arttEmail)
                            '    If slContactEMail <> "" Then
                            '        Print #hmTo, "F" & slContactEMail
                            '    End If
                            'End If
                            If igExportSource = 2 Then DoEvents
                            SQLQuery = "SELECT arttEmail FROM artt WHERE (arttShttCode = '" & tlISCISendInfo(llISCIInfo).iShfCode & "' And arttISCI2Contact = '1')"
                            Set rst = gSQLSelectCall(SQLQuery)
                            Do While Not rst.EOF
                                If igExportSource = 2 Then DoEvents
                                slContactEMail = Trim$(rst!arttEmail)
                                If slContactEMail <> "" Then
                                    Print #hmTo, "F" & slContactEMail
                                End If
                                rst.MoveNext
                            Loop
                            
                        'End If
                    End If
                End If
            End If
            tlISCISendInfo(llISCIInfo).iUpdateDateSent = False   'Can't update dates because of Duplicates
            sISCI = Trim$(gRemoveChar(tlISCISendInfo(llISCIInfo).sISCI, Chr$(9)))
            Print #hmTo, "D" & sISCI & Chr$(9) & mReplaceTab(Trim$(tlISCISendInfo(llISCIInfo).sAdvtName))
            llISCIInfo = llISCIInfo + 1
            If llISCIInfo >= llCountISCISendInfo Then
                Exit Do
            End If
            If (tlISCISendInfo(llISCIInfo).iCommProvArfCode <> tlISCISendInfo(llISCIInfo - 1).iCommProvArfCode) Then
                Exit Do
            End If
            If (tlISCISendInfo(llISCIInfo).iCommProvArfCode = tlISCISendInfo(llISCIInfo - 1).iCommProvArfCode) Then
                If (tlISCISendInfo(llISCIInfo - 1).iShfCode = 0) And (tlISCISendInfo(llISCIInfo).iShfCode <> 0) Then
                    Exit Do
                End If
                If (tlISCISendInfo(llISCIInfo - 1).iShfCode <> 0) And (tlISCISendInfo(llISCIInfo).iShfCode = 0) Then
                    Exit Do
                End If
            End If
            If tlISCISendInfo(llISCIInfo - 1).iVefCode <> tlISCISendInfo(llISCIInfo).iVefCode Then
                imPrintB = True
                imPrintC = True
            End If
            If tlISCISendInfo(llISCIInfo - 1).iShfCode <> tlISCISendInfo(llISCIInfo).iShfCode Then
                imPrintC = True
            End If
            llETime = timeGetTime
            If llETime - llRTime > 10000 Then
                llRTime = llETime
                lacResult.Caption = sName & " " & (llETime - llSTime) / 1000 & " sec elapsed"
            End If
        Loop
        Close #hmTo
    Loop
    On Error Resume Next
    Close #hmTo
'    'Clear the Export Flags
'    For ilLoop = 0 To UBound(imFinalAlertVefCode) - 1 Step 1
'        For llDate = llSDate To llEDate Step 7
'            slDate = Format$(llDate, "m/d/yy")
'            iRet = gAlertClear("A", "F", "I", imFinalAlertVefCode(ilLoop), slDate)
'        Next llDate
'    Next ilLoop
'    For ilLoop = 0 To lbcVehicles.ListCount - 1
'        If lbcVehicles.Selected(ilLoop) Then
'            imVefCode = lbcVehicles.ItemData(ilLoop)
'            For llDate = llSDate To llEDate Step 7
'                slDate = Format$(llDate, "m/d/yy")
'                iRet = gAlertClear("A", "R", "I", imVefCode, slDate)
'            Next llDate
'        End If
'    Next ilLoop
    lacResult.Caption = ""
    mExportISCI_AllFormat = True
    Exit Function
'mExportISCI_AllFormatErr:
'    iRet = 1
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mExportISCI_AllFormats"
    mExportISCI_AllFormat = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mGatherISCI_AllFormat           *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Gather ISCI by Provider, vehicle*
'*                     station.  Create one table of   *
'*                     ISCI code to export             *
'*                                                     *
'*******************************************************
Private Function mGatherISCI_AllFormat() As Integer
    Dim sDate As String
    Dim iNoWeeks As Integer
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim llAdf As Long
    Dim sAdvt As String
    Dim llIndex As Long
    Dim sCallLetters As String
    Dim sTimeZone As String 'TTP 10926
    Dim iShfCode As Integer
    Dim sISCI As String
    'Dim llUpper As Long
    Dim slStr As String
    Dim iFound As Integer
    Dim lFeedDate As Long
    Dim llPrevFeedDate As Long
    Dim ilOkStation As Integer
    Dim iExport As Integer  '0=Don't export as it did not change
                            '1=Export and create aet record
                            '2=Export and don't create aet reord (nothing changed but generating all spot export)
    Dim sRCart As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim sRProd As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim ilRet As Integer
    Dim ilHourNo As Integer
    Dim ilSegmentNo As Integer
    Dim ilPositionNo As Integer
    Dim ilBreakNo As Integer
    Dim llFeedTime As Long
    Dim llPrevFeedTime As Long
    Dim slHour As String
    Dim slSegment As String
    Dim slBreak As String
    Dim slPosition As String
    Dim slYear As String
    Dim slWeekNo As String
    Dim llStartYearDate As Long
    Dim ilVff As Integer
    Dim slProgCode As String
    Dim slACName As String
    Dim slContactEMail As String
    Dim ilDuplicateSpot As Integer
    Dim llDuplTestTime As Long
    Dim ilDuptTest As Integer
    Dim slDay As String '1=Mon; 2=Tue;....7=Sun
    Dim ilFirstDate As Integer
    Dim ilGenB As Integer
    Dim ilStartIndexOfBreak As Integer
    Dim ilPromoOnly As Integer
    Dim ilTest As Integer
    Dim ilNoAirDays As Integer
    Dim llSetKey As Long
    Dim slKey As String
    Dim ilShtt As Integer
    Dim ilAnf As Integer
    Dim blSpotOk As Boolean
    Dim llSmDate As Long
    Dim llEndDate As Long
    '10927
    Dim slTempString As String
    Dim slMilitaryHours As String
    On Error GoTo ErrHand
    
    'lgSTime3 = timeGetTime
    
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    ilFirstDate = True
    ilGenB = True
    '10927 by break and with time zone?  We need time zone stuff
    If udcCriteria.iExportType(1) And udcCriteria.IIncludeCommands(3) = vbChecked Then
        Set myZoneAndDSTHelper = New cDST
        'I don't test success, as it will only fail if can't make sql call...then we have more problems than this!
        myZoneAndDSTHelper.StartGetSite isci_export
    'testing only
            'myZoneAndDSTHelper.TestIsDaylight
            'myZoneAndDSTHelper.TestZoneByMethod
    ' end testing
        myZoneAndDSTHelper.isDSTActive sMoDate, sEndDate
    End If
    'D.S. 11/21/05
'    iRet = gGetMaxAstCode()
'    If Not iRet Then
'        Exit Function
'    End If
    
    Do
        If igExportSource = 2 Then DoEvents
        llSmDate = DateValue(gAdjYear(smDate))
        llEndDate = DateValue(gAdjYear(sEndDate))
        '''Get CPTT so that Stations requiring CP can be obtained
        ''SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
        ''SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        ''SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, mktName"
        ''SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"
        'SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attACName"
        If Not udcCriteria.iExportType(0) Then
            SQLQuery = "SELECT shttCallLetters, cpttShfCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attACName"
            SQLQuery = SQLQuery + " FROM shtt, cptt, att"
            SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
            SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
            '10/29/14: Bypass Service agreements
            SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
            SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery & " Order by shttCallLetters"
        Else
            SQLQuery = "SELECT cpttShfCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attACName"
            SQLQuery = SQLQuery + " FROM cptt, att"
            SQLQuery = SQLQuery + " WHERE (attCode = cpttAtfCode"
            '10/29/14: Bypass Service agreements
            SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
            SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
            SQLQuery = SQLQuery + " AND cpttShfCode = " & imShttCode
            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        End If
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            If igExportSource = 2 Then DoEvents
            'sCallLetters = Trim$(cprst!shttCallLetters)
            'iShfCode = cprst!shttCode
            iShfCode = cprst!cpttshfcode
            ilShtt = gBinarySearchStationInfoByCode(iShfCode)
            If ilShtt <> -1 Then
                sCallLetters = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                'TTP 10926
                sTimeZone = Trim$(Mid(tgStationInfoByCode(ilShtt).sZone, 1, 1))
            Else
                iShfCode = -1
                sTimeZone = ""
            End If
            If igExportSource = 2 Then DoEvents
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For iLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If lbcStation.Selected(iLoop) Then
                        If lbcStation.ItemData(iLoop) = iShfCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next iLoop
            Else
                ilOkStation = True
            End If
            If ilOkStation Then
                '11/16/17
                mCloseDuplInfo
                Set DuplInfo_rst = mInitDuplInfo()
                Print #hmMsg, "Gather ISCI for: " & sCallLetters; " on " & Trim$(smVefName) & " at " & Format$(Now, "m/d/yyyy " & sgShowTimeWSecForm) & " by " & Trim$(sgUserName)
                lacResult.Caption = "Gather ISCI for: " & sCallLetters & " on " & Trim$(smVefName)
                If igExportSource = 2 Then DoEvents
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cprst!cpttCode
                tgCPPosting(0).iStatus = cprst!cpttStatus
                tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cprst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cprst!attTimeType
                tgCPPosting(0).iVefCode = imVefCode
                tgCPPosting(0).iShttCode = iShfCode 'cprst!shttCode
                tgCPPosting(0).sZone = Trim$(tgStationInfoByCode(ilShtt).sZone)      'cprst!shttTimeZone
                tgCPPosting(0).sDate = Format$(sMoDate, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                
                'Create AST records
                slACName = Trim$(cprst!attACName)
                llPrevFeedDate = -1
                If ilNoAirDays > 1 Then
                    ilFirstDate = True
                End If
                ilNoAirDays = 0
                ilHourNo = 0
                ilSegmentNo = 0
                ilBreakNo = 0
                ilPositionNo = 0
                llPrevFeedTime = -1
                igTimes = 1 'By Week
                imAdfCode = -1
                ilVff = gBinarySearchVff(imVefCode)
                If ilVff <> -1 Then
                    'If Trim$(tgVffInfo(ilVff).sXDXMLForm) = "P" Then
                    '    slProgCode = ""
                    'Else
                        slProgCode = Trim$(tgVffInfo(ilVff).sXDProgCodeID)
                    'End Iff
                Else
                    slProgCode = ""
                End If
                If igExportSource = 2 Then DoEvents
                '10927
                If udcCriteria.iExportType(1) And udcCriteria.IIncludeCommands(3) = vbChecked Then
                    myZoneAndDSTHelper.StationZone = sTimeZone
                    myZoneAndDSTHelper.StationHonorDaylight tgStationInfoByCode(ilShtt).iAckDaylight
                    If myZoneAndDSTHelper.StationZone <> myZoneAndDSTHelper.SiteZone Then
                        igTimes = 3
                        tgCPPosting(0).iNumberDays = imNumberDays + 1
                        'right now site hard coded as "E".  But if site "C" and station is "E", we have to get the NEXT day, not the previous.
                        'the test is if we need previous day
                        If myZoneAndDSTHelper.StationZone < myZoneAndDSTHelper.SiteZone Then
                            tgCPPosting(0).sDate = DateAdd("d", -1, smDate)
                        End If
                    End If
                End If
                iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True)
                gFilterAstExtendedTypes tmAstInfo
                mMovePledgeToFeed
                'Output AST
                For iLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    blSpotOk = True
                    ilAnf = gBinarySearchAnf(tmAstInfo(iLoop).iAnfCode)
                    If ilAnf <> -1 Then
                        If tgAvailNamesInfo(ilAnf).sISCIExport = "N" Then
                            blSpotOk = False
                        End If
                    End If
                    If blSpotOk Then
                        lFeedDate = DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate))
                         '10927 Main work is here, calculate the new hour with zone and dst and set as military time. If rolled to next day, handle that too
                        If udcCriteria.iExportType(1) And udcCriteria.IIncludeCommands(3) = vbChecked Then
                            'slMilitaryHours = mConvertToRealHour(iLoop)
                            slMilitaryHours = myZoneAndDSTHelper.ZoneByMethod(Trim$(tmAstInfo(iLoop).sFeedDate), Trim$(tmAstInfo(iLoop).sFeedTime))
                            If myZoneAndDSTHelper.isNextDay Then
                                'these 2 force the 'print to next day'
                                lFeedDate = lFeedDate + 1
                                tmAstInfo(iLoop).sFeedDate = Format$(lFeedDate, sgShowDateForm)
                            ElseIf myZoneAndDSTHelper.isPreviousDay Then
                                lFeedDate = lFeedDate - 1
                                tmAstInfo(iLoop).sFeedDate = Format$(lFeedDate, sgShowDateForm)
                            End If
                        End If
                        'Bypass duplicates
                        ilDuplicateSpot = False
                        '8/6/18: Include adding the first spot to the array
                        'If iLoop <> LBound(tmAstInfo) Then
                            llFeedTime = gTimeToLong(tmAstInfo(iLoop).sFeedTime, False)
                            '11/16/17: Replaced with result set test
                            'For ilDuptTest = iLoop - 1 To LBound(tmAstInfo) Step -1
                            '    If igExportSource = 2 Then DoEvents
                            '    If (DateValue(gAdjYear(tmAstInfo(ilDuptTest).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(ilDuptTest).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(ilDuptTest).iPledgeStatus)).iPledged <> 2) Then
                            '        llDuplTestTime = gTimeToLong(tmAstInfo(ilDuptTest).sFeedTime, False)
                            '        '4/29/15: with igTimes set to 1 the sort is by air date and air time causing duplicate spots not to be adjacent
                            '        'If llFeedTime <> llDuplTestTime Then
                            '        '    Exit For
                            '        'End If
                            '        'If tmAstInfo(ilDuptTest).lSdfCode = tmAstInfo(iLoop).lSdfCode Then
                            '        '    ilDuplicateSpot = True
                            '        '    Exit For
                            '        'End If
                            '        If llFeedTime = llDuplTestTime Then
                            '            If tmAstInfo(ilDuptTest).lSdfCode = tmAstInfo(iLoop).lSdfCode Then
                            '                ilDuplicateSpot = True
                            '                Exit For
                            '            End If
                            '        End If
                            '    End If
                            'Next ilDuptTest
                            ilDuplicateSpot = mTestDuplInfo(tmAstInfo(iLoop).lSdfCode, llFeedTime, lFeedDate, llSmDate, llEndDate, tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iPledgeStatus)).iPledged)
                        'End If
                        '11/16/17
                        'If (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iPledgeStatus)).iPledged <> 2) And (Not ilDuplicateSpot) Then
                        If (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) >= llSmDate) And (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) <= llEndDate) And (tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iPledgeStatus)).iPledged <> 2) And (Not ilDuplicateSpot) Then
                            If llPrevFeedDate <> lFeedDate Then
                                ilNoAirDays = ilNoAirDays + 1
                                If udcCriteria.iExportType(1) Then
                                    If (llPrevFeedTime <> -1) And (udcCriteria.IIncludeCommands(0) = vbChecked) Then
                                        'Test if last segment should be included.
                                        'If the last break contents only Promo spots, then don't include last segment because these
                                        'promo will air at different places and the last avail was just a place holder
                                        ilPromoOnly = True
                                        For ilTest = ilStartIndexOfBreak To iLoop - 1 Step 1
                                            If igExportSource = 2 Then DoEvents
                                            'SQLQuery = "SELECT chfType FROM chf_Contract_Header INNER JOIN sdf_Spot_Detail On chfCode = sdfChfCode WHERE sdfCode = " & tmAstInfo(ilTest).lSdfCode
                                            If tmAstInfo(ilTest).lSdfCode > 0 Then
                                                SQLQuery = "SELECT sdfChfCode FROM sdf_Spot_Detail WHERE sdfCode = " & tmAstInfo(ilTest).lSdfCode
                                                Set chfrst = gSQLSelectCall(SQLQuery)
                                                If Not chfrst.EOF Then
                                                    SQLQuery = "SELECT chfType FROM chf_Contract_Header WHERE chfCode = " & chfrst!sdfChfCode
                                                    Set chfrst = gSQLSelectCall(SQLQuery)
                                                    If Not chfrst.EOF Then
                                                        If chfrst!chfType <> "M" Then
                                                            ilPromoOnly = False
                                                            Exit For
                                                        End If
                                                    Else
                                                        ilPromoOnly = False
                                                        Exit For
                                                    End If
                                                Else
                                                    ilPromoOnly = False
                                                    Exit For
                                                End If
                                            End If
                                        Next ilTest
                                        If igExportSource = 2 Then DoEvents
                                        On Error Resume Next
                                        chfrst.Close
                                        On Error GoTo 0
                                        If Not ilPromoOnly Then
                                            slHour = Trim$(Str$(ilHourNo))
                                            If Len(slHour) = 1 Then
                                                slHour = "0" & slHour
                                            End If
                                            slSegment = Trim$(Str$(ilSegmentNo + 1))
                                            If Len(slSegment) = 1 Then
                                                slSegment = "0" & slSegment
                                            End If
                                            Print #hmISCIbyBreak, "E" & slHour & slSegment & Chr$(9) & slYear & slWeekNo & slDay & slProgCode & "-" & "H" & slHour & "S" & slSegment & Chr$(9) & "Show"
                                        End If
                                    End If
                                    'If ((llPrevFeedDate = -1) And (ilFirstDate)) Or ((llPrevFeedDate <> lFeedDate) And (llPrevFeedDate <> -1)) Then
                                    If ((llPrevFeedDate = -1) And (ilFirstDate) And ((imNoAirDays <> 1) Or (lmFeedDateLastA <> lFeedDate))) Or ((llPrevFeedDate <> lFeedDate) And (llPrevFeedDate <> -1)) Then
                                        Print #hmISCIbyBreak, "A" & Format(gAdjYear(tmAstInfo(iLoop).sFeedDate), "mm/dd/yy")
                                        imNoAirDays = ilNoAirDays
                                        lmFeedDateLastA = lFeedDate
                                        ilGenB = True
                                    End If
                                    If ilGenB Then
                                        Print #hmISCIbyBreak, "B" & smVefName & Chr$(9) & slProgCode
                                        ilGenB = False
                                    End If
                                    ilFirstDate = False
                                    Print #hmISCIbyBreak, "C" & sCallLetters
                                    'TTP 10926
                                    If udcCriteria.iExportType(1) = True Then 'Make sure "ISCI by Break" export type is selected
                                        If udcCriteria.IIncludeCommands(3) = vbChecked Then Print #hmISCIbyBreak, "Z" & sTimeZone
                                    End If
                                    If udcCriteria.IIncludeCommands(1) = vbChecked Then
                                        '4/28/10:  Until E-Mail address added to agreement, use E-Mail address from Station Affiliate Contact
                                        'If slACName = "" Then
                                            'SQLQuery = "SELECT arttEmail FROM artt WHERE (arttShttCode = '" & iShfCode & "' And arttAffContact = '1')"
                                            'Set rst = gSQLSelectCall(SQLQuery)
                                            'If Not rst.EOF Then
                                            '    slContactEMail = Trim$(rst!arttEmail)
                                            '    If slContactEMail <> "" Then
                                            '        Print #hmISCIbyBreak, "F" & slContactEMail
                                            '    End If
                                            'End If
                                            
                                            SQLQuery = "SELECT arttEmail FROM artt WHERE (arttShttCode = '" & iShfCode & "' And arttISCI2Contact = '1')"
                                            Set rst = gSQLSelectCall(SQLQuery)
                                            Do While Not rst.EOF
                                                If igExportSource = 2 Then DoEvents
                                                slContactEMail = Trim$(rst!arttEmail)
                                                If slContactEMail <> "" Then
                                                    Print #hmISCIbyBreak, "F" & slContactEMail
                                                End If
                                                rst.MoveNext
                                            Loop
    
                                        'End If
                                    End If
                                End If
                                ilHourNo = 0
                                ilSegmentNo = 0
                                ilBreakNo = 0
                                ilPositionNo = 0
                                llPrevFeedTime = -1
                                slYear = right$(Year(gObtainNextSunday(gAdjYear(tmAstInfo(iLoop).sFeedDate))), 2)
                                llStartYearDate = DateValue(gObtainYearStartDate(gAdjYear(tmAstInfo(iLoop).sFeedDate)))
                                slWeekNo = Trim$(Str$((lFeedDate - llStartYearDate) \ 7 + 1))
                                If Len(slWeekNo) = 1 Then
                                    slWeekNo = "0" & slWeekNo
                                End If
                                slDay = Trim$(Str$(Weekday(gAdjYear(tmAstInfo(iLoop).sFeedDate), vbMonday)))
                            End If
                            llPrevFeedDate = lFeedDate
                            llFeedTime = gTimeToLong(tmAstInfo(iLoop).sFeedTime, False)
                            If llPrevFeedTime <> llFeedTime Then
                                If llPrevFeedTime <> -1 Then
                                    If Hour(gLongToTime(llPrevFeedTime)) <> Hour(gLongToTime(llFeedTime)) Then
                                        ilHourNo = ilHourNo + 1
                                        ilSegmentNo = 0
                                        ilBreakNo = 0
                                        ilPositionNo = 0
                                    End If
                                Else
                                    ilHourNo = 1
                                End If
                                ilSegmentNo = ilSegmentNo + 1
                                ilBreakNo = ilBreakNo + 1
                                ilPositionNo = 1
                                If (udcCriteria.iExportType(1)) And (udcCriteria.IIncludeCommands(0) = vbChecked) Then
                                    slHour = Trim$(Str$(ilHourNo))
                                    If Len(slHour) = 1 Then
                                        slHour = "0" & slHour
                                    End If
                                    slSegment = Trim$(Str$(ilSegmentNo))
                                    If Len(slSegment) = 1 Then
                                        slSegment = "0" & slSegment
                                    End If
                                    Print #hmISCIbyBreak, "E" & slHour & slSegment & Chr$(9) & slYear & slWeekNo & slDay & slProgCode & "-" & "H" & slHour & "S" & slSegment & Chr$(9) & "Show"
                                End If
                                ilStartIndexOfBreak = iLoop
                            Else
                                ilPositionNo = ilPositionNo + 1
                            End If
                            llPrevFeedTime = llFeedTime
                            sISCI = Trim$(tmAstInfo(iLoop).sISCI)
                            llAdf = gBinarySearchAdf(CLng(tmAstInfo(iLoop).iAdfCode))
                            If llAdf <> -1 Then
                                sAdvt = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                            Else
                                sAdvt = "Missing" & Str(tmAstInfo(iLoop).iAdfCode)
                            End If
                            
                            ''6/12/06- Check if any region copy defined for the spots
                            ''ilRet = gGetRegionCopy(tmAstInfo(iLoop).iShttCode, tmAstInfo(iLoop).lSdfCode, tmAstInfo(iLoop).iVefCode, sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                            'ilRet = gGetRegionCopy(tmAstInfo(iLoop), sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                            'If ilRet Then
                            If tmAstInfo(iLoop).iRegionType > 0 Then
                                sISCI = Trim$(tmAstInfo(iLoop).sRISCI) 'sRISCI
                            End If
                            
                            If igExportSource = 2 Then DoEvents
                            If InStr(1, sAdvt, "Missing", vbTextCompare) = 1 Then
                                Print #hmMsg, Trim$(smVefName) & ": Advertiser " & sAdvt & " on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                                lbcMsg.AddItem Trim$(smVefName) & ": Advertiser " & sAdvt & " on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                            Else
                                If sISCI <> "" Then
                                    If udcCriteria.iExportType(1) Then
                                        slHour = Trim$(Str$(ilHourNo))
                                        If Len(slHour) = 1 Then
                                            slHour = "0" & slHour
                                        End If
                                        slBreak = Trim$(Str$(ilBreakNo))
                                        If Len(slBreak) = 1 Then
                                            slBreak = "0" & slBreak
                                        End If
                                        slPosition = Trim$(Str$(ilPositionNo))
                                        If Len(slPosition) = 1 Then
                                            slPosition = "0" & slPosition
                                        End If
                                        'Print #hmISCIbyBreak, "D" & slHour & slBreak & slPosition & Chr$(9) & sISCI & Chr$(9) & sAdvt
                                        '10927 I redid this
'                                        If (udcCriteria.IIncludeCommands(2) = vbChecked) Then
'                                            Print #hmISCIbyBreak, "D" & slHour & slBreak & slPosition & Chr$(9) & mReplaceTab(sISCI) & Chr$(9) & mReplaceTab(sAdvt) & Chr$(9) & tmAstInfo(iLoop).lCode
'                                        Else
'                                            Print #hmISCIbyBreak, "D" & slHour & slBreak & slPosition & Chr$(9) & mReplaceTab(sISCI) & Chr$(9) & mReplaceTab(sAdvt)
'                                        End If
'                                        '10927
                                        If udcCriteria.iExportType(1) And udcCriteria.IIncludeCommands(3) = vbChecked Then
                                            slHour = slMilitaryHours
                                        End If
                                        slTempString = "D" & slHour & slBreak & slPosition & Chr$(9) & mReplaceTab(sISCI) & Chr$(9) & mReplaceTab(sAdvt)
                                        If (udcCriteria.IIncludeCommands(2) = vbChecked) Then
                                            slTempString = slTempString & Chr$(9) & tmAstInfo(iLoop).lCode
                                        End If
                                        Print #hmISCIbyBreak, slTempString
                                    Else
                                        iFound = False
                                        If imEmbedded Then
                                            tmISCISendInfo_E(lmSendInfoUpper_E).sVehName = smVefName
                                            tmISCISendInfo_E(lmSendInfoUpper_E).sISCI = sISCI
                                            slKey = "!" & smCommProvArfCode & "|" & tmISCISendInfo_E(lmSendInfoUpper_E).sVehName & "|" & tmISCISendInfo_E(lmSendInfoUpper_E).sISCI & "|" & smProducerArfCode
                                            iFound = mBinarySearchSendInfo(0, lmSendInfoUpper_E - 1, slKey, tmISCISendInfo_E())
                                            If Not iFound Then
                                                mAddToISCISendInfo sISCI, iShfCode, sCallLetters, sAdvt, lFeedDate, slACName, lmSendInfoUpper_E, tmISCISendInfo_E()
                                            End If
                                        Else
                                            tmISCISendInfo(lmSendInfoUpper).sVehName = smVefName
                                            tmISCISendInfo(lmSendInfoUpper).sISCI = sISCI
                                            tmISCISendInfo(lmSendInfoUpper).sCallLetters = sCallLetters
                                            slKey = smCommProvArfCode & "|" & tmISCISendInfo(lmSendInfoUpper).sVehName & "|" & tmISCISendInfo(lmSendInfoUpper).sCallLetters & "|" & tmISCISendInfo(lmSendInfoUpper).sISCI
                                            iFound = mBinarySearchSendInfo(0, lmSendInfoUpper - 1, slKey, tmISCISendInfo())
                                            If Not iFound Then
                                                mAddToISCISendInfo sISCI, iShfCode, sCallLetters, sAdvt, lFeedDate, slACName, lmSendInfoUpper, tmISCISendInfo()
                                            End If
                                        End If
                                        
                                        If iFound Then
                                            lmFound = lmFound + 1
                                        Else
                                            lmNotFound = lmNotFound + 1
                                        End If
                                        
                                        'lgETime4 = timeGetTime
                                        'lgTtlTime4 = lgTtlTime4 + (lgETime4 - lgSTime4)
                                        
                                        'end timer 4
                                        'start timer 5
                                        'lgSTime5 = timeGetTime
                                        
                                    End If
                                    'lgETime5 = timeGetTime
                                    'lgTtlTime5 = lgTtlTime5 + (lgETime5 - lgSTime5)
                                Else
                                    Print #hmMsg, Trim$(smVefName) & " " & sAdvt & ": ISCI Missing on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                                    lbcMsg.AddItem Trim$(smVefName) & " " & sAdvt & ": ISCI Missing on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                                End If
                            End If
                        End If
                    End If
                Next iLoop
                If (llPrevFeedTime <> -1) Then
                    If (udcCriteria.iExportType(1)) Then
                        If (udcCriteria.IIncludeCommands(0) = vbChecked) Then
                            'Test if last segment should be included.
                            'If the last break contents only Promo spots, then don't include last segment because these
                            'promo will air at different places and the last avail was just a place holder
                            ilPromoOnly = True
                            For ilTest = ilStartIndexOfBreak To iLoop - 1 Step 1
                                If igExportSource = 2 Then DoEvents
                                If tmAstInfo(ilTest).lSdfCode > 0 Then
                                    SQLQuery = "SELECT sdfChfCode FROM sdf_Spot_Detail WHERE sdfCode = " & tmAstInfo(ilTest).lSdfCode
                                    Set chfrst = gSQLSelectCall(SQLQuery)
                                    If Not chfrst.EOF Then
                                        SQLQuery = "SELECT chfType FROM chf_Contract_Header WHERE chfCode = " & chfrst!sdfChfCode
                                        Set chfrst = gSQLSelectCall(SQLQuery)
                                        If Not chfrst.EOF Then
                                            If chfrst!chfType <> "M" Then
                                                ilPromoOnly = False
                                                Exit For
                                            End If
                                        Else
                                            ilPromoOnly = False
                                            Exit For
                                        End If
                                    Else
                                        ilPromoOnly = False
                                        Exit For
                                    End If
                                    If igExportSource = 2 Then DoEvents
                                End If
                            Next ilTest
                            On Error Resume Next
                            chfrst.Close
                            On Error GoTo 0
                            If Not ilPromoOnly Then
                                slHour = Trim$(Str$(ilHourNo))
                                If Len(slHour) = 1 Then
                                    slHour = "0" & slHour
                                End If
                                slSegment = Trim$(Str$(ilSegmentNo + 1))
                                If Len(slSegment) = 1 Then
                                    slSegment = "0" & slSegment
                                End If
                                Print #hmISCIbyBreak, "E" & slHour & slSegment & Chr$(9) & slYear & slWeekNo & slDay & slProgCode & "-" & "H" & slHour & "S" & slSegment & Chr$(9) & "Show"
                            End If
                        End If
                    End If
                End If
            End If
            cprst.MoveNext
        Wend
        If udcCriteria.iExportType(1) Then
            If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
                gClearASTInfo True
            Else
                gClearASTInfo False
            End If
        End If
        sMoDate = DateAdd("d", 7, sMoDate)
        '10927
        If udcCriteria.iExportType(1) And udcCriteria.IIncludeCommands(3) = vbChecked Then
            sMoDate = DateAdd("d", imNumberDays + 1, sMoDate)
        End If
    Loop While DateValue(gAdjYear(sMoDate)) < DateValue(gAdjYear(sEndDate))

    mGatherISCI_AllFormat = True
    
    'lgETime3 = timeGetTime
    'lgTtlTime3 = lgTtlTime3 + (lgETime3 - lgSTime3)
    
    
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mGatherISCI_AllFormats"
    Resume Next
    mGatherISCI_AllFormat = False
    Exit Function
    
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
Private Function mOpenMsgFile(sMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    On Error GoTo mOpenMsgFileErr:
    ilRet = 0
    slNowDate = Format$(gNow(), sgShowDateForm)
    slToFile = sgMsgDirectory & "ExptISCI.Txt"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
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
                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
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
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    Print #hmMsg, "** Export ISCI: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Sub mFillVehicle()
    Dim iLoop As Integer
    Dim ilRet As Integer
    
    ilRet = gPopVff()
    imEmbeddedAllowed = False
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            If (tgVehicleInfo(iLoop).iCommProvArfCode > 0) Then
                'If (tgVehicleInfo(iLoop).sEmbeddedComm <> "Y") Or ((tgVehicleInfo(iLoop).sEmbeddedComm = "Y") And (tgVehicleInfo(iLoop).iProducerArfCode > 0)) Then
                
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
                    If tgVehicleInfo(iLoop).iProducerArfCode > 0 Then
                        imEmbeddedAllowed = True
                    End If
                'End If
            End If
        'End If
    Next iLoop
End Sub




Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
        If lbcVehicles.ListCount > 1 Then
            edcTitle3.Visible = False
            'chkAllStation.Visible = False
            lbcStation.Visible = False
            lbcStation.Clear
        Else
            edcTitle3.Visible = True
            chkAllStation.Visible = True
            lbcStation.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        'D.S. 9/10/19
        If iValue Then
            lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        End If
        imAllClick = False
    End If
    mFillStations

End Sub

Private Sub chkAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllStationClick Then
        Exit Sub
    End If
    If chkAllStation.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    'lacStationMsg.Visible = Not lbcStation.Visible
    
    If chkAllStation.Value = vbChecked Then
        lbcStation.Visible = False
        lacStationMsg.Visible = True
    Else
        '11/8/19: Fill only during export and all unchecked
        If (rbcFilter(4).Value = True) And (iValue = False) Then
            imAllStationClick = True
            mFillStations True
            imAllStationClick = False
        End If
        lbcStation.Visible = True
        lacStationMsg.Visible = False
    End If
    
    
    If lbcStation.ListCount > 0 Then
        imAllStationClick = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
    End If

End Sub


Private Sub cmdExport_Click()
    Dim iLoop As Integer
    Dim ilRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String
    Dim sNowDate As String
    Dim iIndex As Integer
    Dim ilSort As Integer
    Dim ilPrevCommProvArfCode As Integer
    Dim llSetKey As Long

    On Error GoTo ErrHand
    
    lgTtlTime4 = 0
    lgTtlTime5 = 0
    lgTtlTime8 = 0
    lgTtlTime2 = 0
    lgTtlTime1 = 0
    lmNotFound = 0
    lmFound = 0
    
    tmcFilterDelay.Enabled = False
    lgSTime1 = timeGetTime
    
    smNowDate = Format(gNow(), "m/d/yy")
   
    lbcMsg.Clear
    If lbcVehicles.ListIndex < 0 Then
        igExportReturn = 2
        Exit Sub
    End If
    If edcDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        edcDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        edcDate.SetFocus
        Exit Sub
    Else
        smDate = Format(edcDate.Text, sgShowDateForm)
    End If
    sMoDate = gObtainPrevMonday(smDate)
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    If (udcCriteria.ISpots(0) = False) And (udcCriteria.ISpots(1) = False) And (udcCriteria.ISpots(2) = False) And (udcCriteria.iExportType(0)) Then
        Beep
        gMsgBox "Please Specify Export ISCI Type.", vbCritical
        Exit Sub
    End If
    '11/8/19: Fill only during export and all unchecked
    If (rbcFilter(4).Value = True) And (chkAllStation.Value = vbChecked) Then
        mFillStations True
    End If
    Screen.MousePointer = vbHourglass
    smExportToPath = udcCriteria.IExportToPath
    If right(smExportToPath, 1) <> "\" Then
        smExportToPath = smExportToPath & "\"
    End If
    If Not mOpenMsgFile(sMsgFileName) Then
        igExportReturn = 2
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    
    imExporting = True
    
    
    mSaveCustomValues
    If Not gPopCopy(sMoDate, "Export ISCI") Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Exit Sub
    End If
    'SQLQuery = "SELECT siteDaysResendISCI From Site Where siteCode = 1"
    'Set rst = gSQLSelectCall(SQLQuery)
    'If Not rst.EOF Then
    '    imNoDaysResendISCI = rst!siteDaysResendISCI
    'Else
    '    imNoDaysResendISCI = 0
    'End If
    If (udcCriteria.ISpots(0) = True) And (udcCriteria.iExportType(0)) Then
        imNoDaysResendISCI = Val(udcCriteria.INoDaysResend)
    Else
        imNoDaysResendISCI = 0
    End If
    On Error GoTo 0
    imNoAirDays = -1
    lmFeedDateLastA = 0
    ReDim tmISCIXRef(0 To 20000) As ISCIXREF
    ReDim tmISCISendInfo(0 To 20000) As ISCISENDINFO
    For llSetKey = 0 To UBound(tmISCISendInfo) Step 1
        tmISCISendInfo(llSetKey).sKey = "~~~~~~~~~~"
    Next llSetKey
    ReDim tmISCISendInfo_E(0 To 20000) As ISCISENDINFO
    For llSetKey = 0 To UBound(tmISCISendInfo_E) Step 1
        tmISCISendInfo_E(llSetKey).sKey = "~~~~~~~~~~"
    Next llSetKey
    lmXRefUpper = 0
    lmSendInfoUpper = 0
    lmSendInfoUpper_E = 0
    lacResult.Caption = ""
    Screen.MousePointer = vbHourglass
    plcGauge.Value = 0
    plcGauge.Visible = True
    lmPercent = 0
    lmProcessedNumber = 0
    bgTaskBlocked = False
    sgTaskBlockedName = "ISCI Export"
    If udcCriteria.iExportType(0) Then
        ilRet = mExportUniqueISCI()
    ElseIf udcCriteria.iExportType(1) Then
        ilRet = mExportISCIByBreak()
    End If
    gCloseRegionSQLRst

    ilRet = gCustomEndStatus(lmEqtCode, igExportSource, "")
    Screen.MousePointer = vbDefault
    imExporting = False
    
    
    lgETime1 = timeGetTime
    lgTtlTime1 = (lgETime1 - lgSTime1)

    On Error Resume Next
    If Not imTerminate Then
        Print #hmMsg, "** Total Export time in seconds: " & lgTtlTime1 / 1000 & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Total  mGatherISCI_AllFormat in seconds: " & lgTtlTime3 / 1000 & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Total gGetAstInfo in seconds: " & lgTtlTime2 / 1000 & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Total Region Copy in seconds: " & lgTtlTime8 / 1000 & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Timer 4: " & lgTtlTime4 / 1000 & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Timer 5: " & lgTtlTime5 / 1000 & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Found = " & lmFound & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        'Print #hmMsg, "** Not Found = " & lmNotFound & " " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        Print #hmMsg, "** Completed Export ISCI: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " " & Trim$(sgUserName) & " **"
        
        lacResult.Caption = "Results: " & sMsgFileName
        plcGauge.Value = 100
    End If
    Close #hmMsg
    If bgTaskBlocked And igExportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Export generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    If igExportSource <> 2 Then
        ilRet = gAlertForceCheck()
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
cmdExportErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-cmcExport_Click"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmExportISCI
End Sub


Private Sub edcExportToPath_Change()
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
        plcGauge.Visible = False
    End If
End Sub

Private Sub edcNoDaysResend_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcNoDaysResend_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    chkAllStation.Visible = True
    'chkAllStation.Enabled = True
    If imFirstTime Then
        udcCriteria.Left = lacExportDate.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        'udcCriteria.Top = txtDate.Top + (3 * txtDate.Height) / 4
        udcCriteria.Top = txtNumberDays.Top + txtNumberDays.Height
        udcCriteria.Action 6
        If UBound(tgEvtInfo) > 0 Then
            chkAll.Value = vbUnchecked
            lbcStation.Clear
            lbcVehicles.Clear
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                If llVef <> -1 Then
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgEvtInfo(ilLoop).iVefCode
                End If
            Next ilLoop
            chkAll.Value = vbChecked
            If lbcVehicles.ListCount = 1 Then
                imVefCode = lbcVehicles.ItemData(0)
                edcTitle3.Visible = True
                chkAllStation.Visible = True
                lbcStation.Visible = True
                mFillStations
            End If
        End If
        If igExportSource = 2 Then
            slNowStart = gNow()
            edcDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "ISCIResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "ISCI Result List, Started: " & slNowStart
            ' pass global so glogMsg will write messages to sgExportResultName
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "ISCIResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "ISCI Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "ISCI Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            imTerminate = True
            '6394 clear values
            hgExportResult = 0
            tmcTerminate.Enabled = True
        End If
        lacStationMsg.Move lbcStation.Left, lbcStation.Top, lbcStation.Width, lbcStation.Height
        lacStationMsg.ZOrder
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    frmExportISCI.Caption = "Export ISCI - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcDate.Text = smDate
    imNumberDays = 7
    smSvNumberDays = "7"
    txtNumberDays.Text = Trim$(Str$(imNumberDays))
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    ilRet = gPopAdvertisers()
    
    lbcStation.Clear
    mFillVehicle
    chkAll.Value = vbChecked
    udcCriteria.Embedded = imEmbeddedAllowed
    'If imEmbeddedAllowed Then
    '    rbcUniqueBy(0).Enabled = True
    'Else
    '    rbcUniqueBy(0).Enabled = False
    'End If
    ReDim smFileNamesCreated(0 To 0) As String
    ReDim imPrevVefCode(0 To 0) As Integer
    'If sgExportISCI = "A" Then
    '    frmVeh.Visible = False
    '    rbcSpots(2).Value = True
    'End If
    ilRet = gPopAvailNames()
    'lacStationMsg.Move lbcStation.Left, lbcVehicles.Height, lbcStation.Width, lbcVehicles.Top
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    If imExporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    mCloseDuplInfo
    Erase smFileNamesCreated
    Erase imPrevVefCode
    Erase tmCPDat
    Erase tmISCIXRef
    Erase tmISCISendInfo
    Erase tmISCISendInfo_E
    Erase tmAstInfo
    Set frmExportISCI = Nothing
End Sub


Private Sub lbcStation_Click()
    If imAllStationClick Then
        Exit Sub
    End If
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
        plcGauge.Visible = False
    End If
    If chkAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
End Sub

Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    'Move to mouse up
    'lbcStation.Clear
    'If mGetGrdSelCount() = 1 Then
    '    edcTitle3.Visible = True
    '    chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    'Else
    '    edcTitle3.Visible = False
    '    chkAllStation.Visible = False
    '    lbcStation.Visible = False
    'End If
    imBypassAll = True
    chkAll.Value = vbUnchecked
    imBypassAll = False
    
    
    
    
'    lbcStation.Clear
'    If cmdExport.Enabled = False Then
'        cmdExport.Enabled = True
'        cmdCancel.Caption = "&Cancel"
'        plcGauge.Visible = False
'    End If
'    If chkAllStation.Value = vbChecked Then
'        chkAllStation.Value = vbUnchecked
'    End If
'    If imAllClick Then
'        Exit Sub
'    End If
'    If chkAll.Value = vbChecked Then
'        imAllClick = True
'        chkAll.Value = vbUnchecked
'        imAllClick = False
'    End If
'    For iLoop = 0 To lbcVehicles.ListCount - 1 Step 1
'        If lbcVehicles.Selected(iLoop) Then
'            imVefCode = lbcVehicles.ItemData(iLoop)
'            iCount = iCount + 1
'            If iCount > 1 Then
'                Exit For
'            End If
'        End If
'    Next iLoop
'    If iCount = 1 Then
'        edcTitle3.Visible = True
'        chkAllStation.Visible = True
'        'lbcStation.Visible = True
'        lbcStation.Visible = False
'        mFillStations
'    Else
'        edcTitle3.Visible = False
'        'chkAllStation.Visible = False
'        'lbcStation.Visible = False
'        mFillStations
'    End If
End Sub


Private Sub edcDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
        plcGauge.Visible = False
    End If
End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:mGatherISCI_CompressedFormat    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Gather ISCI by Provider,        *
'*                     station.  Create two tables     *
'*                     A cross Reference table which is*
'*                     station within vehicle by       *
'*                     provider                        *
'*                     The other table is ISCI within  *
'*                     station by provider             *
'*                                                     *
'*******************************************************
Private Function mGatherISCI_CompressedFormat() As Integer
    'dan m 6/4/15 removed a ton of variables that were set but never used

    'Dim sDate As String
   ' Dim iNoWeeks As Integer
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim llAdf As Long
    Dim sAdvt As String
    Dim llIndex As Long
    Dim sCallLetters As String
    Dim iShfCode As Integer
    Dim sISCI As String
    'Dim llUpper As Long
   ' Dim slStr As String
    Dim iFound As Integer
    Dim lFeedDate As Long
  '  Dim llPrevFeedDate As Long
    Dim ilOkStation As Integer
   ' Dim iExport As Integer  '0=Don't export as it did not change
                            '1=Export and create aet record
                            '2=Export and don't create aet reord (nothing changed but generating all spot export)
    'Dim sRCart As String
    Dim sRISCI As String
   ' Dim sRCreative As String
   ' Dim sRProd As String
   ' Dim lRCrfCsfCode As Long
   ' Dim lRCrfCode As Long
   ' Dim ilRet As Integer
   ' Dim ilHourNo As Integer
   ' Dim ilSegmentNo As Integer
   ' Dim ilPositionNo As Integer
   ' Dim ilBreakNo As Integer
    Dim llFeedTime As Long
   ' Dim llPrevFeedTime As Long
    'Dim slHour As String
    'Dim slSegment As String
    'Dim slBreak As String
    'Dim slPosition As String
    'Dim slYear As String
    'Dim slWeekNo As String
    'Dim llStartYearDate As Long
    'Dim ilVff As Integer
    'Dim slProgCode As String
    'Dim slACName As String
    'Dim slContactEMail As String
    Dim ilDuplicateSpot As Integer
    Dim llDuplTestTime As Long
    Dim ilDuptTest As Integer
   ' Dim slDay As String
    'Dim ilFirstDate As Integer
    'Dim ilGenB As Integer
   ' Dim ilStartIndexOfBreak As Integer
   ' Dim ilPromoOnly As Integer
  '  Dim ilTest As Integer
   ' Dim ilNoAirDays As Integer
    Dim llSetKey As Long
    Dim slKey As String
    Dim ilAnf As Integer
    Dim blSpotOk As Boolean
    
    On Error GoTo ErrHand
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
   ' ilFirstDate = True
   ' ilGenB = True
    
    'D.S. 11/21/05
'    iRet = gGetMaxAstCode()
'    If Not iRet Then
'        Exit Function
'    End If
    
    Do
        If igExportSource = 2 Then DoEvents
        ''Get CPTT so that Stations requiring CP can be obtained
        'SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
        'SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        'SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, mktName"
        'SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attACName"
        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
        '10/29/14: Bypass Service agreements
        SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
        'SQLQuery = SQLQuery + " AND attExportType = 2"
        SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " Order by shttCallLetters"
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            If igExportSource = 2 Then DoEvents
            sCallLetters = Trim$(cprst!shttCallLetters)
            iShfCode = cprst!shttCode
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For iLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If lbcStation.Selected(iLoop) Then
                        If lbcStation.ItemData(iLoop) = iShfCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next iLoop
            Else
                ilOkStation = True
            End If
            If ilOkStation Then
                iFound = False
                'For llIndex = 0 To UBound(tmISCIXRef) - 1 Step 1
                For llIndex = 0 To lmXRefUpper - 1 Step 1
                    If (StrComp(Trim$(tmISCIXRef(llIndex).sVehName), smVefName, vbTextCompare) = 0) And (StrComp(Trim$(tmISCIXRef(llIndex).sCallLetters), sCallLetters, vbTextCompare) = 0) And (tmISCIXRef(llIndex).iEmbedded = imEmbedded) Then
                        iFound = True
                        Exit For
                    End If
                Next llIndex
                If Not iFound Then
                    'llUpper = UBound(tmISCIXRef)
                    tmISCIXRef(lmXRefUpper).iCommProvArfCode = imCommProvArfCode
                    tmISCIXRef(lmXRefUpper).iEmbedded = imEmbedded
                    tmISCIXRef(lmXRefUpper).sVehName = smVefName
                    tmISCIXRef(lmXRefUpper).sCallLetters = sCallLetters
                    tmISCIXRef(lmXRefUpper).iVefCode = imVefCode
                    tmISCIXRef(lmXRefUpper).iShfCode = iShfCode
                    If imEmbedded Then
                        tmISCIXRef(lmXRefUpper).sKey = "!" & smCommProvArfCode & "|" & tmISCIXRef(lmXRefUpper).sVehName & "|" & tmISCIXRef(lmXRefUpper).sCallLetters
                    Else
                        tmISCIXRef(lmXRefUpper).sKey = smCommProvArfCode & "|" & tmISCIXRef(lmXRefUpper).sVehName & "|" & tmISCIXRef(lmXRefUpper).sCallLetters
                    End If
                    'ReDim Preserve tmISCIXRef(0 To llUpper + 1) As ISCIXREF
                    If lmXRefUpper >= UBound(tmISCIXRef) Then
                        ReDim Preserve tmISCIXRef(0 To lmXRefUpper + 20000) As ISCIXREF
                    End If
                    lmXRefUpper = lmXRefUpper + 1
                End If
                Print #hmMsg, "Gather ISCI for: " & Trim$(smVefName) & " on " & sCallLetters & " at " & Format$(Now, "m/d/yyyy " & sgShowTimeWSecForm) & " by " & Trim$(sgUserName)
                lacResult.Caption = "Gather ISCI for: " & Trim$(smVefName) & " on " & sCallLetters
                If igExportSource = 2 Then DoEvents
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cprst!cpttCode
                tgCPPosting(0).iStatus = cprst!cpttStatus
                tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cprst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cprst!attTimeType
                tgCPPosting(0).iVefCode = imVefCode
                tgCPPosting(0).iShttCode = cprst!shttCode
                tgCPPosting(0).sZone = cprst!shttTimeZone
                tgCPPosting(0).sDate = Format$(sMoDate, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                'Create AST records
                'slACName = Trim$(cprst!attACName)
 '               llPrevFeedDate = -1
'                If ilNoAirDays > 1 Then
'                    ilFirstDate = True
'                End If
'                ilNoAirDays = 0
 '               ilHourNo = 0
'                ilSegmentNo = 0
'                ilBreakNo = 0
'                ilPositionNo = 0
'                llPrevFeedTime = -1
                igTimes = 1 'By Week
                imAdfCode = -1
'                ilVff = gBinarySearchVff(imVefCode)
'                If ilVff <> -1 Then
''                    'If Trim$(tgVffInfo(ilVff).sXDXMLForm) = "P" Then
''                    '    slProgCode = ""
''                    'Else
''                        slProgCode = Trim$(tgVffInfo(ilVff).sXDProgCodeID)
''                    'End If
'                Else
'                    slProgCode = ""
'                End If
                If igExportSource = 2 Then DoEvents
                iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True)
                gFilterAstExtendedTypes tmAstInfo
                'Output AST
                For iLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    blSpotOk = True
                    ilAnf = gBinarySearchAnf(tmAstInfo(iLoop).iAnfCode)
                    If ilAnf <> -1 Then
                        If tgAvailNamesInfo(ilAnf).sISCIExport = "N" Then
                            blSpotOk = False
                        End If
                    End If
                    If blSpotOk Then
                        lFeedDate = DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate))
                        'Bypass duplicates
                        ilDuplicateSpot = False
                        If iLoop <> LBound(tmAstInfo) Then
                            llFeedTime = gTimeToLong(tmAstInfo(iLoop).sFeedTime, False)
                            For ilDuptTest = iLoop - 1 To LBound(tmAstInfo) Step -1
                                If igExportSource = 2 Then DoEvents
                                If (DateValue(gAdjYear(tmAstInfo(ilDuptTest).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(ilDuptTest).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(ilDuptTest).iPledgeStatus)).iPledged <> 2) Then
                                    llDuplTestTime = gTimeToLong(tmAstInfo(ilDuptTest).sFeedTime, False)
                                    If llFeedTime <> llDuplTestTime Then
                                        Exit For
                                    End If
                                    If tmAstInfo(ilDuptTest).lSdfCode = tmAstInfo(iLoop).lSdfCode Then
                                        ilDuplicateSpot = True
                                        Exit For
                                    End If
                                End If
                            Next ilDuptTest
                        End If
                        If (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iPledgeStatus)).iPledged <> 2) And (Not ilDuplicateSpot) Then
'                            If llPrevFeedDate <> lFeedDate Then
'                                ilNoAirDays = ilNoAirDays + 1
'                                ilHourNo = 0
'                                ilSegmentNo = 0
'                                ilBreakNo = 0
'                                ilPositionNo = 0
'                                llPrevFeedTime = -1
'                                slYear = right$(Year(gObtainNextSunday(gAdjYear(tmAstInfo(iLoop).sFeedDate))), 2)
'                                llStartYearDate = DateValue(gObtainYearStartDate(gAdjYear(tmAstInfo(iLoop).sFeedDate)))
'                                slWeekNo = Trim$(Str$((lFeedDate - llStartYearDate) \ 7 + 1))
'                                If Len(slWeekNo) = 1 Then
'                                    slWeekNo = "0" & slWeekNo
'                                End If
'                                slDay = Trim$(Str$(Weekday(gAdjYear(tmAstInfo(iLoop).sFeedDate), vbMonday)))
'                            End If
'                            llPrevFeedDate = lFeedDate
                            llFeedTime = gTimeToLong(tmAstInfo(iLoop).sFeedTime, False)
                            'dan m 6/4/15 not used
'                            If llPrevFeedTime <> llFeedTime Then
'                                If llPrevFeedTime <> -1 Then
'                                    If Hour(gLongToTime(llPrevFeedTime)) <> Hour(gLongToTime(llFeedTime)) Then
'                                        ilHourNo = ilHourNo + 1
'                                        ilSegmentNo = 0
'                                        ilBreakNo = 0
'                                        ilPositionNo = 0
'                                    End If
'                                Else
'                                    ilHourNo = 1
'                                End If
'                                ilSegmentNo = ilSegmentNo + 1
'                                ilBreakNo = ilBreakNo + 1
'                                ilPositionNo = 1
'                                ilStartIndexOfBreak = iLoop
'                            Else
'                                ilPositionNo = ilPositionNo + 1
'                            End If
'                            llPrevFeedTime = llFeedTime
                            sISCI = Trim$(tmAstInfo(iLoop).sISCI)
                            llAdf = gBinarySearchAdf(CLng(tmAstInfo(iLoop).iAdfCode))
                            If llAdf <> -1 Then
                                sAdvt = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                            Else
                                sAdvt = "Missing" & Str(tmAstInfo(iLoop).iAdfCode)
                            End If
                            
                            ''6/12/06- Check if any region copy defined for the spots
                            ''ilRet = gGetRegionCopy(tmAstInfo(iLoop).iShttCode, tmAstInfo(iLoop).lSdfCode, tmAstInfo(iLoop).iVefCode, sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                            'ilRet = gGetRegionCopy(tmAstInfo(iLoop), sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                            'If ilRet Then
                            If tmAstInfo(iLoop).iRegionType > 0 Then
                                sISCI = Trim$(tmAstInfo(iLoop).sRISCI) 'sRISCI
                            End If
                            
                            If igExportSource = 2 Then DoEvents
                            If InStr(1, sAdvt, "Missing", vbTextCompare) = 1 Then
                                Print #hmMsg, Trim$(smVefName) & ": Advertiser " & sAdvt & " on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                                lbcMsg.AddItem Trim$(smVefName) & ": Advertiser " & sAdvt & " on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                            Else
                                If sISCI <> "" Then
                                    iFound = False
                                    tmISCISendInfo(lmSendInfoUpper).sISCI = sISCI
                                    If imEmbedded Then
                                        tmISCISendInfo(lmSendInfoUpper).sCallLetters = smProducerName
                                        slKey = "!" & smCommProvArfCode & "|" & tmISCISendInfo(lmSendInfoUpper).sCallLetters & "|" & tmISCISendInfo(lmSendInfoUpper).sISCI
                                    Else
                                        tmISCISendInfo(lmSendInfoUpper).sCallLetters = sCallLetters
                                        slKey = smCommProvArfCode & "|" & tmISCISendInfo(lmSendInfoUpper).sCallLetters & "|" & tmISCISendInfo(lmSendInfoUpper).sISCI
                                    End If
                                    iFound = mBinarySearchSendInfo(0, lmSendInfoUpper - 1, slKey, tmISCISendInfo())
                                    If Not iFound Then
                                        'llUpper = UBound(tmISCISendInfo)
                                        tmISCISendInfo(lmSendInfoUpper).iCommProvArfCode = imCommProvArfCode
                                        tmISCISendInfo(lmSendInfoUpper).iEmbedded = imEmbedded
                                        tmISCISendInfo(lmSendInfoUpper).sVehName = smVefName
                                        tmISCISendInfo(lmSendInfoUpper).sISCI = sISCI
                                        If imEmbedded Then
                                            tmISCISendInfo(lmSendInfoUpper).sCallLetters = smProducerName
                                            tmISCISendInfo(lmSendInfoUpper).iShfCode = 0
                                            tmISCISendInfo(lmSendInfoUpper).iProducerArfCode = imProducerArfCode
                                        Else
                                            tmISCISendInfo(lmSendInfoUpper).sCallLetters = sCallLetters
                                            tmISCISendInfo(lmSendInfoUpper).iShfCode = iShfCode
                                            tmISCISendInfo(lmSendInfoUpper).iProducerArfCode = 0
                                        End If
                                        tmISCISendInfo(lmSendInfoUpper).sAdvtName = sAdvt
                                        tmISCISendInfo(lmSendInfoUpper).lLatestFeedDate = lFeedDate
                                        tmISCISendInfo(lmSendInfoUpper).lEarlestFeedDate = lFeedDate
                                        If imEmbedded Then
                                            tmISCISendInfo(lmSendInfoUpper).sKey = "!" & smCommProvArfCode & "|" & tmISCISendInfo(lmSendInfoUpper).sCallLetters & "|" & tmISCISendInfo(lmSendInfoUpper).sISCI
                                        Else
                                            tmISCISendInfo(lmSendInfoUpper).sKey = smCommProvArfCode & "|" & tmISCISendInfo(lmSendInfoUpper).sCallLetters & "|" & tmISCISendInfo(lmSendInfoUpper).sISCI
                                        End If
                                        If ((udcCriteria.ISpots(0)) Or (udcCriteria.ISpots(2))) And (udcCriteria.IVeh = True) Then
                                            SQLQuery = "SELECT eitDateSent, eitLLDRef, eitLLDSent, eitCode"
                                            SQLQuery = SQLQuery & " FROM EIT"
                                            If imEmbedded Then
                                                SQLQuery = SQLQuery + " WHERE eitISCI ='" & sISCI & "' And eitType = 'E'" & " And eitshfCode = 0" & " And eitCommProvArfCode = " & imCommProvArfCode
                                            Else
                                                SQLQuery = SQLQuery + " WHERE eitISCI ='" & sISCI & "' And eitType = 'S'" & " And eitshfCode = " & iShfCode & " And eitCommProvArfCode = " & imCommProvArfCode
                                            End If
                                            Set rst = gSQLSelectCall(SQLQuery)
                                            If Not rst.EOF Then
                                                tmISCISendInfo(lmSendInfoUpper).lEitCode = rst!eitCode
                                                tmISCISendInfo(lmSendInfoUpper).lDateSent = DateValue(gAdjYear(rst!eitDateSent))
                                                tmISCISendInfo(lmSendInfoUpper).lLLDRef = DateValue(gAdjYear(rst!eitLLDRef))
                                                tmISCISendInfo(lmSendInfoUpper).lLLDSent = DateValue(gAdjYear(rst!eitLLDSent))
                                            Else
                                                tmISCISendInfo(lmSendInfoUpper).lEitCode = 0
                                                tmISCISendInfo(lmSendInfoUpper).lLLDRef = 0
                                                tmISCISendInfo(lmSendInfoUpper).lLLDSent = 0
                                            End If
                                        Else
                                            tmISCISendInfo(lmSendInfoUpper).lEitCode = 0
                                            tmISCISendInfo(lmSendInfoUpper).lLLDRef = 0
                                            tmISCISendInfo(lmSendInfoUpper).lLLDSent = 0
                                        End If
                                        tmISCISendInfo(lmSendInfoUpper).iUpdateDateSent = False
                                        tmISCISendInfo(lmSendInfoUpper).iVefCode = imVefCode
                                        'ReDim Preserve tmISCISendInfo(0 To llUpper + 1) As ISCISENDINFO
                                        If lmSendInfoUpper >= UBound(tmISCISendInfo) Then
                                            ReDim Preserve tmISCISendInfo(0 To lmSendInfoUpper + 20000) As ISCISENDINFO
                                            For llSetKey = lmSendInfoUpper + 1 To UBound(tmISCISendInfo) Step 1
                                                tmISCISendInfo(llSetKey).sKey = "~~~~~~~~~~"
                                            Next llSetKey
                                        End If
                                        lmSendInfoUpper = lmSendInfoUpper + 1
                                        If lmSendInfoUpper - 1 >= 0 Then
                                            ArraySortTyp fnAV(tmISCISendInfo(), 0), lmSendInfoUpper, 0, LenB(tmISCISendInfo(0)), 0, LenB(tmISCISendInfo(0).sKey), 0
                                        End If
                                    End If
                                Else
                                    Print #hmMsg, Trim$(smVefName) & " " & sAdvt & ": ISCI Missing on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                                    lbcMsg.AddItem Trim$(smVefName) & " " & sAdvt & ": ISCI Missing on " & Format$(tmAstInfo(iLoop).sFeedDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sFeedTime, "h:mm:ssAM/PM")
                                End If
                            End If
                        End If
                    End If
                Next iLoop
            End If
            cprst.MoveNext
        Wend
        sMoDate = DateAdd("d", 7, sMoDate)
    Loop While DateValue(gAdjYear(sMoDate)) < DateValue(gAdjYear(sEndDate))

    mGatherISCI_CompressedFormat = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mGatherISCI_CompressedFormat"
    mGatherISCI_CompressedFormat = False
    Exit Function
    
End Function

Private Sub tmcFilterDelay_Timer()
    tmcFilterDelay.Enabled = False
    mFillStations
End Sub


Private Sub mFillStations(Optional blFillStations As Boolean = False)

    Dim llRow As Long
    Dim ilFound As Integer
    Dim ilFilterIdx As Integer
    Dim slTemp As String
    Dim ilvehicle As Integer
    Dim ilVefCode As Integer

    On Error GoTo ErrHand
    '11/8/19: Fill only during export and all unchecked
    If (rbcFilter(4).Value = True) And (blFillStations = False) And (chkAllStation.Value = vbChecked) Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    lbcStation.Clear
    For ilvehicle = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 1 Then DoEvents
        If lbcVehicles.Selected(ilvehicle) Then
            ilVefCode = lbcVehicles.ItemData(ilvehicle)
            SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode, shttState, shttMarket, shttFmtCode, shttMetCode"
            SQLQuery = SQLQuery + " FROM shtt, att"
            SQLQuery = SQLQuery + " WHERE (attVefCode = " & ilVefCode
            SQLQuery = SQLQuery + " AND shttCode = attShfCode)"
            SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
            Set rst = gSQLSelectCall(SQLQuery)
            While Not rst.EOF
                '11/8/19: Replaced SendMessage call with BinarySearch which is much faster
                'have we already added the call letters?
                'llRow = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, Trim$(rst!shttCallLetters))
                llRow = mBinarySearchStation(Trim$(rst!shttCallLetters))

                ilFound = 1
                If llRow < 0 Then
                    If rbcFilter(0).Value = True Then   'DMA
                        For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                            If lbcFilter.Selected(ilFilterIdx) Then
                                If Trim(lbcFilter.List(ilFilterIdx)) = Trim$(rst!shttMarket) Then
                                    lbcStation.AddItem Trim$(rst!shttCallLetters)
                                    lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                End If
                            End If
                        Next ilFilterIdx
                    End If
                    If rbcFilter(1).Value = True Then    'Format
                        slTemp = mGetFormat(rst!shttFmtCode)
                        For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                            If lbcFilter.Selected(ilFilterIdx) Then
                                If Trim(lbcFilter.List(ilFilterIdx)) = Trim(slTemp) Then
                                    lbcStation.AddItem Trim$(rst!shttCallLetters)
                                    lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                End If
                            End If
                        Next ilFilterIdx
                    End If
                    If rbcFilter(2).Value = True Then   'MSA
                        slTemp = mGetMSA(rst!shttMetCode)
                        For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                            If lbcFilter.Selected(ilFilterIdx) Then
                                If Trim(lbcFilter.List(ilFilterIdx)) = Trim$(slTemp) Then
                                    lbcStation.AddItem Trim$(rst!shttCallLetters)
                                    lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                End If
                            End If
                        Next ilFilterIdx
                    End If
                    If rbcFilter(3).Value = True Then    'State
                        For ilFilterIdx = 0 To lbcFilter.ListCount - 1
                            If lbcFilter.Selected(ilFilterIdx) Then
                                If Left(lbcFilter.List(ilFilterIdx), 2) = Trim$(rst!shttState) Then
                                    lbcStation.AddItem Trim$(rst!shttCallLetters)
                                    lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                                End If
                            End If
                        Next ilFilterIdx
                    End If
                    If rbcFilter(4).Value = True Then   'Stations
                        lbcStation.AddItem Trim$(rst!shttCallLetters)
                        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
                    End If
                End If
                rst.MoveNext
            Wend
        End If
    Next ilvehicle
'    chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    chkAllStation_Click
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mFillStations"
End Sub

Private Function mCheckLastExportDate() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim slDate As String
    Dim ilLoop As Integer
    'Dim slFields(1 To 15) As String
    Dim slFields(0 To 14) As String
    
    'slFromFile = txtFile.Text
    'ilRet = 0
    'On Error GoTo mCheckLastExportDateErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        If udcCriteria.ISpots(0) Then
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
                    On Error GoTo mCheckLastExportDateErr
                    If slDate = "" Then
                        'slDate = slFields(5)
                        slDate = slFields(4)
                    Else
                        'If DateValue(gAdjYear(slFields(5))) > DateValue(gAdjYear(slDate)) Then
                        If DateValue(gAdjYear(slFields(4))) > DateValue(gAdjYear(slDate)) Then
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
    gHandleError "AffErrorLog.txt", "Export ISCI-mCheckLastExportDate"
End Function


Private Sub rbcFilter_Click(Index As Integer)

    Dim ilRet As Integer
    
    If rbcFilter(Index).Value Then
        Select Case Index
            Case 0
                mPopDMA
                smFilterType = "DMA"
            Case 1
                mPopFormat
                smFilterType = "Format"
            Case 2
                mPopMSA
                smFilterType = "MSA"
            Case 3
                mPopState
                smFilterType = "State"
            Case 4
                smFilterType = "Station"
        End Select
        mFillStations
    End If
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmExportISCI
End Sub

Private Sub txtNumberDays_Change()
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
        plcGauge.Visible = False
    End If
End Sub

Private Function mExportISCI_CompressedFormat() As Integer
    Dim llXRef As Long
    Dim iISCI As Integer
    Dim iNewFile As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim sNowDate As String
    Dim sName As String
    Dim sToFile As String
    Dim sSDate As String
    Dim sEDate As String
    Dim llEDate As Long
    Dim iPrintISCI As Integer
    Dim llISCIInfo As Long
    Dim sDateTime As String
    Dim sISCI As String
    Dim sCallLetters As String
    Dim lTestLLDSent As Long
    Dim ilRet As Integer
    
    sNowDate = Format$(smNowDate, "mmddyy")
    sSDate = Format$(smDate, "mm/dd/yy")
    sEDate = Format$(DateAdd("d", imNumberDays - 1, smDate), "mm/dd/yy")
    llEDate = DateValue(gAdjYear(sEDate))
    iNewFile = True
    llXRef = LBound(tmISCIXRef)
    llISCIInfo = LBound(tmISCISendInfo)
    
    Do While llXRef < UBound(tmISCIXRef)
        If igExportSource = 2 Then DoEvents
        If iNewFile Then
            'Create file
            On Error GoTo ErrHand:
            SQLQuery = "SELECT arfID"
            SQLQuery = SQLQuery & " FROM ARF_Addresses"
            SQLQuery = SQLQuery + " WHERE arfCode = " & tmISCIXRef(llXRef).iCommProvArfCode
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                sName = Trim$(rst!arfID)
            Else
                sName = "Missed"
            End If
            If tmISCIXRef(llXRef).iEmbedded Then
                sToFile = smExportToPath & sName & "_E_" & sNowDate & ".txt"
            Else
                sToFile = smExportToPath & sName & "_" & sNowDate & ".txt"
            End If
            'On Error GoTo mExportISCI_CompressedFormatErr:
            iRet = 0
            'sDateTime = FileDateTime(sToFile)
            ilRet = gFileExist(sToFile)
            If iRet = 0 Then
                Kill sToFile
                If iRet <> 0 Then
                    Close hmTo
                    gMsgBox "Kill File " & sToFile & " error#" & Str$(Err.Number), vbOKOnly
                    mExportISCI_CompressedFormat = False
                    Exit Function
                End If
            End If
            'iRet = 0
            'hmTo = FreeFile
            'Open sToFile For Output As hmTo
            iRet = gFileOpen(sToFile, "Output", hmTo)
            If iRet <> 0 Then
                Close hmTo
                gMsgBox "Open File " & sToFile & " error#" & Str$(Err.Number), vbOKOnly
                mExportISCI_CompressedFormat = False
                Exit Function
            End If
            Print #hmMsg, "** Storing Output into " & sToFile & " **"
            If udcCriteria.ISpots(1) Then
                Print #hmTo, "TResend"
            ElseIf udcCriteria.ISpots(2) Then
                Print #hmTo, "TAll"
            Else
                Print #hmTo, "TUnsent"
            End If
            Print #hmTo, "A" & sSDate & "-" & sEDate
        End If
        If igExportSource = 2 Then DoEvents
        'Process records
        If llXRef > LBound(tmISCIXRef) Then
            If tmISCIXRef(llXRef - 1).iVefCode <> tmISCIXRef(llXRef).iVefCode Then
                Print #hmTo, "B" & tmISCIXRef(llXRef).sVehName
            End If
        Else
            Print #hmTo, "B" & tmISCIXRef(llXRef).sVehName
        End If
        Print #hmTo, "C" & tmISCIXRef(llXRef).sCallLetters
        If igExportSource = 2 Then DoEvents
        'Test if file should be closed
        iNewFile = False
        iPrintISCI = False
        If llXRef + 1 < UBound(tmISCIXRef) Then
            If tmISCIXRef(llXRef).iCommProvArfCode <> tmISCIXRef(llXRef + 1).iCommProvArfCode Then
                'Output ISCI Info
                iNewFile = True
                iPrintISCI = True
            End If
        Else
            'Output ISCI Info
            iPrintISCI = True
        End If
        If iPrintISCI Then
            Do While (tmISCISendInfo(llISCIInfo).iCommProvArfCode = tmISCIXRef(llXRef).iCommProvArfCode) And (llISCIInfo < UBound(tmISCISendInfo))
                If igExportSource = 2 Then DoEvents
                iPrintISCI = False
                If tmISCISendInfo(llISCIInfo).lEitCode = 0 Then
                    iPrintISCI = True
                Else
                    If udcCriteria.ISpots(0) Then   'Unsent
                        If tmISCISendInfo(llISCIInfo).lLLDRef + imNoDaysResendISCI <= tmISCISendInfo(llISCIInfo).lEarlestFeedDate Then
                            iPrintISCI = True
                        End If
                    ElseIf udcCriteria.ISpots(1) Then   'Resend- Logic never completed
'                        If tmISCISendInfo(llISCIInfo).lLLDRef >= DateValue(gAdjYear(sSDate)) Then
'                            iPrintISCI = True
'                        End If
                        lTestLLDSent = tmISCISendInfo(llISCIInfo).lLLDSent
                        Do
                            If igExportSource = 2 Then DoEvents
                            If (lTestLLDSent >= DateValue(gAdjYear(sSDate))) And (lTestLLDSent <= DateValue(gAdjYear(sEDate))) Then
                                iPrintISCI = True
                                Exit Do
                            End If
                            lTestLLDSent = lTestLLDSent - imNoDaysResendISCI
                        Loop While lTestLLDSent > DateValue(gAdjYear(sSDate))
                    Else    'All
                        iPrintISCI = True
                    End If
                End If
                sCallLetters = tmISCISendInfo(llISCIInfo).sCallLetters
                If InStr(1, sCallLetters, "!", vbTextCompare) = 1 Then
                    sCallLetters = Mid(sCallLetters, 2)
                End If
                sCallLetters = Trim$(sCallLetters)
                If iPrintISCI Then
                    tmISCISendInfo(llISCIInfo).iUpdateDateSent = True
                    sISCI = Trim$(gRemoveChar(tmISCISendInfo(llISCIInfo).sISCI, Chr$(9)))
                    If llISCIInfo > LBound(tmISCISendInfo) Then
                        If tmISCISendInfo(llISCIInfo - 1).iShfCode <> tmISCISendInfo(llISCIInfo).iShfCode Then
                            Print #hmTo, "D" & sCallLetters
                        End If
                    Else
                        Print #hmTo, "D" & sCallLetters
                    End If
                    Print #hmTo, "E" & sISCI & Chr$(9) & mReplaceTab(Trim$(tmISCISendInfo(llISCIInfo).sAdvtName))
                Else
                    tmISCISendInfo(llISCIInfo).iUpdateDateSent = False
                    If llISCIInfo > LBound(tmISCISendInfo) Then
                        If tmISCISendInfo(llISCIInfo - 1).iShfCode <> tmISCISendInfo(llISCIInfo).iShfCode Then
                            Print #hmTo, "D" & sCallLetters
                        End If
                    Else
                        Print #hmTo, "D" & sCallLetters
                    End If
                End If
                llISCIInfo = llISCIInfo + 1
            Loop
        End If
        If iNewFile Then
            Close #hmTo
        End If
        llXRef = llXRef + 1
    Loop
    Close #hmTo
    mExportISCI_CompressedFormat = True
    Exit Function
'mExportISCI_CompressedFormatErr:
'    iRet = 1
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mExportISCI_CompressedFormat"
    mExportISCI_CompressedFormat = False
    Exit Function
End Function

Private Function mUpdateEIT() As Integer
    Dim iLoop As Integer
    Dim sNowDate As String
    Dim sISCI As String
    Dim sType As String
    Dim sAddStr As String
    
    On Error GoTo ErrHand:
    sNowDate = Format$(smNowDate, "mm/dd/yy")
    For iLoop = LBound(tmISCISendInfo) To UBound(tmISCISendInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tmISCISendInfo(iLoop).lEitCode = 0 Then
            sISCI = Trim$(gFixQuote(tmISCISendInfo(iLoop).sISCI))
            If InStr(1, tmISCISendInfo(iLoop).sCallLetters, "!", vbTextCompare) = 1 Then
                sType = "E"
            Else
                sType = "S"
            End If
            SQLQuery = "INSERT INTO eit (eitType, eitShfCode, eitCommProvArfCode, eitISCI, eitDateSent, eitLLDRef, eitLLDSent)"
            SQLQuery = SQLQuery & " VALUES ('" & sType & "', " & tmISCISendInfo(iLoop).iShfCode & ", " & tmISCISendInfo(iLoop).iCommProvArfCode & ", '" & sISCI & "', '"
            SQLQuery = SQLQuery & Format$(sNowDate, sgSQLDateForm) & "','" & Format$(tmISCISendInfo(iLoop).lLatestFeedDate, sgSQLDateForm) & "','" & Format$(tmISCISendInfo(iLoop).lLatestFeedDate, sgSQLDateForm) & "')"
            sAddStr = "."
        Else
            sAddStr = ""
            SQLQuery = "UPDATE eit SET "
            If (tmISCISendInfo(iLoop).lLLDRef < tmISCISendInfo(iLoop).lLatestFeedDate) Then
                SQLQuery = SQLQuery + "eitLLDRef = '" & Format$(tmISCISendInfo(iLoop).lLatestFeedDate, sgSQLDateForm) & "'"
                sAddStr = ", "
            End If
            If tmISCISendInfo(iLoop).iUpdateDateSent Then
                If (tmISCISendInfo(iLoop).lLLDSent < tmISCISendInfo(iLoop).lLatestFeedDate) Then
                    SQLQuery = SQLQuery + sAddStr + "eitLLDSent = '" & Format$(tmISCISendInfo(iLoop).lLatestFeedDate, sgSQLDateForm) & "'"
                    sAddStr = ", "
                End If
                If tmISCISendInfo(iLoop).lDateSent < DateValue(gAdjYear(sNowDate)) Then
                    SQLQuery = SQLQuery + sAddStr + "eitDateSent = '" & Format$(sNowDate, sgSQLDateForm) & "'"
                    sAddStr = ", "
                End If
            End If
            SQLQuery = SQLQuery + " WHERE eitCode = " & tmISCISendInfo(iLoop).lEitCode & ""
        End If
        If sAddStr <> "" Then
            If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                'mUpdateEIT = False
                'Exit Function
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "Export IDC-mUpdateEIT"
                mUpdateEIT = False
                Exit Function
            End If
        End If
    Next iLoop
    mUpdateEIT = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mUpdateEIT"
    mUpdateEIT = False
    Exit Function
End Function

Private Sub txtNumberDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


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
Private Function mOpenISCIbyBreakFile(iiCommProvArfCode As Integer) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim slNowDate As String
    
    On Error GoTo ErrHand:
    slNowDate = Format$(smNowDate, "mmddyy")
    SQLQuery = "SELECT arfID"
    SQLQuery = SQLQuery & " FROM ARF_Addresses"
    SQLQuery = SQLQuery + " WHERE arfCode = " & iiCommProvArfCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        slName = Trim$(rst!arfID)
    Else
        slName = "Missed"
    End If
    slToFile = smExportToPath & slName & "_" & slNowDate & "_L" & ".txt"
    'On Error GoTo mOpenISCIbyBreakFileErr:
    ilRet = 0
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        If ilRet <> 0 Then
            Close hmISCIbyBreak
            gMsgBox "Kill File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenISCIbyBreakFile = False
            Exit Function
        End If
    End If
    'ilRet = 0
    'hmISCIbyBreak = FreeFile
    'Open slToFile For Output As hmISCIbyBreak
    ilRet = gFileOpen(slToFile, "Output", hmISCIbyBreak)
    If ilRet <> 0 Then
        Close hmISCIbyBreak
        gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        mOpenISCIbyBreakFile = False
        Exit Function
    End If
    mOpenISCIbyBreakFile = True
    Exit Function
'mOpenISCIbyBreakFileErr:
'    ilRet = 1
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export ISCI-mOpenISCIByBreakFile"
    mOpenISCIbyBreakFile = False
    Exit Function
End Function

Private Function mBinarySearchSendInfo(llInMin As Long, llInMax As Long, slKey, tlISCISendInfo() As ISCISENDINFO) As Integer
    Dim llMiddle As Long
    Dim slCode As String
    Dim llResult As Long
    Dim ilPass As Integer
    Dim llMin As Long
    Dim llMax As Long
    
    mBinarySearchSendInfo = False
    slCode = UCase(Trim$(slKey))
    
    On Error GoTo ErrHand
    For ilPass = 0 To 1 Step 1
        llMin = llInMin
        llMax = llInMax
        Do While llMin <= llMax
            If igExportSource = 2 Then DoEvents
            llMiddle = (llMin + llMax) \ 2
            'llResult = StrComp(Trim(tmISCISendInfo(llMiddle).sKey), Trim$(slKey), vbTextCompare)
            'llResult = StrComp(Trim(tmISCISendInfo(llMiddle).sKey), Trim$(slKey), vbBinaryCompare)
            'vbTestCompare was failing for unknown reasons!!!
            If ilPass = 0 Then
                llResult = StrComp(UCase(Trim(tlISCISendInfo(llMiddle).sKey)), slCode, vbTextCompare)
            Else
                llResult = StrComp(UCase(Trim(tlISCISendInfo(llMiddle).sKey)), slCode, vbBinaryCompare)
            End If
            Select Case llResult
                Case 0:
                    mBinarySearchSendInfo = True
                    Exit Function
                Case 1:
                    llMax = llMiddle - 1
                Case -1:
                    llMin = llMiddle + 1
            End Select
        Loop
    Next ilPass
    Exit Function
ErrHand:
    Exit Function
End Function


Private Function mExportUniqueISCI() As Integer
    Dim ilShtt As Integer
    Dim ilOkStation As Integer
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slSDate As String
    Dim llSDate As Long
    Dim slEDate As String
    Dim llEDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim ilAppendToFile As Integer
    Dim llSetKey As Integer
    Dim ilRet As Integer
    Dim ilVefSelectedCount As Integer
    ReDim illbcVehicleIndex(0 To 0) As Integer
    Dim ilTest As Integer
    Dim ilTest2 As Integer
    Dim ilFound As Integer
    Dim sMoDate As String
    Dim ilVefLoop As Integer
    ReDim ilShttCode(0 To 0) As Integer
    
    mExportUniqueISCI = False
    '5/8/14: Moved here
    ReDim imFinalAlertVefCode(0 To 0) As Integer
    ilAppendToFile = False
    ilVefSelectedCount = 0
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            ilVefSelectedCount = ilVefSelectedCount + 1
            illbcVehicleIndex(UBound(illbcVehicleIndex)) = ilVef
            ReDim Preserve illbcVehicleIndex(0 To UBound(illbcVehicleIndex) + 1) As Integer
        End If
    Next ilVef
    sMoDate = gObtainPrevMonday(smDate)
    lmTotalNumber = UBound(illbcVehicleIndex)
    For ilVefLoop = 0 To UBound(illbcVehicleIndex) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        ReDim ilShttCode(0 To 0) As Integer
        ilVef = illbcVehicleIndex(ilVefLoop)
        imVefCode = lbcVehicles.ItemData(ilVef)
        SQLQuery = "SELECT DISTINCT shttCallLetters, cpttShfCode"
        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
        '10/29/14: Bypass Service agreements
        SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
        SQLQuery = SQLQuery + " AND ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " Order by shttCallLetters"
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            If igExportSource = 2 Then DoEvents
            ilFound = False
            For ilTest2 = 0 To UBound(ilShttCode) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If ilShttCode(ilTest2) = cprst!cpttshfcode Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest2
            If Not ilFound Then
                ilShttCode(UBound(ilShttCode)) = cprst!cpttshfcode
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
            cprst.MoveNext
        Wend
        'For ilShtt = LBound(tgStationInfoByCode) To UBound(tgStationInfoByCode) - 1 Step 1
        For ilTest2 = 0 To UBound(ilShttCode) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            ilShtt = gBinarySearchStationInfoByCode(ilShttCode(ilTest2))
            If ilShtt <> -1 Then
                ilOkStation = True
            Else
                ilOkStation = False
            End If
            If igExportSource = 2 Then DoEvents
            If ilOkStation Then
                lacResult.Caption = "Checking: " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                If igExportSource = 2 Then DoEvents
                If lbcStation.ListCount > 0 Then
                    ilOkStation = False
                    For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        If lbcStation.Selected(ilLoop) Then
                            If lbcStation.ItemData(ilLoop) = tgStationInfoByCode(ilShtt).iCode Then
                                ilOkStation = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                Else
                    ilOkStation = True
                End If
            End If
            If ilOkStation Then
                imShttCode = tgStationInfoByCode(ilShtt).iCode
                If ((udcCriteria.ISpots(0)) Or (udcCriteria.ISpots(2))) And (udcCriteria.iExportType(0)) And (udcCriteria.IVeh = True) Then
                Else
                    ReDim tmISCISendInfo(0 To 20000) As ISCISENDINFO
                    For llSetKey = 0 To UBound(tmISCISendInfo) Step 1
                        tmISCISendInfo(llSetKey).sKey = "~~~~~~~~~~"
                    Next llSetKey
                    lmSendInfoUpper = 0
                End If
                ilVef = illbcVehicleIndex(ilVefLoop)
                imVefCode = lbcVehicles.ItemData(ilVef)
                smVefName = Trim$(lbcVehicles.List(ilVef))
                lacResult.Caption = "Checking: " & Trim$(tgStationInfoByCode(ilShtt).sCallLetters) & " on " & smVefName
                smCommProvArfCode = ""
                imCommProvArfCode = 0
                imEmbedded = False
                If udcCriteria.IUniqueBy(0) = True Then
                    imEmbedded = True
                End If
                smProducerName = "!Missing Producer "
                imProducerArfCode = 0
                smProducerArfCode = Trim$(Str$(imProducerArfCode))
                Do While Len(smProducerArfCode) < 5
                    smProducerArfCode = "0" & smProducerArfCode
                Loop
                ilIndex = gBinarySearchVef(CLng(imVefCode))
                If ilIndex <> -1 Then
                    imCommProvArfCode = tgVehicleInfo(ilIndex).iCommProvArfCode
                    smCommProvArfCode = Trim$(Str$(tgVehicleInfo(ilIndex).iCommProvArfCode))
                    Do While Len(smCommProvArfCode) < 5
                        smCommProvArfCode = "0" & smCommProvArfCode
                    Loop
                    imProducerArfCode = tgVehicleInfo(ilIndex).iProducerArfCode
                    smProducerArfCode = Trim$(Str$(imProducerArfCode))
                    Do While Len(smProducerArfCode) < 5
                        smProducerArfCode = "0" & smProducerArfCode
                    Loop
    
                    If (tgVehicleInfo(ilIndex).iProducerArfCode > 0) And (imEmbedded) Then
                        imEmbedded = True
                        SQLQuery = "SELECT arfName"
                        SQLQuery = SQLQuery & " FROM ARF_Addresses"
                        SQLQuery = SQLQuery + " WHERE arfCode = " & tgVehicleInfo(ilIndex).iProducerArfCode
                        Set rst = gSQLSelectCall(SQLQuery)
                        If Not rst.EOF Then
                            smProducerName = "!" & Trim$(rst!arfName)
                        End If
                    Else
                        imEmbedded = False
                    End If
        
                    If ((udcCriteria.ISpots(0)) Or (udcCriteria.ISpots(2))) And (udcCriteria.iExportType(0)) And (udcCriteria.IVeh = True) Then
                        ilRet = mGatherISCI_CompressedFormat()
                    Else
                        ilRet = mGatherISCI_AllFormat()
                        If (ilRet = True) And (Not imTerminate) And (lmSendInfoUpper > 0) Then
                            ilRet = mExportISCI_AllFormat(lmSendInfoUpper, tmISCISendInfo())
                            ilAppendToFile = True
                            ReDim tmISCISendInfo(0 To 20000) As ISCISENDINFO
                            For llSetKey = 0 To UBound(tmISCISendInfo) Step 1
                                tmISCISendInfo(llSetKey).sKey = "~~~~~~~~~~"
                            Next llSetKey
                            lmSendInfoUpper = 0
                        End If
                    End If
                    If igExportSource = 2 Then DoEvents
                    If (ilRet = False) Then
                        gCloseRegionSQLRst
                        Print #hmMsg, "** Terminated **"
                        'Close #hmMsg
                        imExporting = False
                        Screen.MousePointer = vbDefault
                        'cmdCancel.SetFocus
                        Exit Function
                    End If
                    If imTerminate Then
                        gCloseRegionSQLRst
                        Print #hmMsg, "** User Terminated **"
                        'Close #hmMsg
                        imExporting = False
                        Screen.MousePointer = vbDefault
                        'cmdCancel.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next ilTest2
        If lmTotalNumber > 0 Then
            lmProcessedNumber = lmProcessedNumber + 1
            lmPercent = (lmProcessedNumber * CSng(100)) / lmTotalNumber
            If lmPercent >= 100 Then
                If lmProcessedNumber + 3 < lmTotalNumber Then
                    lmPercent = 99
                Else
                    lmPercent = 100
                End If
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
                If igExportSource = 2 Then DoEvents
            End If
        End If
        gClearASTInfo False
        mClearAbf
    Next ilVefLoop
    ReDim Preserve tmISCIXRef(0 To lmXRefUpper) As ISCIXREF
    ReDim Preserve tmISCISendInfo(0 To lmSendInfoUpper) As ISCISENDINFO
    ReDim Preserve tmISCISendInfo_E(0 To lmSendInfoUpper_E) As ISCISENDINFO

    ''If sgExportISCI <> "A" Then
    'If (rbcSpots(0).Value) Or (rbcSpots(2).Value) Then
    If ((udcCriteria.ISpots(0)) Or (udcCriteria.ISpots(2))) And (udcCriteria.iExportType(0)) And (udcCriteria.IVeh = True) Then
        'Sort Cross Reference
        If UBound(tmISCIXRef) - 1 > 0 Then
            ArraySortTyp fnAV(tmISCIXRef(), 0), UBound(tmISCIXRef), 0, LenB(tmISCIXRef(0)), 0, LenB(tmISCIXRef(0).sKey), 0
        End If
        'Sort ISCI Table
        If UBound(tmISCISendInfo) - 1 > 0 Then
            ArraySortTyp fnAV(tmISCISendInfo(), 0), UBound(tmISCISendInfo), 0, LenB(tmISCISendInfo(0)), 0, LenB(tmISCISendInfo(0).sKey), 0
        End If
        ilRet = mExportISCI_CompressedFormat()
    Else
        If UBound(tmISCISendInfo_E) - 1 > 0 Then
            ArraySortTyp fnAV(tmISCISendInfo_E(), 0), UBound(tmISCISendInfo_E), 0, LenB(tmISCISendInfo_E(0)), 0, LenB(tmISCISendInfo_E(0).sKey), 0
        End If
        ilRet = mExportISCI_AllFormat(lmSendInfoUpper_E, tmISCISendInfo_E())
        'Clear the Export Flags
        slSDate = Format$(smDate, "mm/dd/yy")
        llSDate = DateValue(gAdjYear(slSDate))
        slEDate = Format$(DateAdd("d", imNumberDays - 1, smDate), "mm/dd/yy")
        llEDate = DateValue(gAdjYear(slEDate))
        For ilLoop = 0 To UBound(imFinalAlertVefCode) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            For llDate = llSDate To llEDate Step 7
                If igExportSource = 2 Then DoEvents
                slDate = Format$(llDate, "m/d/yy")
                ilRet = gAlertClear("A", "F", "I", imFinalAlertVefCode(ilLoop), slDate)
            Next llDate
        Next ilLoop
        For ilLoop = 0 To lbcVehicles.ListCount - 1
            If igExportSource = 2 Then DoEvents
            If lbcVehicles.Selected(ilLoop) Then
                imVefCode = lbcVehicles.ItemData(ilLoop)
                For llDate = llSDate To llEDate Step 7
                    If igExportSource = 2 Then DoEvents
                    slDate = Format$(llDate, "m/d/yy")
                    ilRet = gAlertClear("A", "R", "I", imVefCode, slDate)
                Next llDate
            End If
        Next ilLoop
    End If
    If ilRet = False Then
        Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        imExporting = False
        Screen.MousePointer = vbDefault
        'cmdCancel.SetFocus
        Exit Function
    End If
    If imTerminate Then
        Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        imExporting = False
        Screen.MousePointer = vbDefault
        'cmdCancel.SetFocus
        Exit Function
    End If
    'Create or Update EIT
    ''If ((rbcSpots(0).Value) Or (rbcSpots(2).Value)) And (sgExportISCI <> "A") Then
    'If (rbcSpots(0).Value) Or (rbcSpots(2).Value) Then
    If ((udcCriteria.ISpots(0)) Or (udcCriteria.ISpots(2))) And (udcCriteria.IVeh = True) Then
        ilRet = mUpdateEIT()
        If ilRet = False Then
            If (ilRet = False) Then
                Print #hmMsg, "** Terminated **"
                'Close #hmMsg
                imExporting = False
                Screen.MousePointer = vbDefault
                'cmdCancel.SetFocus
                Exit Function
            End If
            If imTerminate Then
                Print #hmMsg, "** User Terminated **"
                'Close #hmMsg
                imExporting = False
                Screen.MousePointer = vbDefault
                'cmdCancel.SetFocus
                Exit Function
            End If
        End If
    End If
    mExportUniqueISCI = True
End Function

Private Function mExportISCIByBreak() As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilPrevCommProvArfCode As Integer
    Dim ilSort As Integer
    Dim ilRet As Integer
    
    mExportISCIByBreak = False
    'Setup sort by provider, then vehicle name
    lbcSort.Clear
    For ilLoop = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilLoop) Then
            imVefCode = lbcVehicles.ItemData(ilLoop)
            smVefName = Trim$(lbcVehicles.List(ilLoop))
            For ilIndex = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If imVefCode = tgVehicleInfo(ilIndex).iCode Then
                    imCommProvArfCode = tgVehicleInfo(ilIndex).iCommProvArfCode
                    smCommProvArfCode = Trim$(Str$(tgVehicleInfo(ilIndex).iCommProvArfCode))
                    Do While Len(smCommProvArfCode) < 5
                        smCommProvArfCode = "0" & smCommProvArfCode
                    Loop
                    lbcSort.AddItem smCommProvArfCode & "|" & smVefName
                    lbcSort.ItemData(lbcSort.NewIndex) = ilLoop
                    Exit For
                End If
            Next ilIndex
        End If
    Next ilLoop
    lmTotalNumber = lbcSort.ListCount
    ilPrevCommProvArfCode = -1
    For ilSort = 0 To lbcSort.ListCount - 1 Step 1
        If igExportSource = 2 Then DoEvents
        ilLoop = lbcSort.ItemData(ilSort)
        'Get hmTo handle
        imVefCode = lbcVehicles.ItemData(ilLoop)
        smVefName = Trim$(lbcVehicles.List(ilLoop))
        smCommProvArfCode = ""
        imCommProvArfCode = 0
        imEmbedded = False
        smProducerName = "!Missing Producer #" & Str(tgVehicleInfo(ilIndex).iProducerArfCode)
        imProducerArfCode = tgVehicleInfo(ilIndex).iProducerArfCode
        smProducerArfCode = Trim$(Str$(imProducerArfCode))
        Do While Len(smProducerArfCode) < 5
            smProducerArfCode = "0" & smProducerArfCode
        Loop
        For ilIndex = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            If imVefCode = tgVehicleInfo(ilIndex).iCode Then
                imCommProvArfCode = tgVehicleInfo(ilIndex).iCommProvArfCode
                smCommProvArfCode = Trim$(Str$(tgVehicleInfo(ilIndex).iCommProvArfCode))
                Exit For
            End If
        Next ilIndex
        Do While Len(smCommProvArfCode) < 5
            smCommProvArfCode = "0" & smCommProvArfCode
        Loop
        Screen.MousePointer = vbHourglass
        If ilPrevCommProvArfCode <> imCommProvArfCode Then
            If ilPrevCommProvArfCode <> -1 Then
                Close hmISCIbyBreak
            End If
            ilRet = mOpenISCIbyBreakFile(imCommProvArfCode)
            If Not ilRet Then
                Print #hmMsg, "** Unable to Open Output File **"
                'Close #hmMsg
                imExporting = False
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            imNoAirDays = -1
            lmFeedDateLastA = 0
        End If
        ilPrevCommProvArfCode = imCommProvArfCode
        ilRet = mGatherISCI_AllFormat()
        lacResult.Caption = ""
        If igExportSource = 2 Then DoEvents
        If (ilRet = False) Then
            Close hmISCIbyBreak
            gCloseRegionSQLRst
            Print #hmMsg, "** Terminated **"
            'Close #hmMsg
            imExporting = False
            Screen.MousePointer = vbDefault
            'cmdCancel.SetFocus
            Exit Function
        End If
        If imTerminate Then
            Close hmISCIbyBreak
            gCloseRegionSQLRst
            Print #hmMsg, "** User Terminated **"
            'Close #hmMsg
            imExporting = False
            Screen.MousePointer = vbDefault
            'cmdCancel.SetFocus
            Exit Function
        End If
        If lmTotalNumber > 0 Then
            lmProcessedNumber = lmProcessedNumber + 1
            lmPercent = (lmProcessedNumber * CSng(100)) / lmTotalNumber
            If lmPercent >= 100 Then
                If lmProcessedNumber + 3 < lmTotalNumber Then
                    lmPercent = 99
                Else
                    lmPercent = 100
                End If
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
                If igExportSource = 2 Then DoEvents
            End If
        End If
        mClearAbf
    Next ilSort
    Close hmISCIbyBreak
    mExportISCIByBreak = True
End Function

Private Sub mAddToISCISendInfo(sISCI As String, iShfCode As Integer, sCallLetters As String, sAdvt As String, lFeedDate As Long, slACName As String, llSendInfoUpper As Long, tlISCISendInfo() As ISCISENDINFO)
    Dim llSetKey As Long
    tlISCISendInfo(llSendInfoUpper).iCommProvArfCode = imCommProvArfCode
    tlISCISendInfo(llSendInfoUpper).iEmbedded = imEmbedded
    tlISCISendInfo(llSendInfoUpper).sVehName = smVefName
    tlISCISendInfo(llSendInfoUpper).sISCI = sISCI
    If imEmbedded Then
        tlISCISendInfo(llSendInfoUpper).sCallLetters = smProducerName
        tlISCISendInfo(llSendInfoUpper).iShfCode = 0
        tlISCISendInfo(llSendInfoUpper).iProducerArfCode = imProducerArfCode
    Else
        tlISCISendInfo(llSendInfoUpper).sCallLetters = sCallLetters
        tlISCISendInfo(llSendInfoUpper).iShfCode = iShfCode
        tlISCISendInfo(llSendInfoUpper).iProducerArfCode = 0
    End If
    tlISCISendInfo(llSendInfoUpper).sAdvtName = sAdvt
    tlISCISendInfo(llSendInfoUpper).lLatestFeedDate = lFeedDate
    tlISCISendInfo(llSendInfoUpper).lEarlestFeedDate = lFeedDate
    If imEmbedded Then
        tlISCISendInfo(llSendInfoUpper).sKey = "!" & smCommProvArfCode & "|" & tlISCISendInfo(llSendInfoUpper).sVehName & "|" & tlISCISendInfo(llSendInfoUpper).sISCI & "|" & smProducerArfCode
    Else
        tlISCISendInfo(llSendInfoUpper).sKey = smCommProvArfCode & "|" & tlISCISendInfo(llSendInfoUpper).sVehName & "|" & tlISCISendInfo(llSendInfoUpper).sCallLetters & "|" & tlISCISendInfo(llSendInfoUpper).sISCI
    End If
    
    If (udcCriteria.iExportType(0)) And ((udcCriteria.ISpots(0)) Or (udcCriteria.ISpots(2))) And (udcCriteria.IVeh = True) Then
        If igExportSource = 2 Then DoEvents
        SQLQuery = "SELECT eitDateSent, eitLLDRef, eitLLDSent, eitCode"
        SQLQuery = SQLQuery & " FROM EIT"
        If imEmbedded Then
            SQLQuery = SQLQuery + " WHERE eitISCI ='" & sISCI & "' And eitType = 'E'" & " And eitshfCode = 0" & " And eitCommProvArfCode = " & imCommProvArfCode
        Else
            SQLQuery = SQLQuery + " WHERE eitISCI ='" & sISCI & "' And eitType = 'S'" & " And eitshfCode = " & iShfCode & " And eitCommProvArfCode = " & imCommProvArfCode
        End If
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            tlISCISendInfo(llSendInfoUpper).lEitCode = rst!eitCode
            tlISCISendInfo(llSendInfoUpper).lDateSent = DateValue(gAdjYear(rst!eitDateSent))
            tlISCISendInfo(llSendInfoUpper).lLLDRef = DateValue(gAdjYear(rst!eitLLDRef))
            tlISCISendInfo(llSendInfoUpper).lLLDSent = DateValue(gAdjYear(rst!eitLLDSent))
        Else
            tlISCISendInfo(llSendInfoUpper).lEitCode = 0
            tlISCISendInfo(llSendInfoUpper).lLLDRef = 0
            tlISCISendInfo(llSendInfoUpper).lLLDSent = 0
        End If
    Else
        tlISCISendInfo(llSendInfoUpper).lEitCode = 0
        tlISCISendInfo(llSendInfoUpper).lLLDRef = 0
        tlISCISendInfo(llSendInfoUpper).lLLDSent = 0
    End If
    tlISCISendInfo(llSendInfoUpper).iUpdateDateSent = False
    tlISCISendInfo(llSendInfoUpper).iVefCode = imVefCode
    If slACName = "" Then
        tlISCISendInfo(llSendInfoUpper).iACExistWithAtt = False
    Else
        tlISCISendInfo(llSendInfoUpper).iACExistWithAtt = True
    End If
    'ReDim Preserve tlISCISendInfo(0 To llUpper + 1) As ISCISENDINFO
    If llSendInfoUpper >= UBound(tlISCISendInfo) Then
        ReDim Preserve tlISCISendInfo(0 To llSendInfoUpper + 20000) As ISCISENDINFO
        For llSetKey = llSendInfoUpper + 1 To UBound(tlISCISendInfo) Step 1
            If igExportSource = 2 Then DoEvents
            tlISCISendInfo(llSetKey).sKey = "~~~~~~~~~~"
        Next llSetKey
    End If
    llSendInfoUpper = llSendInfoUpper + 1
    If llSendInfoUpper - 1 >= 0 Then
        ArraySortTyp fnAV(tlISCISendInfo(), 0), llSendInfoUpper, 0, LenB(tlISCISendInfo(0)), 0, LenB(tlISCISendInfo(0).sKey), 0
    End If

End Sub

Private Sub mMovePledgeToFeed()
    Dim llVpf As Long
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim llTime As Long
    Dim slDate As String
    Dim slTime As String
    Dim slSpotNo As String
    
    If udcCriteria.iExportType(1) = False Then
        Exit Sub
    End If
    llVpf = gBinarySearchVpf(CLng(imVefCode))
    If llVpf = -1 Then
        Exit Sub
    End If
    'exit if using feed date
    If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And EXPORTISCIBYPLEDGE) <> EXPORTISCIBYPLEDGE Then
        Exit Sub
    End If
    lbcPledgeDateTime.Clear
    'Sort by Pledgedate, PledgeTime, Spot #
    For ilLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        llDate = gDateValue(tmAstInfo(ilLoop).sPledgeDate)
        slDate = Trim$(Str$(llDate))
        Do While Len(slDate) < 7
            slDate = "0" & slDate
        Loop
        llTime = gTimeToLong(tmAstInfo(ilLoop).sPledgeStartTime, False)
        slTime = Trim$(Str$(llTime))
        Do While Len(slTime) < 6
            slTime = "0" & slTime
        Loop
        slSpotNo = Trim$(Str$(ilLoop))
        Do While Len(slSpotNo) < 4
            slSpotNo = "0" & slSpotNo
        Loop
        lbcPledgeDateTime.AddItem slDate & slTime & slSpotNo
        lbcPledgeDateTime.ItemData(lbcPledgeDateTime.NewIndex) = ilLoop
    Next ilLoop
    ReDim tlTemp(0 To UBound(tmAstInfo)) As ASTINFO
    For ilLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        tlTemp(ilLoop) = tmAstInfo(ilLoop)
    Next ilLoop
    For ilLoop = 0 To lbcPledgeDateTime.ListCount - 1 Step 1
        If igExportSource = 2 Then DoEvents
        tmAstInfo(ilLoop) = tlTemp(lbcPledgeDateTime.ItemData(ilLoop))
        tmAstInfo(ilLoop).sFeedDate = tmAstInfo(ilLoop).sPledgeDate
        tmAstInfo(ilLoop).sFeedTime = tmAstInfo(ilLoop).sPledgeStartTime
    Next ilLoop
End Sub

Private Sub udcCriteria_ISCIChg()
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
        plcGauge.Visible = False
    End If
End Sub


Private Sub mSaveCustomValues()
    Dim ilLoop As Integer
    ReDim ilVefCode(0 To 0) As Integer
    ReDim ilShttCode(0 To 0) As Integer
    If igExportSource <> 2 Then
        ReDim tgEhtInfo(0 To 1) As EHTINFO
        ReDim tgEvtInfo(0 To 0) As EVTINFO
        ReDim tgEctInfo(0 To 0) As ECTINFO
        lgExportEhtInfoIndex = 0
        tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1
        For ilLoop = 0 To lbcVehicles.ListCount - 1
            If lbcVehicles.Selected(ilLoop) Then
                ilVefCode(UBound(ilVefCode)) = lbcVehicles.ItemData(ilLoop)
                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
            End If
        Next ilLoop
        For ilLoop = 0 To lbcStation.ListCount - 1
            If lbcStation.Selected(ilLoop) Then
                ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
        Next ilLoop
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("I", "ISCI", "I", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
Function mReplaceTab(slInName As String) As String
    'Dim slName As String
    'Dim ilPos As Integer
    'Dim ilFound As Integer
    'slName = slInName
    ''Remove " and '
    'Do
    '    If igExportSource = 2 Then DoEvents
    '    ilFound = False
    '    ilPos = InStr(1, slName, Chr$(9), 1)
    '    If ilPos > 0 Then
    '        Mid$(slName, ilPos, 1) = " "
    '        ilFound = True
    '    End If
    'Loop While ilFound
    'mReplaceTab = slName
    mReplaceTab = Replace(slInName, Chr$(9), " ")
End Function

Private Function mInitDuplInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "sdfCode", adInteger
        .Append "FeedTime", adInteger
        .Append "FeedDate", adInteger
        .Append "StartDate", adInteger
        .Append "EndDate", adInteger
        .Append "Pledge", adInteger
    End With
    rst.Open
    rst!sdfcode.Properties("optimize") = True
    rst.Sort = "sdfCode,FeedTime"
    Set mInitDuplInfo = rst
End Function
Private Sub mCloseDuplInfo()
    On Error Resume Next
    If Not DuplInfo_rst Is Nothing Then
        If (DuplInfo_rst.State And adStateOpen) <> 0 Then
            DuplInfo_rst.Close
        End If
        Set DuplInfo_rst = Nothing
    End If
End Sub

Private Function mTestDuplInfo(llSdfCode As Long, llFeedTime As Long, llFeedDate As Long, llStartDate As Long, llEndDate As Long, ilPledge As Integer) As Integer
    Dim blFound As Boolean
    DuplInfo_rst.Filter = "sdfCode = " & llSdfCode & " And FeedTime = " & llFeedTime
    If DuplInfo_rst.EOF And DuplInfo_rst.BOF Then
        blFound = False
    Else
        blFound = False
        DuplInfo_rst.MoveFirst
        Do While Not DuplInfo_rst.EOF
            If (DuplInfo_rst!FeedDate >= llStartDate) And (DuplInfo_rst!FeedDate <= llEndDate) And (DuplInfo_rst!Pledge = ilPledge) Then
                blFound = True
                Exit Do
            End If
            If (DuplInfo_rst!FeedDate >= llStartDate) And (DuplInfo_rst!FeedDate <= llEndDate) And (DuplInfo_rst!Pledge <> 2) Then
                blFound = True
                Exit Do
            End If
            DuplInfo_rst.MoveNext
        Loop
    End If
    If Not blFound Then
        DuplInfo_rst.AddNew Array("sdfCode", "FeedTime", "FeedDate", "StartDate", "EndDate", "Pledge"), Array(llSdfCode, llFeedTime, llFeedDate, llStartDate, llEndDate, ilPledge)
        mTestDuplInfo = False
    Else
        mTestDuplInfo = True
    End If
End Function

Private Sub mClearAbf()
    Dim sEndDate As String
    Dim sMoDate As String
    
    If (lbcStation.ListCount > 0) And (chkAllStation.Value = vbUnchecked) Then
        Exit Sub
    End If
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    Do
        gClearAbf imVefCode, 0, sMoDate, gObtainNextSunday(sMoDate), True
        sMoDate = DateAdd("d", 7, sMoDate)
    Loop While DateValue(gAdjYear(sMoDate)) < DateValue(gAdjYear(sEndDate))

End Sub

Private Sub mPopDMA()
    Dim ilLoop As Integer
    lbcFilter.Clear
    For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgMarketInfo(ilLoop).sName)
        lbcFilter.ItemData(lbcFilter.NewIndex) = tgMarketInfo(ilLoop).lCode
    Next ilLoop
    'lbcFilter.AddItem "[Defined]", 0
    'lbcFilter.ItemData(lbcFilter.NewIndex) = -1

End Sub

Private Sub mPopMSA()
    Dim ilLoop As Integer
    lbcFilter.Clear
    For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgMSAMarketInfo(ilLoop).sName)
        lbcFilter.ItemData(lbcFilter.NewIndex) = tgMSAMarketInfo(ilLoop).lCode
    Next ilLoop
    'lbcFilter.AddItem "[Defined]", 0
    'lbcFilter.ItemData(lbcFilter.NewIndex) = -1

End Sub

Private Function mGetMSA(lMSACode As Long) As String

    Dim ilLoop As Integer
    
    mGetMSA = ""
    For ilLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        If tgMSAMarketInfo(ilLoop).lCode = lMSACode Then
            mGetMSA = tgMSAMarketInfo(ilLoop).sName
            Exit For
        End If
    Next ilLoop

End Function

Private Sub mPopFormat()
    Dim ilLoop As Integer
    lbcFilter.Clear
    For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgFormatInfo(ilLoop).sName)
        lbcFilter.ItemData(lbcFilter.NewIndex) = tgFormatInfo(ilLoop).lCode
    Next ilLoop
    'lbcFilter.AddItem "[Defined]", 0
    'lbcFilter.ItemData(lbcFilter.NewIndex) = -1

End Sub

Private Function mGetFormat(mFormatCode As Long) As String

    Dim ilLoop As Integer
    
    mGetFormat = ""
    For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        If mFormatCode = tgFormatInfo(ilLoop).lCode Then
            mGetFormat = tgFormatInfo(ilLoop).sName
        End If
    Next ilLoop

End Function


Private Sub mPopState()
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    lbcFilter.Clear
    ilRet = gPopStates()
    For ilRow = 0 To UBound(tgStateInfo) - 1 Step 1
        lbcFilter.AddItem Trim$(tgStateInfo(ilRow).sPostalName) & " (" & Trim$(tgStateInfo(ilRow).sName) & ")"
        lbcFilter.ItemData(lbcFilter.NewIndex) = ilRow    'tgStateInfo(ilRow).iCode
    Next ilRow
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmStation-mPopState"
End Sub


Private Function mGetMarketName(iMarketCode As Integer) As String

    Dim temp_rst As ADODB.Recordset
    
    mGetMarketName = ""
    SQLQuery = "Select mktName, mktRank from Mkt where mktCode = " & rst!shttMktCode
    Set temp_rst = gSQLSelectCall(SQLQuery)
    If Not temp_rst.EOF Then
        mGetMarketName = Trim$(temp_rst!mktName)
    End If

End Function

Private Sub lbcFilter_Click()
    tmcFilterDelay.Enabled = True
End Sub

Private Sub lbcFilter_GotFocus()
    tmcFilterDelay.Enabled = False
End Sub

Private Sub lbcFilter_LostFocus()
    tmcFilterDelay_Timer
    
End Sub

Private Function mBinarySearchStation(slCallLetter As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llResult As Long
    
    llMin = 0
    llMax = lbcStation.ListCount - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        llResult = StrComp(Trim(lbcStation.List(llMiddle)), Trim$(slCallLetter), vbTextCompare)
        Select Case llResult
            Case 0:
                mBinarySearchStation = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    mBinarySearchStation = -1
    Exit Function
    
End Function




           
            
