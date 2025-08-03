VERSION 5.00
Begin VB.Form ExpGP 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3645
   ClientLeft      =   825
   ClientTop       =   2400
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   7095
   Begin VB.Frame frcFileFormat 
      Caption         =   "File Format"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   3735
      Begin VB.OptionButton rbcFileFormat 
         Caption         =   "Multiple Files"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   17
         ToolTipText     =   "(old method) creates two files: ARInv.csv and ARBody.csv"
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton rbcFileFormat 
         Caption         =   "Single File"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "(new method) Creates one file: GPExport-[mmddyyy].csv"
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame frcInstallMethod 
      Caption         =   "Installment Method"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2895
      Begin VB.OptionButton rbcInstallMethod 
         Caption         =   "Billing"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton rbcInstallMethod 
         Caption         =   "Revenue"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6435
      Top             =   60
   End
   Begin VB.TextBox edcSelCFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   0
      Top             =   450
      Width           =   615
   End
   Begin VB.TextBox edcSelCFrom1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   1
      Top             =   450
      Width           =   615
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6435
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5790
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
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
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label lacTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Great Plains G/L Export"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   75
      Width           =   2415
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3240
      TabIndex        =   10
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lacSelCFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Bdcst Month"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   480
      Width           =   1905
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   3120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   3510
   End
End
Attribute VB_Name = "ExpGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExpGp.frm on Fri 3/12/10 @ 11:00 AM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
Dim hmTo2 As Integer   'From file hanle
Dim tmRvf As RVF
Dim hmPrf As Integer        'Prf Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length
Dim tmPrfSrchKey As LONGKEY0  'Prf key record image
Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim smStartCal As String    'Starting date for standard billing
Dim smEndCal As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim lmStartCal As Long
Dim lmEndCal As Long
Dim hmVaf As Integer            'Vehicle Account file handle
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey0 As INTKEY0     'SLF key image
Dim tmSlf As SLF
Dim tmSaf As SAF
Dim hmSaf As Integer
Dim tmSafSrchKey As INTKEY0
Dim imSafRecLen As Integer
Dim lmGPBatchNo As Long
Dim smGPCustomerNo As String
Dim smGPPrefixChar As String
Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer
Dim tmSrchKey As INTKEY0
Dim tmVaf() As VAF
Dim bmUseRevVsBill As Boolean       '11-10-16 true if separate billing and revenue receivables records; Use revenuer records.  If false, billing is same as revenue


Private Sub cmcCancel_Click()

       If imExporting Then
           imTerminate = True
           Exit Sub
       End If
       mTerminate

End Sub

Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim slToFile2 As String
    Dim ilRet As Integer
    Dim slDateTime As String
    
    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    If Not Len(edcSelCFrom.Text) > 0 Then
        ''MsgBox "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
        gAutomationAlertAndLogHandler "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
        edcSelCFrom.SetFocus
        Exit Sub
    End If
    
    If Not Len(edcSelCFrom1.Text) > 0 Then
        ''MsgBox "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
        gAutomationAlertAndLogHandler "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
        edcSelCFrom1.SetFocus
        Exit Sub
    End If
    
    'TTP 10533 - Great Plains export: add single file option
    If rbcFileFormat(0).Value = True Then 'Single File
        slToFile = sgExportPath & "GPExport-" & Format(Date, "MMDDYYYY") & ".csv"
        'If DoesFileExist(slToFile) Then
        If gFileExist(slToFile) = 0 Then
            Kill slToFile
        End If
        If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
            slToFile = sgExportPath & slToFile
        End If
        ilRet = 0
    
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            Exit Sub
        End If
    End If
    
    If rbcFileFormat(1).Value = True Then 'Multiple Files
        slToFile = sgExportPath & "ARINV.CSV"
        'If DoesFileExist(slToFile) Then
        If gFileExist(slToFile) = 0 Then
            Kill slToFile
        End If
        If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
            slToFile = sgExportPath & slToFile
        End If
        ilRet = 0
        'On Error GoTo cmcExportErr:
        'slDateTime = FileDateTime(slToFile)
    
        slToFile2 = sgExportPath & "ARBODY.CSV"
        'If DoesFileExist(slToFile2) Then
        If gFileExist(slToFile2) = 0 Then
            Kill slToFile2
        End If
        If (InStr(slToFile2, ":") = 0) And (Left$(slToFile2, 2) <> "\\") Then
            slToFile2 = sgExportPath & slToFile2
        End If
        ilRet = 0
        'On Error GoTo cmcExportErr:
        'slDateTime = FileDateTime(slToFile2)
        ilRet = gFileExist(slToFile2)
        If ilRet = 0 Then
            'hmTo = FreeFile
            'Open slToFile For Append As hmTo
            ilRet = gFileOpen(slToFile, "Append", hmTo)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         Else
            ilRet = 0
            'hmTo = FreeFile
            'Open slToFile For Output As hmTo
            ilRet = gFileOpen(slToFile, "Output", hmTo)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         End If
    
        If ilRet = 0 Then
            'hmTo2 = FreeFile
            'Open slToFile2 For Append As hmTo2
            ilRet = gFileOpen(slToFile2, "Append", hmTo2)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile2 & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile2 & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         Else
            ilRet = 0
            'hmTo2 = FreeFile
            'Open slToFile2 For Output As hmTo2
            ilRet = gFileOpen(slToFile2, "Output", hmTo2)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile2 & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile2 & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
             End If
         End If
    End If
    Screen.MousePointer = vbHourglass
    imExporting = True
    
    sgMessageFile = sgDBPath & "Messages\" & "ExportGreatPlains.txt"
    '  gLogMsg "Great Plains Export for: " & edcSelCFrom.Text & "/" & edcSelCFrom1.Text, "ExportGreatPlains.txt", False
    
    gAutomationAlertAndLogHandler "** Export Great Plains G/L **"
    gAutomationAlertAndLogHandler "* StartMonth = " & edcSelCFrom.Text
    gAutomationAlertAndLogHandler "* StartYear = " & edcSelCFrom1.Text
    If rbcInstallMethod(0).Value = True Then gAutomationAlertAndLogHandler "* InstallmentMethod = Billing"
    If rbcInstallMethod(1).Value = True Then gAutomationAlertAndLogHandler "* InstallmentMethod = Revenue"
    
    'TTP 10533 - Great Plains export: add single file option
    If rbcFileFormat(0).Value = True Then
        gAutomationAlertAndLogHandler "* FileFormat = Single" 'Single File
        
        '1. Customer number:        from ARInv, column 2; prefix to customer # and rep agency code or rep advertiser code
        '2. Document number:        from ARInv, column 1; three letter division abbreviation plus invoice number
        '3. Document description:   from arinv.csv, column 12; product name (description field)
        '4. Document date:          from arinv.csv, column 4; invoice date
        '5. Sales person ID:        this is new. use salesperson ID (slfCode)
        '6. Amount:                 from arbody, column 2; amount
        '7. GL Account:             from arbody, column 3 (code)
        Print #hmTo, "Customer number,Document number,Document description,Document date,Sales person ID,Amount,GL Account" 'Single File Header
    End If
    
    If rbcFileFormat(1).Value = True Then gAutomationAlertAndLogHandler "* FileFormat = Multiple" 'Multiple Files

    gAutomationAlertAndLogHandler "Exporting..."
    
    ilRet = mLoadRvfAndVaf()
    
    If ilRet = False Then        'error will be an error code
        lacInfo(0).Caption = "Export Failed"
        gLogMsg "Export failed: #" & Trim$(str$(ilRet)), "ExportGreatPlains.txt", False
    Else
        lacInfo(0).Caption = "Export Successfully Completed"
        gLogMsg "Export Successfully Completed, Export Files: " & slToFile & " and " & slToFile2, "ExportGreatPlains.txt", False
    End If
    
    'TTP 10533 - Great Plains export: add single file option
    If rbcFileFormat(0).Value = True Then
        lacInfo(1).Caption = "Export Files: " & slToFile
    Else
        lacInfo(1).Caption = "Export Files: " & slToFile & " and " & slToFile2
    End If
    lacInfo(0).Visible = True
    lacInfo(1).Visible = True
    Close hmTo
    If rbcFileFormat(1).Value = True Then 'Multiple Files
        Close hmTo2
    End If
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    cmcExport.Enabled = False
    Screen.MousePointer = vbDefault
    imExporting = False
    mIncBatchNo
    Exit Sub
'cmcExportErr:
'      ilRet = Err.Number
'       Resume Next

ExportError:
    'gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export: " & err & " - " & Error(err), vbCritical + vbOkOnly, "Export Failed"
    
End Sub

Private Sub Form_Activate()

       If Not imFirstActivate Then
'           DoEvents    'Process events so pending keys are not sent to this
           Me.KeyPreview = True
           Exit Sub
       End If
       imFirstActivate = False
'       DoEvents    'Process events so pending keys are not sent to this
       Me.KeyPreview = True
       Me.Refresh
       edcSelCFrom.SetFocus

End Sub
Private Sub Form_Deactivate()

       Me.KeyPreview = False

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

       If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
           gFunctionKeyBranch KeyCode
       End If

End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

       sgDoneMsg = CmdStr
       igChildDone = True
       Cancel = 0

End Sub
Private Sub Form_Load()

    mInit
    If imTerminate Then
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width      'move off the screen so screen won't flash
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    ilRet = btrClose(hmVaf)
    btrDestroy hmVaf
    
    Set ExpGP = Nothing   'Remove data segment

End Sub

Private Sub rbcInstallMethod_Click(Index As Integer)
    If Index = 0 Then           'billing
        bmUseRevVsBill = False
    Else
        bmUseRevVsBill = True
    End If
End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
Private Sub mInit()

       Dim ilRet As Integer
       Dim slStdDate As String

       imTerminate = False
       imFirstActivate = True
       Screen.MousePointer = vbHourglass
       imExporting = False
       imFirstFocus = True
       imBypassFocus = False
       lmTotalNoBytes = 0
       lmProcessedNoBytes = 0

        hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAgf, "", sgDBPath & "AGF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: AGF.Btr)", ExpGP
        imAgfRecLen = Len(tmAgf)

        hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", ExpGP
        imAdfRecLen = Len(tmAdf)

        hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpGP
        imPrfRecLen = Len(tmPrf)

        hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj
        ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpGP
        imSafRecLen = Len(tmSaf)

        hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpGP
        imSlfRecLen = Len(tmSlf)


        hmVaf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmVaf, "", sgDBPath & "Vaf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpGP

        gCenterStdAlone ExpGP
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStdDate
        edcSelCFrom.Text = Format$(slStdDate, "MMM")
        edcSelCFrom1.Text = Year(slStdDate)
        
        bmUseRevVsBill = False          'default to use billing (vs revenue)
        If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT Then
            rbcInstallMethod(0).Value = True        'default to billing
            If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then        'separate revenue and billing
                bmUseRevVsBill = True
                rbcInstallMethod(1).Value = True        'Installment site is using revenue is earned, default to revenue
            End If
        Else                    'hide the selectivity, default to use billing if not using installment
            frcInstallMethod.Visible = False
        End If

        Screen.MousePointer = vbDefault

        ilRet = mPopVaf
        gAutomationAlertAndLogHandler ""
        gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
        
        Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mTerminate()


    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpGP
    igManUnload = NO

End Sub




Public Function mLoadRvfAndVaf() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilUseCodes                    tlVaf                     *
'*  slCash                        slName                                                  *
'******************************************************************************************



    Dim slStr1 As String
    Dim ilYear As Integer
    Dim ilRet As Integer

    Dim tlTranType As TRANTYPES
    ReDim tlRvf(0 To 0) As RVF
    Dim llRvfLoop As Long
    Dim ilWhichDate As Integer      '0=use tran date, 1 = use date entered
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilCurrentMonth As Integer
    Dim tlGP() As GP
    Dim iCounter As Integer
    Dim iMax As Integer
    Dim iPackageCounter As Integer
    Dim slAmount As String
    Dim slAmount2 As String
    Dim llAmt As Long

    tlTranType.iNTR = True
    tlTranType.iAirTime = True
    tlTranType.iInv = True              'invoices
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iAdj = False              'adjustments

    ilWhichDate = 0                     'default to use tran date vs date entered

    mGetSafItems

    slStr1 = ExpGP!edcSelCFrom.Text             'month in text form (jan..dec)
    ilYear = Val(ExpGP!edcSelCFrom1.Text)

    gGetMonthNoFromString slStr1, ilCurrentMonth         'getmonth #
    If ilCurrentMonth = 0 Then                                 'input isn't text month name, try month #
        ilCurrentMonth = Val(slStr1)
    End If
    slStr1 = Trim$(str(ilCurrentMonth)) & "/15/" & Trim$(ExpGP!edcSelCFrom1.Text)     'form mm/dd/yy

    smStartStd = gObtainStartStd(slStr1)               'obtain std start date for month
    lmStartStd = gDateValue(smStartStd)
    smEndStd = gObtainEndStd(slStr1)                 'obtain std end date for month
    lmEndStd = gDateValue(smEndStd)
    smEndCal = gObtainEndCal(smEndStd)
    smStartCal = gObtainStartCal(smEndCal)
    lmStartCal = gDateValue(smStartCal)
    lmEndCal = gDateValue(smEndCal)

    If (lmStartStd > 0) And (lmStartCal > 0) Then
        If lmStartStd < lmStartCal Then
            slStartDate = Format$(lmStartStd, "m/d/yy")
        Else
            slStartDate = Format$(lmStartCal, "m/d/yy")
        End If
    ElseIf lmStartStd > 0 Then
        slStartDate = Format$(lmStartStd, "m/d/yy")
    Else
        slStartDate = Format$(lmStartCal, "m/d/yy")
    End If
    If (lmEndStd > 0) And (lmEndCal > 0) Then
        If lmEndStd > lmEndCal Then
            slEndDate = Format$(lmEndStd, "m/d/yy")
        Else
            slEndDate = Format$(lmEndCal, "m/d/yy")
        End If
    ElseIf lmEndStd > 0 Then
        slEndDate = Format$(lmEndStd, "m/d/yy")
    Else
        slEndDate = Format$(lmEndCal, "m/d/yy")
    End If


    ilRet = gObtainPhfRvf(ExpGP, slStartDate, slEndDate, tlTranType, tlRvf(), ilWhichDate) '12-14-06 add parm to indicate to use tran date or entry date
    If ilRet = False Then
        mLoadRvfAndVaf = False
        Exit Function
    End If
    ' DM   5/8/08  group and sum based on vaf.sDivisionCode & tlRvf().lInvNo & tlRvf().sCashTrade

    ReDim tlGP(LBound(tlRvf) To UBound(tlRvf)) As GP
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1
    
        ilRet = mBinarySearchVaf(tlRvf(llRvfLoop).iAirVefCode)
        If ilRet >= 0 Then
            If Trim$(tlRvf(llRvfLoop).sCashTrade) = "T" Then
                '4-10-12 strip out blanks
                tlGP(llRvfLoop).sKey = Trim$(tmVaf(ilRet).sDivisionCode) & Trim$(tmVaf(ilRet).sBranchCodeTrade) & tlRvf(llRvfLoop).lInvNo & tlRvf(llRvfLoop).sCashTrade
            Else
                tlGP(llRvfLoop).sKey = Trim$(tmVaf(ilRet).sDivisionCode) & Trim$(tmVaf(ilRet).sBranchCodeCash) & tlRvf(llRvfLoop).lInvNo & tlRvf(llRvfLoop).sCashTrade
            End If
        Else            '7-7-11 no vendor info exists
                        '4-10-12 blank out all fields since client may not be using division codes and/or this may be a package vehicle
            tlGP(llRvfLoop).sKey = "" & tlRvf(llRvfLoop).lInvNo & tlRvf(llRvfLoop).sCashTrade
        End If
        ' Dan M 8-05-08 changed lrvfpointer to long
        tlGP(llRvfLoop).lRvfPointer = llRvfLoop
        tlGP(llRvfLoop).bDelete = False
    Next llRvfLoop
    If LBound(tlGP) > 0 Then
        ArraySortTyp fnAV(tlGP(), 1), UBound(tlGP) - 1, 0, LenB(tlGP(1)), 0, LenB(tlGP(1).sKey), 0
    Else
        ArraySortTyp fnAV(tlGP(), 0), UBound(tlGP), 0, LenB(tlGP(0)), 0, LenB(tlGP(0).sKey), 0
    End If
    iCounter = LBound(tlGP)
    iMax = UBound(tlGP)
    Do While iCounter < iMax      'compare one rows key to the next.  Add values if match.  Continue checking following rows until no match.
        '1/3/08:  Bypass zero amount transactions
        gPDNToLong tlRvf(tlGP(iCounter).lRvfPointer).sNet, llAmt
 
    ' Dan M must distinguish between rvf/phf types...want to exclude "a". This first 'if' will check the current record
'        If (tlRvf(tlGP(iCounter).lRvfPointer).sType = "A") Or (llAmt = 0) Then
        'Determine type of record (for installment) to include the billing or revenue record
        '11-10-16 change to test site to see how installment is billed (by Rev is same as billing, or by Aired, which is separate revenue and billing
        '  A = Installment revenue (Aired - separate rev and billing)
        '   I = Installment Billing record
        '   blank - non installment bill
        If (tlRvf(tlGP(iCounter).lRvfPointer).sType = "A" And bmUseRevVsBill = False) Or (tlRvf(tlGP(iCounter).lRvfPointer).sType = "I" And bmUseRevVsBill = True) Then
            tlGP(iCounter).bDelete = True
    ' added logic to make sure icounter+1 wasn't "A" by skipping addition
        ElseIf tlGP(iCounter).sKey = tlGP(iCounter + 1).sKey Then
'            If tlRvf(tlGP(iCounter + 1).lRvfPointer).sType <> "A" Then
            'Determine type of record (for installment) to include the billing or revenue record
            '11-10-16 change to test site to see how installment is billed (by Rev is same as billing, or by Aired, which is separate revenue and billing
            '  A = Installment revenue (Aired - separate rev and billing)
            '  blank or I = Installment Billing record
            If (tlRvf(tlGP(iCounter + 1).lRvfPointer).sType = " ") Or (tlRvf(tlGP(iCounter + 1).lRvfPointer).sType = "A" And bmUseRevVsBill = True) Or (tlRvf(tlGP(iCounter + 1).lRvfPointer).sType = "I" And bmUseRevVsBill = False) Then

                gPDNToStr tlRvf(tlGP(iCounter).lRvfPointer).sNet, 2, slAmount
                gPDNToStr tlRvf(tlGP(iCounter + 1).lRvfPointer).sNet, 2, slAmount2
                slAmount = gAddStr(slAmount, slAmount2)
                gStrToPDN slAmount, 2, 6, tlRvf(tlGP(iCounter).lRvfPointer).sNet
                gPDNToStr tlRvf(tlGP(iCounter).lRvfPointer).sGross, 2, slAmount
                gPDNToStr tlRvf(tlGP(iCounter + 1).lRvfPointer).sGross, 2, slAmount2
                slAmount = gAddStr(slAmount, slAmount2)
                gStrToPDN slAmount, 2, 6, tlRvf(tlGP(iCounter).lRvfPointer).sGross
            End If
            tlGP(iCounter + 1).bDelete = True
            iPackageCounter = 1
            Do While iPackageCounter + iCounter < iMax
                If tlGP(iCounter).sKey = tlGP(iCounter + iPackageCounter + 1).sKey Then
'                    If tlRvf(tlGP(iCounter + iPackageCounter + 1).lRvfPointer).sType <> "A" Then        'as above, added logic to not add if type = A
                    'Determine type of record (for installment) to include the billing or revenue record
                    '11-10-16 change to test site to see how installment is billed (by Rev is same as billing, or by Aired, which is separate revenue and billing
                    '  A = Installment revenue (Aired - separate rev and billing)
                    '  blank or I = Installment Billing record
                    If (tlRvf(tlGP(iCounter + iPackageCounter + 1).lRvfPointer).sType = " ") Or (tlRvf(tlGP(iCounter + iPackageCounter + 1).lRvfPointer).sType = "A" And bmUseRevVsBill = True) Or (tlRvf(tlGP(iCounter + iPackageCounter + 1).lRvfPointer).sType = "I" And bmUseRevVsBill = False) Then

                        gPDNToStr tlRvf(tlGP(iCounter).lRvfPointer).sNet, 2, slAmount
                        gPDNToStr tlRvf(tlGP(iCounter + iPackageCounter + 1).lRvfPointer).sNet, 2, slAmount2
                        slAmount = gAddStr(slAmount, slAmount2)
                        gStrToPDN slAmount, 2, 6, tlRvf(tlGP(iCounter).lRvfPointer).sNet
                        gPDNToStr tlRvf(tlGP(iCounter).lRvfPointer).sGross, 2, slAmount
                        gPDNToStr tlRvf(tlGP(iCounter + iPackageCounter + 1).lRvfPointer).sGross, 2, slAmount2
                        slAmount = gAddStr(slAmount, slAmount2)
                        gStrToPDN slAmount, 2, 6, tlRvf(tlGP(iCounter).lRvfPointer).sGross
                    End If
                    tlGP(iCounter + iPackageCounter + 1).bDelete = True
                    iPackageCounter = iPackageCounter + 1
                Else
                    iCounter = iCounter + iPackageCounter
                    Exit Do
                End If
            Loop
            If iPackageCounter + iCounter >= iMax Then
                Exit Do
            End If
        End If          'stype = "A" and elseif .skey = .skey
        iCounter = iCounter + 1
    Loop
    For llRvfLoop = LBound(tlGP) To UBound(tlGP) - 1
        If Not tlGP(llRvfLoop).bDelete Then
            tmRvf = tlRvf(tlGP(llRvfLoop).lRvfPointer)
            mMakeExportRec
        End If
    Next llRvfLoop
    mLoadRvfAndVaf = True

    Exit Function

End Function

Function mBinarySearchVaf(ilCode As Integer) As Integer

    Dim ilMin As Integer

    Dim ilMax As Integer

    Dim ilMiddle As Integer

    ilMin = LBound(tmVaf)

    ilMax = UBound(tmVaf) - 1

    Do While ilMin <= ilMax

        ilMiddle = (ilMin + ilMax) \ 2

        If ilCode = tmVaf(ilMiddle).iVefCode Then
            'found the match

            mBinarySearchVaf = ilMiddle

            Exit Function

        ElseIf ilCode < tmVaf(ilMiddle).iVefCode Then

            ilMax = ilMiddle - 1

        Else

            'search the right half

            ilMin = ilMiddle + 1

        End If

    Loop

    mBinarySearchVaf = -1

End Function



Private Sub mMakeExportRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  tlVaf                                                                                 *
'******************************************************************************************


    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slProd As String
    Dim slCustCode As String
    Dim slInvDate As String
    Dim slTranDate As String
    Dim slInvNo As String
    Dim slNet As String
    Dim slGross As String
    Dim slCash As String
    Dim slRecord As String
    Dim slRecord2 As String
    Dim slAType As String
    Dim slDueDate As String
    Dim slAirVeh As String
    Dim slCntrNo As String
    Dim slSlsp As String
    Dim slBranchCodeCash As String
    Dim slBranchCodeTrade As String
    Dim slPCAgyCommCash As String
    Dim slPCGrossSalesCash As String
    Dim slPCGrossSalesTrade As String
    Dim slPCRecvCash As String
    Dim slPCRecvTrade As String
    Dim slAmount As String
    Dim slCheckNo As String
    Dim slCode As String
    Dim ilType As Integer
    Dim tlInvRec(0 To 11) As ComDelRec
    Dim tlBodyRec(0 To 3) As ComDelRec
    Dim tlSingleFileRec(0 To 7) As ComDelRec
    Dim llAmt As Long
    
 On Error GoTo mmakeexportrecerr
    gPDNToLong tmRvf.sNet, llAmt
    If llAmt = 0 Then                   '4-16-07 request to ignore $0 transactions
        Exit Sub
    End If

    If tmRvf.lInvNo > 0 Then
       slInvNo = Trim$(str$(tmRvf.lInvNo))
    Else
       slInvNo = ""
    End If

    If tmRvf.iAgfCode > 0 Then
        slAType = "AGNY"
     Else
        slAType = "ADVT"
     End If

    slCustCode = ""
    If slAType = "ADVT" Then
        ilLoop = gBinarySearchAdf(tmRvf.iAdfCode)
        If ilLoop <> -1 Then
         tmSrchKey.iCode = tmRvf.iAdfCode
              ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
              If ilRet = BTRV_ERR_NONE Then
                slCustCode = Trim$(tmAdf.sCodeRep)
              End If

        End If
     Else
        ilLoop = gBinarySearchAgf(tmRvf.iAgfCode)
        If ilLoop <> -1 Then
         tmSrchKey.iCode = tmRvf.iAgfCode
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
              If ilRet = BTRV_ERR_NONE Then
                slCustCode = Trim$(tmAgf.sCodeRep)
              End If

        End If
     End If


    slProd = ""
    If tmRvf.lPrfCode > 0 Then
       tmPrfSrchKey.lCode = tmRvf.lPrfCode
       ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
       If (ilRet = BTRV_ERR_NONE) Then
          slProd = Trim$(tmPrf.sName)
       End If
    End If

    slAirVeh = ""
    If tmRvf.iAirVefCode > 0 Then
          ilLoop = gBinarySearchVef(tmRvf.iAirVefCode)
          If ilLoop <> -1 Then
              slAirVeh = Trim$(tgMVef(ilLoop).sName)
           End If
    End If
    slCntrNo = ""
    If tmRvf.lCntrNo > 0 Then
      slCntrNo = Trim$(str$(tmRvf.lCntrNo))
    Else
      slCntrNo = ""
    End If
    gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slTranDate
    gUnpackDate tmRvf.iInvDate(0), tmRvf.iInvDate(1), slInvDate
    slInvDate = gAdjYear(slInvDate)
    slDueDate = DateAdd("d", 15, slInvDate)

    slCheckNo = ""
    '6/9/15: Check number changed to string
    'If tmRvf.lCheckNo > 0 Then
    '  slCheckNo = Trim$(str$(tmRvf.lCheckNo))
    'Else
    '  slCheckNo = ""
    'End If
    slCheckNo = Trim$(str$(tmRvf.sCheckNo))

    gPDNToStr tmRvf.sGross, 2, slGross
    gPDNToStr tmRvf.sNet, 2, slNet
    slCash = tmRvf.sCashTrade
    slSlsp = ","
    tmSlfSrchKey0.iCode = tmRvf.iSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        slSlsp = Trim$(tmSlf.sCodeStn)
    End If
' Dan
    ilRet = mBinarySearchVaf(tmRvf.iAirVefCode)
        If tmVaf(ilRet).iVefCode = tmRvf.iAirVefCode Then
            slInvNo = Trim$(tmVaf(ilRet).sDivisionCode) & slInvNo
            slBranchCodeCash = Trim$(tmVaf(ilRet).sBranchCodeCash)
            slBranchCodeTrade = Trim$(tmVaf(ilRet).sBranchCodeTrade)
            slPCAgyCommCash = Trim$(tmVaf(ilRet).sPCAgyCommCash)
            slPCGrossSalesCash = Trim$(tmVaf(ilRet).sPCGrossSalesCash)
            slPCGrossSalesTrade = Trim$(tmVaf(ilRet).sPCGrossSalesTrade)
            slPCRecvCash = Trim$(tmVaf(ilRet).sPCRecvCash)
            slPCRecvTrade = Trim$(tmVaf(ilRet).sPCRecvTrade)




'     ilRet = gObtainVaf(ExpGP, tmRvf.iAirVefCode, tlVaf, hmVaf)
'        If tlVaf.iVefCode = tmRvf.iAirVefCode Then
'            slInvNo = Trim$(tlVaf.sDivisionCode) & slInvNo
'            slBranchCodeCash = Trim$(tlVaf.sBranchCodeCash)
'            slBranchCodeTrade = Trim$(tlVaf.sBranchCodeTrade)
'            slPCAgyCommCash = Trim$(tlVaf.sPCAgyCommCash)
'            slPCGrossSalesCash = Trim$(tlVaf.sPCGrossSalesCash)
'            slPCGrossSalesTrade = Trim$(tlVaf.sPCGrossSalesTrade)
'            slPCRecvCash = Trim$(tlVaf.sPCRecvCash)
'            slPCRecvTrade = Trim$(tlVaf.sPCRecvTrade)
        Else
            slInvNo = slInvNo
            slBranchCodeCash = ""
            slBranchCodeTrade = ""
            slPCAgyCommCash = ""
            slPCGrossSalesCash = ""
            slPCGrossSalesTrade = ""
            slPCRecvCash = ""
            slPCRecvTrade = ""
        End If

    tlInvRec(0).sFieldValue = slInvNo
    tlInvRec(1).sFieldValue = smGPPrefixChar & slCustCode 'tmSaf.sGPPrefixChar + tmAdf.sCodeRep
    tlInvRec(2).sFieldValue = lmGPBatchNo
    tlInvRec(3).sFieldValue = slInvDate
    tlInvRec(4).sFieldNum = slGross
    tlInvRec(5).sFieldNum = slGross - slNet
    tlInvRec(6).sFieldNum = 1
    tlInvRec(7).sFieldValue = slDueDate
    tlInvRec(8).sFieldValue = slSlsp
    tlInvRec(9).sFieldNum = 1
    If Trim$(slCash) = "T" Then
        tlInvRec(10).sFieldNum = slBranchCodeTrade
    Else
        tlInvRec(10).sFieldNum = slBranchCodeCash
    End If

    tlInvRec(11).sFieldValue = slProd

    If rbcFileFormat(1).Value = True Then 'Multiple Files
        slRecord = gCreateComDelRec(1233, tlInvRec())
        Print #hmTo, slRecord
    End If
    For ilLoop = 1 To 3
        If slCash <> "T" Then
            Select Case ilLoop
                Case Is = 1 'Sales
                    slAmount = -slGross
                    slCode = "1-" & slBranchCodeCash & "-" & slPCGrossSalesCash & "-"
                    ilType = 9
                Case Is = 2 'Receivables
                    slAmount = slNet
                    slCode = "1-" & slBranchCodeCash & "-" & slPCRecvCash & "-"
                    ilType = 3
                Case Is = 3 'Trade or Commission
                    If slGross - slNet = 0 Then
                        Exit For
                    Else
                        slAmount = slGross - slNet
                        slCode = "1-" & slBranchCodeCash & "-" & slPCAgyCommCash & "-"
                        ilType = 10
                    End If
            End Select
        Else
            Select Case ilLoop
            ' dan M
                Case Is = 1 'Sales
                    slAmount = -slGross
                    slCode = "1-" & slBranchCodeTrade & "-" & slPCGrossSalesTrade & "-"
                    ilType = 9
                Case Is = 2 'Receivables
                    slAmount = slNet
                    slCode = "1-" & slBranchCodeTrade & "-" & slPCRecvTrade & "-"
                    ilType = 9
                Case Is = 3 'Trade or Commission
                ' Dan M  6-16-08 trade needs comission option
                    If slGross - slNet = 0 Then
                        Exit For
                    Else
                        slAmount = slGross - slNet
                        slCode = "1-" & slBranchCodeTrade & "-" & slPCAgyCommCash & "-"
                        ilType = 9
                    End If



                    'Exit For
            End Select
        End If

        tlBodyRec(0).sFieldValue = slInvNo
        tlBodyRec(1).sFieldNum = slAmount
        tlBodyRec(2).sFieldValue = slCode
        tlBodyRec(3).sFieldNum = ilType
        
        'TTP 10533 - Great Plains export: add single file option
        If rbcFileFormat(0).Value = True Then 'Single File
            'Columns:
            '1. Customer number:        from ARInv, column 2; prefix to customer # and rep agency code or rep advertiser code
            'tlSingleFileRec(0).sFieldValue = slCustCode
            tlSingleFileRec(0).sFieldValue = smGPPrefixChar & slCustCode 'Fix per Jason Email v81 Great Plains testing 10/3
            
            '2. Document number:        from ARInv, column 1; three letter division abbreviation plus invoice number
            'tlSingleFileRec(1).sFieldNum = smGPPrefixChar & slInvNo
            tlSingleFileRec(1).sFieldNum = slInvNo 'Fix per Jason Email v81 Great Plains testing 10/3
            
            '3. Document description:   from arinv.csv, column 12; product name (description field)
            tlSingleFileRec(2).sFieldValue = slProd
            
            '4. Document date:          from arinv.csv, column 4; invoice date
            tlSingleFileRec(3).sFieldValue = slInvDate
            
            '5. Sales person ID:        this is new. use salesperson ID (slfCode)
            tlSingleFileRec(4).sFieldNum = tmRvf.iSlfCode
            
            '6. Amount:                 from arbody, column 2; amount
            tlSingleFileRec(5).sFieldNum = slAmount
            
            '7. GL Account:             from arbody, column 3 (code)
            tlSingleFileRec(6).sFieldValue = slCode
            
            slRecord = gCreateComDelRec(123, tlSingleFileRec())
            Print #hmTo, slRecord
        Else
            slRecord2 = gCreateComDelRec(123, tlBodyRec())
            Print #hmTo2, slRecord2
        End If
    Next ilLoop
    Exit Sub
mmakeexportrecerr:
Resume Next
End Sub

Private Function mGetSafItems() As Integer
    Dim ilRet As Integer


        imSafRecLen = Len(tmSaf) 'btrRecordLength(hmSaf)  'Get and save record length
        ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            lmGPBatchNo = tmSaf.lGPBatchNo
            smGPCustomerNo = tmSaf.sGPCustomerNo
            smGPPrefixChar = tmSaf.sGPPrefixChar
        End If
        
        'one client wants to always use Billing vs Revenue, even tho their site is defined as separate billing and revenue
'        '11-10-16 for installment records, determine whether to use Billing or revenue
'        bmUseRevVsBill = False
'        If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then        'separate revenue and billing
'            bmUseRevVsBill = True
'        End If

End Function

Private Function mIncBatchNo()
    Dim ilRet As Integer

    imSafRecLen = Len(tmSaf)
    tmSafSrchKey.iCode = tmSaf.iCode
    ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)

    If ilRet = BTRV_ERR_NONE Then

        tmSaf.lGPBatchNo = tmSaf.lGPBatchNo + 1
        ilRet = btrUpdate(hmSaf, tmSaf, imSafRecLen)
    End If
End Function


Private Function mPopVaf() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRecPos                                                                              *
'******************************************************************************************

Dim ilRet As Integer
Dim ilRecLen As Integer

ReDim tmVaf(0 To 0) As VAF
ilRecLen = Len(tmVaf(UBound(tmVaf)))
ilRet = btrGetFirst(hmVaf, tmVaf(UBound(tmVaf)), ilRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)

Do While ilRet = BTRV_ERR_NONE
    ReDim Preserve tmVaf(0 To UBound(tmVaf) + 1) As VAF
    ilRet = btrGetNext(hmVaf, tmVaf(UBound(tmVaf)), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)

Loop

mPopVaf = True



End Function
