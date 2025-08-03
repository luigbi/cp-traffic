VERSION 5.00
Begin VB.Form ExpInv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
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
   ScaleHeight     =   2805
   ScaleWidth      =   7095
   Begin VB.ListBox lbcVehicles 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "expinv.frx":0000
      Left            =   2730
      List            =   "expinv.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   1740
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5595
      Top             =   360
   End
   Begin VB.TextBox edcSelCFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2235
      MaxLength       =   3
      TabIndex        =   0
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox edcSelCFrom1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   1
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6390
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5115
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5715
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1245
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
      Top             =   2355
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
      Top             =   2355
      Width           =   1050
   End
   Begin VB.Label lacTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Export"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   75
      Width           =   2325
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3180
      TabIndex        =   10
      Top             =   660
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
      Top             =   675
      Width           =   1905
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   420
      TabIndex        =   8
      Top             =   1770
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   2310
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   7
      Top             =   1485
      Visible         =   0   'False
      Width           =   5550
   End
End
Attribute VB_Name = "ExpInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim lmStartCal As Long
Dim lmEndCal As Long
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey0 As INTKEY0     'SLF key image
Dim tmSlf As SLF

Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer
Dim tmSrchKey As INTKEY0

Dim tmChf As CHF
Dim imCHFRecLen As Integer
Dim hmCHF As Integer
Dim tmChfSrchKey As LONGKEY0

Dim tmClf As CLF
Dim imClfRecLen As Integer
Dim hmClf As Integer
Dim tmClfSrchKey0 As CLFKEY0

Dim tmCff As CFF
Dim imCffRecLen As Integer
Dim hmCff As Integer

Dim tmPnf As PNF
Dim hmPnf As Integer
Dim imPnfRecLen As Integer
Dim tmPnfSrchKey0 As INTKEY0     'PNF key image

Dim tmRaf As RAF
Dim imRafRecLen As Integer
Dim hmRaf As Integer
Dim tmRafSrchKey0 As LONGKEY0

Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim hmSdf As Integer
Dim tmSdfSrchKey3 As LONGKEY0

Dim tmCif As CIF
Dim imCifRecLen As Integer
Dim hmCif As Integer
Dim tmCifSrchKey As LONGKEY0

Dim tmCpf As CPF
Dim imCpfRecLen As Integer
Dim hmCpf As Integer
Dim tmCpfSrchKey As LONGKEY0

Dim tmEff As EFF
Dim imEffRecLen As Integer
Dim hmEff As Integer
Dim tmEffSrchKey As LONGKEY0

Dim tmMnf As MNF
Dim imMnfRecLen As Integer
Dim hmMnf As Integer
Dim tmMnfSrchKey As INTKEY0

Dim hmSmf As Integer

Dim hmVsf As Integer

Dim hmVef As Integer

Dim tmSdfInfo() As SDFSORTBYLINE
Dim tmSpotTypes As SPOTTYPES     'spot types to include
Dim tmExportInfo As ExportInfo

Private Type ExportInfo
    sInvType As String * 16   'hard coded to Channel 1 (2301)
    sInvDate As String * 8     'std bdcst end date (MM/DD/YY)
    sSBUSubCode As String * 4 'hard coded to 2301
    sProjectCode As String * 10 'from cnt header
    sSMS As String * 10     'from cnt header
    sSlsp As String * 40      'salesperson name
    sClient As String      'client name & address
    sVehicle As String * 40   'vehicle name
    sProduct As String * 35   'product name from copy
    sContact As String * 20   'agency contact (buyer)
    sLength As String      'spot length
    sRegion As String * 80     'region name
    sSpotType As String * 10   'mg, bonus
    sISCI As String * 20       'ISCI code
    sAirDate As String * 8    'sched air date (MM/DD/YY)
    sNetAmt As String       'spot rate (net, decimal included)
    sRetReq As String * 3     'return requested - hard coded to Yes
    sInternalAlloc As String * 1  'hard coded to blank
    sSBU As String * 1          'hard coded to blank
    sInvNumber As String * 1    'hard coded to blank
End Type
Private Sub cmcCancel_Click()
       If imExporting Then
           imTerminate = True
           Exit Sub
       End If
       mTerminate
End Sub

Private Sub cmcExport_Click()
        Dim slToFile As String
        Dim ilRet As Integer
        Dim slDateTime As String
        Dim slRepeat As String
        Dim ilFileExists As Integer
        Dim ilShowMesg As Integer
        Dim slClientName As String

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


        slClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slClientName = Trim$(tmMnf.sName)
            End If
        End If


        slRepeat = ""
        slToFile = sgExportPath & Trim$(slClientName) & "_" & edcSelCFrom & edcSelCFrom1 & ".csv"
        ilFileExists = False
        Do
            If DoesFileExist(slToFile) Then
                ilFileExists = True
                    'Kill slToFile
                If slRepeat = "" Then       'first time doesnt have the alpha appended for conseutive runs
                    slRepeat = "A"
                Else
                    slRepeat = Chr(Asc(slRepeat) + 1)
                End If
                slToFile = sgExportPath & Trim$(slClientName) & "_" & edcSelCFrom & edcSelCFrom1 & Trim$(slRepeat) & ".csv"
            Else
                ilFileExists = False
            End If
        Loop While ilFileExists = True
        If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
            slToFile = sgExportPath & slToFile
        End If
        ilRet = 0
        'On Error GoTo cmcExportErr:
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
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

      Screen.MousePointer = vbHourglass
      imExporting = True

      gLogMsg "Invoice Export for: " & edcSelCFrom.Text & "/" & edcSelCFrom1.Text, "ExportInvoice.txt", False

      ilShowMesg = mGetSpotsAndExportInv()

      If ilShowMesg = True Then        'error will be an error code
          lacInfo(0).Caption = "Call Counterpoint: Export has errors, see ExportInvoice.txt"
      Else
          lacInfo(0).Caption = "Export Successfully Completed"
          gLogMsg "Export Successfully Completed, Export Files: " & slToFile, "ExportInvoice.txt", False
      End If
      lacInfo(1).Caption = "Export Files: " & slToFile

      lacInfo(0).Visible = True
      lacInfo(1).Visible = True
      Close hmTo
      cmcCancel.Caption = "&Done"
      cmcCancel.SetFocus
      cmcExport.Enabled = False
      Screen.MousePointer = vbDefault
      imExporting = False

       Exit Sub
'cmcExportErr:
'      ilRet = Err.Number
'       Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub

Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
End Sub

Private Sub Form_Activate()

       If Not imFirstActivate Then
           DoEvents    'Process events so pending keys are not sent to this
           Me.KeyPreview = True
           Exit Sub
       End If
       imFirstActivate = False
       DoEvents    'Process events so pending keys are not sent to this
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
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmPnf)
    btrDestroy hmPnf
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmEff)
    btrDestroy hmEff
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf

    Set ExpInv = Nothing   'Remove data segment

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
        gBtrvErrorMsg ilRet, "mInit (btrOpen: AGF.Btr)", ExpInv
        imAgfRecLen = Len(tmAgf)

        hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", ExpInv
        imAdfRecLen = Len(tmAdf)

        hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf)", ExpInv
        imSlfRecLen = Len(tmSlf)

        hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf)", ExpInv
        imCHFRecLen = Len(tmChf)

        hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf)", ExpInv
        imClfRecLen = Len(tmClf)

        hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff)", ExpInv
        imCffRecLen = Len(tmCff)

        hmPnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Pnf)", ExpInv
        imPnfRecLen = Len(tmPnf)

        hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf)", ExpInv
        imRafRecLen = Len(tmRaf)

        hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif)", ExpInv
        imCifRecLen = Len(tmCif)

        hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf)", ExpInv
        imCpfRecLen = Len(tmCpf)

        hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf)", ExpInv
        imSdfRecLen = Len(tmSdf)

        hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf)", ExpInv

        hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf)", ExpInv

        hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef)", ExpInv


        hmEff = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmEff, "", sgDBPath & "Eff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Eff)", ExpInv
        imEffRecLen = Len(tmEff)

        hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf)", ExpInv
        imMnfRecLen = Len(tmMnf)

        gCenterStdAlone ExpInv
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStdDate
        edcSelCFrom.Text = Format$(slStdDate, "MMM")
        edcSelCFrom1.Text = Year(slStdDate)
        Screen.MousePointer = vbDefault

        ilRet = gPopUserVehicleBox(ExpInv, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVehicles, tgCSVNameCode(), sgCSVNameCodeTag)
        If ilRet <> BTRV_ERR_NONE Then
            On Error GoTo mInitErr
            gBtrvErrorMsg ilRet, "mInit (gPopUserVehiclebox", ExpInv
            Exit Sub
        End If
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
        Unload ExpInv
        igManUnload = NO

End Sub
'
'       Export spots for Invoice Export (Channel One) customized format
'       Spots are exported for a given standard month.  Only spots aired
'       are gathered. Missed/cancelled spots are excluded.
'
Public Function mGetSpotsAndExportInv() As Integer


    Dim ilLoop As Integer
    Dim slStr1 As String
    Dim ilYear As Integer
    Dim ilRet As Integer

    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilCurrentMonth As Integer
    Dim ilWhichKey As Integer
    ReDim llSdfCodes(0 To 0) As Long
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer
    Dim llLoopOnSpots As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilErr As Integer
    Dim ilShowErrorMessage As Integer

    slStr1 = ExpInv!edcSelCFrom.Text             'month in text form (jan..dec)
    ilYear = Val(ExpInv!edcSelCFrom1.Text)

    gGetMonthNoFromString slStr1, ilCurrentMonth         'getmonth #
    If ilCurrentMonth = 0 Then                                 'input isn't text month name, try month #
        ilCurrentMonth = Val(slStr1)
    End If
    slStr1 = Trim$(str(ilCurrentMonth)) & "/15/" & Trim$(ExpInv!edcSelCFrom1.Text)     'form mm/dd/yy

    smStartStd = gObtainStartStd(slStr1)               'obtain std start date for month
    lmStartStd = gDateValue(smStartStd)
    smEndStd = gObtainEndStd(slStr1)                 'obtain std end date for month
    lmEndStd = gDateValue(smEndStd)

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

    tmSpotTypes.iSched = True
    tmSpotTypes.iMissed = False
    tmSpotTypes.iMG = True
    tmSpotTypes.iOutside = True
    tmSpotTypes.iHidden = False
    tmSpotTypes.iCancel = False
    tmSpotTypes.iFill = True

    ilWhichKey = INDEXKEY1              'use key 1 vehicle, date search


    For ilVehicle = 0 To lbcVehicles.ListCount - 1      'loop on all conv/sellingvehicles
        slNameCode = tgCSVNameCode(ilVehicle).sKey             'pick up vehicle code
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        'return array of spot codes, ignore the array of SDFInfo used for sorting
        gObtainSDFByKey hmSdf, ilVefCode, smStartStd, smEndStd, llChfCode, ilWhichKey, llSdfCodes(), tmSdfInfo(), tmSpotTypes

        'If UBound(tmSdfInfo) - 1 > 1 Then
        '    ArraySortTyp fnAV(tmSdfInfo(), 0), UBound(tmSdfInfo), 0, LenB(tmSdfInfo(0)), 0, LenB(tmSdfInfo(0).sKey), 0
        'End If

        ilShowErrorMessage = False
        For llLoopOnSpots = LBound(llSdfCodes) To UBound(llSdfCodes) - 1
            tmSdfSrchKey3.lCode = llSdfCodes(llLoopOnSpots)
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                gLogMsg "Invalid Spot ID: " & Trim$(str(llSdfCodes(llLoopOnSpots))), "ExportInvoice.txt", False
                ''MsgBox "Invalid SpotID: " & Trim$(str(llSdfCodes(llLoopOnSpots)))
                gAutomationAlertAndLogHandler "Invalid SpotID: " & Trim$(str(llSdfCodes(llLoopOnSpots))), vbOkOnly, "ExportInvoice"
            Else
                ilErr = mGetAllTables()
                If ilErr Then
                    ilShowErrorMessage = True
                End If
            End If
        Next llLoopOnSpots
    Next ilVehicle

    mGetSpotsAndExportInv = ilShowErrorMessage

    Erase llSdfCodes
    Erase tmSdfInfo
    Exit Function

End Function
'           mGetallTables - read all supporting files to build export record for spot
'           <input>  none
'           <output> tmExportInfo record contains info from supporting files
'           Return : false = OK, true= some error recorded in invoiceexport.txt
Private Function mGetAllTables() As Integer

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slRecord As String
    Dim llAmt As Long
    Dim slName As String
    Dim slAddr1 As String
    Dim slAddr2 As String
    Dim slAddr3 As String
    Dim ilError As Integer
    Dim slPrice As String

    Dim slSharePct As String
    Dim ilCommPct As Integer
    Dim slStr As String

    On Error GoTo mGetAllTablesErr

        ilError = False             'assume all OK with retrieval of supporting files
        tmExportInfo.sInvType = "Channel 1 (2301)"
        tmExportInfo.sInvDate = smEndStd            'std month end date
        tmExportInfo.sSBUSubCode = "2301"
        tmExportInfo.sProjectCode = ""
        tmExportInfo.sProduct = ""
        tmExportInfo.sSMS = ""
        tmExportInfo.sClient = ""
        tmExportInfo.sVehicle = ""
        tmExportInfo.sProduct = ""
        tmExportInfo.sContact = ""
        tmExportInfo.sLength = ""
        tmExportInfo.sRegion = ""
        tmExportInfo.sSpotType = ""
        tmExportInfo.sISCI = ""
        tmExportInfo.sAirDate = ""
        tmExportInfo.sNetAmt = ".00"
        tmExportInfo.sRetReq = "YES"
        tmExportInfo.sInternalAlloc = ""
        tmExportInfo.sSBU = ""
        tmExportInfo.sInvNumber = ""


        'get contract header
        slName = "Unknown Client"
        slAddr1 = ""
        slAddr2 = ""
        slAddr3 = ""
        tmChfSrchKey.lCode = tmSdf.lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then   'if no contract, do not continue
            ilError = True
            gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Contract ID : " & Trim$(str(tmSdf.lChfCode)), "ExportInvoice.txt", False
        End If

        ''3/13/19: Fields replaced with EDI Client Code and EDI Product Code as the Project and SMS for Channel client that did not use our Traffic system
        'tmExportInfo.sProjectCode = ""
        'tmExportInfo.sSMS = ""
        ''If tmChf.lEffCode > 0 Then
        ''    tmEffSrchKey.lCode = tmChf.lEffCode
        ''    ilRet = btrGetEqual(hmEff, tmEff, imEffRecLen, tmEffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ''    If ilRet = BTRV_ERR_NONE Then
        ''        tmExportInfo.sProjectCode = tmEff.sString(0)
        ''        tmExportInfo.sSMS = tmEff.sString(1)
        ''    Else
        ''        ilError = True
        ''        gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Project Code ID: " & Trim$(str(tmChf.lEffCode)), "ExportInvoice.txt", False
        ''    End If
        ''End If
        tmExportInfo.sProjectCode = ""
        tmExportInfo.sSMS = ""
        If tmChf.lEffCode > 0 Then
            tmEffSrchKey.lCode = tmChf.lEffCode
            ilRet = btrGetEqual(hmEff, tmEff, imEffRecLen, tmEffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                tmExportInfo.sProjectCode = tmEff.sString(0)
                tmExportInfo.sSMS = tmEff.sString(1)
            Else
                ilError = True
                gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Project Code ID: " & Trim$(str(tmChf.lEffCode)), "ExportInvoice.txt", False
            End If
        End If

        tmPnfSrchKey0.iCode = tmChf.iPnfBuyer
        ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tmExportInfo.sContact = tmPnf.sName
        Else
            tmExportInfo.sContact = ""
        End If
        tmExportInfo.sProduct = tmChf.sProduct  'get the product from contract in case it doesnt exist from copy
        'get billing address from either agency or direct advt
        If tmChf.iAgfCode = 0 Then          'test for direct advertiser
            ilLoop = gBinarySearchAdf(tmChf.iAdfCode)
            If ilLoop <> -1 Then
                tmSrchKey.iCode = tmChf.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slName = tmAdf.sName
                    slAddr1 = tmAdf.sBillAddr(0)
                    slAddr2 = tmAdf.sBillAddr(1)
                    slAddr3 = tmAdf.sBillAddr(2)
                    If Trim$(tmAdf.sBillAddr(0)) = "" Then         'use contract address
                        slAddr1 = tmAdf.sCntrAddr(0)
                        slAddr2 = tmAdf.sCntrAddr(1)
                        slAddr3 = tmAdf.sCntrAddr(2)
                    End If
                Else
                    ilError = True
                    gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Advertiser ID : " & Trim$(str(tmChf.iAdfCode)), "ExportInvoice.txt", False
                End If
            End If
        Else
            ilLoop = gBinarySearchAgf(tmChf.iAgfCode)
            If ilLoop <> -1 Then
                tmSrchKey.iCode = tmChf.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slName = tmAgf.sName
                    slAddr1 = tmAgf.sBillAddr(0)
                    slAddr2 = tmAgf.sBillAddr(1)
                    slAddr3 = tmAgf.sBillAddr(2)
                    If Trim$(tmAgf.sBillAddr(0)) = "" Then        'use contract address if no bill address
                        slAddr1 = tmAgf.sCntrAddr(0)
                        slAddr2 = tmAgf.sCntrAddr(1)
                        slAddr3 = tmAgf.sCntrAddr(2)
                    End If
                Else
                    ilError = True
                    gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Agency ID : " & Trim$(str(tmChf.iAdfCode)), "ExportInvoice.txt", False
                End If
            End If
        End If

        tmExportInfo.sClient = Trim$(slName) & "," & Trim$(slAddr1) & "," & Trim$(slAddr2) & "," & Trim$(slAddr3)

        If tmSdf.iVefCode > 0 Then
            ilLoop = gBinarySearchVef(tmSdf.iVefCode)
            If ilLoop <> -1 Then
                tmExportInfo.sVehicle = Trim$(tgMVef(ilLoop).sName)
            Else
                ilError = True
                gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid vehicle ID : " & Trim$(str(tmSdf.iVefCode)), "ExportInvoice.txt", False
             End If
        End If

        tmSlfSrchKey0.iCode = tmChf.iSlfCode(0)
        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If (ilRet = BTRV_ERR_NONE) Then
            tmExportInfo.sSlsp = Trim$(tmSlf.sFirstName) + " " + Trim$(tmSlf.sLastName)
        Else
            ilError = True
            gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Slsp ID : " & Trim$(str(tmSdf.lChfCode)), "ExportInvoice.txt", False

        End If

        'get the schedule line to get the region, if exists
        tmClfSrchKey0.lChfCode = tmChf.lCode
        tmClfSrchKey0.iLine = tmSdf.iLineNo
        tmClfSrchKey0.iCntRevNo = tmChf.iCntRevNo
        tmClfSrchKey0.iPropVer = tmChf.iPropVer
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) Then
            'get spot rate
            ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
            If (InStr(slPrice, ".") <> 0) Then    'found spot cost
                llAmt = gStrDecToLong(slPrice, 2)    'get the actual spot value
                'calc net amt
                If tmChf.iAgfCode = 0 Then          'direct
                    ilCommPct = 10000                'no commission
                Else
                    ilCommPct = (10000 - tmAgf.iComm)
                End If

                slSharePct = gIntToStrDec(ilCommPct, 4)
                slStr = gMulStr(slSharePct, slPrice)                       ' gross portion of possible split
                tmExportInfo.sNetAmt = gRoundStr(slStr, ".01", 2)
            ElseIf Trim$(slPrice) = "ADU" Then
                tmExportInfo.sSpotType = slPrice
            ElseIf Trim$(slPrice) = "Bonus" Then
                tmExportInfo.sSpotType = slPrice
            ElseIf Trim$(slPrice) = "+ Fill" Then
                tmExportInfo.sSpotType = "Bonus"
            ElseIf Trim$(slPrice) = "- Fill" Then
                tmExportInfo.sSpotType = "Bonus"
            'ElseIf Trim$(slPrice) = "N/C" Then

            ElseIf Trim$(slPrice) = "Recapturable" Then
                tmExportInfo.sSpotType = slPrice
            ElseIf Trim$(slPrice) = "Spinoff" Then
                tmExportInfo.sSpotType = slPrice
            ElseIf Trim$(slPrice) = "MG" Then
                tmExportInfo.sSpotType = "MG"

            End If



            'see if region exists
            If tmClf.lRafCode > 0 Then
                tmRafSrchKey0.lCode = tmClf.lRafCode
                ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    tmExportInfo.sRegion = tmRaf.sName
                Else
                    ilError = True
                    gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Region ID : " & Trim$(str(tmClf.lRafCode)), "ExportInvoice.txt", False
                End If
            End If
        Else
            ilError = True
            gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Line does not exist : " & Trim$(str(tmSdf.iLineNo)), "ExportInvoice.txt", False
        End If

        tmExportInfo.sLength = Trim$(str(tmSdf.iLen))       'spot length


        If tmSdf.sPtType = "1" Then                 'copy instr
            tmCifSrchKey.lCode = tmSdf.lCopyCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                tmCpfSrchKey.lCode = tmCif.lcpfCode     'product/isci
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) Then
                    tmExportInfo.sProduct = Trim$(tmCpf.sName)
                    tmExportInfo.sISCI = Trim$(tmCpf.sISCI)
                Else
                    ilError = True
                    gLogMsg "Spot ID " & Trim$(tmSdf.lCode) & ": " & "Invalid Copy ID : " & Trim$(str(tmSdf.lCopyCode)), "ExportInvoice.txt", False
                End If
            End If
        End If

        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), tmExportInfo.sAirDate

        slRecord = """" & Trim$(tmExportInfo.sInvType) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sInvDate) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sSBUSubCode) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sProjectCode) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sSMS) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sSlsp) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sClient) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sVehicle) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sProduct) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sContact) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sLength) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sRegion) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sSpotType) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sISCI) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sAirDate) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sNetAmt) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sRetReq) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sInternalAlloc) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sSBU) & """" & ","
        slRecord = slRecord & """" & Trim$(tmExportInfo.sInvNumber) & """" & ","
        Print #hmTo, slRecord

        mGetAllTables = ilError     'return error flag
        Exit Function
mGetAllTablesErr:
        Resume Next
End Function



