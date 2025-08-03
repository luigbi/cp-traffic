VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmArchSpotRetrv 
   Caption         =   "Archived Spot Retriever"
   ClientHeight    =   8145
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14610
   Icon            =   "ArchSpotRetrv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFound 
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   5280
      Width           =   2100
   End
   Begin VB.TextBox txtRecsSearched 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   5280
      Width           =   2100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   7200
      Width           =   2685
   End
   Begin ArchiveSpotRetriever.CSI_Calendar CSI_Calendar2 
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Text            =   "2/11/15"
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
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin ArchiveSpotRetriever.CSI_Calendar CSI_Calendar1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Text            =   "2/11/15"
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
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.CheckBox ckcAllAdv 
      Caption         =   "All Advertisers"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox lbcAdvertiser 
      Height          =   3375
      ItemData        =   "ArchSpotRetrv.frx":08CA
      Left            =   360
      List            =   "ArchSpotRetrv.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   13860
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   12360
      TabIndex        =   6
      Top             =   6480
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11160
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   7200
      Width           =   2685
   End
   Begin VB.TextBox txtBrowse 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6480
      Width           =   11655
   End
   Begin VB.Label txtRecsFound 
      Caption         =   "Total Records Found:"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Caption         =   "Total Records Scanned:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label lblBrowse 
      Caption         =   "Save Results To:"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblEndDate 
      Caption         =   "End Date:"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblStartDate 
      Caption         =   "Start Date:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmArchSpotRetrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'*  frmArchSpotRetrv- Retrieve Archived Spots and Write to a CSV file
'*
'*  Doug Smith
'*
'*  Copyright 2015 Counterpoint Software, Inc.
'***********************************************************************
Option Explicit
Option Compare Text

Private Const FORMNAME As String = "frmArchSpotRetrv"
Private bmNoPervasive As Boolean
Private imAllAdvertisersClick As Integer
Private smStartDate As String
Private smEndDate As String
Private smAdvSelected As String
Private smArchivePath As String
Private smSelArchiveMonths() As String
Private smSelArchiveFiles() As String
Private smFileTitleName As String
Private lmStartDate As Long
Private lmEndDate As Long
Private lmRecSearched As Long
Private smExpDirLen As Integer
Private smAdvLen As Integer
Private smDashStartDate As String
Private smDashEndDate As String
Private smExportPath As String
Private smSelAdv() As String
Private lmTtlSpotsFound As Long
Private lmTtlScanned As Long
Private smAdvertiser As String

Private Sub ckcAllAdv_Click()

    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    On Error GoTo ErrHandler
    If imAllAdvertisersClick Then
        Exit Sub
    End If
    If ckcAllAdv.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcAdvertiser.ListCount > 0 Then
        imAllAdvertisersClick = True
        lRg = CLng(lbcAdvertiser.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcAdvertiser.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllAdvertisersClick = False
    End If
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mGetArchiveFolders"
End Sub

Private Sub cmcBrowse_Click()
    Dim slCurDir As String
    Dim ilPos As Integer
    
    If smFileTitleName = "" Then
        MsgBox "Please Select an Advertiser "
        Exit Sub
    End If
    slCurDir = CurDir
    txtBrowse.Text = ""
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "CSV Files (*.csv)"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    mDashStartDate
    mDashEndDate
    ' Display the Open dialog box
    CommonDialog1.fileName = smFileTitleName
    CommonDialog1.InitDir = smExportPath
    CommonDialog1.ShowOpen
    ' Display name of selected file
    txtBrowse.Text = Trim$(CommonDialog1.fileName)
   
   ilPos = InStrRev(CommonDialog1.fileName, "\")
   smExportPath = Left(CommonDialog1.fileName, ilPos)
   lbcAdvertiser_Click

    DoEvents
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Function InitPervAndGlobals() As Boolean
    
    Dim sBuffer As String
    Dim lSize As Long
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim ilValue As Integer
    Dim ilValue8 As Integer
    Dim slDate As String
    Dim ilDatabase As Integer
    Dim ilLocation As Integer
    Dim ilSQL As Integer
    Dim ilForm As Integer
    Dim sMsg As String
    Dim iLoop As Integer
    Dim sCurDate As String
    Dim sAutoLogin As String
    Dim slTimeOut As String
    Dim slDSN As String
    Dim slStartIn As String
    Dim slStartStdMo As String
    Dim slTemp As String
    ReDim sWin(0 To 13) As String * 1
    Dim ilIsTntEmpty As Integer
    Dim ilIsShttEmpty As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    Dim EmailExists_rst As ADODB.Recordset
    Dim blAddGuide As Boolean
    Dim blNeedToCloseCnn As Boolean
    Dim slXMLINIInputFile As String
    Dim slRootPath As String
    Dim slPhotoPath As String
    Dim slBitmapPath As String
    Dim slLogoPath As String
    
    InitPervAndGlobals = False
    sgCommand = Command$
    blNeedToCloseCnn = False
    'Warning: One thing to remember is that if you are expecting a return value from a gMsgBox
    'and you turn gMsgBox off then you need to make sure that you handle that case.
    'example:   ilRet = gMsgBox "xxxx"
    igShowMsgBox = True
    igDemoMode = False
    If InStr(sgCommand, "Demo") Then
        igDemoMode = True
    End If
    'Used to speed-up testing exports with multiple files reduce record count needed to create a new file
    igSmallFiles = False
    If InStr(sgCommand, "SmallFiles") Then
        igSmallFiles = True
    End If
    
    igAutoImport = False
    If InStr(sgCommand, "AutoImport") Then
        igAutoImport = True
        igShowMsgBox = False
    End If
    slStartIn = CurDir$
    sgCurDir = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommand, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    sgBS = Chr$(8)  'Backspace
    sgTB = Chr$(9)  'Tab
    sgLF = Chr$(10) 'Line Feed (New Line)
    sgCR = Chr$(13) 'Carriage Return
    sgCRLF = sgCR + sgLF
    ilRet = 0
    ilLocation = False
    ilDatabase = False
    sgDatabaseName = ""
    sgReportDirectory = ""
    sgExportDirectory = ""
    sgImportDirectory = ""
    sgExeDirectory = ""
    sgLogoDirectory = ""
    sgPasswordAddition = ""
    sgSQLDateForm = "yyyy-mm-dd"
    sgCrystalDateForm = "yyyy,mm,dd"
    sgSQLTimeForm = "hh:mm:ss"
    igSQLSpec = 1               'Pervasive 2000
    sgShowDateForm = "m/d/yyyy"
    sgShowTimeWOSecForm = "h:mma/p"
    sgShowTimeWSecForm = "h:mm:ssa/p"
    igWaitCount = 10
    igTimeOut = -1
    sgWallpaper = ""
    bgIgnoreDuplicateError = False
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    sgLogoName = "rptlogo.bmp"
    sgNowDate = ""
    ilPos = InStr(1, sgCommand, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommand, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommand, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gIsDate(slDate) Then
            sgNowDate = slDate
        End If
    End If
    bgReportQueue = False
    ilPos = InStr(1, sgCommand, "/Q", 1)
    If ilPos > 0 Then
        bgReportQueue = True
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    If Not gLoadOption("Database", "Name", sgDatabaseName) Then
        gMsgBox "Affiliat.Ini [Database] 'Name' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    End If
    If Not gLoadOption("Locations", "Reports", sgReportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Reports' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    End If
    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Export' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    End If
    If Not gLoadOption("Locations", "Exe", sgExeDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Exe' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    End If
    'Import is optional
    If gLoadOption("Locations", "Import", sgImportDirectory) Then
        sgImportDirectory = gSetPathEndSlash(sgImportDirectory, True)
    Else
        sgImportDirectory = ""
    End If
    
    If gLoadOption("Locations", "ContractPDF", sgContractPDFPath) Then
        sgContractPDFPath = gSetPathEndSlash(sgContractPDFPath, True)
    Else
        sgContractPDFPath = ""
    End If
    '5676 sgRootDrive.  If drive C doesn't exist, look for RootDrive in ini file, then test that value to make sure it exists.
    If Dir("c:\") = "" Then
        If gLoadOption("Locations", "RootDrive", sgRootDrive) Then
            sgRootDrive = gSetPathEndSlash(sgRootDrive, True)
            If Dir(sgRootDrive) = "" Then
                sgRootDrive = "C:\"
            End If
        Else
            sgRootDrive = "C:\"
        End If
    Else
        sgRootDrive = "C:\"
    End If
    sgReportDirectory = gSetPathEndSlash(sgReportDirectory, True)
    sgExportDirectory = gSetPathEndSlash(sgExportDirectory, True)
    smExpDirLen = Len(sgExportDirectory)
    sgExeDirectory = gSetPathEndSlash(sgExeDirectory, True)
    sgLogoDirectory = gSetPathEndSlash(sgLogoDirectory, True)
    Call gLoadOption("SQLSpec", "Date", sgSQLDateForm)
    Call gLoadOption("SQLSpec", "Time", sgSQLTimeForm)
    If gLoadOption("SQLSpec", "System", sBuffer) Then
        If sBuffer = "P7" Then
            igSQLSpec = 0
        End If
    End If
    If gLoadOption("Locations", "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption("Locations", "Wallpaper", sgWallpaper)
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    If Not gLoadOption("Locations", "DBPath", sgDBPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload frmArchSpotRetrv
        Exit Function
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    'Set Message folder
    If Not gLoadOption("Locations", "DBPath", sgMsgDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Exit Function
    Else
        sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory, True) & "Messages\"
    End If
    On Error GoTo ErrHandler
    Set cnn = New ADODB.Connection
    slDSN = sgDatabaseName
    On Error GoTo ERRNOPERVASIVE
    ilRet = 0
    cnn.Open "DSN=" & slDSN
    On Error GoTo ErrHandler
    If ilRet = 1 Then
        Sleep 2000
        cnn.Open "DSN=" & slDSN
    End If
    Set rst = New ADODB.Recordset
    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(slDSN, Len(slDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        slDSN = Left(slDSN, Len(slDSN) - 1)
    End If
    'Start the Pervasive API engine
    If Not mOpenPervasiveAPI Then
        Unload frmArchSpotRetrv
        Exit Function
    End If
Exit Function
    
TableDoesNotExist:
    ilRet = False
    Resume Next

mTrafficStartUpErr:
    ilRet = Err.Number
    Resume Next

mReadFileErr:
    ilRet = Err.Number
    Resume Next
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-InitPervAndGlobals"
    bmNoPervasive = True
    If blNeedToCloseCnn Then
        cnn.RollbackTrans
    End If
End Function

Private Sub cmcBrowse_LostFocus()
    'lbcAdvertiser_Click
End Sub

Private Sub cmdCancel_Click()
    
    Dim ilYesNo As Integer
    
    If cmdCancel.Caption <> "Done" Then
        ilYesNo = gMsgBox("Are you sure that you want to cancel the program?", vbYesNo)
        If ilYesNo = vbYes Then
            gLogMsg "** User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **", "SpotArchiveRetrieval.Txt", False
            gLogMsg " ", "SpotArchiveRetrieval.Txt", False
            Unload frmArchSpotRetrv
        End If
    Else
        Unload frmArchSpotRetrv
    End If
End Sub

Private Sub cmdStart_Click()

    Dim ilLoop As Integer
    Dim slTemp As String
    Dim ilRet As Boolean
    Dim blFound As Boolean
    
    On Error GoTo ErrHandler
    
    lmRecSearched = 0
    cmdCancel.Caption = "Cancel"
    mCheckGG
    If (igGGFlag = 0) And (igRptGGFlag = 0) Then
        Exit Sub
    End If
    lmTtlSpotsFound = 0
    lmTtlScanned = 0
    blFound = False
    For ilLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
        If lbcAdvertiser.Selected(ilLoop) Then
            smAdvertiser = Trim(lbcAdvertiser.List(ilLoop))
            blFound = True
            Exit For
        End If
    Next ilLoop
    If Not blFound Then
        MsgBox "Please Select an Advertiser Prior to Pressing Start."
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    smStartDate = CSI_Calendar1.Text
    If smStartDate = "" Or Not gIsDate(smStartDate) Then
        MsgBox "Start Date Must Have a Valid Date"
        CSI_Calendar1.SetFocus
        Exit Sub
    End If
    
    smEndDate = CSI_Calendar2.Text
    If smEndDate = "" Or Not gIsDate(smEndDate) Then
        MsgBox "End Date Must Have a Valid Date"
        CSI_Calendar2.SetFocus
        Exit Sub
    End If
    slTemp = mGetSubFolders(smArchivePath)
    If InStr(slTemp, "Sorry") Then
        MsgBox slTemp & sgCR & sgLF & sgCR & sgLF & "Currently, there are No archived spots available for retrieval."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    smExpDirLen = Len(smExportPath)
    smFileTitleName = Mid(txtBrowse.Text, smExpDirLen + 1, Len(txtBrowse.Text))
    If Not mCheckFileName(smFileTitleName) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lbcAdvertiser.Enabled = False
    txtBrowse.Enabled = False
    cmdStart.Enabled = False
    cmcBrowse.Enabled = False
    lblStatus.Caption = "Scanning record number: " & CStr(txtRecsSearched) & " for " & Trim(lbcAdvertiser.List(ilLoop)) & " Spots"
    gLogMsg " Starting Archived Spot Retrieval For: " & Trim(lbcAdvertiser.List(ilLoop)) & " For the date range: " & smStartDate & " - " & smEndDate, "SpotArchiveRetrieval.Txt", False
    gLogMsg "    Building File: " & smFileTitleName, "SpotArchiveRetrieval.Txt", False
    mGetArchiveFolders
    mProcessSelArchFiles
    cmdCancel.Caption = "Done"
    lbcAdvertiser.Enabled = True
    txtBrowse.Enabled = True
    cmdStart.Enabled = True
    cmcBrowse.Enabled = True
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-cmdStart_Click"
End Sub

Private Sub CSI_Calendar1_LostFocus()
    mDashStartDate
    lbcAdvertiser_Click
End Sub

Private Sub CSI_Calendar2_LostFocus()
    mDashEndDate
    lbcAdvertiser_Click
End Sub

Private Sub Form_Initialize()
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2.5
    gSetFonts Me
End Sub

Private Sub Form_Load()

    Dim sIniValue As String
    Dim slLocation As String
    Dim sFileName As String
    Dim slTemp As String

    On Error GoTo ErrHandler
    InitPervAndGlobals
    slTemp = sgIniPathFileName
    sgIniPathFileName = sgStartupDirectory & "\Traffic.Ini"
    Call gLoadOption("Locations", "ARCHIVE", smArchivePath)
    CSI_Calendar1.Text = gNow
    CSI_Calendar1_LostFocus
    CSI_Calendar2.Text = gNow
    CSI_Calendar2_LostFocus
    txtBrowse.Text = sgExportDirectory
    smExportPath = sgExportDirectory
    sgIniPathFileName = slTemp
    Call gPopAdvertisers
    imAllAdvertisersClick = False
    ReDim smSelAdv(0 To 0) As String
    mLoadAdv
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-Form_Load"
End Sub

Private Sub mLoadAdv()

    Dim ilAdf As Integer
    
    On Error GoTo ErrHandler
    For ilAdf = 0 To UBound(tgAdvtInfo) - 1 Step 1
        lbcAdvertiser.AddItem tgAdvtInfo(ilAdf).sAdvtName
        lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = tgAdvtInfo(ilAdf).iCode
    Next ilAdf
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mLoadAdv"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase smSelArchiveMonths
    Erase smSelArchiveFiles
    Set frmArchSpotRetrv = Nothing
    cnn.Close
    
    btrStopAppl
    End
End Sub

Private Sub lbcAdvertiser_Click()

    Dim ilLoop As Integer
    'Dim slStartDate As String
    'Dim slEndDate As String
    Dim slAdvtAbbr As String
    Dim llCode As Long
    Dim llAdfCode As Long
    
    On Error GoTo ErrHandler
    If imAllAdvertisersClick Then
        Exit Sub
    End If
    txtBrowse.Text = ""
    txtBrowse.Text = smExportPath
    If ckcAllAdv.Value = vbChecked Then
        imAllAdvertisersClick = True
        ckcAllAdv = vbUnchecked
        imAllAdvertisersClick = False
    End If
    For ilLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
        If lbcAdvertiser.Selected(ilLoop) Then
            smAdvSelected = Trim(lbcAdvertiser.List(ilLoop))
            llCode = lbcAdvertiser.ItemData(ilLoop)
            smFileTitleName = smAdvSelected
            llAdfCode = gBinarySearchAdf(llCode)
            slAdvtAbbr = tgAdvtInfo(llAdfCode).sAdvtAbbr
        End If
    Next ilLoop
    If smFileTitleName <> "" Then
        txtBrowse.Text = txtBrowse.Text & smFileTitleName & "_" & smDashStartDate & "_" & smDashEndDate & ".csv"
    End If
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-lbcAdvertiser_Click"
End Sub

Private Function mGetArchiveFolders() As Boolean

    Dim slUserStartDate As String
    Dim slUserEndDate As String
    Dim slArchStartDate As String
    Dim slArchEndDate As String
    Dim blFoundFirst As Boolean
    Dim blNeedToCloseCnn As Boolean
    Dim ilLen As Integer
    Dim ilIdx As Integer
    Dim slfolderName As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim fs, f, f1, fc, s
    
    On Error GoTo ErrHandler
    
    'Debug
    'smStartDate = "8/5/13"
    'smEndDate = "9/1/13"
    mGetArchiveFolders = False
    ReDim smSelArchiveMonths(0 To 0) As String
    ilIdx = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(smArchivePath)
    Set fc = f.SubFolders
    blFoundFirst = False
    
    blFoundFirst = False
    For Each f1 In fc
        slfolderName = f1.Name
        ilLen = Len(slfolderName)
        slYear = Mid(slfolderName, 8, 4)
        slMonth = Mid(slfolderName, 12, 2)
        slDay = Mid(slfolderName, 14, 2)
        slArchEndDate = slMonth & "/" & slDay & "/" & slYear
        slArchStartDate = gObtainStartStd(slArchEndDate)
        slUserStartDate = smStartDate
        slUserEndDate = smEndDate
        If Not blFoundFirst And gDateValue(smStartDate) <= gDateValue(slArchEndDate) Then
            smSelArchiveMonths(ilIdx) = slfolderName
            ReDim Preserve smSelArchiveMonths(0 To UBound(smSelArchiveMonths) + 1)
            ilIdx = ilIdx + 1
            blFoundFirst = True
        Else
            If blFoundFirst And gDateValue(smEndDate) >= gDateValue(slArchStartDate) Then
                smSelArchiveMonths(ilIdx) = slfolderName
                ReDim Preserve smSelArchiveMonths(0 To UBound(smSelArchiveMonths) + 1)
                ilIdx = ilIdx + 1
            End If
        End If
    Next
    Set fs = Nothing
    mGetArchiveFolders = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mGetArchiveFolders"
End Function

Private Function mProcessSelArchFiles()

    Dim ilLoop As Integer
    Dim ilIdx As Integer
    Dim ilIndex As Integer
    Dim ilLen As Integer
    Dim llArchDate As Long
    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim tlTxtStream2 As TextStream
    Dim fs2 As New FileSystemObject
    Dim olRetString As TextStream
    Dim alRecordsArray() As String
    Dim slTmpStr As String
    Dim slRetString As String
    Dim slTemp As String
    Dim slTemp2 As String
    Dim slTempStartDate As String
    Dim slTempEndDate As String
    Dim fso, A
    
    On Error GoTo ErrHandler
    ilIdx = 0
    'For speed purposes load all of the selected adv. into an array.  This will normally be a subset of the total advertisers
    'so we should have considearbly less loops to test
    For ilLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
        If lbcAdvertiser.Selected(ilLoop) Then
            smSelAdv(ilIdx) = Trim(lbcAdvertiser.List(ilLoop))
            ilIdx = ilIdx + 1
            ReDim Preserve smSelAdv(0 To ilIdx)
        End If
    Next ilLoop
    slTemp = txtBrowse.Text
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set A = fso.CreateTextFile(slTemp, True)
    Screen.MousePointer = vbHourglass
    lmStartDate = gDateValue(smStartDate)
    lmEndDate = gDateValue(smEndDate)
    For ilLoop = 0 To UBound(smSelArchiveMonths) - 1 Step 1
        slTemp = smArchivePath & "\" & smSelArchiveMonths(ilLoop)
        mGetMonthFileList slTemp, ilLoop
        For ilIndex = 0 To UBound(smSelArchiveFiles) - 1
            DoEvents
            slTemp2 = slTemp & "\" & smSelArchiveFiles(ilIndex)
            If fs.FILEEXISTS(slTemp2) Then
                Set tlTxtStream = fs.OpenTextFile(slTemp2, ForReading, False)
            Else
                MsgBox "** No Data Available **"
                Exit Function
            End If
            If ilLoop = 0 Then
                slRetString = tlTxtStream.ReadLine
                A.WriteLine slRetString
            End If
            If ilLoop > 0 Then
                tlTxtStream.ReadLine
                
            End If
            Do While tlTxtStream.AtEndOfStream <> True
                slRetString = tlTxtStream.ReadLine
                alRecordsArray = Split(slRetString, ",")
                If Not IsArray(alRecordsArray) Then
                    Exit Function
                End If
                If UBound(alRecordsArray) < 1 Then
                    Exit Function
                End If
                'Used to compare feed date
                'llArchDate = gDateValue(alRecordsArray(11))
                'Used to compare feed date
                llArchDate = gDateValue(alRecordsArray(26))
                ilLen = Len(alRecordsArray(7))
                For ilIdx = 0 To UBound(smSelAdv) - 1 Step 1
                    DoEvents
                    lmTtlScanned = lmTtlScanned + 1
                    If lmTtlScanned Mod 1000 = 0 Then
                        txtRecsSearched.Text = lmTtlScanned
                    End If
                    slTmpStr = Mid(alRecordsArray(7), 2, ilLen - 2)
                    If StrComp(Trim(smSelAdv(ilIdx)), Trim(slTmpStr)) = 0 Then
                        If llArchDate >= lmStartDate And llArchDate <= lmEndDate Then
                            A.WriteLine slRetString
                            lmTtlSpotsFound = lmTtlSpotsFound + 1
                            If lmTtlSpotsFound Mod 100 = 0 Then
                                txtFound.Text = lmTtlSpotsFound
                            End If
                        End If
                    End If
                Next ilIdx
            Loop
        Next ilIndex
    Next ilLoop
    Set fso = Nothing
    txtFound.Text = lmTtlSpotsFound
    txtRecsSearched.Text = lmTtlScanned
    gLogMsg "    Within the Date Range of: " & smStartDate & " - " & smEndDate & ".  " & CStr(lmTtlSpotsFound) & "  Archived Spots Were Found", "SpotArchiveRetrieval.Txt", False
    gLogMsg "    Total Records Scanned: " & CStr(lmTtlScanned), "SpotArchiveRetrieval.Txt", False
    gLogMsg " Ending Archived Spot Retrieval For: " & Trim(smAdvertiser) & " For the date range: " & smStartDate & " - " & smEndDate, "SpotArchiveRetrieval.Txt", False
    gLogMsg "  ", "SpotArchiveRetrieval.Txt", False
    DoEvents
    If ilLoop = 0 Then
        MsgBox "No Archive Data was Found in the Date Range of:  " & smStartDate & " - " & smEndDate
    Else
        MsgBox "Within the Date Range of:  " & smStartDate & " - " & smEndDate & "  " & CStr(lmTtlSpotsFound) & "  Archived Spots Were Found"
    End If
    DoEvents
    Screen.MousePointer = vbDefault
    tlTxtStream.Close
    A.Close
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mProcessSelArchFiles"
End Function

Private Function mGetSubFolders(folderspec As String) As String
  
    'Get all of the sub folders under the passed in path
    Dim fso, f, f1, s, sf
  
    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 In sf
        s = s & f1.Name
        s = s & sgCRLF
    Next
    Set fso = Nothing
    mGetSubFolders = s
    Exit Function
ErrHandler:
    mGetSubFolders = "Sorry, folder: " & folderspec & " was not found."
    Screen.MousePointer = vbDefault
    'gHandleError "", FORMNAME & "-mGetSubFolders"
End Function

Sub mGetMonthFileList(sFolderLoc As String, iIdx As Integer)
    'The Archive Program is putting out a single file for the entire month.
    Dim fs, f, f1, fc, s
        
    On Error GoTo ErrHandler
    ReDim Preserve smSelArchiveFiles(0 To 0)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(sFolderLoc)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        smSelArchiveFiles(UBound(smSelArchiveFiles)) = s
        ReDim Preserve smSelArchiveFiles(0 To UBound(smSelArchiveFiles) + 1)
    Next
    Set fs = Nothing
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mGetMonthFileList"
End Sub

Sub mGetDayFileList(sFolderLoc As String, iIdx As Integer)
    'The Archive Program is putting out a file for each day of the month.
    Dim fs, f, f1, fc, s
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim llDate As Long
        
    On Error GoTo ErrHandler
    ReDim Preserve smSelArchiveFiles(0 To 0)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(sFolderLoc)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        slDay = Mid(s, 15, 2)
        slYear = Mid(s, 9, 4)
        slMonth = Mid(s, 13, 2)
        slDate = slMonth & "/" & slDay & "/" & slYear
        llDate = gDateValue(slDate)
        If llDate >= lmStartDate And llDate <= lmEndDate Then
            smSelArchiveFiles(UBound(smSelArchiveFiles)) = s
            ReDim Preserve smSelArchiveFiles(0 To UBound(smSelArchiveFiles) + 1)
        End If
    Next
    Set fs = Nothing
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mGetMonthFileList"
End Sub


Private Function mCheckFileName(mFileName As String) As Boolean

    On Error GoTo ErrHandler
    mCheckFileName = False
    If InStr(mFileName, ":") Then
        MsgBox ("The file name has an illegal Colon Character: " & "  : " & vbCrLf & "Please delete the Colon Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, "<") Then
        MsgBox ("The file name has an illegal Less Than Character: " & " < " & vbCrLf & "Please delete the Less Than Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, ">") Then
        MsgBox ("The file name has an illegal Greater Than Character: " & " > " & vbCrLf & "Please delete the Greater Than Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, """") Then
        MsgBox ("The file name has an illegal Double Quote Character: " & """" & vbCrLf & "Please delete the Double Quote Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, "/") Then
        MsgBox ("The file name has an illegal Forward Slash Character: " & " / " & vbCrLf & "Please delete the Forward Slash Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, "\") Then
        MsgBox ("The file name has an illegal Back Slash Character: " & " \ " & vbCrLf & "Please delete the Back Slash Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, "|") Then
        MsgBox ("The file name has an illegal Pipe Sign Character: " & " | " & vbCrLf & "Please delete the Pipe Sign Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, "?") Then
        MsgBox ("The file name has an illegal Question Mark Character: " & " ? " & vbCrLf & "Please delete the Question Mark Character or use another Character")
        Exit Function
    End If
    If InStr(mFileName, "*") Then
        MsgBox ("The file name has an illegal Asterisk Character: " & " * " & vbCrLf & "Please delete the Asterisk Character or use another Character")
        Exit Function
    End If
    mCheckFileName = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-mCheckFileName"
End Function


Private Sub mDashStartDate()
    smDashStartDate = Replace(CSI_Calendar1.Text, "/", "-")
End Sub


Private Sub mDashEndDate()
    smDashEndDate = Replace(CSI_Calendar2.Text, "/", "-")
End Sub

Private Sub mCheckGG()
    Dim c As Integer
    Dim slName As String
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llNow As Long
    
    Dim gg_rst As ADODB.Recordset
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    SQLQuery = "Select safName From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set gg_rst = cnn.Execute(SQLQuery)
    If Not gg_rst.EOF Then
        slName = Trim$(gg_rst!safName)
        ilField1 = Asc(slName)
        slStr = Mid$(slName, 2, 5)
        llDate = Val(slStr)
        llNow = gDateValue(Format$(Now, "m/d/yy"))
        ilField2 = Asc(Mid$(slName, 11, 1))
        If (ilField1 = 0) And (ilField2 = 1) Then
            If llDate <= llNow Then
                ilField2 = 0
            End If
        End If
        If (ilField1 = 0) And (ilField2 = 0) Then
            igGGFlag = 0
        End If
        gSetRptGGFlag slName
    End If
    gg_rst.Close
End Sub

