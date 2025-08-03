VERSION 5.00
Begin VB.Form EngrLogIn 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Icon            =   "EngrLogIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.PictureBox plcSignon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   -15
      Picture         =   "EngrLogIn.frx":030A
      ScaleHeight     =   3555
      ScaleWidth      =   4335
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   4395
      Begin VB.TextBox edcPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1575
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1080
         Width           =   2130
      End
      Begin VB.PictureBox plcPassword 
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
         Height          =   345
         Left            =   1545
         ScaleHeight     =   285
         ScaleWidth      =   2130
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2190
      End
      Begin VB.TextBox edcName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1575
         TabIndex        =   3
         Top             =   690
         Width           =   2130
      End
      Begin VB.PictureBox plcName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1545
         ScaleHeight     =   285
         ScaleWidth      =   2130
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   2190
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   45
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   930
         Width           =   105
      End
      Begin VB.Label lacTestMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Test System       Test System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   45
         TabIndex        =   14
         Top             =   3135
         Visible         =   0   'False
         Width           =   3390
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   3405
         TabIndex        =   13
         Top             =   1695
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   3210
      End
      Begin VB.Label lacUserName 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   345
         TabIndex        =   1
         Top             =   690
         Width           =   1170
      End
      Begin VB.Label lacPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   435
         TabIndex        =   4
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label lacStart 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         Top             =   1785
         Width           =   1350
      End
      Begin VB.Label lacExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   9
         Top             =   2295
         Width           =   1335
      End
   End
   Begin VB.PictureBox pbc256LogIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   165
      Picture         =   "EngrLogIn.frx":34E34
      ScaleHeight     =   630
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmcStart 
      Appearance      =   0  'Flat
      Caption         =   "&START System"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   8
      Top             =   1785
      Width           =   1365
   End
   Begin VB.CommandButton cmcExit 
      Appearance      =   0  'Flat
      Caption         =   "&START System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      TabIndex        =   10
      Top             =   2295
      Width           =   1365
   End
End
Attribute VB_Name = "EngrLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrLogIn - basic log-on form for SSQL server
'*
'*  Created Aug, 2004
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Dim smSpecial As String
Const USERINDEX = 12



'--------------------------------------------------
'
'   Multi-Automation allowed:  see EngrMain for design ideas
'
'--------------------------------------------------

Private Sub cmcCSLogo_Click()
    imcOutline.Visible = False
    EngrAbout.Show vbModal
End Sub

Private Sub cmcExit_Click()
    mExit
End Sub

Private Sub cmcExit_GotFocus()
    imcOutline.Move 1560, 2280, 1365, 360
    imcOutline.Visible = True
End Sub

Private Sub cmcExit_LostFocus()
    imcOutline.Visible = False
End Sub

Private Sub cmcExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
    imcOutline.Visible = True
End Sub

Private Sub cmcStart_Click()
    mStart
End Sub

Private Sub cmcStart_GotFocus()
    imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
    imcOutline.Visible = True
End Sub

Private Sub cmcStart_LostFocus()
    imcOutline.Visible = False
End Sub

Private Sub cmcStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
    imcOutline.Visible = True
End Sub

Private Sub Form_Load()
    Dim slBuffer As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim ilDatabase As Integer
    Dim ilLocation As Integer
    Dim ilSQL As Integer
    Dim ilForm As Integer
    Dim slAutoLogin As String
    Dim slTimeOut As String
    Dim slDSN As String
    Dim slStartIn As String
    Dim slTime As String
    Dim slDatabase As String
    Dim slLocations As String
    Dim slCartUnloadTime As String
    
    sgCommand = Command$
    
    If App.PrevInstance Then
        MsgBox "Only one copy of Engineering can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        End
    End If
    
    igOperationMode = 0

    slStartIn = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
        slLocations = "Locations"
        slDatabase = "Database"
        lacTestMsg.Visible = False
    Else
        igTestSystem = True
        slLocations = "TestLocations"
        slDatabase = "TestDatabase"
        lacTestMsg.Visible = True
    End If
    
    mGetResolution
    
    gInitVar
   
   
    ilRet = 0
    ilLocation = False
    ilDatabase = False
    sgDatabaseName = ""
    sgReportDirectory = ""
    sgExportDirectory = ""
    sgImportDirectory = ""
    sgLogoDirectory = ""
    sgSQLDateForm = "yyyy-mm-dd"
    sgSQLTimeForm = "hh:mm:ss"
    igSQLSpec = 1               'Pervasive 2000
    sgShowDateForm = "m/d/yyyy"
    sgShowTimeWOSecForm = "hh:mm"
    sgShowTimeWSecForm = "hh:mm:ss"
    igWaitCount = 10
    igTimeOut = -1
    igBkgdProg = 0
    lgCartUnloadTime = 10
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Engineer.Ini"
    
    ilRet = 0
    On Error GoTo mReadFileErr
    slTime = FileDateTime(sgIniPathFileName)
    If ilRet <> 0 Then
        MsgBox "Engineer.Ini missing from " & sgStartupDirectory, vbCritical
        Unload EngrLogIn
        Exit Sub
    End If
    sgNowDate = ""
    ilPos = InStr(1, sgCommand, "/D:", vbTextCompare)
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
    
    sgClientFields = "A"
    ilPos = InStr(1, sgCommand, "/Demo", vbTextCompare)
    If ilPos > 0 Then
        sgClientFields = ""
    End If
    ilPos = InStr(1, sgCommand, "/WWO", vbTextCompare)
    If ilPos > 0 Then
        sgClientFields = "W"
    End If
    
    igRunningFrom = 1
    ilPos = InStr(1, sgCommand, "/Server", 1)
    If ilPos > 0 Then
        igRunningFrom = 0
    End If
    
    If Not gLoadOption(slDatabase, "Name", sgDatabaseName) Then
        MsgBox "Engineer.Ini [" & slDatabase & "] 'Name' key is missing.", vbCritical
        Unload EngrLogIn
        Exit Sub
    End If
    If Not gLoadOption(slLocations, "Reports", sgReportDirectory) Then
        MsgBox "Engineer.Ini [" & slLocations & "] 'Reports' key is missing.", vbCritical
        Unload EngrLogIn
        Exit Sub
    End If
    If Not gLoadOption(slLocations, "Export", sgExportDirectory) Then
        MsgBox "Engineer.Ini [" & slLocations & "] 'Export' key is missing.", vbCritical
        Unload EngrLogIn
        Exit Sub
    End If
    If Not gLoadOption(slLocations, "Logo", sgLogoDirectory) Then
        MsgBox "Engineer.Ini [" & slLocations & "] 'Logo' key is missing.", vbCritical
        Unload EngrLogIn
        Exit Sub
    End If
    'Import is optional
    If gLoadOption(slLocations, "Import", sgImportDirectory) Then
        sgImportDirectory = gSetPathEndSlash(sgImportDirectory)
    Else
        sgImportDirectory = ""
    End If
    If gLoadOption(slLocations, "Exe", sgExeDirectory) Then
        sgExeDirectory = gSetPathEndSlash(sgExeDirectory)
    Else
        sgExeDirectory = ""
    End If
    If gLoadOption(slLocations, "CartUnloadTime", slCartUnloadTime) Then
        lgCartUnloadTime = Val(slCartUnloadTime)
    End If
    
    
    'Commented out below because I can't see why you would need a backslash
    'on the end of a DSN name
    'sgDatabaseName = gSetPathEndSlash(sgDatabaseName)
    sgReportDirectory = gSetPathEndSlash(sgReportDirectory)
    sgExportDirectory = gSetPathEndSlash(sgExportDirectory)
    'sgImportDirectory = gSetPathEndSlash(sgImportDirectory)
    sgLogoDirectory = gSetPathEndSlash(sgLogoDirectory)
    
    Call gLoadOption("SQLSpec", "Date", sgSQLDateForm)
    Call gLoadOption("SQLSpec", "Time", sgSQLTimeForm)
    If gLoadOption("SQLSpec", "System", slBuffer) Then
        If slBuffer = "P7" Then
            igSQLSpec = 0
        End If
    End If
    If gLoadOption(slLocations, "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    
    If Not gLoadOption(slLocations, "DBPath", sgDBPath) Then
        MsgBox "Engineer.Ini [" & slLocations & "] 'DBPath' key is missing.", vbCritical
        Unload EngrLogIn
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath)
    End If
    sgMsgDirectory = sgDBPath & "Messages\"

    If Not gLoadOption(slLocations, "ServerDatabase", sgServerDatabase) Then
        sgServerDatabase = ""
    Else
        sgServerDatabase = gSetPathEndSlash(sgServerDatabase)
    End If

    On Error GoTo ErrHand
'    Set env = rdoEnvironments(0)
'    env.CursorDriver = rdUseOdbc
'
'    ' The default timeout is 15 seconds. This always fails on my PC the first time I run this program.
'    env.LoginTimeout = 30  ' Increase from the default of 15 to 30 seconds.

    slDSN = sgDatabaseName
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(slDSN, Len(slDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        slDSN = Left(slDSN, Len(slDSN) - 1)
    End If
'    Set cnn = env.OpenConnection(dsName:=slDSN, Prompt:=rdDriverCompleteRequired)
'    If igTimeOut >= 0 Then
'        cnn.QueryTimeout = igTimeOut
'    End If
    
    Set cnn = New ADODB.Connection
    cnn.Open "DSN=" & slDSN
    
    Set rst = New ADODB.Recordset
    
    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
    
    hgDB = CBtrvMngrInit(0, "", "", sgDBPath, 0, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    
    If Not gCheckDDFDates() Then
        Unload EngrLogIn
        Set EngrLogIn = Nothing   'Remove data segment
        End
    End If
    
    'Code modified for testing
    edcName.text = ""
    edcPassword.text = ""
    
    mInit
    mChkForGuide
    gGetSiteOption
    
    
    
    If Trim$(sgNowDate) = "" Then
        If InStr(1, sgClientName, "XYZ Broadcasting", vbTextCompare) > 0 Then
            sgNowDate = "12/15/1999"
        End If
    End If

    

    Call gLoadOption("Database", "AutoLogin", slAutoLogin)
    If slAutoLogin = "AutoLogin" Then
        edcName.text = "guide"
        edcPassword.text = "radio1234"
        Call mStart
    End If
    
    ilRet = mInitAPIReport()      '4-19-04
    ilRet = gPopReportNames()
        
    Exit Sub
mReadFileErr:
    ilRet = Err.Number
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    'env.RollbackTrans
    On Error Resume Next
    cnn.RollbackTrans
    If gMsg = "" Then
        MsgBox "Error at Start-up " & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub


Private Sub mExit()
    On Error Resume Next
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    btrStopAppl
    Unload EngrLogIn
'    Set EngrLogIn = Nothing
End Sub


Private Sub mStart()
    Dim ilRet As Integer
    Dim slName As String
    Dim slPass As String
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilGetStatus As Integer
    Dim ilIndex As Integer
    Dim tlUte As UTE
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    igPasswordOk = False
    lgSchTopRow = -1
    lgLibTopRow = -1
    lgTempTopRow = -1
    ilGetStatus = False
    slName = Trim$(edcName.text)
    slPass = Trim$(edcPassword.text)
    If (StrComp(slName, "Counterpoint", vbTextCompare) <> 0) Or (StrComp(slPass, "JD#41", vbTextCompare) <> 0) Then
        '7/26/11: Disallow Guide with special password
        'If ((StrComp(slName, "Guide", vbTextCompare) <> 0) Or (StrComp(slPass, smSpecial, vbTextCompare) <> 0)) And ((StrComp(slName, "CSI", vbTextCompare) <> 0) Or (StrComp(slPass, smSpecial, vbTextCompare) <> 0)) Then
        If ((StrComp(slName, "CSI", vbTextCompare) <> 0) Or (StrComp(slPass, smSpecial, vbTextCompare) <> 0)) Then
            For ilLoop = LBound(tgCurrUIE) To UBound(tgCurrUIE) - 1 Step 1
                If (StrComp(Trim$(tgCurrUIE(ilLoop).sSignOnName), slName, vbTextCompare) = 0) And (StrComp(Trim$(tgCurrUIE(ilLoop).sPassword), slPass, vbTextCompare) = 0) Then
                    LSet tgUIE = tgCurrUIE(ilLoop)
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                Beep
                Screen.MousePointer = vbDefault
                edcName.SetFocus
                Exit Sub
            End If
            ilGetStatus = True
            sgUserName = slName
            ilRet = gPutUpdate_UIE_UserInfo(2, tgUIE, "User Option-mStart: UIE")
        Else
            sgUserName = "Guide" 'slName
            For ilLoop = LBound(igJobStatus) To UBound(igJobStatus) Step 1
                igJobStatus(ilLoop) = 2
            Next ilLoop
            For ilLoop = LBound(igListStatus) To UBound(igListStatus) Step 1
                igListStatus(ilLoop) = 2
            Next ilLoop
            If Len(slPass) = 9 Then
                sgSpecialPassword = Mid$(slPass, 6)
                If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
                   StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                    igPasswordOk = True
                End If
            End If
            tgUIE.sShowName = "Guide"
            tgUIE.iCode = 1
        End If
    Else
        sgUserName = slName
        For ilLoop = LBound(igJobStatus) To UBound(igJobStatus) Step 1
            igJobStatus(ilLoop) = 2
        Next ilLoop
        For ilLoop = LBound(igListStatus) To UBound(igListStatus) Step 1
            igListStatus(ilLoop) = 2
        Next ilLoop
        If Len(slPass) = 5 Then
            sgSpecialPassword = slPass
            If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
               StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                igPasswordOk = True
            End If
        End If
        tgUIE.sShowName = "Guide"
        tgUIE.iCode = 1
    End If
    mChkTask
    If (ilGetStatus) Or (UBound(tgCurrUIE) <= 1) Then
        ilRet = gGetRecs_UTE_UserTasks(sgCurrUTEStamp, tgUIE.iCode, "Login-mStart: Get UTE", tgCurrUTE())
        If UBound(tgCurrUTE) <= LBound(tgCurrUTE) Then
            'Add Task Jobs
            For ilLoop = 0 To UBound(tgJobTaskNames) Step 1
                tlUte.iCode = 0
                tlUte.iUieCode = tgUIE.iCode
                tlUte.iTneCode = tgJobTaskNames(ilLoop).iCode
                If (StrComp(slName, "Guide", vbTextCompare) = 0) Then
                    tlUte.sTaskStatus = "E"
                Else
                    tlUte.sTaskStatus = "V"
                End If
                tlUte.sUnused = ""
                ilRet = gPutInsert_UTE_UserTasks(tlUte, "Login-mStart: Insert UTE-Job")
            Next ilLoop
             'Add Task Lists
            For ilLoop = 0 To UBound(tgListTaskNames) Step 1
                tlUte.iCode = 0
                tlUte.iUieCode = tgUIE.iCode
                tlUte.iTneCode = tgListTaskNames(ilLoop).iCode
                If (StrComp(slName, "Guide", vbTextCompare) = 0) Then
                    tlUte.sTaskStatus = "E"
                Else
                    tlUte.sTaskStatus = "V"
                End If
                tlUte.sUnused = ""
                ilRet = gPutInsert_UTE_UserTasks(tlUte, "Login-mStart: Insert UTE-List")
            Next ilLoop
            ilRet = gGetRecs_UTE_UserTasks(sgCurrUTEStamp, tgUIE.iCode, "Login-mStart: Get UTE", tgCurrUTE())
        End If
        If ilGetStatus Then
            For ilLoop = LBound(tgCurrUTE) To UBound(tgCurrUTE) - 1 Step 1
                For ilIndex = 0 To UBound(tgJobTaskNames) Step 1
                    If tgCurrUTE(ilLoop).iTneCode = tgJobTaskNames(ilIndex).iCode Then
                        If tgCurrUTE(ilLoop).sTaskStatus = "E" Then
                            igJobStatus(ilIndex) = 2
                        ElseIf tgCurrUTE(ilLoop).sTaskStatus = "V" Then
                            igJobStatus(ilIndex) = 1
                        Else
                            igJobStatus(ilIndex) = 0
                        End If
                        Exit For
                    End If
                Next ilIndex
                For ilIndex = 0 To UBound(tgListTaskNames) Step 1
                    If tgCurrUTE(ilLoop).iTneCode = tgListTaskNames(ilIndex).iCode Then
                        If tgCurrUTE(ilLoop).sTaskStatus = "E" Then
                            igListStatus(ilIndex) = 2
                        ElseIf tgCurrUTE(ilLoop).sTaskStatus = "V" Then
                            igListStatus(ilIndex) = 1
                        Else
                            igListStatus(ilIndex) = 0
                        End If
                        Exit For
                    End If
                Next ilIndex
            Next ilLoop
            '7/12/11: Only allow guide to change or add users
            If (StrComp(slName, "Guide", vbTextCompare) = 0) Then
                For ilLoop = LBound(igJobStatus) To UBound(igJobStatus) Step 1
                    igJobStatus(ilLoop) = 0
                Next ilLoop
                For ilLoop = LBound(igListStatus) To UBound(igListStatus) Step 1
                    igListStatus(ilLoop) = 0
                Next ilLoop
                igListStatus(USERINDEX) = 2
            End If
        End If
    End If
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLogIn-mStart Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrLogIn-mStart Event Properties", tgCurrEPE())
    gGetEPEUsedSummary
    gGetEPEManSummary
    gGetAuto
    'Need to Unload so that Controls will be set with correct properties (List Buttons)
    Unload EngrList
    'Need to Unload so that Controls will be set with correct properties (Job Buttons)
    Unload EngrJob
    'Need to load Job as EngrMain is loaded when getting Task Names and will not load EngrJob
    Load EngrJob
    
    Screen.MousePointer = vbDefault
    EngrMain.Show
    Set EngrLogIn = Nothing
    Exit Sub
ErrHand:
    gMsg = ""
    Resume Next
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in EngrLogIn - cmdOK Click: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in EngrLogIn - cmdOK Click: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Screen.MousePointer = vbDefault
    Unload EngrLogIn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrLogIn = Nothing
End Sub


Private Sub edcPassword_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'           mInitAPIReport - Gather all the filenames from File.ddf.  Required
'           if converting a Btrieve report to ODBC.  If aliases on filenames are
'           used in the report, need to get the real name of the filename to
'           store in the database/tables/location.
'           4-20-04
'
Public Function mInitAPIReport() As Integer
    Dim sFileName As String * 20
    Dim ilUpper As Integer
    Dim ilPos As Integer
    'Dim ddf_rst As rdoResultset
    Dim ddf_rst As ADODB.Recordset
    Dim SQLQuery As String
    
    On Error GoTo ErrHand

    ReDim tgDDFFileNames(0 To 0) As DDFFILENAMES
    ilUpper = UBound(tgDDFFileNames)
    SQLQuery = "SELECT Xf$Name FROM X$File"
    'Set ddf_rst = cnn.OpenResultset(SQLQuery)
    Set ddf_rst = cnn.Execute(SQLQuery)

    If Not ddf_rst.EOF Then
        While Not ddf_rst.EOF
            If Mid(ddf_rst(0).Value, 1, 2) <> "X$" Then
                tgDDFFileNames(ilUpper).sLongName = Trim$(ddf_rst(0).Value)
                sFileName = Trim$(tgDDFFileNames(ilUpper).sLongName)
                ilPos = InStr(sFileName, "_")
                If ilPos = 0 Then
                    tgDDFFileNames(ilUpper).sShortName = Trim$(sFileName)
                Else
                    tgDDFFileNames(ilUpper).sShortName = Mid(sFileName, 1, ilPos - 1)
                End If
                ilUpper = ilUpper + 1
                ReDim Preserve tgDDFFileNames(0 To ilUpper)
            End If
            ddf_rst.MoveNext
        Wend
    Else
        MsgBox "DDF Open Failed"
    End If
    Exit Function

ErrHand:
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in EngrLogIn - mInitAPIReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A SQL error has occured in EngrLogIn - mInitAPIReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Screen.MousePointer = vbDefault
    Unload EngrLogIn

End Function

Public Sub mGetResolution()
    'Get current setting
    lgCurrHRes = GetDeviceCaps(pbc256LogIn.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(pbc256LogIn.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(pbc256LogIn.hdc, BITSPIXEL)
    If lgCurrBPP <= 8 Then
        EngrLogIn.Picture = pbc256LogIn
    End If
    
End Sub

Private Sub lacExit_Click()
    cmcExit_Click
End Sub

Private Sub lacExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Move 1560, 2280, 1365, 360 '465, 2190, 3405, 405
    imcOutline.Visible = True
End Sub

Private Sub lacStart_Click()
    cmcStart_Click
End Sub

Private Sub lacStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imcOutline.Move 1560, 1770, 1365, 360 '465, 1770, 3405, 405
    imcOutline.Visible = True
End Sub

Private Sub pbcClickFocus_Click()
    imcOutline.Visible = False
End Sub

Private Sub mChkForGuide()
    Dim ilRet As Integer
    
    sgCurrUIEStamp = ""
    
    ilRet = gGetTypeOfRecs_UIE_UserInfo("C", sgCurrUIEStamp, "LogIn-mChkForGuide", tgCurrUIE())
    
    Exit Sub

End Sub


Private Sub mChkTask()
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilTne As Integer
    Dim ilRet As Integer
    Dim tlTne As TNE
    
    ilRet = gGetAll_TNE_TaskName("LogIn-mChkTask")
    ReDim tgListTaskNames(0 To EngrList!lacTask.UBound) As TNE
    For ilLoop = 0 To EngrList!lacTask.UBound Step 1
        tgListTaskNames(ilLoop).iCode = 0
        tgListTaskNames(ilLoop).sType = "L"
        tgListTaskNames(ilLoop).sName = EngrList!lacTask(ilLoop)
        tgListTaskNames(ilLoop).sUnused = ""
    Next ilLoop
    ReDim tgJobTaskNames(0 To EngrJob.lacTask.UBound) As TNE
    For ilLoop = 0 To EngrJob!lacTask.UBound Step 1
        tgJobTaskNames(ilLoop).iCode = 0
        tgJobTaskNames(ilLoop).sType = "J"
        tgJobTaskNames(ilLoop).sName = EngrJob!lacTask(ilLoop)
        tgJobTaskNames(ilLoop).sUnused = ""
    Next ilLoop
''    'Include names which don't have List or Job buttons
''    ReDim tgExtraTaskNames(0 To 4) As TNE
''    tgExtraTaskNames(0).iCode = 0
''    tgExtraTaskNames(0).sType = "E"
''    tgExtraTaskNames(0).sName = "Audio Types"
''    tgExtraTaskNames(0).sUnused = ""
''    tgExtraTaskNames(1).iCode = 0
''    tgExtraTaskNames(1).sType = "E"
''    tgExtraTaskNames(1).sName = "Bus Groups"
''    tgExtraTaskNames(1).sUnused = ""
''    tgExtraTaskNames(2).iCode = 0
''    tgExtraTaskNames(2).sType = "E"
''    tgExtraTaskNames(2).sName = "Sub-Library Names"
''    tgExtraTaskNames(2).sUnused = ""
''    tgExtraTaskNames(3).iCode = 0
''    tgExtraTaskNames(3).sType = "E"
''    tgExtraTaskNames(3).sName = "Sub-Template Names"
''    tgExtraTaskNames(3).sUnused = ""
''    tgExtraTaskNames(4).iCode = 0
''    tgExtraTaskNames(4).sType = "E"
''    tgExtraTaskNames(4).sName = "Control Char"
''    tgExtraTaskNames(4).sUnused = ""
'    ReDim tgExtraTaskNames(0 To 3) As TNE
'    tgExtraTaskNames(0).iCode = 0
'    tgExtraTaskNames(0).sType = "E"
'    tgExtraTaskNames(0).sName = "Sub-Library Names"
'    tgExtraTaskNames(0).sUnused = ""
'    tgExtraTaskNames(1).iCode = 0
'    tgExtraTaskNames(1).sType = "E"
'    tgExtraTaskNames(1).sName = "Sub-Template Names"
'    tgExtraTaskNames(1).sUnused = ""
'    tgExtraTaskNames(2).iCode = 0
'    tgExtraTaskNames(2).sType = "E"
'    tgExtraTaskNames(2).sName = "Control Char"
'    tgExtraTaskNames(2).sUnused = ""
    ReDim tgExtraTaskNames(0 To 1) As TNE
    tgExtraTaskNames(0).iCode = 0
    tgExtraTaskNames(0).sType = "E"
    tgExtraTaskNames(0).sName = "Sub-Library Names"
    tgExtraTaskNames(0).sUnused = ""
    tgExtraTaskNames(1).iCode = 0
    tgExtraTaskNames(1).sType = "E"
    tgExtraTaskNames(1).sName = "Sub-Template Names"
    tgExtraTaskNames(1).sUnused = ""
    ReDim tgAlertTaskNames(0 To 0) As TNE
    tgAlertTaskNames(0).iCode = 0
    tgAlertTaskNames(0).sType = "A"
    tgAlertTaskNames(0).sName = "Schedule Not Retrieved"
    tgAlertTaskNames(0).sUnused = ""
    
    ReDim tgNoticeTaskNames(0 To 2) As TNE
    tgNoticeTaskNames(0).iCode = 0
    tgNoticeTaskNames(0).sType = "N"
    tgNoticeTaskNames(0).sName = "Merge Errors"
    tgNoticeTaskNames(0).sUnused = ""
    tgNoticeTaskNames(1).iCode = 0
    tgNoticeTaskNames(1).sType = "N"
    tgNoticeTaskNames(1).sName = "Schedule Not Retrieved"
    tgNoticeTaskNames(1).sUnused = ""
    tgNoticeTaskNames(2).iCode = 0
    tgNoticeTaskNames(2).sType = "N"
    tgNoticeTaskNames(2).sName = "Commercial Test Error"
    tgNoticeTaskNames(2).sUnused = ""
    
    For ilLoop = 0 To UBound(tgJobTaskNames) Step 1
        ilFound = False
        For ilTne = LBound(tgCurrTNE) To UBound(tgCurrTNE) - 1 Step 1
            If tgCurrTNE(ilTne).sType = "J" Then
                If StrComp(Trim$(tgCurrTNE(ilTne).sName), Trim$(tgJobTaskNames(ilLoop).sName), vbTextCompare) = 0 Then
                    ilFound = True
                    tgJobTaskNames(ilLoop).iCode = tgCurrTNE(ilTne).iCode
                    Exit For
                End If
            End If
        Next ilTne
        If Not ilFound Then
            ilRet = gPutInsert_TNE_TaskName(tgJobTaskNames(ilLoop), "LogIn-mChkTask")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgListTaskNames) Step 1
        ilFound = False
        For ilTne = LBound(tgCurrTNE) To UBound(tgCurrTNE) - 1 Step 1
            If tgCurrTNE(ilTne).sType = "L" Then
                If StrComp(Trim$(tgCurrTNE(ilTne).sName), Trim$(tgListTaskNames(ilLoop).sName), vbTextCompare) = 0 Then
                    ilFound = True
                    tgListTaskNames(ilLoop).iCode = tgCurrTNE(ilTne).iCode
                    Exit For
                End If
            End If
        Next ilTne
        If Not ilFound Then
            ilRet = gPutInsert_TNE_TaskName(tgListTaskNames(ilLoop), "LogIn-mChkTask")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtraTaskNames) Step 1
        ilFound = False
        For ilTne = LBound(tgCurrTNE) To UBound(tgCurrTNE) - 1 Step 1
            If tgCurrTNE(ilTne).sType = "E" Then
                If StrComp(Trim$(tgCurrTNE(ilTne).sName), Trim$(tgExtraTaskNames(ilLoop).sName), vbTextCompare) = 0 Then
                    ilFound = True
                    tgExtraTaskNames(ilLoop).iCode = tgCurrTNE(ilTne).iCode
                    Exit For
                End If
            End If
        Next ilTne
        If Not ilFound Then
            ilRet = gPutInsert_TNE_TaskName(tgExtraTaskNames(ilLoop), "LogIn-mChkTask")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgAlertTaskNames) Step 1
        ilFound = False
        For ilTne = LBound(tgCurrTNE) To UBound(tgCurrTNE) - 1 Step 1
            If tgCurrTNE(ilTne).sType = "A" Then
                If StrComp(Trim$(tgCurrTNE(ilTne).sName), Trim$(tgAlertTaskNames(ilLoop).sName), vbTextCompare) = 0 Then
                    ilFound = True
                    tgAlertTaskNames(ilLoop).iCode = tgCurrTNE(ilTne).iCode
                    Exit For
                End If
            End If
        Next ilTne
        If Not ilFound Then
            ilRet = gPutInsert_TNE_TaskName(tgAlertTaskNames(ilLoop), "LogIn-mChkTask")
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgNoticeTaskNames) Step 1
        ilFound = False
        For ilTne = LBound(tgCurrTNE) To UBound(tgCurrTNE) - 1 Step 1
            If tgCurrTNE(ilTne).sType = "N" Then
                If StrComp(Trim$(tgCurrTNE(ilTne).sName), Trim$(tgNoticeTaskNames(ilLoop).sName), vbTextCompare) = 0 Then
                    ilFound = True
                    tgNoticeTaskNames(ilLoop).iCode = tgCurrTNE(ilTne).iCode
                    Exit For
                End If
            End If
        Next ilTne
        If Not ilFound Then
            ilRet = gPutInsert_TNE_TaskName(tgNoticeTaskNames(ilLoop), "LogIn-mChkTask")
        End If
    Next ilLoop
    ilRet = gGetAll_TNE_TaskName("LogIn-mChkTask")
    Exit Sub

End Sub


Private Sub mInit()
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String
    
    slDate = Format$(Now(), "ddddd")
    slMonth = Month(slDate)
    slYear = Year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(Str$(ilValue))
    Do While Len(slStr) < 4
        slStr = "0" & slStr
    Loop
    smSpecial = "Login" & slStr
    lgLastServiceDate = -1
    lgLastServiceTime = -1
End Sub
