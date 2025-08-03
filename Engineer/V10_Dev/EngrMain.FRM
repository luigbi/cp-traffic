VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm EngrMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Counterpoint ®"
   ClientHeight    =   5385
   ClientLeft      =   255
   ClientTop       =   2520
   ClientWidth     =   10335
   Icon            =   "EngrMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "HelpMsg"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pbcList 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1980
      Left            =   0
      ScaleHeight     =   1950
      ScaleWidth      =   10305
      TabIndex        =   0
      Top             =   285
      Visible         =   0   'False
      Width           =   10335
      Begin VB.PictureBox pbc256List 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   8715
         Picture         =   "EngrMain.frx":08CA
         ScaleHeight     =   480
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   1125
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox pbc256Job 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   8715
         Picture         =   "EngrMain.frx":482E
         ScaleHeight     =   450
         ScaleWidth      =   540
         TabIndex        =   7
         Top             =   630
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.PictureBox pbc256SignOn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   8820
         Picture         =   "EngrMain.frx":8792
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   6
         Top             =   15
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Timer tmcFlashAlert 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   7215
         Top             =   30
      End
      Begin VB.PictureBox pbcTestFont 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   165
         Left            =   7485
         ScaleHeight     =   135
         ScaleWidth      =   1335
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1365
      End
      Begin VB.PictureBox pbcSymb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Enabled         =   0   'False
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
         Height          =   675
         Left            =   1215
         ScaleHeight     =   675
         ScaleWidth      =   690
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   690
      End
      Begin VB.Timer tmcCheckAlert 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   7770
         Top             =   405
      End
      Begin VB.PictureBox pbcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   675
         Left            =   2730
         ScaleHeight     =   675
         ScaleWidth      =   2985
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   645
         Width           =   2985
         Begin VB.PictureBox pbcArial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1485
            ScaleHeight     =   225
            ScaleWidth      =   1005
            TabIndex        =   5
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Timer tmcClock 
         Interval        =   60000
         Left            =   7740
         Top             =   -60
      End
      Begin MSComDlg.CommonDialog cdcSetup 
         Left            =   6480
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin VB.Image imcTrashOpened 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5730
         Picture         =   "EngrMain.frx":C9E9
         Top             =   615
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imcTrashClosed 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5700
         Picture         =   "EngrMain.frx":CCF3
         Top             =   1155
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imcKey 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   5730
         Picture         =   "EngrMain.frx":CFFD
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imcPrinter 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6540
         Picture         =   "EngrMain.frx":D307
         Top             =   1245
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imcInsert 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6585
         Picture         =   "EngrMain.frx":DBD1
         Top             =   615
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox pbcMsgArea 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   10335
      TabIndex        =   2
      Top             =   0
      Width           =   10335
      Begin VB.Label imcServiceStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SERVICE STATUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   14
         Top             =   15
         Width           =   3285
      End
      Begin VB.Label imcReport 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9495
         TabIndex        =   13
         Top             =   15
         Width           =   825
      End
      Begin VB.Label imcTask 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SCHEDULE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3375
         TabIndex        =   12
         Top             =   15
         Width           =   2730
      End
      Begin VB.Label imcTask 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TEMPLATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2250
         TabIndex        =   11
         Top             =   15
         Width           =   1110
      End
      Begin VB.Label imcTask 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LIBRARY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1125
         TabIndex        =   10
         Top             =   15
         Width           =   1110
      End
      Begin VB.Label imcTask 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIME FINDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   15
         Width           =   1110
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Task"
      Begin VB.Menu mnuTaskJob 
         Caption         =   "&Jobs"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuTaskList 
         Caption         =   "&Lists"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuTaskSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaskBackup 
         Caption         =   "&Backup"
         Begin VB.Menu mnuTaskBackupSystem 
            Caption         =   "System Backup"
         End
         Begin VB.Menu mnuTaskBackupCopy 
            Caption         =   "Copy Data to Test"
         End
      End
      Begin VB.Menu mnuTaskSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuImportComm 
            Caption         =   "&Commercial"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuImportFiles 
            Caption         =   "&Audio, Bus, Netcue, Relay"
         End
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuExportLoad 
            Caption         =   "&Load Automation"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuExportReload 
            Caption         =   "&Reload Automation"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuCreateConflictTables 
         Caption         =   "Create Conflict &Tables"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTaskSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaskExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAccessory 
      Caption         =   "&Accessories"
      Begin VB.Menu mnuAccessoryDate 
         Caption         =   "Counterpoint &Date"
      End
      Begin VB.Menu mnuAccessorySpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoryUtilities 
         Caption         =   "Utilities"
         Begin VB.Menu mnuAccessoryItemCheck 
            Caption         =   "&Check Item ID's"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAccessoryFileCheck 
            Caption         =   "&File Check/Fix"
         End
         Begin VB.Menu mnuAccessoryFileCleanUp 
            Caption         =   "File Clean-&Up"
         End
      End
      Begin VB.Menu mnuAccessorySpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessoryMessages 
         Caption         =   "&Messages"
      End
   End
   Begin VB.Menu mnuBlank 
      Caption         =   "                                                                "
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuDate 
      Caption         =   "  "
   End
End
Attribute VB_Name = "EngrMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software® All rights reserved.
' Proprietary Software, Do not copy
' This is a test
' File Name: EngrMain.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the EngrMain input screen code
Option Explicit
Option Compare Text

'------------------------------------------------------
'
'    Multi-Automation Equipment Changes:
'
'   1.  Add to ETE_Event_Types and DHE_Day_Header a reference to AEE_Auto_Equip.
'       To ETE because each piece of equipment could have different used and mandatory
'       definitions.
'       To DHE because the event types could be different (allow libraries from
'       one automation equipment to be copied into the other).  Would also
'       need at startup module to copy all libraries activate libraries from one
'       automation equipment to the other
'   2.  On event screen ask which piece of automation equipment.
'       if coming from the Library screen, this would be obtained from the one
'       selected with the Libraries
'   3.  On the Library screen ask which automation equipment
'   4.  On the schedule screen ask which automation equipment
'   5.  Once automation equipment picked, then setup the public Mandatory and Used
'       tables
'   6.  Create a Public variable that indicates the number of Active Automation
'       equipment in use.
'   7.  On the Library selection screen, show which automation equipment it is for.
'
'   Alternative design for Libraries:  Use only one structure but require all
'                                      fields as defined in ETE even if the
'                                      field in one piece of automation equipment
'                                      is not required but it is in the other
'                                      piece of equipment.  i.e. use the must
'                                      severe requires for each field.
'                                      Example:  Equipment 1 does not require Bus
'                                                control but Equipment 2 does, then
'                                                require that it be defined.
'                                      If using this alternative design, then
'                                      Automation selection is required only for
'                                      Event Type definitions and Schedule.
'                                      When generating Schedule, If Field not used, then
'                                      ignore value from Library
'   Templates would have to work like Libraries
'
'-------------------------------------------------------

Dim hmAlert As Long
Dim lmMenuID As Long
Dim hmRedAlertBitmap As Long
Dim hmWhiteAlertBitmap As Long

Dim imTerminate As Integer  'True= terminating form; False= OK
Dim smCommand As String

Dim lm1970 As Long

Const RESOURCEINDEX = 0
Const LIBRARYINDEX = 1
Const TEMPLATEINDEX = 2
Const SCHEDULEINDEX = 3

Private Sub mSetCtrlEnable(ilCtrlIndex As Integer, ilTaskNo As Integer)
    If igJobStatus(ilTaskNo) = 0 Then
        imcTask(ilCtrlIndex).Enabled = False
    Else
        imcTask(ilCtrlIndex).Enabled = True
    End If
End Sub

Private Sub mSetControls()
    Dim ilCount As Integer
    
    mSetCtrlEnable RESOURCEINDEX, RESOURCEJOB
    mSetCtrlEnable LIBRARYINDEX, LIBRARYJOB
    mSetCtrlEnable TEMPLATEINDEX, TEMPLATEJOB
    mSetCtrlEnable SCHEDULEINDEX, SCHEDULEJOB
    On Error Resume Next
    ilCount = gGetCheckStatus()
    If ilCount > 0 Then
        If igJobVisible Then
            EngrJob!lacTask(SCHEDULEJOB).ForeColor = vbRed
        End If
        imcTask(SCHEDULEJOB).ForeColor = vbRed
    Else
        If igJobVisible Then
            EngrJob!lacTask(SCHEDULEJOB).ForeColor = vbButtonText
        End If
        imcTask(SCHEDULEJOB).ForeColor = vbButtonText
    End If
    gGetSchDates
End Sub

Private Sub mLoadPicture()
        
    Dim slCommand As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    
    If igTestSystem Then
        EngrMstPict.Show vbModeless
        EngrMstPict.Enabled = False
    Else
        On Error GoTo mRetryPicture:
        If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
            ilRet = 0
            EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI640e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI640e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI640e.bmp")
                End If
            End If
        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
            ilRet = 0
            EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI800e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI800e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI800e.bmp")
                End If
            End If
        ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
            ilRet = 0
            EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI1024e.bmp")
                End If
            End If
        Else
            ilRet = 0
            EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI1280e.jpg")
            If ilRet <> 0 Then
                ilRet = 0
                EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI1280e.gif")
                If ilRet <> 0 Then
                    On Error GoTo mNoPicture:
                    ilRet = 0
                    EngrMain.Picture = LoadPicture(sgLogoDirectory & "CSI1280e.bmp")
                End If
            End If
        End If
    End If
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture:
    On Error GoTo 0
    Resume Next
End Sub



Private Function mCopyPictureImage(SourceImage As PictureBox) As Long
    Dim bm As BITMAP
    Dim newbm&
    Dim tdc&, oldbm&
    Dim di&

    ' First get the information about the image bitmap
    di = GetObjectAPI(SourceImage.Image, Len(bm), bm)
    bm.bmBits = 0
    ' Create a new bitmap with the same structure and size
    ' of the image bitmap
    newbm = CreateBitmapIndirect(bm)

    ' Create a temporary memory device context to use
    tdc = CreateCompatibleDC(SourceImage.hdc)
    ' Select in the newly created bitmap
    oldbm = SelectObject(tdc, newbm)

    ' Now copy the bitmap from the persistant bitmap in
    ' picture 2 (note that picture2 has AutoRedraw set TRUE
    di = BitBlt(tdc, 0, 0, bm.bmWidth, bm.bmHeight, SourceImage.hdc, 0, 0, SRCCOPY)
    ' Select out the bitmap and delete the memory DC
    oldbm = SelectObject(tdc, oldbm)
    di = DeleteDC(tdc)

    ' And return the new bitmap
    mCopyPictureImage = newbm
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mCheckVersion                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check Database version          *
'*                                                     *
'*******************************************************
Private Function mCheckVersion() As Integer
    Dim ilRet As Integer
    Dim hlVersion As Integer
    Dim slVersion As String
    Dim ilPos As Integer
    Dim slChar As String
    Dim slLine As String
    
    slVersion = ""
    ilPos = InStr(sgEngrVersion, " ") + 1
    slChar = Mid$(sgEngrVersion, ilPos, 1)
    Do While slChar <> " "
        slVersion = slVersion & slChar
        ilPos = ilPos + 1
        If ilPos > Len(sgEngrVersion) Then
            Exit Do
        End If
        slChar = Mid$(sgEngrVersion, ilPos, 1)
    Loop
    ilRet = 0
    On Error GoTo mWrongVersion:
    hlVersion = FreeFile
    Open sgDBPath & "EngrVer.Csi" For Input Access Read As hlVersion
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Programs and Database versions don't match", vbOKOnly + vbCritical, "Error"
        imTerminate = True
        mCheckVersion = False
        Exit Function
    End If
    On Error GoTo mWrongVersion:
    Line Input #hlVersion, slLine
    On Error GoTo 0
    Close hlVersion
    If (Asc(slLine) = 26) Or (ilRet <> 0) Or (StrComp(slVersion, slLine, 1) <> 0) Then    'Ctrl Z
        Screen.MousePointer = vbDefault
        MsgBox "Programs and Database versions don't match", vbOKOnly + vbCritical, "Error"
        imTerminate = True
        'Unload EngrMain
        'Set EngrMain = Nothing   'Remove data segment
        mCheckVersion = False
        Exit Function
    End If
    mCheckVersion = True
    Exit Function
mWrongVersion:
    On Error GoTo 0
    ilRet = 1
    Resume Next
End Function
Private Sub imcReport_Click()
    Dim slStr As String
    Dim ilRptType As Integer
    
    On Error Resume Next
    EngrReports.Show 'vbModal    'vbModeless  'This is required to activate form if previousely activated

End Sub

Private Sub imcTask_Click(Index As Integer)
    Dim slStr As String
    Dim ilRptType As Integer
    Dim ilAskMkt As Integer
    Dim ilLoop As Integer
    
    Select Case Index
        Case 0  'Resource
            If igJobStatus(RESOURCEJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If Not igJobShowing(RESOURCEJOB) Then
                EngrRS.Show vbModeless  'This is required to activate form if previousely activated
            Else
                EngrRS.SetFocus
            End If
        Case 1  'Library
            If igJobStatus(LIBRARYJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If igJobShowing(LIBRARYJOB) = 0 Then
                EngrLib.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igJobShowing(LIBRARYJOB) = 2 Then
                EngrLibDef.SetFocus
            Else
                EngrLib.SetFocus
            End If
        Case 2  'Template
            If igJobStatus(TEMPLATEJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If igJobShowing(TEMPLATEJOB) = 0 Then
                EngrTemp.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igJobShowing(TEMPLATEJOB) = 2 Then
                EngrTempDef.SetFocus   'Show vbModeless  'This is required to activate form if previousely activated
            Else
                EngrTemp.SetFocus
            End If
        Case 3  'Schedule
            If igJobStatus(SCHEDULEJOB) = 0 Then
                Exit Sub
            End If
            On Error Resume Next
            If igJobShowing(SCHEDULEJOB) = 0 Then
                EngrSchd.Show vbModeless  'This is required to activate form if previousely activated
            ElseIf igJobShowing(SCHEDULEJOB) = 2 Then
                EngrSchdDef.SetFocus
            ElseIf igJobShowing(SCHEDULEJOB) = 3 Then
                EngrAsAirCompare.SetFocus
            ElseIf igJobShowing(SCHEDULEJOB) = 4 Then
                EngrImportAsAir.SetFocus
            Else
                EngrSchd.SetFocus
            End If
    End Select
End Sub

Private Sub MDIForm_Activate()
    DoEvents
    mSetControls
End Sub
Private Sub MDIForm_Load()
    Dim ilRet As Integer
    Dim ilChgMode As Integer
    Dim ilLoop As Integer
    Dim slCommand As String
    Dim slDate As String
    Dim ilParse As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim ilPass As Integer
    Dim ilDelay As Integer
    Dim slSetDate As String
    Dim ilShell As Integer
    Dim ilErr As Integer
    Dim slParameter As String
    Dim slDateTime1 As String
    Dim slStartIn As String
    Dim slWinDir As String
    Dim llWinLen As Long
    Dim llHdc As Long
    Dim slUserName As String
    Dim slWindowsDir As String * MAX_PATH
    
    imTerminate = False
    'Only allow one instance
    If App.PrevInstance Then
        MsgBox "Only one copy of Engineering can be run at a time, sorry", vbOKOnly + vbInformation, "Counterpoint"
        End
    End If
    sgCommand = Command$
    slStartIn = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    
    
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        slUserName = "System"
    Else
        slUserName = sgUserName
    End If
    
    EngrMain.Caption = slUserName & " on Counterpoint ®, " & "V" & App.Major & "." & App.Minor & " for " & sgClientName

    EngrMain.WindowState = vbMaximized
    EngrMain.Show 0
    lm1970 = gDateValue("1/1/1970")
    imcServiceStatus.ForeColor = vbWhite
    
    slSetDate = ""
    pbcMsgArea.Move pbcMsgArea.Left, pbcMsgArea.Top, EngrMain.Width
    'imcReport.Move imcTask(3).Left + imcTask(3).Width + 15, imcReport.Top, pbcMsgArea.Width - imcTask(3).Left - imcTask(3).Width - 45
    imcReport.Move imcServiceStatus.Left + imcServiceStatus.Width + 15, imcReport.Top, pbcMsgArea.Width - imcServiceStatus.Left - imcServiceStatus.Width - 45
    sgNowDate = ""
    If StrComp(sgCommand, "Debug1", 1) <> 0 Then
        ilPos = InStr(1, sgCommand, "/D:", 1)
        If ilPos > 0 Then
            ilSpace = InStr(ilPos, sgCommand, " ")
            If ilSpace = 0 Then
                slDate = Trim$(Mid$(sgCommand, ilPos + 3))
            Else
                slDate = Trim$(Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3))
            End If
            If gIsDate(slDate) Then
                slDate = gAdjYear(slDate)
                slSetDate = slDate
            End If
            sgNowDate = slSetDate
        End If
'        If Not gCheckDDFDates() Then
'            Unload EngrMain
'            Set EngrMain = Nothing   'Remove data segment
'            Exit Sub
'        End If
    End If
    slCommand = sgCommand
    Screen.MousePointer = vbHourglass
    'slCommand = Command$
    smCommand = ""
    If StrComp(slCommand, "Debug1", 1) = 0 Then
        'If Not gCheckDDFDates() Then
        '    Unload EngrMain
        '    Set EngrMain = Nothing   'Remove data segment
        '    Exit Sub
        'End If
    Else
    End If
'    If Not mCheckVersion() Then
'        Unload EngrMain
'        Set EngrMain = Nothing   'Remove data segment
'        Exit Sub
'    End If
    If Not mEngrMainStartUp() Then
        Unload EngrMain
        'Set EngrMain = Nothing
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If Trim$(sgNowDate) = "" Then
'        If InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0 Then
'            slSetDate = "12/15/1999"
'            slDate = slSetDate
'            sgNowDate = slDate
'        End If
    End If
    If Trim$(slSetDate) <> "" Then
        tmcClock_Timer
    End If
    tmcClock_Timer
    On Error Resume Next
    On Error GoTo mRetryPicture:
    ilRet = 0
    pbcLogo.Picture = LoadPicture(sgLogoDirectory & "CustLogo.jpg")
    If ilRet <> 0 Then
        ilRet = 0
        pbcLogo.Picture = LoadPicture(sgLogoDirectory & "CustLogo.gif")
        If ilRet <> 0 Then
            On Error GoTo mNoPicture1:
            ilRet = 0
            pbcLogo.Picture = LoadPicture(sgLogoDirectory & "CustLogo.Bmp")
        End If
    End If
    On Error GoTo mRetryPicture:
    ilRet = 0
    pbcSymb.Picture = LoadPicture(sgLogoDirectory & "CustSymb.jpg")
    If ilRet <> 0 Then
        ilRet = 0
        pbcSymb.Picture = LoadPicture(sgLogoDirectory & "CustSymb.gif")
        If ilRet <> 0 Then
            On Error GoTo mNoPicture2:
            ilRet = 0
            pbcSymb.Picture = LoadPicture(sgLogoDirectory & "CustSymb.Bmp")
        End If
    End If
    On Error GoTo 0
    If igTestSystem Then
        EngrMain.BackColor = &HC0C0C0
    Else
        EngrMain.BackColor = &H80000001
    End If

    mLoadPicture
    'DoEvents

    EngrJob.Show
    Unload EngrLogIn

'    ilRet = gAlertForceCheck()
    Screen.MousePointer = vbDefault  'Wait
    
    Exit Sub
mRetryPicture:
    ilRet = 1
    Resume Next
mNoPicture1:
    pbcLogo.Picture = LoadPicture()
    Resume Next
mNoPicture2:
    pbcSymb.Picture = LoadPicture()
    Resume Next
ExeTestErr:
    ilRet = False
    Resume Next
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Dim ilShell As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    igStopCancel = False
    mEraseArrays
    mUnloadForms
    If igStopCancel Then    'User stop the unload
        Cancel = 1  'Stop unload
        Exit Sub
    End If
    If igJobShowing(RESOURCEJOB) Then
        Unload EngrRS
    End If
    If igJobShowing(LIBRARYJOB) = 1 Then
        Unload EngrLib
    End If
    If igJobShowing(LIBRARYJOB) = 2 Then
        Unload EngrLibDef
    End If
    If igJobShowing(SCHEDULEJOB) = 1 Then
        Unload EngrSchd
    End If
    If igJobShowing(SCHEDULEJOB) = 2 Then
        Unload EngrSchdDef
    End If
    If igJobShowing(SCHEDULEJOB) = 3 Then
        Unload EngrAsAirCompare
    End If
    If igJobShowing(SCHEDULEJOB) = 4 Then
        Unload EngrImportAsAir
    End If
    If igJobShowing(TEMPLATEJOB) = 1 Then
        Unload EngrTemp
    End If
    If igJobShowing(TEMPLATEJOB) = 2 Then
        Unload EngrTempDef
    End If
    tmcCheckAlert.Enabled = False
    tmcClock.Enabled = False
    If hmAlert <> 0 Then
        Call DeleteObject(hmAlert)
    End If
    On Error Resume Next
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    btrStopAppl
    Set EngrMain = Nothing   'Remove data segment
    End
ExeOkErr:
    ilRet = False
    Resume Next
End Sub
Private Sub mnuAboutHelp_Click()
    Dim iAns As Integer
    Dim ilShell As Integer
    Dim ilRet As Integer
    Dim ilErr As Integer
    ''sgEngrVersion set within InitSubs.Bas-gInitGlobalVar
    ''iAns = MsgBox(sgEngrVersion, 0, "Counterpoint Software")
    'ilShell = 0
    'ilErr = False
    'On Error GoTo mnuAboutHelpErr
    'If igTestSystem Then
    '    ilShell = Shell(sgExePath & "CSIAbout.Exe EngrMain^Test\" & sgUserName, 1)
    'Else
    '    ilShell = Shell(sgExePath & "CSIAbout.Exe EngrMain^Prod\" & sgUserName, 1)
    'End If
    'If ilErr Then
    '    Exit Sub
    'End If
    'While GetModuleUsage(ilShell) > 0
    '    ilRet = DoEvents()
    'Wend
    On Error Resume Next
    EngrAbout.Show vbModal
    Exit Sub
mnuAboutHelpErr:
    ilErr = True
    iAns = MsgBox(sgEngrVersion, 0, "Counterpoint Software")
    Resume Next
End Sub




Private Sub mnuAccessoryDate_Click()
    Dim slDate As String
    Dim slTime As String
    Dim slMsg As String
    Dim slInputDate As String
    Dim ilRet As Integer
    slDate = Format$(gNow(), "ddddd")
    slTime = Format$(gNow(), "h:mm:ss AM/PM")
    slMsg = "Counterpoint Date: " & slDate & "; Counterpoint Time: " & slTime
'    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
'       ((StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) And (sgSpecialPassword <> "") And (Len(sgSpecialPassword) = 4)) Then
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Then
        slInputDate = InputBox$(slMsg, "Date/Time", slDate)
        If Len(slInputDate) <> 0 Then
            If gIsDate(slInputDate) Then
                If gDateValue(slDate) <> gDateValue(slInputDate) Then
                    'Add year if missing
                    slInputDate = gAdjYear(slInputDate)
                    sgNowDate = slInputDate
                    mSetControls
                    tmcClock_Timer
                Else
                    sgNowDate = ""
                End If
            End If
        End If
    Else
        MsgBox slMsg, vbOKOnly + vbInformation, "Date/Time"
    End If
End Sub

Private Sub mnuAccessoryEditor_Click()
    Dim ilShell As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim ilPos As Integer
    Dim ilPFE As Integer
    
    sgGenMsg = "Select Editor or Cancel"
    sgCMCTitle(0) = "&NotePad"
    sgCMCTitle(1) = "&WordPad"
    sgCMCTitle(2) = "&Cancel"
    sgCMCTitle(3) = ""
    igDefCMC = 0
    igEditBox = 0
    sgEditValue = ""
    On Error Resume Next
    EngrGenMsg.Show vbModal
    If igAnsCMC = 2 Then
        Exit Sub
    End If
    ilShell = 0
    On Error GoTo mnuEditorFileErr
    'ilShell = Shell("NotePad.Exe", 3)
    'While GetModuleUsage(ilShell) > 0
    '    ilRet = DoEvents()
    'Wend
    If igAnsCMC = 0 Then
        gShellAndWait EngrMain, "NotePad.Exe", vbMaximizedFocus
    ElseIf igAnsCMC = 1 Then
        gShellAndWait EngrMain, "WordPad.Exe", vbMaximizedFocus
    End If
    Exit Sub
mnuEditorFileErr:
    If igAnsCMC = 0 Then
        MsgBox "The Editor Notepad.Exe not in Search Path, Please copy into system or system32 folder", vbOKOnly + vbInformation, "Error"
    ElseIf igAnsCMC = 1 Then
        MsgBox "The Editor Wordpad.Exe not in Search Path, Please copy into system or system32 folder", vbOKOnly + vbInformation, "Error"
    ElseIf igAnsCMC = 2 Then
        MsgBox "The Editor PFE32.Exe is Missing", vbOKOnly + vbInformation, "Error"
    End If
    Resume Next
mnuAccessoryEditErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub mnuAccessoryFileCheck_Click()
    EngrFileCheck.Show vbModal
End Sub

Private Sub mnuAccessoryFileCleanUp_Click()
    EngrCompress.Show vbModal
End Sub

Private Sub mnuAccessoryMessages_Click()
    EngrMessages.Show vbModal
    
End Sub


Private Sub mnuAlert_Click()
    'AlertVw.Show vbModal
End Sub








Private Sub mnuContentHelp_Click()
'    cdcSetup.HelpFile = sgHelpPath & "EngrMain.hlp"
'    cdcSetup.HelpCommand = HELP_INDEX
'    cdcSetup.Action = 6
End Sub

Private Sub mnuCreateConflictTables_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim hlCME As Integer
    Dim hlSEE As Integer
    
    slDate = Format$(gNow(), "ddddd")
    ilRet = MsgBox("Recreate Conflict Tables from " & slDate & ", Continue?", vbQuestion + vbYesNo, "Create Conflict Files")
    If ilRet = vbYes Then
        Screen.MousePointer = vbHourglass
        slStartDate = Format$(gDateValue(slDate), "ddddd")
        slEndDate = "12/31/2069"
        hlCME = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hlCME, "", sgDBPath & "CME.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        hlSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hlSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        ilRet = gConflictTableCreations(slStartDate, slEndDate, hlCME, hlSEE)
        btrDestroy hlCME
        btrDestroy hlSEE
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuCreateSchd_Click()

End Sub

Private Sub mnuDate_Click()
'    Calendar.Show vbModal
End Sub

Private Sub mnuExportAuto_Click()
    Dim slStr As String
'    If igTestSystem Then
'        slStr = "EngrMain^Test\" & sgUserName & "\" & Trim$(Str$(CALLNONE))
'    Else
'        slStr = "EngrMain^Prod\" & sgUserName & "\" & Trim$(Str$(CALLNONE))
'    End If
'    sgCommandStr = slStr
'    On Error Resume Next
'    ExptGen.Show vbModal
'    ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)
'        ChDir sgCurDir
'    End If

End Sub

Private Sub mnuSearchHelp_Click()
'    cdcSetup.HelpFile = sgHelpPath & "EngrMain.hlp"
'    cdcSetup.HelpKey = ""
'    cdcSetup.HelpCommand = HELP_PARTIALKEY
'    cdcSetup.Action = 6
End Sub

Private Sub mnuImportFiles_Click()
    EngrImport.Show vbModal
End Sub

Private Sub mnuTaskBackupCopy_Click()
    EngrBUDataToTest.Show vbModal
End Sub

Private Sub mnuTaskBackupSystem_Click()
    EngrBUZip.Show vbModal
End Sub

Private Sub mnuTaskExit_Click()
    
    Dim ilRet As Integer
    
    
    Unload EngrMain
    If Not igStopCancel Then    'User stop the unload
        'Set EngrMain = Nothing   'Remove data segment
    End If
End Sub

Private Sub mnuTaskJob_Click()
    If igJobVisible Then
        EngrJob.SetFocus
    Else
        EngrJob.Show vbModeless
    End If
End Sub

Private Sub mnuTaskList_Click()
    If igListVisible Then
        EngrList.SetFocus
    Else
        EngrList.Show vbModeless
    End If
End Sub


Private Sub mnuUsingHelp_Click()
'    cdcSetup.HelpFile = sgHelpPath & "EngrMain.hlp"
'    cdcSetup.HelpCommand = HELP_HELPONHELP
'    cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulateListBoxes              *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate all list boxes on      *
'*                     EngrMain form.  This list boxes  *
'*                     used to populate list boxes on  *
'*                     forms.                          *
'*                                                     *
'*******************************************************
Private Sub mPopulateListBoxes()
'
'   mPopulateListBoxes
'   Where:
'
'    On Error GoTo PopListErr
    Exit Sub
PopListErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEngrMainStartUp                 *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Start up                        *
'*                                                     *
'*******************************************************
Private Function mEngrMainStartUp() As Integer
    Dim slVehType As String
    Dim ilRet As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    Dim llNoRec As Long
    Dim ilSetTest As Integer    '0=Test; 1=Set
    Dim llNow As Long
    Dim slStr As String
    Dim slNow As String
    Dim ilLoop As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llDate As Long
    Dim slAppName As String
    Dim ilPos As Integer
    
    'ilRet = 0
    'On Error GoTo mEngrMainStartUpErr:
    'slDateTime1 = FileDateTime(sgExePath & "EngrMain.Exe")
    'If ilRet = 0 Then
    '    sgBrannerMsg = sgBrannerMsg & " (Made:" & Format$(slDateTime1, "ddddd") & ") "
    'End If
    'EngrMain.Caption = sgBrannerMsg
    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"
    'sgExeDirectory = App.Path
    'sgExeDirectory = gSetPathEndSlash(sgExeDirectory)
    ilRet = 0
    On Error GoTo mEngrMainStartUpErr:
    'slDateTime1 = FileDateTime(sgExePath & "EngrMain.Exe")
    slDateTime1 = FileDateTime(sgExeDirectory & slAppName)
    If ilRet = 0 Then
        sgEngrVersion = sgEngrVersion & " created " & Format$(slDateTime1, "ddddd") & " at " & Format$(slDateTime1, "ttttt")
    End If
'    MstPict.Show vbModeless
'    'Mstpict.Move Mstpict.Left, Mstpict.Top, pbcMsgArea.Width, pbcMsgArea.Top
'    If igTaskLoc = 0 Then
'        MstPict.Move 0, 0, pbcMsgArea.Width, Screen.Height
'    Else
'        MstPict.Move 0, 0, pbcMsgArea.Width, pbcMsgArea.Top
'    End If
    If igExitEngrMain Then
        Unload EngrMain
        Set EngrMain = Nothing   'Remove data segment
        Exit Function
    End If
    'gSetUsingEngrMain
    
    Screen.MousePointer = vbHourglass  'Wait
    
    hmRedAlertBitmap = mCopyPictureImage(EngrIcon!pbcRedAlert)
    hmWhiteAlertBitmap = mCopyPictureImage(EngrIcon!pbcWhiteAlert)
    
'    If tgSpf.iGAlertInterval <> 0 Then
'        hgAuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'        ilRet = btrOpen(hgAuf, "", sgDBPath & "AUF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        If ilRet <> BTRV_ERR_NONE Then
'            MsgBox "Unable to Open Alert File, Error = " & Str$(ilRet), vbOKOnly + vbInformation, "Warning"
'        End If
'    End If
    
'    If (tgUrf(0).sCompAlert = "Y") Or (tgUrf(0).sReprintLogAlert = "Y") Then
'        If tgSpf.iGAlertInterval > 0 Then
'            igAlertTimer = tgSpf.iGAlertInterval
'        Else
'            igAlertTimer = 5
'        End If
'        tmcCheckAlert.Enabled = True
'        igAlertFlash = -1
'    End If

    
    Screen.MousePointer = vbHourglass  'Wait
    gEraseVar
    
    Screen.MousePointer = vbHourglass  'Wait
    
    For ilLoop = 0 To 3 Step 1
        imcTask(ilLoop).Visible = True
    Next ilLoop
    imcServiceStatus.Visible = True
    For ilLoop = LBound(igJobShowing) To UBound(igJobShowing) Step 1
        igJobShowing(ilLoop) = False
    Next ilLoop
    igJobShowing(LIBRARYJOB) = 0
    igJobShowing(TEMPLATEJOB) = 0
    igJobShowing(SCHEDULEJOB) = 0
    imcReport.Visible = True
    If igJobStatus(RESOURCEJOB) = 0 Then
        imcTask(0).Enabled = False
    Else
        If igJobStatus(RESOURCEJOB) = 1 Then
            imcTask(0).FontBold = False
        Else
            'imcTask(0).BackColor = vbButtonFace
        End If
    End If
    If igJobStatus(LIBRARYJOB) = 0 Then
        imcTask(1).Enabled = False
    Else
        If igJobStatus(LIBRARYJOB) = 1 Then
            imcTask(1).FontBold = False
        Else
            'imcTask(1).BackColor = vbButtonFace
        End If
    End If
    If igJobStatus(TEMPLATEJOB) = 0 Then
        imcTask(2).Enabled = False
    Else
        If igJobStatus(TEMPLATEJOB) = 1 Then
            imcTask(2).FontBold = False
        Else
            'imcTask(2).BackColor = vbButtonFace
        End If
    End If
    If igJobStatus(SCHEDULEJOB) = 0 Then
        imcTask(3).Enabled = False
    Else
        If igJobStatus(SCHEDULEJOB) = 1 Then
            imcTask(3).FontBold = False
        Else
            'imcTask(3).BackColor = vbButtonFace
        End If
    End If
    'imcReport.Left = imcTask(3).Left + imcTask(3).Width + 15
    imcReport.Left = imcServiceStatus.Left + imcServiceStatus.Width + 15
    imcReport.Width = lgCurrHRes * 15 - imcReport.Left
    
    
    'Move report logo to local C drice (c:\csi\rptlogo.bmp)
    ilRet = 0
    On Error GoTo mEngrMainStartUpErr:
    slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
    If ilRet <> 0 Then
        ilRet = 0
        MkDir "C:\CSI"
        If ilRet = 0 Then
            FileCopy sgLogoDirectory & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
        Else
            FileCopy sgLogoDirectory & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
        End If
    Else
        ilRet = 0
        slDateTime2 = FileDateTime(sgLogoDirectory & "RptLogo.Bmp")
        If ilRet = 0 Then
            If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
                FileCopy sgLogoDirectory & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
            End If
        End If
    End If
    On Error GoTo 0
'    If (Not igShowPicture) Or (igRemoteUser) Or (igTestSystem) Then
'        Basic10.Show vbModeless
'    End If
    
    
    mEngrMainStartUp = True
    Exit Function
mEngrMainStartUpErr:
    ilRet = Err.Number
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnloadForms                    *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove any active form.  Test if*
'*                     the user stopped the unload     *
'*                                                     *
'*******************************************************
Private Sub mUnloadForms()
'
'   mUnloadForms iSignon
'   Where:
'       iSignon (I)- True = Unload Signon, False = leave Signon loaded
'
    If igTestSystem Then
        Unload EngrMstPict
        If igStopCancel Then    'User stop the unload
            Exit Sub
        End If
        Set EngrMstPict = Nothing   'Remove data segment
    End If
End Sub


Private Sub pbcMsgArea_GotFocus()
'    'This control will get focus if no other form is active
'    'This control also get focus at various times- therefore the test on
'    'the number of forms running
'    If Forms.Count = igNoMinForms + igCalActive + igCalcActive Then 'Show basic 10 if only this form to showing
'        If Not igShowPicture Then
'            If igBasicOrAux = 2 Then
'                EngrMain!cmcAuxWnd.Value = True
'            Else
'                EngrMain!cmcBasicWnd.Value = True
'            End If
'        Else
'            If (igCalActive = 0) And (igCalcActive = 0) Then
'                'cmcJobs.Enabled = True
'                'cmcLists.Enabled = True
'                igShowPicture = False
'                'EngrMain.plcHelp.Caption = " press JOBS or LISTS button to continue"
'                'EngrMain!plcHelp.Cls
'                'EngrMain!plcHelp.CurrentX = 0
'                'EngrMain!plcHelp.CurrentY = 0
'                'EngrMain!plcHelp.Print " press JOBS or LISTS button to continue"
'            End If
'        End If
'    End If
End Sub




Private Sub tmcCheckAlert_Timer()
'    Dim ilRet As Integer
'
'    ilRet = gAlertCheck()
End Sub

Private Sub tmcClock_Timer()
    Dim slStr As String
    Dim slDateTime As String
    Dim ilPos As Integer
    
    slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    
    mnuDate.Caption = slDateTime & "                       "
    'Check Service status
    mCheckServiceStatus
End Sub

Private Sub mShowAlert(hlBitmap As Long)
    Dim ilRet As Integer
    Dim llMenuID As Long
    
    ' Get a handle to the top level menu
    hmAlert = GetMenu(EngrMain.hwnd)
    lmMenuID = GetMenuItemID(hmAlert, 3)
    
    'If hmBitmap = 0 Then hmBitmap = CopyPictureImage(IconTraf!pbcAlert)
    ' And replace it with a bitmap.
    ilRet = ModifyMenuBynum(hmAlert, 3, MF_BITMAP Or MF_BYPOSITION Or MF_ENABLED, llMenuID, hlBitmap)
    
End Sub

Private Sub tmcFlashAlert_Timer()
    igAlertFlash = igAlertFlash + 1
'    If igAlertFlash = 10 Then
'        If ((tgSpf.iGAlertInterval <> 0) And (tgSpf.iGAlertInterval <= igAlertTimer)) Then
'            tmcFlashAlert.Enabled = False
'            mnuAlert.Visible = False
'            DoEvents
'            mnuAlert.Visible = True
'            mShowAlert hmRedAlertBitmap
'            tmcCheckAlert.Enabled = True
'        Else
'            mnuAlert.Visible = False
'            DoEvents
'            mnuAlert.Visible = True
'            mShowAlert hmRedAlertBitmap
'            tmcFlashAlert.Interval = 60000  'Every minute
'            igAlertFlash = 0
'            igAlertTimer = igAlertTimer + 1
'        End If
'        Exit Sub
'    Else
'        tmcFlashAlert.Interval = 2000  'Every 2 seconds
'        If igAlertFlash Mod 2 = 0 Then
'            mnuAlert.Visible = False
'            DoEvents
'            mnuAlert.Visible = True
'            mShowAlert hmRedAlertBitmap
'        Else
'            mnuAlert.Visible = False
'            DoEvents
'            mnuAlert.Visible = True
'            mShowAlert hmWhiteAlertBitmap
'        End If
'        DoEvents
'    End If
End Sub

Private Sub mEraseArrays()
    Erase tgCurrACE
    Erase tgCurrADE
    Erase tgCurrAEE
    Erase tgCurrAFE
    Erase tgCurrAPE
    Erase tgCurrANE
    Erase tgBothANE
    Erase tgUsedANE
    Erase tgCurrASE
    Erase tgBothASE
    Erase tgCurrATE
    Erase tgBothATE
    Erase tgUsedATE
    Erase tgCurrBDE
    Erase tgBothBDE
    Erase tgUsedBDE
    Erase tgCurrBGE
    Erase tgCurrBSE
    Erase tgCurrCCE
    Erase tgCurrAudioCCE
    Erase tgUsedAudioCCE
    Erase tgCurrBusCCE
    Erase tgUsedBusCCE
    Erase tgCurrCTE
    Erase tgCurrDEE
    Erase tgCurrDHE
    Erase tgCurrLibDHE
    Erase tgBothLibDHE
    Erase tgCurrTempDHE
    Erase tgCurrDNE
    Erase tgCurrLibDNE
    Erase tgCurrTempDNE
    Erase tgCurrDSE
    Erase tgCurrEBE
    Erase tgCurrEPE
    Erase tgCurrETE
    Erase tgUsedETE
    Erase tgCurrFNE
    Erase tgUsedFNE
    Erase tgCurrITE
    Erase tgCurrMTE
    Erase tgUsedMTE
    Erase tgCurrNNE
    Erase tgUsedNNE
    Erase tgCurrRNE
    Erase tgUsedRNE
    Erase tgCurrSCE
    Erase tgUsedSCE
    Erase tgCurrSGE
    Erase tgCurrSOE
    Erase tgCurrSPE
    Erase tgCurrTSE
    Erase tgAirInfoTSE
    Erase tgCurrTTE
    Erase tgCurrStartTTE
    Erase tgUsedStartTTE
    Erase tgCurrEndTTE
    Erase tgUsedEndTTE
    Erase tgCurrTNE
    Erase tgCurrUIE
    Erase tgCurrUTE
    Erase tgSpotCurrSEE
    
    Erase lgLibDheUsed
    
    Erase tgJobTaskNames
    Erase tgListTaskNames
    Erase tgExtraTaskNames
    Erase tgAlertTaskNames
    Erase tgNoticeTaskNames
    
    Erase tgDDFFileNames
    
    Erase tgReportNames
    
    Erase tgFilterValues
    Erase tgFilterFields
    
    Erase tgSchdReplaceValues
    Erase tgReplaceFields
    Erase tgLibReplaceValues
    
    Erase tgItemIDChk
End Sub

Private Sub mCheckServiceStatus()
    Dim ilRet As Integer
    Dim llServiceDate As Long
    Dim llServiceTime As Long
    
    ilRet = gGetServiceStatus_MIE_MessageInfo("Main", tgMie)
    If (ilRet) And (tgMie.lCode > 0) Then
        llServiceDate = gDateValue(tgMie.sEnteredDate)
        llServiceTime = gTimeToLong(tgMie.sEnteredTime, True)
        If lgLastServiceTime < 0 Then
            lgLastServiceDate = llServiceDate
            lgLastServiceTime = llServiceTime
            igCountTimeNotChanged = 0
            imcServiceStatus.Caption = "SERVICE STATUS: NOT DETERMINED"
            imcServiceStatus.BackColor = vbButtonText
            Exit Sub
        End If
        If lgLastServiceTime <> llServiceTime Then
            igCountTimeNotChanged = 0
            If lgLastServiceTime + 300 >= llServiceTime Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
                imcServiceStatus.Caption = "SERVICE STATUS: RUNNING"
                imcServiceStatus.BackColor = DARKGREEN
                Exit Sub
            End If
            If (lgLastServiceTime > 85800) And (llServiceTime < 600) Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
                imcServiceStatus.Caption = "SERVICE STATUS: RUNNING"
                imcServiceStatus.BackColor = DARKGREEN
                Exit Sub
            End If
            If lm1970 <> llServiceDate Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
                imcServiceStatus.Caption = "SERVICE STATUS: RUNNING"
                imcServiceStatus.BackColor = DARKGREEN
                Exit Sub
            End If
        Else
            igCountTimeNotChanged = igCountTimeNotChanged + 1
            If igCountTimeNotChanged < 5 Then
                Exit Sub
            End If
        End If
    End If
    imcServiceStatus.Caption = "SERVICE STATUS: NOT RUNNING"
    imcServiceStatus.BackColor = vbRed
    If tgMie.lCode > 0 Then
        tgMie.sEnteredDate = Format$("1/1/1970", sgShowDateForm)
        ilRet = gPutUpdate_MIE_MessageInfo(tgMie, "EngrServiveMain")
    End If
    igCountTimeNotChanged = 0
End Sub

