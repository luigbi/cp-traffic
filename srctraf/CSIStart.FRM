VERSION 5.00
Begin VB.Form CSIStart 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4695
   ClientLeft      =   2190
   ClientTop       =   2490
   ClientWidth     =   8775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "CSIStart.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4695
   ScaleWidth      =   8775
   Begin VB.Timer tmcMoveLine 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   8400
      Top             =   3825
   End
   Begin VB.FileListBox lbcFontNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   7380
      MultiSelect     =   2  'Extended
      Pattern         =   "*.ttf"
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4365
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.FileListBox lbcFontNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   5745
      MultiSelect     =   2  'Extended
      Pattern         =   "*.ttf"
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4350
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.FileListBox lbcFileNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3375
      MultiSelect     =   2  'Extended
      Pattern         =   "*.dll;*ocx;*.msi;*.bat"
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.DriveListBox cbcDrives 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   7
      Top             =   4365
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox plcCSIStart 
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
      Height          =   3960
      Left            =   495
      ScaleHeight     =   3930
      ScaleWidth      =   7755
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   7785
      Begin VB.PictureBox pbcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   540
         Picture         =   "CSIStart.frx":08CA
         ScaleHeight     =   645
         ScaleWidth      =   735
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox plcWName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   1665
         Picture         =   "CSIStart.frx":1801
         ScaleHeight     =   570
         ScaleWidth      =   4515
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   4515
      End
      Begin VB.PictureBox plcWelcome 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
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
         Height          =   3225
         Left            =   -15
         Picture         =   "CSIStart.frx":5022
         ScaleHeight     =   3195
         ScaleWidth      =   7755
         TabIndex        =   1
         Top             =   735
         Width           =   7785
         Begin VB.Label lacWCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Welcome"
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
            Height          =   225
            Left            =   90
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to Counterpoint Software System"
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
            Height          =   300
            Index           =   0
            Left            =   255
            TabIndex        =   2
            Top             =   390
            Width           =   7185
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Index           =   1
            Left            =   375
            TabIndex        =   3
            Top             =   1080
            Width           =   6945
         End
         Begin VB.Label lacWMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Index           =   2
            Left            =   315
            TabIndex        =   4
            Top             =   1515
            Width           =   7110
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   105
         X2              =   195
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   0
         X2              =   7755
         Y1              =   705
         Y2              =   705
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6855
      Top             =   4245
   End
End
Attribute VB_Name = "CSIStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy

'**********************************************************
'                Btrieve Match MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************


'
'
'This requires to be compiled 3 times:
'Once as Traffic
'Once as Affiliat
'Once as Engineer
'
'



Option Explicit
Option Compare Text
Dim hmmsg As Integer   'From file hanle
Dim hmFile As Integer           'file handle
Dim hmTo As Integer
Dim smAppName As String
Dim smExePath As String
Dim smFontPath As String
Dim smSystemPath As String
Dim smDBPath As String
Dim smMessagePath As String
Dim smCrystal10Path As String
Dim smWinDir As String
Dim imUseCrystal10 As Integer

Dim smNowDate As String   'Todays date
Dim imCancelled As Integer
Dim imTerminate As Integer
Dim imProcessing As Integer
Dim smTodaysDate As String
Dim smCurParamters As String

Dim imErrorMsg As Integer
Dim imShowErrMsg As Integer
'5676
Dim smRootDrive As String

Private Const MAX_PATH = 260
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1   'Unicode null
Private Const STATUSNOLOG = 0
Private Const STATUSLOGUNKNOWN = 3
Private Const STATUSGOOD = 1
Private Const STATUSERROR = 2
Private Const CSIVERSION = 81


Private Sub Form_Activate()
    plcWelcome.BorderStyle = 0
    plcWelcome.PaintPicture plcWelcome.Picture, 0, 0
    tmcStart.Enabled = True
End Sub

Private Sub Form_Load()
    Dim ilPos As Integer
    smCurParamters = Command$
    smExePath = App.Path
    smAppName = App.EXEName
    ilPos = InStr(1, smExePath, "\Src", vbTextCompare)
    If ilPos > 0 Then
        smExePath = left$(smExePath, ilPos) & "exe"
    End If
    If App.PrevInstance Then
        End
    Else
        mInit
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    'Reset used instead of Close to cause # Clients on network to be decrement
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
    End
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
    
    Dim ilRet As Integer   'Return from btrieve calls
    Dim slDate As String
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilPass As Integer
    Dim ilPos As Integer
    Dim slVersion As String
    Dim slTest As String
    Dim slStr As String
    
    
    
    imTerminate = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    plcCSIStart.Move 495, 510
    plcWelcome.Move 0, 735
    
    ilPos = InStr(1, smExePath, "\exe", vbTextCompare)
    If ilPos > 0 Then
        smFontPath = left$(smExePath, ilPos) & "Setup\Fonts"
        smSystemPath = left$(smExePath, ilPos) & "Setup\ClientInstall\System32"
        slStr = left$(smExePath, ilPos - 1)
        ilPos = InStrRev(slStr, "\", -1, vbTextCompare)
        If ilPos > 0 Then
            smCrystal10Path = left$(slStr, ilPos) & "Crystal10"
        Else
            imTerminate = True
            Exit Sub
        End If
    Else
        imTerminate = True
        Exit Sub
    End If
    smDBPath = CurDir$
    ilPos = InStr(1, smDBPath, "Data", vbTextCompare)
    If ilPos > 0 Then
        smMessagePath = smDBPath & "\Messages"
        On Error GoTo mMakeFolderErr:
        ilRet = 0
        MkDir smMessagePath
        If (ilRet <> 0) And (ilRet <> 75) Then
            smMessagePath = ""
        End If
    Else
        ilPos = InStr(1, smExePath, "\exe", vbTextCompare)
        If ilPos > 0 Then
            smMessagePath = left$(smExePath, ilPos) & "data\messages"
            On Error GoTo mMakeFolderErr:
            ilRet = 0
            MkDir smMessagePath
            If (ilRet <> 0) And (ilRet <> 75) Then
                smMessagePath = ""
            End If
        Else
            smMessagePath = ""
        End If
    End If
    '5676
    If Dir("c:\") = "" Then
        gLoadFromIni "Locations", "RootDrive", CurDir$ & "\Traffic.ini", smRootDrive
        If smRootDrive = "Not Found" Then
            smRootDrive = "C:\"
        Else
            smRootDrive = gSetPathEndSlash(smRootDrive)
        End If
    Else
        smRootDrive = "c:\"
    End If
    CSIStart.Move (Screen.Width - CSIStart.Width) \ 2, (Screen.Height - CSIStart.Height) \ 2 + 175 '+ Screen.Height \ 10
    CSIStart.Show
    
    smNowDate = Format$(Now, "m/d/yy")
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    mTerminate
    Exit Sub
mInitPath:
    ilRet = 1
    Resume Next
mMakeFolderErr:
    ilRet = Err.Number
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInstall                        *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create folders, Copy Images,   *
'*                      unzip images                   *
'*                                                     *
'*******************************************************
Private Function mInstall() As Integer
'    Dim slSource As String
'    Dim slDestination As String
'    Dim ilRet As Integer
'    Dim ilSourceCount As Integer
'    Dim ilIndex As Integer
'    Dim ilLoop As Integer
'    Dim ilBase As Integer
'    Dim ilPercent As Integer
'    Dim slDrivePath As String
'    Dim slMsgFile As String
'    Dim ilError As Integer
'    Dim slName As String
'    Dim slToName As String
'    Dim slDateTime As String
'    Dim llPercent As Long
'    Dim ilCount As Integer
'    ReDim ilPos(0 To 1) As Integer
'
'    imProcessing = True
'    Screen.MousePointer = vbHourglass
'    lacFinish(4).Caption = ""
'    lacProgress(1).Visible = False
'    lacProgress(2).Visible = False
'    lbcDatabases.Clear
'    lncProgress.X2 = lncProgress.X1
'    DoEvents
'    'Determine number of modules to be copied and unzipped
'    slSource = edcSource.Text
'    If right(slSource, 1) = "\" Then
'        slSource = slSource & "csi" '\prod"
'    ElseIf right(slSource, 1) = ":" Then
'        slSource = slSource & "\csi"    '\prod"
'    ElseIf Len(slSource) > 1 Then
'        slSource = slSource & "\csi"    '\prod"
'    Else
'        slSource = slSource & ":\csi"   '\prod"
'    End If
'    If right$(slSource, 1) <> "\" Then
'        slSource = slSource & "\"
'    End If
'    If Not imCSIExist Then
'        If rbcProdTest(0).Value Then
'            smDProd = "prod\"
'        Else
'            smDProd = "test\"
'        End If
'    End If
''    If rbcProdTest(1).Value Then
''        smDProd = "Test\"
''    End If
'    slDestination = edcDestination.Text
'    If right$(slDestination, 1) <> "\" Then
'        slDestination = slDestination & "\"
'    End If
'    slDestination = slDestination & smDCsi
'    ilRet = 0
'    On Error GoTo mMakeFolderErr:
'    lbcSource.Path = slSource & smSProd  'left$(slSource, Len(slSource) - 5)
'    If lbcSource.ListCount <= 0 Then
'        lacFinish(3).Caption = ""
'        imProcessing = False
'        Screen.MousePointer = vbDefault
'        MsgBox "Unable to find " & slSource & smSProd & " Error # " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'        mInstall = False
'        Exit Function
'    End If
'    lbcSource.Path = left$(slSource, Len(slSource) - 1)
'    ReDim smSourceFolder(0 To lbcSource.ListCount) As String
'
'    For ilLoop = 0 To lbcSource.ListCount - 1 Step 1
'        smSourceFolder(ilLoop) = lbcSource.List(ilLoop)
'    Next ilLoop
'    ilLoop = 0
'    Do
'        lbcSource.Path = smSourceFolder(ilLoop)
'        If lbcSource.ListCount > 0 Then
'            ilBase = UBound(smSourceFolder)
'            ReDim Preserve smSourceFolder(0 To ilBase + lbcSource.ListCount) As String
'            For ilIndex = 0 To lbcSource.ListCount - 1 Step 1
'                smSourceFolder(ilIndex + ilBase) = lbcSource.List(ilIndex)
'            Next ilIndex
'        End If
'        ilLoop = ilLoop + 1
'    Loop While ilLoop < UBound(smSourceFolder)
'
'    For ilLoop = 0 To UBound(smSourceFolder) - 1 Step 1
'        lbcSource.Path = smSourceFolder(ilLoop)
'        lbcSourceFiles.Path = lbcSource.Path
'        ilSourceCount = ilSourceCount + lbcSourceFiles.ListCount
'        'If lbcSource.ListCount > 0 Then
'        '    ilBase = UBound(smSourceFolder)
'        '    ReDim Preserve smSourceFolder(0 To ilBase + lbcSource.ListCount) As String
'        '    For ilIndex = 0 To lbcSource.ListCount - 1 Step 1
'        '        smSourceFolder(ilIndex + ilBase) = lbcSource.List(ilIndex)
'        '    Next ilIndex
'        'End If
'    Next ilLoop
'    'Determine if base path previously created. If not, then create
'    slDrivePath = left$(slDestination, Len(slDestination) - 1)
'    ilPos(0) = InStr(1, slDrivePath, "\", 1)
'    If ilPos(0) > 0 Then
'        Do
'            ilPos(1) = InStr(ilPos(0) + 1, slDrivePath, "\", 1)
'            If ilPos(1) > 0 Then
'                On Error GoTo mMakeFolderErr:
'                ilRet = 0
'                MkDir left$(slDrivePath, ilPos(1) - 1)
'                If (ilRet <> 0) And (ilRet <> 75) Then
'                    imProcessing = False
'                    Screen.MousePointer = vbDefault
'                    MsgBox "Unable to create " & slDrivePath & " Error # " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                    smResultFilePath = slDestination & smDProd
'                    mInstall = False
'                    Exit Function
'                End If
'                ilPos(0) = ilPos(1)
'            Else
'                If Len(slDrivePath) > ilPos(0) + 1 Then
'                    On Error GoTo mMakeFolderErr:
'                    ilRet = 0
'                    MkDir slDrivePath
'                    If (ilRet <> 0) And (ilRet <> 75) Then
'                        imProcessing = False
'                        Screen.MousePointer = vbDefault
'                        MsgBox "Unable to create " & slDrivePath & " Error # " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                        smResultFilePath = slDestination & smDProd
'                        mInstall = False
'                        Exit Function
'                    End If
'                End If
'                Exit Do
'            End If
'        Loop While ilPos(1) > 0
'    End If
'    smResultFilePath = slDestination
'    If Not mOpenMsgFile(slMsgFile) Then
'        imProcessing = False
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'    Print #hmMsg, "CSI Install " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'    Print #hmMsg, ""
'    lacFinish(3).Caption = "Results of Install stored into " & slMsgFile
'    'Test if traffic/Affiliat are being used, copy images to safe folder (Updates), then try to copy it back into
'    'exe folder.  If unable to copy it back in, then it is in use.  First copying it to Updates saves the image
'    'in case it is not part of the release.
'    If imCSIExist Then
'        lacProgress(3).Caption = "Checking if Traffic System in Use"
'        DoEvents
'        ilRet = 0
'        On Error GoTo mMakeFolderErr:
'        slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Traffic.exe")
'        If ilRet = 0 Then
'            ilRet = 0
'            On Error GoTo mMakeFolderErr:
'            FileCopy smResultFilePath & smDProd & "Exe\Traffic.exe", smResultFilePath & smDProd & "Updates\Traffic.Exe"
'            If ilRet <> 0 Then
'                imProcessing = False
'                Screen.MousePointer = vbDefault
'                Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Traffic.exe" & " To " & smResultFilePath & smDProd & "Updates\Traffic.Exe" & " Error #" & Str$(ilRet)
'                MsgBox "Unable to copy traffic.exe to updates folder.  See " & slMsgFile & " for Error messages", vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                smResultFilePath = slDestination & smDProd
'                mInstall = False
'                Exit Function
'            End If
'            ilRet = 0
'            FileCopy smResultFilePath & smDProd & "Updates\Traffic.exe", smResultFilePath & smDProd & "Exe\Traffic.Exe"
'            If ilRet <> 0 Then
'                imProcessing = False
'                Screen.MousePointer = vbDefault
'                Print #hmMsg, "Traffic in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Traffic.exe" & " To " & smResultFilePath & smDProd & "Exe\Traffic.Exe" & " Error #" & Str$(ilRet)
'                MsgBox "Traffic in use.  See " & slMsgFile & " for Error messages", vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                smResultFilePath = slDestination & smDProd
'                mInstall = False
'                Exit Function
'            End If
'            On Error Resume Next
'            Kill smResultFilePath & smDProd & "Updates\Traffic.exe"
'        End If
'        lacProgress(3).Caption = "Checking if Affiliate System in Use"
'        DoEvents
'        ilRet = 0
'        On Error GoTo mMakeFolderErr:
'        slDateTime = FileDateTime(smResultFilePath & smDProd & "Exe\Affiliat.exe")
'        If ilRet = 0 Then
'            ilRet = 0
'            On Error GoTo mMakeFolderErr:
'            FileCopy smResultFilePath & smDProd & "Exe\Affiliat.exe", smResultFilePath & smDProd & "Updates\Affiliat.Exe"
'            If ilRet <> 0 Then
'                imProcessing = False
'                Screen.MousePointer = vbDefault
'                Print #hmMsg, "Unable to Copy " & smResultFilePath & smDProd & "Exe\Affiliat.exe" & " To " & smResultFilePath & smDProd & "Updates\Affiliat.Exe" & " Error #" & Str$(ilRet)
'                MsgBox "Unable to copy traffic.exe to updates folder.  See " & slMsgFile & " for Error messages", vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                smResultFilePath = slDestination & smDProd
'                mInstall = False
'                Exit Function
'            End If
'            ilRet = 0
'            FileCopy smResultFilePath & smDProd & "Updates\Affiliat.exe", smResultFilePath & smDProd & "Exe\Affiliat.Exe"
'            If ilRet <> 0 Then
'                imProcessing = False
'                Screen.MousePointer = vbDefault
'                Print #hmMsg, "Traffic in Use, Unable to Copy " & smResultFilePath & smDProd & "Updates\Affiliat.exe" & " To " & smResultFilePath & smDProd & "Exe\Affiliat.Exe" & " Error #" & Str$(ilRet)
'                MsgBox "Traffic in use.  See " & slMsgFile & " for Error messages", vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                smResultFilePath = slDestination & smDProd
'                mInstall = False
'                Exit Function
'            End If
'            On Error Resume Next
'            Kill smResultFilePath & smDProd & "Updates\Affiliat.exe"
'        End If
'    End If
'    'Move Zero BTR if updating system to temporary folder so it can be used to place into databases prior to running DDFReorg
'    If imCSIExist Then
'        lacProgress(3).Caption = "Copying Zero_Btr to temporary Folder"
'        DoEvents
'        mRemoveTempBTR slDestination & smDProd & "Updates"
'        ilRet = mMoveZeroBTR(slDestination & smDProd & "Zero_BTR", slDestination & smDProd & "Updates")
'        If Not ilRet Then
'            imProcessing = False
'            Screen.MousePointer = vbDefault
'            MsgBox "Unable to copy zero_btr files to updates folder.  See " & slMsgFile & " for Error messages", vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'            smResultFilePath = slDestination & smDProd
'            mInstall = False
'            Exit Function
'        End If
'    End If
'    'Create All Folders
'    lacProgress(3).Caption = "Creating Folders and Subfolders"
'    DoEvents
'    ilRet = 0
'    ilError = False
'    For ilLoop = 0 To UBound(smSourceFolder) - 1 Step 1
'        slToName = Mid$(smSourceFolder(ilLoop), Len(slSource) + 1)
'        ''If (InStr(1, slToName, "Zip", vbTextCompare) <= 0) Then
'        'If (StrComp(slToName, "Zip", vbTextCompare) = 0) Then
'        If (InStr(1, slToName, "Zip", vbTextCompare) <= 0) Or (InStr(1, slToName, "WinZip", vbTextCompare) > 0) Then
'            If InStr(1, slToName, "Prod", vbTextCompare) > 0 Then
'                If StrComp(slToName, smSProd, vbTextCompare) = 0 Then
'                    slToName = slDestination & left$(smDProd, Len(smDProd) - 1)
'                Else
'                    slToName = slDestination & smDProd & Mid$(slToName, Len(smSProd) + 2)
'                End If
'            Else
'                slToName = slDestination & slToName
'            End If
'
'            'slDrivePath = slDestination & Mid$(smSourceFolder(ilLoop), Len(slSource) + 1)
'            slDrivePath = slToName
'            On Error GoTo mMakeFolderErr:
'            ilRet = 0
'            MkDir slDrivePath
'            If ilRet = 0 Then
'                Print #hmMsg, slDrivePath & " Created"
'            ElseIf ilRet = 75 Then
'                Print #hmMsg, slDrivePath & " Previously Existed"
'            Else
'                Print #hmMsg, "Unable to create " & slDrivePath & " Error #" & Str$(ilRet)
'                ilError = True
'            End If
'        End If
'    Next ilLoop
'    If ilError Then
'        imProcessing = False
'        Screen.MousePointer = vbDefault
'        MsgBox "Unable to create required Folders.  See " & slMsgFile & " for Error messages", vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'        smResultFilePath = slDestination & smDProd
'        mInstall = False
'        Exit Function
'    End If
'    lacProgress(1).Caption = "Installing into:"
'    lacProgress(2).Caption = "File:"
'    lacProgress(1).Visible = True
'    lacProgress(2).Visible = True
'    lacProgress(3).Caption = ""
'    DoEvents
'    'Copy or unzip files into folders
'    ilRet = 0
'    ilCount = 0
'    For ilLoop = 0 To UBound(smSourceFolder) - 1 Step 1
'        smModFilePath = smSourceFolder(ilLoop) & "\"
'        'lbcSource.Path = smSourceFolder(ilLoop)
'        lbcSourceFiles.Path = smSourceFolder(ilLoop)
'        For ilIndex = 0 To lbcSourceFiles.ListCount - 1 Step 1
'            slName = lbcSourceFiles.List(ilIndex)
'            lacProgreFileListBox.Caption = slName
'            slToName = Mid$(smSourceFolder(ilLoop), Len(slSource) + 1)
'            ''If (InStr(1, slToName, "Zip", vbTextCompare) > 0) Then
'            'If (StrComp(slToName, "Zip", vbTextCompare) <> 0) Then
'            If (InStr(1, slToName, "Zip", vbTextCompare) > 0) And (InStr(1, slToName, "WinZip", vbTextCompare) <= 0) Then
'                smResultFilePath = slDestination & smDProd
'                ilRet = mUnZipFiles(slName)
'            Else
'                If InStr(1, slToName, smSProd, vbTextCompare) > 0 Then
'                    slToName = slDestination & smDProd & Mid$(slToName, Len(smSProd) + 2) & "\"
'                Else
'                    slToName = slDestination & slToName & "\"
'                End If
'                'smResultFilePath = slToName & Mid$(smSourceFolder(ilLoop), Len(slSource) + 1) & "\"
'                smResultFilePath = slToName
'                lacProgressFolder.Caption = left$(smResultFilePath, Len(smResultFilePath) - 1)
'                ilRet = mCopyFile(slName)
'            End If
'            If imCancelled Then
'                Screen.MousePointer = vbDefault
'                plcProgress.Visible = False
'                plcCancel.Visible = True
'                cmcNext.Enabled = True
'                imCancelAns = 0
'                Do
'                    DoEvents
'                Loop While imCancelAns = 0
'                plcProgress.Visible = True
'                plcCancel.Visible = False
'                cmcNext.Enabled = False
'                If imCancelAns = 2 Then
'                    imProcessing = False
'                    Screen.MousePointer = vbDefault
'                    Print #hmMsg, ""
'                    Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'                    Close #hmMsg
'                    smResultFilePath = slDestination & smDProd
'                    mInstall = False
'                    Exit Function
'                Else
'                    Screen.MousePointer = vbHourglass
'                    imCancelled = False
'                End If
'            End If
'            ilCount = ilCount + 1
'            lacProgress(3).Caption = Str$(ilCount) & " of" & Str$(ilSourceCount)
'            llPercent = (CLng(ilCount) * 100) \ ilSourceCount
'            If llPercent >= 100 Then
'                llPercent = 100
'            End If
'            lncProgress.X2 = ((pbcProgress.Width - 2 * lncProgress.X1) * llPercent) / 100
'            If llPercent > 0 Then
'                lncProgress.Visible = True
'            End If
'            DoEvents
'        Next ilIndex
'    Next ilLoop
'    smResultFilePath = slDestination & smDProd
'    mCheckDDF
'    If lbcDatabases.ListCount <= 0 Then
'        mRemoveTempBTR smResultFilePath & "Updates"
'        Print #hmMsg, "CSI File Images Install Successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'        Print #hmMsg, ""
'        Close #hmMsg
'        lacFinish(4).Caption = "Installed" & Str$(ilCount) & " of" & Str$(ilSourceCount) & " files"
'    Else
'        Print #hmMsg, "CSI File Images Install Successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'        lacFinish(4).Caption = "Installed" & Str$(ilCount) & " of" & Str$(ilSourceCount) & " files"
'    End If
'    imProcessing = False
'    mInstall = True
'    Exit Function
'mMakeFolderErr:
'    ilRet = Err
'    Resume Next
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
Private Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    imShowErrMsg = False
    If smMessagePath = "" Then
        slMsgFile = ""
        hmmsg = -1
        mOpenMsgFile = False
        Exit Function
    End If
    On Error GoTo mOpenMsgFileErr:
    ilRet = 0
    'slToFile = smMessagePath & "\CSIStart.Txt"
    '5676
   ' slToFile = "c:\csi\CSIStart.Txt"
    slToFile = smRootDrive & "csi\CSIStart.Txt"
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
'Create new file each time-Jim 11/5/03
'It appears that the DataValue routine is not functioning as it is being used
'because the section is always being appended
'        slFileDate = Format$(slDateTime, "m/d/yy")
'        If DateValue(slFileDate) = DateValue(smNowDate) Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                'Screen.MousePointer = vbDefault
'                'MsgBox "Open " & slToFile & " Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'                'mOpenMsgFile = False
'                slMsgFile = ""
'                hmMsg = -1
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmmsg = FreeFile
            Open slToFile For Output As hmmsg
            If ilRet <> 0 Then
                'Screen.MousePointer = vbDefault
                'MsgBox "Open " & slToFile & " Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                'mOpenMsgFile = False
                slMsgFile = ""
                hmmsg = -1
                mOpenMsgFile = False
                Exit Function
            End If
            'imShowErrMsg = True
'        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmmsg = FreeFile
        Open slToFile For Output As hmmsg
        If ilRet <> 0 Then
            'Screen.MousePointer = vbDefault
            'MsgBox "Open " & slToFile & " Error #" & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            'mOpenMsgFile = False
            slMsgFile = ""
            hmmsg = -1
            mOpenMsgFile = False
            Exit Function
        End If
        imShowErrMsg = True
    End If
    On Error GoTo 0
    Print #hmmsg, "Checking DLL, OCX and Fonts on " & Format(Now, "m/d/yy") & " at " & Format(Now, "h:mm:ssAM/PM")
    Print #hmmsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/17/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate Links                *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbDefault
    Unload CSIStart
    Set CSIStart = Nothing
    End
End Sub

Private Sub tmcMoveLine_Timer()
    If Line1.X1 + Line1.X2 >= Line3.X1 + Line3.X2 Then
        Line1.X1 = Line3.X1
        Line1.X2 = Line1.X1 + 90
    Else
        Line1.X1 = Line1.X1 + Line1.X2
        Line1.X2 = Line1.X1 + 90
    End If
End Sub

Private Sub tmcStart_Timer()
    Dim ilRet As Integer
    Dim llRet As Long
    Dim slFileName As String
    Dim llWinLen As Long
    Dim slWindowsDir As String * MAX_PATH
    
    tmcStart.Enabled = False
    If imTerminate Then
        MsgBox "Please call Counterpoint as this program is not in the exe folder but in " & smExePath, vbOKOnly, "Setup Error"
        mTerminate
        Exit Sub
    End If
    On Error GoTo mTMCStartErr:
    imErrorMsg = False
    Screen.MousePointer = vbHourglass
    Line1.X1 = Line3.X1
    Line1.X2 = Line1.X1 + 90
    Line1.Y1 = Line3.Y1
    Line1.Y2 = Line1.Y1
    Line1.ZOrder
    Line1.Visible = True
    tmcMoveLine.Enabled = True
    mCreateCSI
    ilRet = mOpenMsgFile(slFileName)
    llWinLen = GetWindowsDirectory(slWindowsDir, MAX_PATH)
    smWinDir = left$(slWindowsDir, llWinLen)
    'Dan M 5/05/11  Citrix writes out: C:\Documents and Settings\kwatts\WINDOWS..so just test for 'windows'
'    If (InStr(1, smWinDir, ":\winnt", vbTextCompare) <> 0) Or (InStr(1, smWinDir, ":\windows", vbTextCompare) <> 0) Then
    If (InStr(1, smWinDir, ":\winnt", vbTextCompare) <> 0) Or (InStr(1, smWinDir, "windows", vbTextCompare) <> 0) Then
        'Test if Fonts exist, if not then not standard window
        'Dan M 3/21/2011 old method blocked continuing if something odd with font directory.  Now, let it continue installing.
'        ilRet = 0
        'lbcFontNames(0).Path = smWinDir & "\Fonts"
'        If ilRet = 0 Then
'            ilRet = mCheckSystem32()
'            'vcRedist no longer used. Dan 8/27/10
''            If ilRet = 44 Then
''                MsgBox "Missing VcRedist executable. Please call Counterpoint", vbOKOnly, "Setup Error"
''                mTerminate
''                Exit Sub
''            ElseIf ilRet = 0 Then
'             If ilRet = 0 Then
'                Print #hmmsg, "          Certain products were not installed properly."
'            End If
'            'Dan M 8/30/2010 get csiNetReporter file version
'            ilRet = mGetVersionInfo(smWinDir & "\system32\csiNetReporter.exe")
'            If (hmmsg <> -1) Then
'                Print #hmmsg, "             CsiNetReporter V:" & ilRet / 10
'            End If
'            ilRet = mGetVersionInfo(smWinDir & "\system32\csiNetReporterAlternate.exe")
'            If (hmmsg <> -1) Then
'                Print #hmmsg, "             CsiNetReporterAlternate V:" & ilRet / 10
'            End If
'
'            ilRet = mCheckFonts()
'        Else
'            If (hmmsg <> -1) Then
'                Print #hmmsg, "Unable to check Components because non standard windows fonts folder " & smWinDir & "\Fonts"
'            End If
'        End If
        ilRet = mCheckSystem32()
        If ilRet = 0 Then
            Print #hmmsg, "          Certain products were not installed properly."
        End If
        'Dan M 8/30/2010 get csiNetReporter file version
        ilRet = mGetVersionInfo(smWinDir & "\system32\csiNetReporter.exe")
        If (hmmsg <> -1) Then
            Print #hmmsg, "             CsiNetReporter V:" & ilRet / 10
        End If
        ilRet = mGetVersionInfo(smWinDir & "\system32\csiNetReporterAlternate.exe")
        If (hmmsg <> -1) Then
            Print #hmmsg, "             CsiNetReporterAlternate V:" & ilRet / 10
        End If
        'Dan M moved testing fonts to down here so don't skip above.
        ilRet = 0
        lbcFontNames(0).Path = smWinDir & "\Fonts"
        If ilRet = 0 Then
            ilRet = mCheckFonts()
        Else
            If (hmmsg <> -1) Then
                Print #hmmsg, "Unable to check fonts because non standard windows fonts folder " & smWinDir & "\Fonts"
            End If
        End If
    Else
        If (hmmsg <> -1) Then
            Print #hmmsg, "Unable to check Components because non standard windows system folder " & smWinDir
        End If
    End If
    tmcMoveLine.Enabled = False
    If (hmmsg <> -1) Then
        Close #hmmsg
    End If
    If (imShowErrMsg) And (imErrorMsg) And (hmmsg <> -1) Then
        If InStr(1, smAppName, "Affiliat", vbTextCompare) > 0 Then
            MsgBox "Please call Counterpoint and ask them to review, Affiliate System will Start when you press OK " & slFileName, vbOKOnly, "Check"
        ElseIf InStr(1, smAppName, "Engineer", vbTextCompare) > 0 Then
            MsgBox "Please call Counterpoint and ask them to review, Engineering System will Start when you press OK " & slFileName, vbOKOnly, "Check"
        Else
            MsgBox "Please call Counterpoint and ask them to review, Traffic System will Start when you press OK " & slFileName, vbOKOnly, "Check"
        End If
    End If
    'Else
        If Not imTerminate Then
            If InStr(1, smAppName, "Affiliat", vbTextCompare) > 0 Then
                llRet = gShellAndWait(CSIStart, smExePath & "\Affiliat2.exe " & smCurParamters, vbNormal, False)
                If llRet <> 0 Then
                    llRet = gShellAndWait(CSIStart, smExePath & "\Affiliat.exe " & smCurParamters, vbNormal, False)
                End If
            ElseIf InStr(1, smAppName, "Engineer", vbTextCompare) > 0 Then
                llRet = gShellAndWait(CSIStart, smExePath & "\Engineer2.exe " & smCurParamters, vbNormal, False)
                If llRet <> 0 Then
                    llRet = gShellAndWait(CSIStart, smExePath & "\Engineer.exe " & smCurParamters, vbNormal, False)
                End If
            Else
                llRet = gShellAndWait(CSIStart, smExePath & "\Traffic2.exe " & smCurParamters, vbNormal, False)
                If llRet <> 0 Then
                    llRet = gShellAndWait(CSIStart, smExePath & "\Traffic.exe " & smCurParamters, vbNormal, False)
                End If
            End If
        End If
    'End If
    mTerminate
    Screen.MousePointer = vbDefault
    Exit Sub
mTMCStartErr:
    llRet = Err.Number
    ilRet = Err.Number
    Resume Next
End Sub
Private Function mGetVersionInfo(FullFileName As String) As Integer
    Dim llRc As Long
    Dim llDummy As Long
    Dim slBuffer() As Byte
    Dim llBufferLen As Long
    Dim llVerPointer As Long
    Dim tlVerBuffer As VS_FIXEDFILEINFO
    Dim llVerbufferLen As Long
    Dim slFileVersion As String
    
    llBufferLen = GetFileVersionInfoSize(FullFileName, llDummy)
    If llBufferLen < 1 Then
       mGetVersionInfo = 0
       Exit Function
    End If
    ReDim slBuffer(llBufferLen)
    llRc = GetFileVersionInfo(FullFileName, 0&, llBufferLen, slBuffer(0))
    llRc = VerQueryValue(slBuffer(0), "\", llVerPointer, llVerbufferLen)
    MoveMemory tlVerBuffer, llVerPointer, Len(tlVerBuffer)
    'gets all parts of version number
    'slFileVersion = Format$(tlVerBuffer.dwFileVersionMSh) & "." & Format$(tlVerBuffer.dwFileVersionMSl) & "." & Format$(tlVerBuffer.dwFileVersionLSh) & "." & Format$(tlVerBuffer.dwFileVersionLSl)
    slFileVersion = Format$(tlVerBuffer.dwFileVersionMSh) & Format$(tlVerBuffer.dwFileVersionMSl)
    ' mGetVersionInfo = mConvertVersion(slFileVersion)
    mGetVersionInfo = CInt(slFileVersion)
End Function



Private Sub mCheckDDF()
'    Dim ilRet As Integer
'    Dim slNewDDFDateTime As String
'    Dim slNewDDFDate As String
'    Dim slDataDateTime As String
'    Dim slDataDate As String
'    Dim ilLoop As Integer
'
'    ilRet = 0
'    lacProgressFolder.Caption = ""
'    lacProgreFileListBox.Caption = ""
'    lacProgress(3).Caption = "Checking DDF Status"
'    DoEvents
'    lbcDatabases.Clear
'    On Error GoTo mCheckDDFErr:
'    slNewDDFDateTime = FileDateTime(smResultFilePath & "NewDDF\Field.ddf")
'    slNewDDFDate = gAdjYear(Format$(slNewDDFDateTime, "m/d/yy"))
'    If ilRet = 0 Then
'        lbcDDFCheckDir.Path = left$(smResultFilePath, Len(smResultFilePath) - 1)
'        For ilLoop = 0 To lbcDDFCheckDir.ListCount - 1 Step 1
'            ilRet = 0
'            On Error GoTo mCheckDDFErr:
'            slDataDateTime = FileDateTime(lbcDDFCheckDir.List(ilLoop) & "\" & "Field.ddf")
'            If ilRet = 0 Then
'                slDataDate = gAdjYear(Format$(slDataDateTime, "m/d/yy"))
'                If DateValue(slNewDDFDate) <> DateValue(slDataDate) Then
'                    lbcDatabases.AddItem lbcDDFCheckDir.List(ilLoop)
'                End If
'            End If
'        Next ilLoop
'    End If
'    Exit Sub
'mCheckDDFErr:
'    ilRet = Err
'    Resume Next

End Sub

Private Function mRunDDFReorg() As Integer
'    Dim ilLoop As Integer
'    Dim slExePath As String
'    Dim ilPos As Integer
'    Dim ilCount As Integer
'    Dim ilCountTotal As Integer
'    Dim llPercent As Long
'    Dim ilRet As Integer
'    Dim slName As String
'    Dim llLen As Long
'    Dim ilRunDDFOffst As Integer
'
'    imProcessing = True
'    Screen.MousePointer = vbHourglass
'    lncProgress.X2 = lncProgress.X1
'    lncProgress.Visible = False
'    DoEvents
'    ilCountTotal = 0
'    For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
'        If lbcDatabases.Selected(ilLoop) Then
'            ilCountTotal = ilCountTotal + 1
'        End If
'    Next ilLoop
'    'Determine if DDFOffst should be run
'    ilRet = 0
'    On Error GoTo mRunDDFReorgErr:
'    llLen = FileLen(smResultFilePath & "Exe\" & "csi_io32.dll")
'    If ilRet <> 0 Then
'        llLen = 800000
'    End If
'    'Two different csi_io32 routines.  One that uses Classic cbtrv432 and the other that
'    'Jeff wrote and does not use cbtrv432.  Jeff does not require the DDFOffst file and
'    'is about 900000.  The other is about 90000.
'    If llLen > 200000 Then
'        ilRunDDFOffst = False
'    Else
'        ilRunDDFOffst = True
'    End If
'
'    ilCount = 0
'    For ilLoop = 0 To lbcDatabases.ListCount - 1 Step 1
'        If lbcDatabases.Selected(ilLoop) Then
'            lacProgressFolder.Caption = lbcDatabases.List(ilLoop)
'            lacProgreFileListBox.Caption = "DDFReorg"
'            DoEvents
'            On Error GoTo mRunDDFReorgErr:
'            ilRet = 0
'            slName = lbcDatabases.List(ilLoop) & "\Messages"
'            MkDir slName
'            If (ilRet <> 0) And (ilRet <> 75) Then
'                lacProgressFolder.Caption = ""
'                lacProgreFileListBox.Caption = ""
'                imProcessing = False
'                Screen.MousePointer = vbDefault
'                Print #hmMsg, "Unable to create " & slName & " Error # " & Str$(ilRet)
'                MsgBox "Unable to create " & slName & " Error # " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
'                mRunDDFReorg = False
'                Exit Function
'            End If
'            ilRet = mMoveZeroBTR(smResultFilePath & "Updates", lbcDatabases.List(ilLoop))
'            slExePath = smResultFilePath & "Exe\DDFReorg /N:" & smResultFilePath & "NewDDF" & " /O:" & lbcDatabases.List(ilLoop) & " /Exit /Run"
'            gShellAndWait CSIStart, slExePath, vbMinimizedFocus 'vbNormalFocus    'vbMinimizedFocus 'vbHide
'            'Move the zero_btr files into database folder
'            If imCancelled Then
'                Screen.MousePointer = vbDefault
'                plcProgress.Visible = False
'                plcCancel.Visible = True
'                cmcNext.Enabled = True
'                imCancelAns = 0
'                Do
'                    DoEvents
'                Loop While imCancelAns = 0
'                plcProgress.Visible = True
'                plcCancel.Visible = False
'                cmcNext.Enabled = False
'                If imCancelAns = 2 Then
'                    lacProgressFolder.Caption = ""
'                    lacProgreFileListBox.Caption = ""
'                    imProcessing = False
'                    Screen.MousePointer = vbDefault
'                    Print #hmMsg, ""
'                    Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'                    Close #hmMsg
'                    mRunDDFReorg = False
'                    Exit Function
'                Else
'                    Screen.MousePointer = vbHourglass
'                    imCancelled = False
'                End If
'            End If
'            ilRet = mMoveZeroBTR(smResultFilePath & "Zero_BTR", lbcDatabases.List(ilLoop))
'            If ilRunDDFOffst Then
'                lacProgreFileListBox.Caption = "DDFOffst"
'                DoEvents
'                slExePath = smResultFilePath & "Exe\DDFOffst " & " /F:" & lbcDatabases.List(ilLoop) & " /Exit /Run"
'                gShellAndWait CSIStart, slExePath, vbMinimizedFocus 'vbNormalFocus    'vbMinimizedFocus 'vbHide
'                'Move the zero_btr files into database folder
'                If imCancelled Then
'                    Screen.MousePointer = vbDefault
'                    plcProgress.Visible = False
'                    plcCancel.Visible = True
'                    cmcNext.Enabled = True
'                    imCancelAns = 0
'                    Do
'                        DoEvents
'                    Loop While imCancelAns = 0
'                    plcProgress.Visible = True
'                    plcCancel.Visible = False
'                    cmcNext.Enabled = False
'                    If imCancelAns = 2 Then
'                        lacProgressFolder.Caption = ""
'                        lacProgreFileListBox.Caption = ""
'                        imProcessing = False
'                        Screen.MousePointer = vbDefault
'                        Print #hmMsg, ""
'                        Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'                        Close #hmMsg
'                        mRunDDFReorg = False
'                        Exit Function
'                    Else
'                        Screen.MousePointer = vbHourglass
'                        imCancelled = False
'                    End If
'                End If
'            End If
'            ilCount = ilCount + 1
'            lacProgress(3).Caption = Str$(ilCount) & " of" & Str$(ilCountTotal)
'            llPercent = (CLng(ilCount) * 100) \ ilCountTotal
'            If llPercent >= 100 Then
'                llPercent = 100
'            End If
'            lncProgress.X2 = ((pbcProgress.Width - 2 * lncProgress.X1) * llPercent) / 100
'            If llPercent > 0 Then
'                lncProgress.Visible = True
'            End If
'            DoEvents
'        End If
'    Next ilLoop
'    lacProgressFolder.Caption = ""
'    lacProgreFileListBox.Caption = ""
'    imProcessing = False
'    Print #hmMsg, "CSI Databases Converted Successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'    Print #hmMsg, ""
'    Close #hmMsg
'    lacFinish(4).Caption = lacFinish(4).Caption & ", Converted" & Str$(ilCount) & " of" & Str$(ilCountTotal) & " databases"
'    mRunDDFReorg = True
'    Exit Function
'mRunDDFReorgErr:
'    ilRet = Err
'    Resume Next
End Function

Public Function mMoveZeroBTR(slFromSubfolder As String, slToSubfolder As String) As Integer
'    Dim ilLoop As Integer
'    Dim slName As String
'    Dim ilRet As Integer
'    Dim ilOk As Integer
'    Dim ilTest As Integer
'
'    mMoveZeroBTR = True
'    lbcZeroBTR.Path = slFromSubfolder
'    For ilLoop = 0 To lbcZeroBTR.ListCount - 1 Step 1
'        slName = lbcZeroBTR.List(ilLoop)
'        ilOk = False
'        If InStr(1, slName, "r.btr", vbTextCompare) = 3 Then
'            ilOk = True
'        Else
'            For ilTest = 1 To UBound(smZeroNames) Step 1
'                If InStr(1, slName, smZeroNames(ilTest), vbTextCompare) = 1 Then
'                    ilOk = True
'                    Exit For
'                End If
'            Next ilTest
'        End If
'        If ilOk Then
'            slName = slFromSubfolder & "\" & lbcZeroBTR.List(ilLoop)
'            ilRet = 0
'            On Error GoTo mMoveZeroBTRErr:
'            FileCopy slName, slToSubfolder & "\" & lbcZeroBTR.List(ilLoop)
'            If ilRet <> 0 Then
'                Print #hmMsg, "Copying " & Trim$(slName) & " To " & Trim$(slToSubfolder) & " Error #" & Str$(ilRet)
'                mMoveZeroBTR = False
'            End If
'        End If
'    Next ilLoop
'    Exit Function
'mMoveZeroBTRErr:
'    ilRet = Err
'    Resume Next
    
End Function
Private Function mIsUserChecker(slFileName As String) As Boolean
    Dim slOldName As String
    Dim blRet As Boolean
    
 On Error GoTo ERRBOX
    blRet = True
    If Not Dir(slFileName) > "" Then
        slOldName = smRootDrive & "csi\csicomponentChecker.txt"
        If Dir(slOldName) > "" Then
            FileCopy slOldName, slFileName
        Else
            blRet = False
        End If
    End If
    mIsUserChecker = blRet
    Exit Function
ERRBOX:
    mIsUserChecker = False
End Function

Private Function mCheckSystem32() As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilRet As Integer
    Dim llRet As Long
    Dim llCurLen As Long
    Dim llDestLen As Long
    Dim llTempLen As Long
    Dim slDateTime As String
    Dim ilCrystal10File As Integer
    Dim tlLocalMSI() As MSI
    Dim ilLocalMSI As Integer
    Dim ilLocalDLL As Integer
    Dim ilLocalBAT As Integer
'    Dim ilLocalEXE As Integer
    Dim tlSetupMSI() As MSI
    Dim ilSetupMSI As Integer
    Dim ilSetupDLL As Integer
    Dim ilSetupBAT As Integer
 '   Dim ilSetupEXE As Integer
    Dim slCurDir As String * MAX_PATH
    Dim slDestDir As String * MAX_PATH
    Dim slTempDir As String * MAX_PATH
    Dim slLogString As String
    Dim slLogPath As String
    Dim slLogCommand As String
    Dim slReinstallOption As String
    Dim slTotalString As String
    Dim slSetupPath As String
        'Dan M 7/13/11
    Dim blIsCitrix As Boolean
    Dim blBlockOtherMsi As Boolean
    Dim blClientForCitrix As Boolean
    Dim slClientRun As String
    Dim slFileName As String
    
    'Dan 12/05/11
    'Dan 8/27/2010 1)Write out more msi version to csiStart. 2)  Stop using gDecToInt to convert version number. 3) get rid of exe
    'Dan 6/14/2010 MsiCsiCrystalInstall now includes the vcRedist_x86.exe in it.  That means that "exe:v1" is no longer necessary as there are no executables.
    'Dan 10/19/2009  1 "MSI:V2" is now "MSICSIINSTALL:V2" each msi must be formatted this way:  start with MSI; include 'V:XX'; and the version can only contain one period: '2.9' is valid, '2.9.1' is not
    'if an msi is not written to clientInstall\system32, it will still be written out, but it will be checked every time.  Currently, a message will be written to csiStart that this is wrong, but this can be changed if as msi ever needs to be run every time.
    'csiStart will write if no log written(msi not launched,usually an error), if it is attempting an install or 'reconfigure'(already run, anything needs repair?), and finally whether it was successful.
   ' ilRet = mGetMSIandDLLVersion("c:\csi\csiComponentChecker.txt", tlLocalMSI, ilLocalDLL, ilLocalBAT, ilLocalEXE)
    slSetupPath = smSystemPath
    If Not mDirExists(slSetupPath) Then
        If hmmsg <> -1 Then
            Print #hmmsg, "          " & "Folder " & smSystemPath & " cannot be found"
        End If
        imErrorMsg = True
        mCheckSystem32 = False
        Exit Function
    End If
    slFileName = "csiComponentCheckerv" & CSIVERSION & ".txt"
    slSetupPath = slSetupPath & "\" & slFileName
    slFileName = smRootDrive & "csi\" & slFileName
    mIsUserChecker slFileName
    'slSetupPath = slSetupPath & "\csiComponentCheckerv" & CSIVERSION & ".txt"
    'ilRet = mGetMSIandDLLVersion("C:\csi\csiComponentCheckerv" & CSIVERSION & ".txt", tlLocalMSI, ilLocalDLL, ilLocalBAT)
    ilRet = mGetMSIandDLLVersion(slFileName, tlLocalMSI, ilLocalDLL, ilLocalBAT)
    If (tlLocalMSI(0).iVersion <> -1) Or (ilLocalDLL <> -1) Then
        'Dan M 8/27/2010 read if false..write and quit.
        'ilRet = mGetMSIandDLLVersion(smSystemPath & "\" & "csiComponentChecker.Txt", tlSetupMSI, ilSetupDLL, ilSetupBAT, ilSetupEXE)
        If Not mGetMSIandDLLVersion(slSetupPath, tlSetupMSI, ilSetupDLL, ilSetupBAT) Then
            If hmmsg <> -1 Then
                Print #hmmsg, "    " & slSetupPath & " doesn't exist."
            End If
            imErrorMsg = True
            mCheckSystem32 = False
            Exit Function
        End If
'        If (tlLocalMSI = tlSetupMSI) And (ilLocalDLL = ilSetupDLL) And (ilLocalBAT = ilSetupBAT) Then
       ' If (tlLocalMSI = tlSetupMSI) And (ilLocalDLL = ilSetupDLL) And (ilLocalBAT = ilSetupBAT) And (ilLocalEXE = ilSetupEXE) Then
       ' If (mCompareMsi(tlLocalMSI, tlSetupMSI)) And (ilLocalDLL = ilSetupDLL) And (ilLocalBAT = ilSetupBAT) And (ilLocalEXE = ilSetupEXE) Then
        If (mCompareMsi(tlLocalMSI, tlSetupMSI)) And (ilLocalDLL = ilSetupDLL) And (ilLocalBAT = ilSetupBAT) Then
            mCheckSystem32 = True
            If hmmsg <> -1 Then
                Print #hmmsg, "    " & " Updating of Msis DLLs and OCXs not needed"
                Dim c As Integer
                For c = 0 To UBound(tlSetupMSI) - 1 Step 1
                    Print #hmmsg, "             " & tlSetupMSI(c).sName & " V:" & tlSetupMSI(c).iVersion / 10
                Next c
            End If
            Exit Function
        End If
    Else
        ReDim tlSetupMSI(0)
        tlSetupMSI(0).iVersion = 0
        tlSetupMSI(0).sName = " "
        ilSetupDLL = 0
        ilSetupBAT = 0
        'Dan removed 8/27/10
        'ilSetupEXE = 0
    End If
    
    lacWMessage(1).Caption = "Confirming Client Components . . ."
    DoEvents
    If hmmsg <> -1 Then
        Print #hmmsg, "     Checking Msis, DLLs and OCXs"
    End If
    
    On Error GoTo mCheckSystem32Err:
    
'    llWinLen = GetWindowsDirectory(slWindowsDir, MAX_PATH)
'    smWinDir = left$(slWindowsDir, llWinLen)
    llRet = 0
    ilRet = 0
    lbcFileNames.Path = smSystemPath
    If (llRet <> 0) Then
    'Dan M 8/27/10 tested already that folder exists.
       ' If hmmsg <> -1 Then
        '    Print #hmmsg, "          " & "Folder " & smSystemPath & " is missing, Error #" & Str$(llRet)
       ' End If
        imErrorMsg = True
        mCheckSystem32 = False
        Exit Function
    End If
    If (lbcFileNames.ListCount <= 0) Then
        If hmmsg <> -1 Then
            Print #hmmsg, "          " & "Folder " & smSystemPath & " is empty"
        End If
        imErrorMsg = True
        mCheckSystem32 = False
        Exit Function
    End If
    'Scan file names for crqe.  If found, then place all crystal modules into csi\crystal10
    imUseCrystal10 = False
    For ilLoop = 0 To lbcFileNames.ListCount - 1 Step 1
        DoEvents
        slName = lbcFileNames.List(ilLoop)
        If InStr(1, slName, "crqe", vbTextCompare) = 1 Then
            imUseCrystal10 = True
            Exit For
        End If
    Next ilLoop
    If imUseCrystal10 Then
        'Check that folder exist
        mCreateCrystal10
    End If
    '7/13/11 dan m so far, only citrix has failed this test. C:\Documents and Settings\kwatts\WINDOWS
    ' on Salem, msiexec is started, but then never runs msi and causes the program to stall.  I'm not sure why this is happening, as our citrix box doesn't
    'have this issue.  Also, if the msi doesn't exist, a message saying so is displayed.  In any event, to avoid stalling csiStart, citrix will get a work around
    ' that will watch how long it is waiting for the csiClientInstall msi to run.  If it's taking too long, it will ask the user if they want to quit.
    If (InStr(1, smWinDir, ":\winnt", vbTextCompare) = 0) And (InStr(1, smWinDir, ":\windows", vbTextCompare) = 0) Then
        blIsCitrix = True
    End If
    For ilLoop = 0 To lbcFileNames.ListCount - 1 Step 1
        DoEvents
        slName = lbcFileNames.List(ilLoop)
        If (InStr(1, slName, ".msi", vbTextCompare) > 1) Then
            'Dan added 8/27/10 changed printing info, so added blret and moved things a bit.
          '  If hmmsg <> -1 Then
           '     Print #hmmsg, "            Testing MSI file " & slName
          '  End If
'            blRet = mCompareMsiVersions(slName, tlLocalMSI, tlSetupMSI, ilSetupMSI)
'            If hmmsg <> -1 Then
'                Print #hmmsg, "             " & slName & "'s Version #:" & ilSetupMSI & " is current? " & blRet
'            End If
            If Not mCompareMsiVersions(slName, tlLocalMSI, tlSetupMSI, ilSetupMSI) Then
                If ilSetupMSI = 0 And hmmsg <> -1 Then
                'msi doesn't exist in csiComponentChecker...if all msi's are rogue, then csiComponentChecker.txt doesn't exist!
                    Print #hmmsg, "                     " & slName & " is a rogue msi. Please let a CSI service person know that csiComponentChecker needs to be updated."
                End If
            'If ilLocalMSI <> ilSetupMSI Then
            ' Dan M 5/07/09 write log csiIntallLog.  Read status of install and write to #hmmsg ("csiStart").  CsiInstallLog then destroyed.
                'slLogCommand = " /Lmew "
                ' slLogCommand = " /Li "
                '5676 forget this log
'                If gDriveExists("C") Then
'                    mCreateCSI
'                    slLogPath = "C:\csi\CsiInstallLog.txt"
'                Else
'                    slLogPath = App.Path & "\CsiInstallLog.txt"
'                End If
'                slLogString = slLogCommand & slLogPath
                slLogString = ""
                'Dan 6/14/10 vcredist_x86.exe no longer needed
                ' If mVcredistAsNeeded(ilLocalEXE, ilSetupEXE, slName) Then
                'dan M 09-03-09 if same version, reinstall automatically
                   '  llRet = gShellAndWait(CSIStart, "msiExec /i " & smSystemPath & "\" & slName & slLogString, vbNormal, True)
                '5676
                If smRootDrive = "C:\" Then
                    slReinstallOption = "  /qb! "
                Else
                    slReinstallOption = ""
                End If
                'slReinstallOption = "  /qb! " 'where qb is quiet mode with progress bar, ! is no cancel button
                'dan M 05-18-11 add quotes
                'dan M 10-10-11 if clientInstall, add path to exefolder
                ' dan 10-26-12 don't need this anymore
'                If InStr(1, slName, "ClientInstall", vbTextCompare) > 0 Then
'                    slTotalString = "msiExec /i " & """" & smSystemPath & "\" & slName & """" & " EXEFOLDER=""" & smExePath & """" & slReinstallOption & slLogString
'                Else
'                    slTotalString = "msiExec /i " & """" & smSystemPath & "\" & slName & """" & slReinstallOption & slLogString
'                End If
                slTotalString = "msiExec /i " & """" & smSystemPath & "\" & slName & """" & slReinstallOption & slLogString
                If blIsCitrix Then
                    If InStr(1, slName, "ClientInstall", vbTextCompare) > 0 Then
                        blClientForCitrix = True
                        'test only
                       ' If MsgBox("running shell and wait on " & slName & ".  Return true?", vbYesNo) = vbNo Then
                       'Dan M 6/14/12 misspelling: had "misExec"
                        If Not gShellAndWaitAwhile(Me, slTotalString, "msiExec", vbNormal, 5, True) Then
                            MsgBox "Please contact Counterpoint service to help with installation of needed modules.", vbInformation, "Installation needed"
                            'to block writing of csiComponentChecker
                            imErrorMsg = True
                            blBlockOtherMsi = True
                        End If
                    'client was ok, let's try crystal, if needed:
                    ElseIf Not blBlockOtherMsi Then
                        If blClientForCitrix Then
                            'MsgBox "run " & slName
                            llRet = gShellAndWait(CSIStart, slTotalString, vbNormal, True)
                        'client wasn't run...but we need to run it first to make sure program won't stall.
                        Else
                            If InStr(1, slName, "CrystalInstall", vbTextCompare) > 0 Then
                                slClientRun = Replace(slTotalString, "CrystalInstall", "ClientInstall", , , vbTextCompare)
                                'If MsgBox("running shell and wait because client not run on " & slName & ".  Return true?", vbYesNo) = vbYes Then
                                If gShellAndWaitAwhile(Me, slClientRun, "msiExec", vbNormal, 1, True) Then
                                   ' MsgBox "run " & slName
                                    llRet = gShellAndWait(CSIStart, slTotalString, vbNormal, True)
                                Else
                                    MsgBox "Please contact Counterpoint service to help with installation of needed modules.", vbInformation, "Installation needed"
                                    imErrorMsg = True
                                End If
                            'in case we add another msi, don't have to run clientInstall first.
                            Else
                               ' MsgBox "run " & slName & ". but shouldn't be here"
                                llRet = gShellAndWait(CSIStart, slTotalString, vbNormal, True)
                            End If
                        End If
                    End If
                Else
                    llRet = gShellAndWait(CSIStart, slTotalString, vbNormal, True)
                End If
                '5676
'                ilRet = mTempLogToRealLog(slLogPath)
'                ' log: can't find, couldn't find outcome of install, install good, install error. If first, assume this means error installing. Second, assume success
'                If hmmsg <> -1 And ilRet <> STATUSGOOD Then
'                    If ilRet = STATUSNOLOG Then  'no msi log to read, or log says error
'                        Print #hmmsg, "                    " & slName & " MsiLog not written. This usually means the msi was not installed."
'                        imErrorMsg = True
'                    ElseIf ilRet = STATUSLOGUNKNOWN Then
'                        Print #hmmsg, "                    " & slName & "MsiLog not written. This usually means the msi was not installed."
'                    ElseIf ilRet = STATUSERROR Then
'                        imErrorMsg = True
'                    End If
'                End If
            '  Else
              'vcredist needed and doesn't exit!
             '     If hmmsg <> -1 Then
              '        Print #hmmsg, "Missing Vcredist_x86.exe.  Installation halted."
               '       imErrorMsg = True
                '      mCheckSystem32 = 44
                 '     Exit Function
                 ' End If
             ' End If
            ElseIf hmmsg <> -1 Then
                ilRet = mGetVersionInfo(smSystemPath & "\" & slName)
                Print #hmmsg, "                     " & Mid(slName, 1, Len(slName) - 3) & " Product Version: " & ilSetupMSI / 10 & ".  Local and Setup versions match.  No need to reinstall."
            End If
        ElseIf (InStr(1, slName, ".bat", vbTextCompare) > 1) Then
            If ilLocalBAT <> ilSetupBAT Then
                llRet = gShellAndWait(CSIStart, smSystemPath & "\" & slName & " " & smSystemPath, vbNormal, True)
            End If
        ' get rid of exe
       ' ElseIf (InStr(1, slName, ".exe", vbTextCompare) > 1) Then
        Else
            If ilLocalDLL <> ilSetupDLL Then
                ilCrystal10File = False
                If imUseCrystal10 Then
                    If (InStr(1, slName, "U", vbTextCompare) = 1) Or (InStr(1, slName, "Cr", vbTextCompare) = 1) Or (InStr(1, slName, "Cr", vbTextCompare) = 1) Or (InStr(1, slName, "P2", vbTextCompare) = 1) Or (InStr(1, slName, "Cx", vbTextCompare) = 1) Or (InStr(1, slName, "ExportM", vbTextCompare) = 1) Or (InStr(1, slName, "QueryB", vbTextCompare) = 1) Then
                        ilCrystal10File = True
                    End If
                End If
                If ilCrystal10File Then
                    llRet = 0
                    On Error GoTo mCheckSystem32Err:
                    slDateTime = FileDateTime(smCrystal10Path & "\" & slName)
                    If llRet <> 0 Then
                        'File does not exist
                        llRet = 0
                    End If
                Else
                    llCurLen = MAX_PATH
                    llDestLen = MAX_PATH
                    llRet = VerFindFile(VFFF_ISSHAREDFILE, slName, smWinDir, smWinDir & "\system32", slCurDir, llCurLen, slDestDir, llDestLen)
                End If
                If llRet = VFF_FILEINUSE Then
                ElseIf llRet = VFF_BUFFTOOSMALL Then
                Else
                    DoEvents
                    If ilCrystal10File Then
                        FileCopy smSystemPath & "\" & slName, smCrystal10Path & "\" & slName
                    Else
                        llTempLen = MAX_PATH
                        'llRet = VerInstallFile(VIFF_DONTDELETEOLD, slName, slName, smSystemPath, slDestDir, slCurDir, slTempDir, llTempLen)
                        llRet = VerInstallFile(VIFF_DONTDELETEOLD, slName, slName, smSystemPath, slDestDir, slCurDir, slTempDir, llTempLen)
                    End If
                    If (llRet <> 0) And (Not ilCrystal10File) Then
                        If ((llRet And VIF_DIFFCODEPG) = VIF_DIFFCODEPG) Or ((llRet And VIF_DIFFLANG) = VIF_DIFFLANG) Or ((llRet And VIF_DIFFTYPE) = VIF_DIFFTYPE) Or ((llRet And VIF_MISMATCH) = VIF_MISMATCH) Or ((llRet And VIF_WRITEPROT) = VIF_WRITEPROT) Then
                            If (llRet And VIF_SRCOLD) = VIF_SRCOLD Then
                            Else
                                llRet = VerInstallFile(VIFF_FORCEINSTALL + VIFF_DONTDELETEOLD, slName, slName, smSystemPath, slDestDir, slCurDir, slTempDir, llTempLen)
                            End If
                        End If
                    End If
                    If llRet <> 0 Then
                        If ilCrystal10File Then
                            If hmmsg <> -1 Then
                                Print #hmmsg, "          " & "Copy File from " & smSystemPath & "\" & slName & "to " & smCrystal10Path & "\" & slName & ", error # " & Trim$(Str$(llRet))
                                imErrorMsg = True
                            End If
                        Else
                            If (llRet And VIF_SRCOLD) = VIF_SRCOLD Then
            '                    'The file to install is older than the preexisting file
            '                    If hmMsg <> -1 Then
            '                        Print #hmMsg, "          " & slName & ": The file to install is older than the preexisting file"
            '                    End If
                            Else
                                If (llRet And VIF_TEMPFILE) = VIF_TEMPFILE Then
                                    'The temporary copy of the new file is in the destination directory.  The cause of failure is reflected in other flags
                                End If
                                If (llRet And VIF_MISMATCH) = VIF_MISMATCH Then
                                    'The new and preexisting files differ in one or more attributes
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The new and preexisting files differ in one or more attributes"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_DIFFLANG) = VIF_DIFFLANG Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The new and preexisting files have different language or code-page values"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_DIFFCODEPG) = VIF_DIFFCODEPG Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The new file requires a code page that cannot be displayed by the version of the system currently running"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_DIFFTYPE) = VIF_DIFFTYPE Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The new file has a different type, subtype or operating system from the preexisting file"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_WRITEPROT) = VIF_WRITEPROT Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The preexisting file is write-protected"
                                    End If
                                End If
                                If (llRet And VIF_FILEINUSE) = VIF_FILEINUSE Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The preexisting file is in use by the system and cannot be deleted"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_OUTOFSPACE) = VIF_OUTOFSPACE Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot complete the requested operation due to insufficient disk space on the destination drive"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_ACCESSVIOLATION) = VIF_ACCESSVIOLATION Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": A read, create, delete or rename operation failed due to an access violation"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_SHARINGVIOLATION) = VIF_SHARINGVIOLATION Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": A read, create, delete or rename operation failed due to a sharing violation"
                                    End If
                                End If
                                If (llRet And VIF_CANNOTCREATE) = VIF_CANNOTCREATE Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot create the temporary file"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_CANNOTDELETE) = VIF_CANNOTDELETE Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot delete the destination file, or cannot delete the existing version of the file located in another directory"
                                    End If
                                End If
                                If (llRet And VIF_CANNOTRENAME) = VIF_CANNOTRENAME Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot rename the temporary file, but already deleted the destination file"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_CANNOTDELETECUR) = VIF_CANNOTDELETECUR Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The existing version of the file could not be deleted and VIFF_DONTDELETEOLD was not specified"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_OUTOFMEMORY) = VIF_OUTOFMEMORY Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot complete the requested operation due to insufficient memory"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_CANNOTREADSRC) = VIF_CANNOTREADSRC Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot read the source file"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_CANNOTREADDST) = VIF_CANNOTREADDST Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The function cannot read the destination (existing) files"
                                        imErrorMsg = True
                                    End If
                                End If
                                If (llRet And VIF_BUFFTOOSMALL) = VIF_BUFFTOOSMALL Then
                                    If hmmsg <> -1 Then
                                        Print #hmmsg, "          " & slName & ": The szTmpFile buffer was too small to contain the name of the temporary source file"
                                        imErrorMsg = True
                                    End If
                                End If
                            End If
                        End If
                    Else
                        DoEvents
                        llRet = 0
                        If StrComp(slName, "Comctl32.ocx", vbTextCompare) = 0 Then
                            llRet = RegComctl32()
                        ElseIf StrComp(slName, "Comdlg32.ocx", vbTextCompare) = 0 Then
                            llRet = RegComdlg32()
                        ElseIf StrComp(slName, "Craxdrt.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = RegCraxdrt()
                            End If
                        ElseIf StrComp(slName, "Crtslv.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = RegCrtslv()
                            End If
                        ElseIf StrComp(slName, "Crviewer.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = Regrviewer()
                            End If
                        ElseIf StrComp(slName, "crqe.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = RegCrqe()
                            End If
                        ElseIf StrComp(slName, "pdfsplitmerge.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = Regpdfsplitmerge()
                            End If
                        ElseIf StrComp(slName, "Duzactx.dll", vbTextCompare) = 0 Then
                            llRet = RegDuzactx()
                        ElseIf StrComp(slName, "Dzactx.dll", vbTextCompare) = 0 Then
                            llRet = RegDzactx()
                        ElseIf StrComp(slName, "ExportModeller.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = RegExportModeller()
                            End If
                        ElseIf StrComp(slName, "Mci32.ocx", vbTextCompare) = 0 Then
                            llRet = RegMci32()
                        ElseIf StrComp(slName, "Mfc40.dll", vbTextCompare) = 0 Then
                            llRet = RegMfc40()
                        ElseIf StrComp(slName, "Mfc42.dll", vbTextCompare) = 0 Then
                            llRet = RegMfc42()
                        ElseIf StrComp(slName, "Mshflxgd.ocx", vbTextCompare) = 0 Then
                            llRet = RegMshflxgd()
                        ElseIf StrComp(slName, "Msmask32.ocx", vbTextCompare) = 0 Then
                            llRet = RegMsmask32()
                        ElseIf StrComp(slName, "MsComCtl.ocx", vbTextCompare) = 0 Then
                            llRet = RegMsComCtl()
                        ElseIf StrComp(slName, "Msrdo20.dll", vbTextCompare) = 0 Then
                            llRet = RegMsrdo20()
                        ElseIf StrComp(slName, "Msvcrt.dll", vbTextCompare) = 0 Then
                            llRet = RegMsvcrt()
                        ElseIf StrComp(slName, "MsComm32.ocx", vbTextCompare) = 0 Then
                            llRet = RegMsComm32()
                        ElseIf StrComp(slName, "Oleaut32.dll", vbTextCompare) = 0 Then
                            llRet = RegOleaut32()
                        ElseIf StrComp(slName, "Olepro32.dll", vbTextCompare) = 0 Then
                            llRet = RegOlepro32()
                        ElseIf StrComp(slName, "Resize32.ocx", vbTextCompare) = 0 Then
                            llRet = RegResize32()
                        ElseIf StrComp(slName, "Ssdw3b32.ocx", vbTextCompare) = 0 Then
                        'Not used
                        '    llRet = RegSsdw3b32()
                        ElseIf StrComp(slName, "Tabctl32.ocx", vbTextCompare) = 0 Then
                        'Not used
                        '    llRet = Regtabctl32()
                        ElseIf StrComp(slName, "cruflCSI.dll", vbTextCompare) = 0 Then
                            If imUseCrystal10 Then
                                llRet = gShellAndWait(CSIStart, "regsvr32 /s " & smCrystal10Path & "\" & slName, vbNormal, True)
                            Else
                                llRet = RegCruflcsi()
                            End If
                        ElseIf (InStr(1, slName, "U2", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "U3", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "Crxf_", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "P2", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "Cr", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "Cx", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "QueryB", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        ElseIf (InStr(1, slName, "Uf", vbTextCompare) = 1) And (InStr(1, slName, ".dll", vbTextCompare) > 1) Then
                            'Don't register
                        Else
                            llRet = gShellAndWait(CSIStart, "regsvr32 /s " & left$(slDestDir, llDestLen - 1) & "\" & slName, vbNormal, True)
                        End If
                        If llRet <> 0 Then
                            If hmmsg <> -1 Then
                                Print #hmmsg, "          " & slName & ": Unable to register, error #" & Str$(llRet) & " file copied from " & smSystemPath & " to " & smWinDir & "\system32"
                                If llRet <> 453 Then
                                    imErrorMsg = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next ilLoop
    If Not imErrorMsg Then
       ' FileCopy smSystemPath & "\" & "csiComponentChecker.txt", "c:\csi\csiComponentChecker.txt"
        FileCopy slSetupPath, slFileName
        mCheckSystem32 = True
    End If
    Exit Function
mCheckSystem32Err:
    llRet = Err.Number
    Resume Next
End Function
Private Function mCompareMsi(tlLocal() As MSI, tlSetup() As MSI) As Boolean
    Dim c As Integer
    Dim j As Integer
    Dim ilLocalBound As Integer
    Dim ilSetupBound As Integer
    Dim blFound As Boolean
    
    ilLocalBound = UBound(tlLocal)
    ilSetupBound = UBound(tlSetup)
    If (ilSetupBound <> ilLocalBound) Or (ilLocalBound = 0) Or (ilSetupBound = 0) Then
        mCompareMsi = False
        Exit Function
    End If
    For c = 0 To ilSetupBound - 1
        For j = 0 To ilLocalBound - 1
            If StrComp(tlSetup(c).sName, tlLocal(j).sName, vbTextCompare) = 0 Then
                If tlSetup(c).iVersion <= tlLocal(j).iVersion Then
                    blFound = True
                End If
                Exit For
            End If
        Next j
        If Not blFound Then
            mCompareMsi = False
            Exit Function
        End If
        blFound = False
    Next c
    mCompareMsi = True
End Function
Private Function mCompareMsiVersions(slName As String, tlLocal() As MSI, tlSetup() As MSI, Optional ByRef ilSetup As Integer) As Boolean
    Dim c As Integer
    Dim ilLocal As Integer
    For c = 0 To UBound(tlSetup) - 1
        If StrComp(slName, tlSetup(c).sName & ".MSI", vbTextCompare) = 0 Then
            ilSetup = tlSetup(c).iVersion
            Exit For
        End If
    Next c
    For c = 0 To UBound(tlLocal) - 1
        If StrComp(slName, tlLocal(c).sName & ".MSI", vbTextCompare) = 0 Then
            ilLocal = tlLocal(c).iVersion
            Exit For
        End If
    Next c
    If ilSetup <> 0 And ilSetup <= ilLocal Then
        mCompareMsiVersions = True
    End If
End Function
Private Function mTempLogToRealLog(slLogPath As String) As Integer
    Dim hlLog As Integer
    Dim slStringBuffer As String
    Dim slErrorString As String
    Dim ilRet As Integer
    Dim ilAfterFindNextLine As Integer

On Error GoTo errorbox
    mTempLogToRealLog = STATUSLOGUNKNOWN
    hlLog = FreeFile
    Open slLogPath For Input As #hlLog
    Do While Not EOF(hlLog)
        Line Input #hlLog, slStringBuffer
        If ilAfterFindNextLine > 0 Then 'read second line after found line for more status info.
            If ilAfterFindNextLine = 1 Then
                ilAfterFindNextLine = 2
            ElseIf ilAfterFindNextLine = 2 Then
                ilRet = mTestLogLine(slStringBuffer, True)
                If ilRet > 0 Then
                    'mLogStatusLine slStringBuffer, ilRet
                    mPrintStatusToLog slStringBuffer
                End If
                Exit Do
            End If
        Else
            ilRet = mTestLogLine(slStringBuffer, False)
            If ilRet > 0 Then   'got the right line
                mPrintLogLine slStringBuffer, ilRet + 2
                ilAfterFindNextLine = 1
                mTempLogToRealLog = mLogLineSaysSuccess(slStringBuffer, ilRet)
            End If
        End If
    Loop
finish:
    Close #hlLog
    Kill slLogPath
    Exit Function
errorbox:
    mTempLogToRealLog = STATUSNOLOG
End Function
Private Function mTestLogLine(slLogLine As String, blSecondStatusLine As Boolean) As Integer
    Dim slTestLine As String
    If blSecondStatusLine Then
        slTestLine = "success or error status:"
    Else
        slTestLine = "]: Product: "
    End If
    mTestLogLine = InStr(1, slLogLine, slTestLine, vbTextCompare)
End Function
Private Function mLogStatusLine(slLogLine As String, ilLocation) As Integer
' this is set up to read the status number, but currently isn't used: if the first log line says successful, that should be good enough.
    Dim slStatus As String
    
    If ilLocation > 0 Then
        ilLocation = InStr(ilLocation, slLogLine, ":", vbTextCompare)
        If ilLocation > 0 Then ' ready to read status
            slStatus = Trim(Mid(slLogLine, ilLocation + 1, Len(slLogLine) - ilLocation - 1))
            If IsNumeric(slStatus) Then
                mLogStatusLine = CInt(slStatus)
            End If
            mPrintStatusToLog slLogLine
        End If
    End If

End Function

Private Function mLogLineSaysSuccess(slLogLine As String, ilLocation) As Integer
    If ilLocation > 0 Then
        ilLocation = InStr(ilLocation, slLogLine, "successfully", vbTextCompare)
        If ilLocation > 0 Then ' ready to read status
            mLogLineSaysSuccess = STATUSGOOD
        Else
            mLogLineSaysSuccess = STATUSERROR
        End If
    End If

End Function

Private Sub mPrintLogLine(slStatString As String, ilLocation As Integer)
    If hmmsg <> -1 Then
        If ilLocation > 0 Then
            Print #hmmsg, "                    " & Mid(slStatString, ilLocation)
        End If
    End If

End Sub
Private Sub mPrintStatusToLog(slStatusString As String)
    Dim slWriteLine As String
    Dim ilLocation As Integer
    
    If hmmsg <> -1 Then
        ilLocation = InStr(1, slStatusString, "Product Version", vbTextCompare)
        If ilLocation > 0 Then
            slWriteLine = Mid(slStatusString, ilLocation)
            slWriteLine = Replace(slWriteLine, "Product Language: 1033.", " ")
            Print #hmmsg, "                          " & slWriteLine
        End If
    End If
        
End Sub

Private Function mCheckFonts() As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilRet As Integer
    Dim llRet As Long
    Dim ilFont As Integer
    Dim ilTest As Integer
    Dim ilGetFont As Integer
    Dim ilAddResource As Integer
    Dim slFontName As String
    Dim hlKey As Long
    
    lacWMessage(1).Caption = "Confirming Client Components . . ."
    DoEvents
    If hmmsg <> -1 Then
        Print #hmmsg, "     Checking Fonts"
    End If
    On Error GoTo mCheckFontsErr:
    ilRet = 0
    llRet = 0
    lbcFontNames(0).Path = smWinDir & "\Fonts"
    If (llRet <> 0) Then
        If hmmsg <> -1 Then
            Print #hmmsg, "          " & "Folder " & smWinDir & "\Fonts" & " is missing, Error #" & Str$(llRet)
        End If
        imErrorMsg = True
        mCheckFonts = False
        Exit Function
    End If
    lbcFontNames(1).Path = smFontPath
    If (llRet <> 0) Then
        If hmmsg <> -1 Then
            Print #hmmsg, "          " & "Folder " & smFontPath & " is missing, Error #" & Str$(llRet)
        End If
        imErrorMsg = True
        mCheckFonts = False
        Exit Function
    End If
    For ilLoop = 0 To lbcFontNames(1).ListCount - 1 Step 1
        ilTest = True
        slName = lbcFontNames(1).List(ilLoop)
        If StrComp(slName, "arialn.ttf", vbTextCompare) = 0 Then
            slFontName = "Arial Narrow"
        ElseIf StrComp(slName, "mtsorts.ttf", vbTextCompare) = 0 Then
            slFontName = "Monotype Sorts"
        ' Dan M 2/2/10 Century not used and causing traffic to hang...issue: schlbk.ttf is not bold
'        ElseIf StrComp(slName, "schlbk.ttf", vbTextCompare) = 0 Then
'            slFontName = "Century Schoolbook Bold"
'        ElseIf StrComp(slName, "schlbkbi.ttf", vbTextCompare) = 0 Then
'            slFontName = "Century Schoolbook Bold Italic"
        ElseIf StrComp(slName, "mistral.ttf", vbTextCompare) = 0 Then
            slFontName = "Mistral"
        Else
            ilTest = False
        End If
        If ilTest Then
            ilAddResource = True
            For ilFont = 0 To Screen.FontCount - 1 Step 1
                ilRet = StrComp(slFontName, Screen.Fonts(ilFont), vbTextCompare)
                If ilRet = 0 Then
                    ilAddResource = False
                    Exit For
                End If
            Next ilFont
            ilGetFont = True
            For ilFont = 0 To lbcFontNames(0).ListCount - 1 Step 1
                If StrComp(slName, lbcFontNames(0).List(ilFont), vbTextCompare) = 0 Then
                    ilGetFont = False
                    Exit For
                End If
            Next ilFont
            If ilGetFont Then
                'Copy font
                ilRet = 0
                FileCopy smFontPath & "\" & slName, smWinDir & "\Fonts\" & slName
                If (ilRet <> 0) And (hmmsg <> -1) Then
                    Print #hmmsg, "          " & "Unable to Copy " & smFontPath & "\" & slName & " To " & smWinDir & "\Fonts\" & slName & " Error #" & Str$(ilRet)
                    imErrorMsg = True
                End If
            End If
            If (ilGetFont) Or (ilAddResource) Then
                'Add font to resource
                llRet = AddFontResource(smWinDir & "\Fonts\" & slName)
                If llRet = 0 Then
                    llRet = GetLastError()
                    If (llRet <> 0) And (hmmsg <> -1) Then
                        Print #hmmsg, "          " & "Unable to Add Font "; slName & " to Resource table." & " Error #" & Str$(llRet)
                        imErrorMsg = True
                    End If
                Else
                    llRet = RegOpenKey(HKEY_LOCAL_MACHINE, "software\microsoft\windows\currentversion\fonts", hlKey)
                    If (llRet <> 0) And (hmmsg <> -1) Then
                        llRet = RegOpenKey(HKEY_LOCAL_MACHINE, "software\microsoft\windows nt\currentversion\fonts", hlKey)
                    End If
                    If llRet = 0 Then
                        llRet = RegSetValueEx(hlKey, slFontName & " (TrueType)", 0, REG_SZ, slName, Len(slName))
                        If (llRet <> 0) And (hmmsg <> -1) Then
                            Print #hmmsg, "          " & "Unable to Register " & slFontName & " as " & slName & ", Error #" & Str$(llRet)
                        End If
                        llRet = RegCloseKey(hlKey)
                    End If
                    llRet = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
                    DoEvents
                End If
            End If
        End If
    Next ilLoop
    mCheckFonts = True
    Exit Function
mCheckFontsErr:
    llRet = Err.Number
    ilRet = Err.Number
    Resume Next

End Function


Private Sub mCreateCSI()
    Dim ilRet As Integer
    
    ilRet = 0
    On Error GoTo mCreateCSIErr:
    MkDir "C:\CSI"
    Exit Sub
mCreateCSIErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub mCreateCrystal10()
    Dim ilRet As Integer
    
    ilRet = 0
    On Error GoTo mCreateCrystal10Err:
    MkDir smCrystal10Path
    Exit Sub
mCreateCrystal10Err:
    ilRet = 1
    Resume Next
End Sub

'Private Function mGetMSIandDLLVersion(slPathFile As String, tlVersionMSI As Integer, ilVersionDLL As Integer, ilVersionBAT As Integer, ilVersionEXE As Integer) As Integer
'    Dim hlFrom As Integer
'    Dim ilRet As Integer
'    Dim ilEof As Integer
'    Dim slLine As String
'    Dim ilPos As Integer
'    Dim slVersion As String
'
'    On Error GoTo mGetMSIandDLLVersionErr:
'    tlVersionMSI = -1
'    ilVersionDLL = -1
'    ilVersionBAT = -1
'    'dan added EXE file 7/06/09
'    ilVersionEXE = -1
'    ilRet = 0
'    hlFrom = FreeFile
'    Open slPathFile For Input Access Read Shared As hlFrom
'    If ilRet = 0 Then
'        Do
'            ilRet = 0
'            On Error GoTo mGetMSIandDLLVersionErr:
'            Line Input #hlFrom, slLine
'            On Error GoTo 0
'            If ilRet = 62 Then
'                Exit Do
'            End If
'            If Len(slLine) > 0 Then
'                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                    ilEof = True
'                Else
'                    ilPos = InStr(1, slLine, "MSI:", vbTextCompare)
'                    If ilPos = 1 Then
'                        slVersion = Trim$(Mid$(slLine, 5))
'                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
'                        If ilPos > 0 Then
'                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
'                        End If
'                        tlVersionMSI = Val(slVersion)
'                    End If
'                    ilPos = InStr(1, slLine, "DLL:", vbTextCompare)
'                    If ilPos = 1 Then
'                        slVersion = Trim$(Mid$(slLine, 5))
'                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
'                        If ilPos > 0 Then
'                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
'                        End If
'                        ilVersionDLL = Val(slVersion)
'                    End If
'                    ilPos = InStr(1, slLine, "BAT:", vbTextCompare)
'                    If ilPos = 1 Then
'                        slVersion = Trim$(Mid$(slLine, 5))
'                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
'                        If ilPos > 0 Then
'                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
'                        End If
'                        ilVersionBAT = Val(slVersion)
'                    End If
'                    ilpost = InStr(1, slLine, "EXE:", vbTextCompare)
'                    If ilPos = 1 Then
'                        slVersion = Trim$(Mid$(slLine, 5))
'                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
'                        If ilPos > 0 Then
'                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
'                        End If
'                        ilVersionEXE = Val(slVersion)
'                    End If
'                End If
'            End If
'        Loop Until ilEof
'        mGetMSIandDLLVersion = True
'    Else
'        mGetMSIandDLLVersion = False
'    End If
'    Close hlFrom
'    Exit Function
'mGetMSIandDLLVersionErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
'Private Function mGetMSIandDLLVersion(slPathFile As String, ByRef tlVersionMSI() As MSI, ByRef ilVersionDLL As Integer, ByRef ilVersionBAT As Integer, ByRef ilVersionEXE As Integer) As Integer
'
'    ' Dan M 7/07/09 read decimals
'    Dim hlFrom As Integer
'    Dim ilRet As Integer
'    Dim ilEof As Integer
'    Dim slLine As String
'    Dim ilPos As Integer
'    Dim slVersion As String
'    Dim ilUBound As Integer
'    On Error GoTo mGetMSIandDLLVersionErr:
'    ReDim tlVersionMSI(0)
'    tlVersionMSI(0).iVersion = -1
'    ilVersionDLL = -1
'    ilVersionBAT = -1
'    'dan added EXE file 7/06/09
'    ilVersionEXE = -1
'    ilRet = 0
'    hlFrom = FreeFile
'    Open slPathFile For Input Access Read Shared As hlFrom
'    If ilRet = 0 Then
'        Do
'            ilRet = 0
'            On Error GoTo mGetMSIandDLLVersionErr:
'            Line Input #hlFrom, slLine
'            On Error GoTo 0
'            If ilRet = 62 Then
'                Exit Do
'            End If
'            If Len(slLine) > 0 Then
'                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                    ilEof = True
'                Else
'                    ilPos = InStr(1, slLine, "MSI", vbTextCompare)
'                    If ilPos = 1 Then
'                        ilUBound = UBound(tlVersionMSI)
'                        ReDim Preserve tlVersionMSI(ilUBound + 1)
'                        ilRet = InStr(1, slLine, ":", vbTextCompare)
'                        tlVersionMSI(ilUBound).sName = Mid(slLine, 4, ilRet - 4)
'                        tlVersionMSI(ilUBound).iVersion = mConvertVersion(slLine)
''                        slVersion = Trim$(Mid$(slLine, 5))
''                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
''                        If ilPos > 0 Then
''                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
''                        End If
''                        tlVersionMSI = Val(slVersion)
'                    End If
'                    ilPos = InStr(1, slLine, "DLL:", vbTextCompare)
'                    If ilPos = 1 Then
'                        ilVersionDLL = mConvertVersion(slLine)
''                        slVersion = Trim$(Mid$(slLine, 5))
''                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
''                        If ilPos > 0 Then
''                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
''                        End If
''                        ilVersionDLL = Val(slVersion)
'                    End If
'                    ilPos = InStr(1, slLine, "BAT:", vbTextCompare)
'                    If ilPos = 1 Then
'                        ilVersionBAT = mConvertVersion(slLine)
''                        slVersion = Trim$(Mid$(slLine, 5))
''                        ilPos = InStr(1, slVersion, "V", vbTextCompare)
''                        If ilPos > 0 Then
''                            slVersion = Trim$(Mid$(slVersion, ilPos + 1))
''                        End If
''                        ilVersionBAT = Val(slVersion)
'                    End If
'                    ilPos = InStr(1, slLine, "EXE:", vbTextCompare)
'                    If ilPos = 1 Then
'                        ilVersionEXE = mConvertVersion(slLine)
'                    End If
'                End If
'            End If
'        Loop Until ilEof
'        mGetMSIandDLLVersion = True
'    Else
'        mGetMSIandDLLVersion = False
'    End If
'    Close hlFrom
'    Exit Function
'mGetMSIandDLLVersionErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
Private Function mGetMSIandDLLVersion(slPathFile As String, ByRef tlVersionMSI() As MSI, ByRef ilVersionDLL As Integer, ByRef ilVersionBAT As Integer) As Integer
' Dan M 8/27/10 lose exe, change reading of version number
    
    ' Dan M 7/07/09 read decimals
    Dim hlFrom As Integer
    Dim ilRet As Integer
    Dim ilEof As Integer
    Dim slLine As String
    Dim ilPos As Integer
    Dim slVersion As String
    Dim ilUBound As Integer
    
    On Error GoTo mGetMSIandDLLVersionErr:
    ReDim tlVersionMSI(0)
    tlVersionMSI(0).iVersion = -1
    ilVersionDLL = -1
    ilVersionBAT = -1
    ilRet = 0
    hlFrom = FreeFile
    Open slPathFile For Input Access Read Shared As hlFrom
    If ilRet = 0 Then
        Do
            ilRet = 0
            On Error GoTo mGetMSIandDLLVersionErr:
            Line Input #hlFrom, slLine
            On Error GoTo 0
            If ilRet = 62 Then
                Exit Do
            End If
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    ilEof = True
                Else
                    If InStr(1, slLine, "MSI", vbTextCompare) = 1 Then
                        ilUBound = UBound(tlVersionMSI)
                        ReDim Preserve tlVersionMSI(ilUBound + 1)
                        ilRet = InStr(1, slLine, ":", vbTextCompare)
                        tlVersionMSI(ilUBound).sName = Mid(slLine, 4, ilRet - 4)
                        tlVersionMSI(ilUBound).iVersion = mConvertVersion(slLine)
                    ElseIf InStr(1, slLine, "DLL:", vbTextCompare) = 1 Then
                        ilVersionDLL = mConvertVersion(slLine)
                    ElseIf InStr(1, slLine, "BAT:", vbTextCompare) = 1 Then
                        ilVersionBAT = mConvertVersion(slLine)
                    End If
                End If
            End If
        Loop Until ilEof
        mGetMSIandDLLVersion = True
    Else
        mGetMSIandDLLVersion = False
    End If
    Close hlFrom
    Exit Function
mGetMSIandDLLVersionErr:
    ilRet = Err.Number
    Resume Next
End Function
'Dan M replaced 8/27/10. Don't want to treat as decimal don't want to round up.
'Private Function mConvertVersion(slLine As String) As Integer
'Dim slVersion As String
'Dim ilPos As Integer
'    slVersion = Trim$(Mid$(slLine, 5))
'    ilPos = InStr(1, slVersion, "V", vbTextCompare)
'    If ilPos > 0 Then
'        slVersion = Trim$(Mid$(slVersion, ilPos + 1))
'    End If
'    mConvertVersion = gStrDecToInt(slVersion, 2)  'Val(slVersion) * 100
'
'End Function
Private Function mConvertVersion(ByVal slLine As String) As Integer
'12.02.20.2 becomes 1202 4.1 becomes 41
    Dim ilFirstDot As Integer
    Dim ilSecondDot As Integer
    Dim ilPos As Integer
    
    ilPos = InStrRev(slLine, "V", , vbTextCompare)
    If ilPos > 0 Then
        slLine = Trim$(Mid$(slLine, ilPos + 1))
    End If
    ilFirstDot = InStr(1, slLine, ".")
    If ilFirstDot > 0 Then
       If ilFirstDot <> InStrRev(slLine, ".") Then
          'crystal version or net platform. trim to '12.2'
          ilSecondDot = InStr(ilFirstDot + 1, slLine, ".")
          slLine = Mid(slLine, 1, ilSecondDot - 1)
       End If
       slLine = Replace(slLine, ".", "")
    End If
    If IsNumeric(slLine) Then
       mConvertVersion = CInt(slLine)
    Else
       mConvertVersion = 0
    End If
 End Function
Private Function mDirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    mDirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
'Dan 6/14/10 don't need vcredist_x86 anymore
'Private Function mVcredistAsNeeded(ilLocalEXE As Integer, ilSetupEXE As Integer, slName As String) As Boolean
'Dim llRet As Long
'Dim slExePath As String
'If StrComp(slName, "MSICSICRYSTALINSTALL.MSI", vbTextCompare) <> 0 Then
'    mVcredistAsNeeded = True    'vcredist not needed for this msi
'    Exit Function
'End If
'slExePath = smSystemPath & "\vcredist_x86.exe"
'On Error GoTo errorbox
'    mVcredistAsNeeded = True
'    If ilLocalEXE <> ilSetupEXE Then
'        If Dir(slExePath) > " " Then
'           llRet = gShellAndWait(CSIStart, slExePath & " /q:a", vbNormal, True)
'        Else
'            mVcredistAsNeeded = False
'        End If
'    End If
'    Exit Function
'errorbox:
'    mVcredistAsNeeded = False
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:gStrDecToInt                    *
''*                                                     *
''*             Created:5/6/93        By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Convert a String with decimal   *
''*                     to integer (round up)           *
''*                                                     *
''*******************************************************
'Function gStrDecToInt(slInNumber As String, ilNoDecPlaces As Integer) As Integer
''
''   ilOut = gStrDecToInt(slInNumber, ilNoDecPlaces)
''   Where:
''       slInNumber(I)- Number to be converted to long
''       ilNoDecPlaces(I)-Number of decimal places (can be zero)
''       ilOut(O)- Output long
''                 slInNumber    # Dec Places   llOut
''                 234.56        2               23456
''                 234.56        0               235
''                 234.56        1               2346
''
''
'    Dim ilPos As Integer
'    Dim slStr As String
'    Dim slLeft As String
'    Dim slRight As String
'    Dim ilAddOne As Integer
'    Dim llNum As Long
'    Dim ilLoop As Integer
'
'    slStr = Trim$(slInNumber)
'    ilPos = InStr(slStr, ".")
'    If ilPos = 0 Then
'        slStr = slStr & "."
'    End If
'    ilPos = InStr(slStr, ".")
'    Do While Len(slStr) - ilPos < ilNoDecPlaces
'        slStr = slStr & "0"
'    Loop
'    ilPos = InStr(slStr, ".")
'    'Check Rounding
'    ilAddOne = False
'    If (Len(slStr) - ilPos) > ilNoDecPlaces Then
'        slRight = right$(slStr, Len(slStr) - ilPos - ilNoDecPlaces)
'        llNum = 5
'        For ilLoop = 2 To Len(slRight) Step 1
'            llNum = 10 * llNum
'        Next ilLoop
'        If Val(slRight) >= llNum Then
'            ilAddOne = True
'        End If
'        slStr = left(slStr, ilPos + ilNoDecPlaces)
'    End If
'    ilPos = InStr(slStr, ".")
'    If ilPos > 0 Then
'        slRight = Mid$(slStr, ilPos + 1)
'        slLeft = left$(slStr, ilPos - 1)
'        slStr = slLeft & slRight
'    End If
'    If Val(slStr) > 32766! Then
'        gStrDecToInt = 0
'    Else
'        If ilAddOne Then
'            gStrDecToInt = Val(slStr) + 1
'        Else
'            gStrDecToInt = Val(slStr)
'        End If
'    End If
'    Exit Function
'End Function
