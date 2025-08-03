VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "Dzactx.dll"
Object = "{0FB90DC1-97D1-11D1-87C0-444553540000}#1.0#0"; "Duzactx.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DBUnzip 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6930
   ClientLeft      =   915
   ClientTop       =   3225
   ClientWidth     =   9495
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6930
   ScaleWidth      =   9495
   Begin VB.FileListBox lbcZeroBTR 
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
      Left            =   405
      MultiSelect     =   2  'Extended
      Pattern         =   "*.btr"
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6555
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox lbcResult 
      Appearance      =   0  'Flat
      Height          =   1080
      ItemData        =   "DBUNZIP.frx":0000
      Left            =   105
      List            =   "DBUNZIP.frx":0002
      TabIndex        =   29
      Top             =   4695
      Width           =   8985
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   195
      Left            =   2205
      TabIndex        =   27
      Top             =   5880
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmcCreate 
      Appearance      =   0  'Flat
      Caption         =   "Create &Folder"
      Height          =   285
      Left            =   7275
      TabIndex        =   24
      Top             =   6690
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox plcResult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1755
      Left            =   105
      ScaleHeight     =   1695
      ScaleWidth      =   9225
      TabIndex        =   9
      Top             =   2865
      Width           =   9285
      Begin VB.Frame frcDDFReorg 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         TabIndex        =   13
         Top             =   1455
         Width           =   3630
         Begin VB.OptionButton rbcRunDDFReorg 
            Caption         =   "Yes"
            Height          =   210
            Index           =   0
            Left            =   2205
            TabIndex        =   15
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton rbcRunDDFReorg 
            Caption         =   "No"
            Height          =   210
            Index           =   1
            Left            =   2850
            TabIndex        =   16
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "Run DDFReorg if Required"
            Height          =   210
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   2205
         End
      End
      Begin VB.Frame frcZero 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   4215
         TabIndex        =   17
         Top             =   1455
         Width           =   4950
         Begin VB.OptionButton rbcZero 
            Caption         =   "None"
            Height          =   210
            Index           =   2
            Left            =   3945
            TabIndex        =   21
            Top             =   0
            Width           =   750
         End
         Begin VB.OptionButton rbcZero 
            Caption         =   "RNF && RTF Only"
            Height          =   210
            Index           =   1
            Left            =   2370
            TabIndex        =   20
            Top             =   0
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton rbcZero 
            Caption         =   "All"
            Height          =   210
            Index           =   0
            Left            =   1770
            TabIndex        =   19
            Top             =   0
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lacZero 
            Caption         =   "Copy Zero_Btr Files"
            Height          =   195
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.DriveListBox cbcResultDir 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   45
         Width           =   2550
      End
      Begin VB.DirListBox lbcResultDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   5430
         TabIndex        =   12
         Top             =   45
         Width           =   3720
      End
      Begin VB.Label lacResult 
         Appearance      =   0  'Flat
         Caption         =   "Unzip database into Subfolder:"
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
         Height          =   390
         Left            =   60
         TabIndex        =   10
         Top             =   75
         Width           =   3150
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   6600
   End
   Begin VB.PictureBox plcExe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   105
      ScaleHeight     =   2475
      ScaleWidth      =   9225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   9285
      Begin VB.FileListBox lbcModFile 
         Appearance      =   0  'Flat
         Height          =   1710
         Left            =   3315
         MultiSelect     =   2  'Extended
         Pattern         =   "*.Zip"
         TabIndex        =   8
         Top             =   540
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.PictureBox pbcLbcModList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   3330
         ScaleHeight     =   2100
         ScaleWidth      =   5565
         TabIndex        =   28
         Top             =   135
         Width           =   5565
      End
      Begin VB.TextBox edcPW 
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
         Left            =   2220
         MaxLength       =   5
         TabIndex        =   3
         Top             =   105
         Width           =   765
      End
      Begin VB.DriveListBox cbcMod 
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
         Left            =   90
         TabIndex        =   5
         Top             =   735
         Width           =   2895
      End
      Begin VB.DirListBox lbcMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   105
         TabIndex        =   6
         Top             =   1095
         Width           =   2865
      End
      Begin VB.ListBox lbcModList 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   3315
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   120
         Width           =   5835
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Unzip Database from Subfolder:"
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
         Left            =   105
         TabIndex        =   4
         Top             =   510
         Width           =   3045
      End
      Begin VB.Label lacPW 
         Appearance      =   0  'Flat
         Caption         =   "Zip Password (DBVxx)"
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
         Height          =   210
         Left            =   105
         TabIndex        =   2
         Top             =   135
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmcUnzip 
      Appearance      =   0  'Flat
      Caption         =   "&Unzip"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3165
      TabIndex        =   22
      Top             =   6495
      Width           =   1425
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4875
      TabIndex        =   23
      Top             =   6510
      Width           =   1425
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   2160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2160
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8745
      OleObjectBlob   =   "DBUNZIP.frx":0004
      Top             =   6390
   End
   Begin DUZACTXLibCtl.duzactxctrl zpcDUnZip 
      Left            =   7995
      OleObjectBlob   =   "DBUNZIP.frx":00FA
      Top             =   6450
   End
   Begin VB.Label lacResults 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3210
      TabIndex        =   26
      Top             =   0
      Width           =   5805
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2160
      TabIndex        =   25
      Top             =   6225
      Width           =   5130
   End
End
Attribute VB_Name = "DBUnzip"
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
Option Explicit
Option Compare Text
Dim hmMsg As Integer   'From file hanle
Dim hmFile As Integer           'file handle
Dim hmTo As Integer
Dim tmModName() As ZIPNAME
Dim tmUnZipName() As UNZIPNAME
Dim tmZippedNames() As ZIPPEDNAMES
Dim smZeroNames(1 To 6) As String
Dim smPattern As String
Dim smPassword As String
Dim smExport As String
Dim smImport As String
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim smNowDate As String   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim smModFilePath As String
Dim smResultFilePath As String
Dim smTodaysDate As String
Dim imListField(1 To 4) As Integer
Dim smMessagesPath As String

Public Function mCopyZeroBTR() As Integer
    Dim slSrcPathFileName As String
    Dim slDstPathFileName As String
    Dim slZeroBtr As String
    Dim ilPos As Integer

    On Error GoTo mErrHand:
    slZeroBtr = left$(smResultFilePath, Len(smResultFilePath) - 1)
    ilPos = InStrRev(slZeroBtr, "\", -1, vbTextCompare)
    slZeroBtr = left$(slZeroBtr, ilPos) & "zero_btr"
    
    ' Copy only the two report definition files. Zero_Btr files are being handled differently now.
    slSrcPathFileName = slZeroBtr & "\RTF.BTR"
    slDstPathFileName = smResultFilePath & "RTF.BTR"
    FileCopy slSrcPathFileName, slDstPathFileName
    
    slSrcPathFileName = slZeroBtr & "\RNF.BTR"
    slDstPathFileName = smResultFilePath & "RNF.BTR"
    FileCopy slSrcPathFileName, slDstPathFileName
    
    Exit Function
    
mErrHand:
    Resume Next
End Function

Private Function mCheckDDF() As Integer
    Dim ilRet As Integer
    Dim slNewDDFDateTime As String
    Dim slNewDDFDate As String
    Dim slDataDateTime As String
    Dim slDataDate As String
    Dim ilLoop As Integer
    Dim slCSIFolder As String
    Dim ilPos As Integer
    
    mCheckDDF = False
    ilRet = 0
    lacProcessing.Caption = "Checking DDF's"
    DoEvents
    ilPos = InStrRev(sgExePath, "\", Len(sgExePath) - 1)
    slCSIFolder = left$(sgExePath, ilPos)
    On Error GoTo mCheckDDFErr:
    slNewDDFDateTime = FileDateTime(slCSIFolder & "NewDDF\Field.ddf")
    slNewDDFDate = gAdjYear(Format$(slNewDDFDateTime, "m/d/yy"))
    If ilRet = 0 Then
        ilRet = 0
        On Error GoTo mCheckDDFErr:
        slDataDateTime = FileDateTime(smResultFilePath & "Field.ddf")
        If ilRet = 0 Then
            slDataDate = gAdjYear(Format$(slDataDateTime, "m/d/yy"))
            If DateValue(slNewDDFDate) <> DateValue(slDataDate) Then
                mCheckDDF = True
            End If
        End If
    End If
    Exit Function
mCheckDDFErr:
    ilRet = Err
    Resume Next

End Function


Private Function mRunDDFReorg() As Integer
    Dim ilLoop As Integer
    Dim slExePath As String
    Dim ilPos As Integer
    Dim ilCount As Integer
    Dim ilCountTotal As Integer
    Dim llPercent As Long
    Dim ilRet As Integer
    Dim slName As String
    Dim llLen As Long
    Dim ilRunDDFOffst As Integer
    Dim slCSIFolder As String
    
    imProcessing = True
    Screen.MousePointer = vbHourglass
    DoEvents
    'Determine if DDFOffst should be run
    ilRet = 0
    On Error GoTo mRunDDFReorgErr:
    llLen = FileLen(sgExePath & "\" & "csi_io32.dll")
    If ilRet <> 0 Then
        llLen = 800000
    End If
    'Two different csi_io32 routines.  One that uses Classic cbtrv432 and the other that
    'Jeff wrote and does not use cbtrv432.  Jeff does not require the DDFOffst file and
    'is about 900000.  The other is about 90000.
    If llLen > 200000 Then
        ilRunDDFOffst = False
    Else
        ilRunDDFOffst = True
    End If

    On Error GoTo mRunDDFReorgErr:
    ilRet = 0
    slName = smResultFilePath & "Messages"
    MkDir slName
    If (ilRet <> 0) And (ilRet <> 75) Then
        lacProcessing.Caption = ""
        imProcessing = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, "Unable to create " & slName & " Error # " & Str$(ilRet)
        MsgBox "Unable to create " & slName & " Error # " & Str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Create Error"
        mRunDDFReorg = False
        Exit Function
    End If
    lbcResult.AddItem "DDFReorg run successfully"
    DoEvents
    Print #hmMsg, "DDFReorg run successfully"
    ilPos = InStrRev(sgExePath, "\", Len(sgExePath) - 1)
    slCSIFolder = left$(sgExePath, ilPos)
    lacProcessing.Caption = "Running DDFReorg"
    DoEvents
'    ilRet = mMoveZeroBTR(smResultFilePath & "Updates", lbcDatabases.List(ilLoop))
    slExePath = sgExePath & "\DDFReorg /N:" & slCSIFolder & "NewDDF" & " /O:" & left$(smResultFilePath, Len(smResultFilePath) - 1) & " /Exit /Run"
    gShellAndWait DBUnzip, slExePath, vbMinimizedFocus 'vbNormalFocus    'vbMinimizedFocus 'vbHide
    'Move the zero_btr files into database folder
    If imCancelled Then
        Screen.MousePointer = vbDefault
        lacProcessing.Caption = ""
        imProcessing = False
        Screen.MousePointer = vbDefault
        Print #hmMsg, ""
        Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        mRunDDFReorg = False
        Exit Function
    End If
'    ilRet = mMoveZeroBTR(smResultFilePath & "Zero_BTR", lbcDatabases.List(ilLoop))
    If ilRunDDFOffst Then
        lacProcessing.Caption = "Running DDFOffst"
        DoEvents
        slExePath = sgExePath & "\DDFOffst " & " /F:" & left$(smResultFilePath, Len(smResultFilePath) - 1) & " /Exit /Run"
        gShellAndWait DBUnzip, slExePath, vbMinimizedFocus 'vbNormalFocus    'vbMinimizedFocus 'vbHide
        'Move the zero_btr files into database folder
        If imCancelled Then
            Screen.MousePointer = vbDefault
            lacProcessing.Caption = ""
            imProcessing = False
            Screen.MousePointer = vbDefault
            Print #hmMsg, ""
            Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
            mRunDDFReorg = False
            Exit Function
        End If
        lbcResult.AddItem "DDFOffst run successfully"
        Print #hmMsg, "DDFOffst run successfully"
        DoEvents
    End If
    lacProcessing.Caption = ""
    DoEvents
    Print #hmMsg, "DDFReorg completed successfully " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    mRunDDFReorg = True
    Exit Function
mRunDDFReorgErr:
    ilRet = Err
    Resume Next
End Function


Private Sub cbcMod_Change()
    lbcMod.Path = cbcMod.Drive
End Sub
Private Sub cbcResultDir_Change()
    lbcResultDir.Path = cbcResultDir.Drive
End Sub
Private Function mGetIniInfo() As Integer

'    Dim ilRet As Integer
'    Dim ilFound As Integer
'    Dim slFileName As String
'    Dim slReturn As String * 130
'    Dim slLocation As String
'    Dim ilPos As Integer
'    Dim slData As String
'    Dim slDataPath As String
'
'    slLocation = "Locations"
'    On Error GoTo gObtainIniValuesErr
'    slFileName = CurDir$ & "\Traffic.Ini"
'    ilRet = GetPrivateProfileString("Locations", "Database", "Not Found", slReturn, 128, slFileName)
'    If left$(slReturn, ilRet) = "Not Found" Then
'        MsgBox "Database is not properly defined under [Locations] in the Traffic.ini file."
'        mGetIniInfo = False
'        Exit Function
'    End If
'    slDataPath = left$(slReturn, ilRet)
'    slDataPath = slDataPath & "\"
'    smMessagesPath = slDataPath & "Messages\"
'    mGetIniInfo = True
'    Exit Function
'
'gObtainIniValuesErr:
'    ilFound = False
'    Resume Next

End Function



Private Sub cmcCancel_Click()
    If imProcessing Then
        imCancelled = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    'gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCreate_Click()
    CSIDir.Show vbModal
End Sub
Private Sub cmcUnzip_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slName As String
    Dim slToName As String
    Dim slToFile As String
    Dim llPercent As Long
    Dim slMsgFile As String
    Dim slError As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slChar As String * 1
    Dim slPrevChar As String * 1
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilPos As Integer
    Dim slOldName As String
    Dim llCount As Long
    Dim llNoRead As Long
    Dim slFileExt As String
    Dim ilChg As Integer
    Dim ilAnyChg As Integer
    Dim ilAddFile As Integer
    Dim llProcessed As Long
    smPassword = UCase$(edcPW.Text)
    If left$(smPassword, 1) = "" Then
        ilRet = MsgBox("Password not defined, Ok to continue?", vbYesNo + vbQuestion, "Password")
        If ilRet = vbNo Then
            Exit Sub
        End If
    Else
        If InStr(1, smPassword, "DBV", vbTextCompare) <= 0 Then
            ilRet = MsgBox("Password form not correct, Ok to continue anyway?", vbYesNo + vbQuestion, "Password")
            If ilRet = vbNo Then
                Exit Sub
            End If
        End If
    End If
    smTodaysDate = Format$(Now, "m/d/yy")
    smResultFilePath = lbcResultDir.Path & "\"
    sgDBPath = smResultFilePath
    Screen.MousePointer = vbHourglass
    imCancelled = False
    ilRet = 0
    On Error GoTo cmcChkErr:
    slDateTime = FileDateTime(slMsgFile)
    If Not mOpenMsgFile(slMsgFile) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lacResults.Caption = "Results Stored into: " & slMsgFile
    On Error GoTo 0
    Print #hmMsg, "CSI Unzipping Database Files " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    
    llCount = 0
    For ilLoop = 0 To lbcModList.ListCount - 1 Step 1
        If lbcModList.Selected(ilLoop) Then
            llCount = llCount + 1
        End If
    Next ilLoop
    lbcResult.Clear
    imProcessing = True
    plcGauge.Value = 0
    llProcessed = 0
    ilAddFile = False
    slPrevChar = ""
    ReDim tmZippedNames(0 To 0) As ZIPPEDNAMES
    For ilLoop = 0 To lbcModList.ListCount - 1 Step 1
        If lbcModList.Selected(ilLoop) Then
            DoEvents
            slName = tmModName(ilLoop).sName
            ilPos = InStr(slName, ".")
            If ilPos > 0 Then
                lacProcessing.Caption = slName
                ilRet = mUnZipFiles(slName)
                If imCancelled Then
                    Print #hmMsg, ""
                    Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                    Close #hmMsg
                    lacProcessing.Caption = "Processing- Cancelled"
                    imProcessing = False
                    cmcCancel.Caption = "&Done"
                    'cmcUnzip.Enabled = False
                    cmcCancel.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If ilRet = False Then
                    Print #hmMsg, ""
                    Print #hmMsg, "Unzip Error, Operation terminated " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                    Close #hmMsg
                    lacProcessing.Caption = "Processing- Terminated"
                    imProcessing = False
                    cmcCancel.Caption = "&Done"
                    'cmcUnzip.Enabled = False
                    cmcCancel.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                llProcessed = llProcessed + 1
                llPercent = (llProcessed * 100) \ llCount
                If llPercent >= 100 Then
                    llPercent = 100
                End If
                plcGauge.Value = llPercent
            End If
        End If
    Next ilLoop
    Print #hmMsg, ""
    If rbcRunDDFReorg(0).Value Then
        If mCheckDDF() Then
            ilRet = mRunDDFReorg()
        Else
            lbcResult.AddItem "DDFReorg not required to be run"
            Print #hmMsg, "DDFReorg not required to be run"
        End If
    Else
        lbcResult.AddItem "DDFReorg specified not to be run"
        Print #hmMsg, "DDFReorg specified not to be run"
    End If
    DoEvents
    If (rbcZero(0).Value) Or (rbcZero(1).Value) Then
        ilRet = mCopyZeroBTR()
    Else
        lbcResult.AddItem "Zero_BTR Files Not Copied"
        Print #hmMsg, "Zero_BTR Files Not Copied"
    End If
    DoEvents
    Print #hmMsg, "CSI Unzip Completed " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    
    Close #hmMsg
    plcGauge.Value = 100
    lacProcessing.Caption = "Processing- Done"
    imProcessing = False
    cmcCancel.Caption = "&Done"
    'cmcUnzip.Enabled = False
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
cmcChkErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcPW_Change()
    tmcStart.Enabled = False
    tmcStart.Enabled = True
End Sub
Private Sub edcPW_GotFocus()
    tmcStart.Enabled = False
End Sub
Private Sub edcPW_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
End Sub
Private Sub Form_Load()
    Dim ilPos As Integer
    sgExePath = App.Path
    ilPos = InStr(1, sgExePath, "\Src", vbTextCompare)
    If ilPos > 0 Then
        sgExePath = left$(sgExePath, ilPos) & "exe"
    End If
    mInit
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    'Reset used instead of Close to cause # Clients on network to be decrement
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
    End
End Sub
Private Sub lbcMod_Change()
    Dim ilLoop As Integer
    tmcStart.Enabled = False
    smModFilePath = lbcMod.Path & "\"
    Screen.MousePointer = vbHourglass
    lbcModFile.Path = lbcMod.Path
    tmcStart.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
Private Sub lbcMod_GotFocus()
    tmcStart.Enabled = False
End Sub

Private Sub lbcModList_Click()
    Dim ilLoop As Integer
    pbcLbcModList_Paint
    cmcUnzip.Enabled = False
    For ilLoop = 0 To lbcModList.ListCount - 1 Step 1
        If lbcModList.Selected(ilLoop) Then
            cmcUnzip.Enabled = True
            Exit For
        End If
    Next ilLoop
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
    Dim llValue As Long
    Dim llRg As Long
    Dim llRet As Long
    
    imTerminate = False
    imCancelled = False
    imProcessing = False
    imSetAll = True
    imAllClicked = False
    Screen.MousePointer = vbHourglass
    
    fgBoxGridH = 180      'Height of grid area (distance from bottom of form letter to bottom of form box)
    fgBoxInsetX = 30    'X Margin from box outline to text
    fgListHTArial825 = 210  '225  'Height of list item (arial 8.25 bold; DivideStyle=1-Black line; ListStyle=0-2D)
    fgListHtSerif825 = 195  '255  'Height of list item (san serif 8.25; Divide style=3=Raised; ListStyle=1-3D)
    igAlignCharWidth = 90
    igScrollBarWidth = 255
    
    imListField(1) = 15
    imListField(2) = 40 * igAlignCharWidth 'igAvgCharWidth * 5
    imListField(3) = 50 * igAlignCharWidth 'igAvgCharWidth * 12
    imListField(4) = 80 * igAlignCharWidth 'igAvgCharWidth * 12

    smZeroNames(1) = "cbf.btr"
    smZeroNames(2) = "grf.btr"
    smZeroNames(3) = "odf.btr"
    smZeroNames(4) = "rnf.btr"
    smZeroNames(5) = "rtf.btr"
    smZeroNames(6) = "swf.btr"

'    llValue = 1000
'    llRg = 0
'    llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    

    'ilRet = mGetIniInfo
    mParseCmmdLine
    'edcPW.Text = "DBV"
    DBUnzip.Height = cmcUnzip.top + 5 * cmcUnzip.Height / 3
    'gCenterStdAlone BtrMatch
    DBUnzip.Move (Screen.Width - DBUnzip.Width) \ 2, (Screen.Height - DBUnzip.Height) \ 2 ' + 175 '+ Screen.Height \ 10
    DBUnzip.Show
    
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    smNowDate = Format$(Now, "m/d/yy")
    'smPattern = "*.E*;*.R*;*.A*;*.B*;*.N*;*.Z*;*.G*"
    smPattern = "*.Zip"
    lbcModFile.Pattern = smPattern
    
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    mTerminate
    Exit Sub
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
Private Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim fs As New FileSystemObject
    Dim slFilePath As String
    
    
    slFilePath = smResultFilePath & "Messages\"
    If fs.FolderExists(slFilePath) Then
        slToFile = slFilePath
    Else
        slToFile = smResultFilePath
    End If
    
    On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpStnFd.Txt"
    'slToFile = smResultFilePath & "DBUnzip.Txt"
    'slToFile = smMessagesPath & "DBUnzip.Txt"
    slToFile = slToFile & "DBUnzip.Txt"
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        slFileDate = Format$(slDateTime, "m/d/yy")
        If DateValue(slFileDate) = DateValue(smTodaysDate) Then  'Append
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmMsg = FreeFile
            Open slToFile For Append As hmMsg
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            On Error GoTo mOpenMsgFileErr:
            hmMsg = FreeFile
            Open slToFile For Output As hmMsg
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim slChar As String
    slCommand = Command$
    smPassword = ""
    ilPos = InStr(1, slCommand, "/PW:", 1)
    If ilPos > 0 Then
        ilPos = ilPos + 4
        slChar = Trim$(Mid$(slCommand, ilPos, 1))
        Do While slChar = " "
            ilPos = ilPos + 1
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        Do While slChar <> " "
            smPassword = smPassword & slChar
            ilPos = ilPos + 1
            If ilPos > Len(slCommand) Then
                Exit Do
            End If
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
    End If
    smPassword = UCase$(Trim$(smPassword))
    edcPW.Text = smPassword
    smExport = ""
    ilPos = InStr(1, slCommand, "/E:", 1)
    If ilPos > 0 Then
        ilPos = ilPos + 3
        slChar = Trim$(Mid$(slCommand, ilPos, 1))
        Do While slChar = " "
            ilPos = ilPos + 1
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        Do While slChar <> " "
            smExport = smExport & slChar
            ilPos = ilPos + 1
            If ilPos > Len(slCommand) Then
                Exit Do
            End If
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        cbcResultDir.Drive = left$(smExport, 1)
        lbcResultDir.Path = smExport
    End If
    smImport = ""
    ilPos = InStr(1, slCommand, "/I:", 1)
    If ilPos > 0 Then
        ilPos = ilPos + 3
        slChar = Trim$(Mid$(slCommand, ilPos, 1))
        Do While slChar = " "
            ilPos = ilPos + 1
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        Do While slChar <> " "
            smImport = smImport & slChar
            ilPos = ilPos + 1
            If ilPos > Len(slCommand) Then
                Exit Do
            End If
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        cbcMod.Drive = left$(smImport, 1)
        lbcMod.Path = smImport
    End If
End Sub
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
    Erase tmModName
    Erase tmUnZipName
    Erase tmZippedNames
    Screen.MousePointer = vbDefault
    Unload DBUnzip
    Set DBUnzip = Nothing
    End
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mUnZipFiles                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unzip files                     *
'*                                                     *
'*******************************************************
Private Function mUnZipFiles(slFromFile As String) As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim slFileName As String
    Dim slDate As String
    Dim slTime As String
    Dim slChar As String
    Dim ilUpper As Integer
    Dim slFTime As String
    Dim llFTime As Long
    Dim slUTime As String
    Dim llUTime As Long
    Dim slFIle As String
    Dim ilAddFile As Integer
    Dim ilFileExist As Integer
    Dim llSize As Long
    Dim ilOvMsg As Integer
    'Build array of files to be unzipped
    ReDim tmUnZipName(1 To 1) As UNZIPNAME
    ilUpper = 0
    ilPos = InStr(slFromFile, ".")
    If ilPos > 0 Then
        initUNZIPCmdStruct zpcDUnZip
        zpcDUnZip.ZIPFile = smModFilePath & slFromFile
        zpcDUnZip.QuietFlag = True
        zpcDUnZip.Filespec = "*.*"
        zpcDUnZip.NoDirectoryNamesFlag = True
        zpcDUnZip.ActionDZ = UNZIP_COUNTNAMEDZIPMEMBERS
        If zpcDUnZip.ErrorCode <> UE_OK Then
            lbcResult.AddItem "Not Valid Zip File " & Trim$(smModFilePath & slFromFile) & " Error #" & Str$(zpcDUnZip.ErrorCode)
            Print #hmMsg, "     " & "Not Valid Zip File " & Trim$(smModFilePath & slFromFile) & " Error #" & Str$(zpcDUnZip.ErrorCode)
            DoEvents
            mUnZipFiles = False
            Exit Function
        Else
            ilCount = zpcDUnZip.ReturnCount
            For ilLoop = 1 To ilCount Step 1
                zpcDUnZip.NoDirectoryNamesFlag = True
                zpcDUnZip.ActionDZ = UNZIP_GETNEXTNAMEDZIPINFO
                ilUpper = UBound(tmUnZipName)
                slFileName = zpcDUnZip.zi_FileName
                ilPos = Len(slFileName)
                slChar = Mid$(slFileName, ilPos, 1)
                slFIle = ""
                Do While slChar <> "\"
                    slFIle = slChar & slFIle
                    ilPos = ilPos - 1
                    If ilPos <= 0 Then
                        Exit Do
                    End If
                    slChar = Mid$(slFileName, ilPos, 1)
                Loop
                If (InStr(1, slFIle, "Copy", vbTextCompare) <= 0) And (InStr(1, slFIle, "Old", vbTextCompare) <= 0) And (InStr(1, slFIle, "Disable", vbTextCompare) <= 0) Then
                    tmUnZipName(ilUpper).sName = slFIle
                    tmUnZipName(ilUpper).sZipName = gGetTrueFileExt(slFIle)
                    tmUnZipName(ilUpper).sDate = gAdjYear(Format$(zpcDUnZip.zi_DateTime, "m/d/yy"))
                    tmUnZipName(ilUpper).sTime = Format$(zpcDUnZip.zi_DateTime, "h:mm:ssAM/PM")
                    tmUnZipName(ilUpper).lSize = zpcDUnZip.zi_oSize
                    tmUnZipName(ilUpper).sOldDate = ""
                    tmUnZipName(ilUpper).sOldTime = ""
                    tmUnZipName(ilUpper).lOldSize = 0
                    ilAddFile = False
                    ilFileExist = False
                    ilOvMsg = False
                    ilRet = 0
                    On Error GoTo mUnZipFilesErr:
                    slDateTime = FileDateTime(smResultFilePath & Trim$(tmUnZipName(ilUpper).sZipName))
                    If ilRet = 0 Then
                        ilFileExist = True
                        ilAddFile = True
                    Else
                        'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                        ilAddFile = True
                    End If
                    If ilOvMsg Then
                        ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sZipName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
                        If ilRet = 6 Then
                            'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                            ilAddFile = True
                        End If
                    End If
                    If ilAddFile Then
                        If ilFileExist Then
                            ilRet = GetAttr(smResultFilePath & Trim$(tmUnZipName(ilUpper).sZipName)) And 1
                            If ilRet <> 0 Then
                                SetAttr smResultFilePath & Trim$(tmUnZipName(ilUpper).sZipName), 0
                            End If
                        End If
                        ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    End If
                End If
            Next ilLoop
            
    
            For ilLoop = 1 To UBound(tmUnZipName) - 1 Step 1
'                If (rbcSave(0).Value) And (tmUnZipName(ilLoop).lOldSize > 0) Then
'                    ilRet = mZipFile(slFromFile, smResultFilePath & slSubfolder & "\" & tmUnZipName(ilLoop).sZipName, tmUnZipName(ilLoop).sOldDate, tmUnZipName(ilLoop).sOldTime, tmUnZipName(ilLoop).lOldSize)
'                End If
                initUNZIPCmdStruct zpcDUnZip
                zpcDUnZip.ZIPFile = smModFilePath & slFromFile
                zpcDUnZip.Destination = left$(smResultFilePath, Len(smResultFilePath) - 1)
                zpcDUnZip.NoDirectoryNamesFlag = True
                zpcDUnZip.Filespec = Trim$(tmUnZipName(ilLoop).sName)
                zpcDUnZip.OverwriteFlag = True
                zpcDUnZip.DiagnosticFlag = True
                zpcDUnZip.DecryptFlag = True
                zpcDUnZip.DecryptCode = smPassword
                zpcDUnZip.QuietFlag = True
                zpcDUnZip.ActionDZ = UNZIP_EXTRACT
                ilRet = zpcDUnZip.ErrorCode
                If ilRet = 0 Then
                    lbcResult.AddItem Trim$(tmUnZipName(ilLoop).sZipName) & " Unzipped into " & left$(smResultFilePath, Len(smResultFilePath) - 1)
                    If Trim$(tmUnZipName(ilLoop).sOldDate) = "" Then
                        Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sZipName) & " " & Trim$(tmUnZipName(ilLoop).sDate) & " " & Trim$(tmUnZipName(ilLoop).sTime) & " Size" & Str$(tmUnZipName(ilLoop).lSize) & " Unzipped into " & left$(smResultFilePath, Len(smResultFilePath) - 1)
                    Else
                        Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sZipName) & " " & Trim$(tmUnZipName(ilLoop).sDate) & " " & Trim$(tmUnZipName(ilLoop).sTime) & " Size" & Str$(tmUnZipName(ilLoop).lSize) & " Unzipped into " & left$(smResultFilePath, Len(smResultFilePath) - 1) & " Replaced: " & Trim$(tmUnZipName(ilLoop).sOldDate) & " " & Trim$(tmUnZipName(ilLoop).sOldTime) & " Size" & Str$(tmUnZipName(ilLoop).lOldSize)
                    End If
                Else
                    lbcResult.AddItem Trim$(tmUnZipName(ilLoop).sName) & " Error #" & Str$(ilRet) & " " & gUnzipError(ilRet)
                    Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sName) & " Error #" & Str$(ilRet) & " " & gUnzipError(ilRet)
                    'If (ilRet = 40) Or (ilRet = 41) Or (ilRet = 44) Or (ilRet = 48) Or (ilRet = 49) Then
                    '    mUnZipFiles = False
                    '    Exit Function
                    'End If
                End If
                DoEvents
            Next ilLoop
        End If
    End If
    mUnZipFiles = True
    Exit Function
mUnZipFilesErr:
    ilRet = Err
    Resume Next
End Function
Private Function mZipFile(slZipName As String, slInName As String, slInDate As String, slInTime As String, llSize As Long) As Integer
'    Dim slDateTime As String
'    Dim ilRet As Integer
'    Dim slName As String
'    Dim slZipPathName As String
'    Dim ilPos As Integer
'    Dim ilNameFd As Integer
'    Dim ilLoop As Integer
'    initZIPCmdStruct zpcDZip
'    On Error GoTo mZipFilesErr:
'    slZipPathName = smSaveFilePath & Trim$(slZipName)
'    ilNameFd = False
'    For ilLoop = LBound(tmZippedNames) To UBound(tmZippedNames) - 1 Step 1
'        If StrComp(Trim$(slZipName), Trim$(tmZippedNames(ilLoop).sName), 1) = 0 Then
'            ilNameFd = True
'            Exit For
'        End If
'    Next ilLoop
'    If Not ilNameFd Then
'        ilRet = 0
'        slDateTime = FileDateTime(slZipPathName)
'        If ilRet = 0 Then
'            Kill slZipPathName
'        End If
'        tmZippedNames(UBound(tmZippedNames)).sName = Trim$(slZipName)
'        ReDim Preserve tmZippedNames(0 To UBound(tmZippedNames) + 1) As ZIPPEDNAMES
'    End If
'    On Error GoTo 0
'    zpcDZip.ZIPFile = Trim$(slZipPathName)
'    zpcDZip.ItemList = Trim$(slInName)
'    zpcDZip.OldAsLatestFlag = True
'    zpcDZip.NoDirectoryNamesFlag = True
'    zpcDZip.EncryptFlag = True
'    zpcDZip.EncryptCode = smPassword
'    If Not ilNameFd Then
'        zpcDZip.ActionDZ = ZIP_ADD
'    Else
'        zpcDZip.ActionDZ = ZIP_UPDATE
'    End If
'    ilRet = zpcDZip.ErrorCode
'    If ilRet = 0 Then
'        Print #hmMsg, "     " & Trim$(slInName) & " " & Trim$(slInDate) & " " & Trim$(slInTime) & " Size" & Str$(llSize) & " saved into " & Trim$(slZipPathName)
'        'lbcResult.AddItem Trim$(slInName) & " zipped into " & slName
'    Else
'        Print #hmMsg, "     " & Trim$(slInName) & " saved into " & Trim$(slZipPathName) & " Failed: " & Str$(ilRet)
'        'lbcResult.AddItem Trim$(slInName) & " not zipped"
'    End If
'    Exit Function
'mZipFilesErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub lbcModList_Scroll()
    pbcLbcModList_Paint
End Sub

Private Sub pbcLbcModList_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilModEnd As Integer
    Dim ilField As Integer
    Dim slFields(1 To 3) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    
    ilModEnd = lbcModList.TopIndex + lbcModList.Height \ fgListHTArial825
    If ilModEnd > lbcModList.ListCount Then
        ilModEnd = lbcModList.ListCount
    End If
    If lbcModList.ListCount <= lbcModList.Height \ fgListHTArial825 Then
        llWidth = lbcModList.Width - 30
    Else
        llWidth = lbcModList.Width - igScrollBarWidth - 30
    End If
    pbcLbcModList.Width = llWidth
    pbcLbcModList.Cls
    llFgColor = pbcLbcModList.ForeColor
    For ilLoop = lbcModList.TopIndex To ilModEnd - 1 Step 1
        pbcLbcModList.ForeColor = llFgColor
        If lbcModList.MultiSelect = 0 Then
            If lbcModList.ListIndex = ilLoop Then
                gPaintArea pbcLbcModList, CSng(0), CSng((ilLoop - lbcModList.TopIndex) * fgListHTArial825), CSng(pbcLbcModList.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcModList.ForeColor = vbWhite
            End If
        Else
            If lbcModList.Selected(ilLoop) Then
                gPaintArea pbcLbcModList, CSng(0), CSng((ilLoop - lbcModList.TopIndex) * fgListHTArial825), CSng(pbcLbcModList.Width), CSng(fgListHTArial825) - 15, vbHighlight 'WHITE
                pbcLbcModList.ForeColor = vbWhite
            End If
        End If
        slStr = lbcModList.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        For ilField = LBound(slFields) To UBound(slFields) Step 1
            pbcLbcModList.CurrentX = imListField(ilField)
            pbcLbcModList.CurrentY = (ilLoop - lbcModList.TopIndex) * fgListHTArial825 + 15
            slStr = slFields(ilField)
            gAdjShowLen pbcLbcModList, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcModList.Print slStr
        Next ilField
        pbcLbcModList.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub tmcStart_Timer()
    Dim slName As String
    Dim slDateTime As String
    Dim slDate As String
    Dim llDate As Long
    Dim slTime As String
    Dim llTime As Long
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilPos As Integer
    Dim ilShow As Integer
    Dim slStr As String
    tmcStart.Enabled = False
    Screen.MousePointer = vbHourglass
    lbcModList.Clear
    pbcLbcModList_Paint
    smPassword = UCase$(edcPW.Text)
    ReDim tmModName(0 To 0) As ZIPNAME
    For ilLoop = 0 To lbcModFile.ListCount - 1 Step 1
        slName = lbcModFile.List(ilLoop)
        ilShow = True
'        ilPos = InStr(slName, ".")
'        If ilPos > 0 Then
'            slStr = Mid$(slName, ilPos + 1)
'            If StrComp(slStr, "Exe", 1) = 0 Then
'                ilShow = False
'            End If
'            If StrComp(slStr, "Zip", 1) = 0 Then
'                ilShow = False
'            End If
'        Else
'            ilShow = False
'        End If
        If ilShow Then
            'initUNZIPCmdStruct zpcDUnZip
            'zpcDUnZip.ZIPFile = smModFilePath & slName
            'zpcDUnZip.QuietFlag = True
            'zpcDUnZip.Filespec = "*.*"
            'zpcDUnZip.Action = UNZIP_COUNTNAMEDZIPMEMBERS
            'If zpcDUnZip.ErrorCode <> UE_OK Then
            '    ilShow = False
            'Else
'                If smPassword <> "" Then
'                    ilPos = InStr(slName, ".")
'                    If ilPos > 0 Then
'                        slStr = Mid$(slName, ilPos + 2)
'                        If InStr(1, smPassword, slStr, 1) = 0 Then
'                            ilShow = False
'                        End If
'                    Else
'                        ilShow = False
'                    End If
'                End If
            'End If
        End If
        If ilShow Then
            slDateTime = FileDateTime(smModFilePath & slName)
            slDate = Format$(slDateTime, "mm/dd/yyyy")
            llDate = DateValue(slDate)
            slDate = Trim$(Str$(99999 - llDate))
            Do While Len(slDate) < 5
                slDate = "0" & slDate
            Loop
            slTime = Format$(slDateTime, "hh:mm:ss")
            llTime = 3600 * Hour(slTime) + 60 * Minute(slTime) + Second(slTime)
            slTime = Trim$(Str$(99999 - llTime))
            Do While Len(slTime) < 5
                slTime = "0" & slTime
            Loop
            ilUpper = UBound(tmModName)
            ilPos = InStr(slName, ".")
            If ilPos > 0 Then
                tmModName(ilUpper).sKey = slDate & slTime & Mid$(slName, ilPos + 1, 1) & slName
                tmModName(ilUpper).sName = slName
                tmModName(ilUpper).sDate = Format$(slDateTime, "m/d/yy")
                tmModName(ilUpper).sTime = Format$(slDateTime, "h:mm:ssam/pm")
                ReDim Preserve tmModName(0 To ilUpper + 1) As ZIPNAME
            End If
        End If
    Next ilLoop
    ilUpper = UBound(tmModName)
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tmModName(), 0), ilUpper, 0, LenB(tmModName(0)), 0, Len(tmModName(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmModName) - 1 Step 1
        'ilPos = InStr(tmModName(ilLoop).sName, ".")
        'lbcModList.AddItem Trim$(Left$(tmModName(ilLoop).sName, ilPos - 1) & " " & tmModName(ilLoop).sDate & " " & tmModName(ilLoop).sTime)
        slName = Trim$(tmModName(ilLoop).sName)
        lbcModList.AddItem slName & "|" & Trim$(tmModName(ilLoop).sDate) & "|" & Trim$(tmModName(ilLoop).sTime)
    Next ilLoop
    pbcLbcModList_Paint
    Screen.MousePointer = vbDefault
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Database Unzip Files"
End Sub

Private Sub zpcDUnZip_UnZipRenameCallback(ItemName As String, iDate As Long, iTime As Long, lAttrib As Long, OrigItemName As String, rc As Long)
            
    Dim ilPos As Integer
    Dim slTempStr As String
    
    ilPos = InStrRev(ItemName, ".", Len(ItemName) - 1)
    slTempStr = UCase(Mid$(ItemName, ilPos + 1))
    
    If slTempStr = "FWE" Then
        ItemName = left$(ItemName, ilPos) & "exe"
    End If
    
    If slTempStr = "FWL" Then
        ItemName = left$(ItemName, ilPos) & "dll"
    End If

    If slTempStr = "FWA" Then
        ItemName = left$(ItemName, ilPos) & "oca"
    End If

    If slTempStr = "FWX" Then
        ItemName = left$(ItemName, ilPos) & "ocx"
    End If

End Sub


