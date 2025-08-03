VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CSIZip 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5730
   ClientLeft      =   915
   ClientTop       =   1755
   ClientWidth     =   12900
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
   ScaleHeight     =   5730
   ScaleWidth      =   12900
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   225
      Left            =   7755
      TabIndex        =   29
      Top             =   4140
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
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
      Height          =   3630
      Left            =   7755
      ScaleHeight     =   3570
      ScaleWidth      =   4950
      TabIndex        =   19
      Top             =   375
      Width           =   5010
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
         Left            =   1035
         TabIndex        =   21
         Top             =   45
         Width           =   3765
      End
      Begin VB.DirListBox lbcResultDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1050
         Left            =   60
         TabIndex        =   22
         Top             =   435
         Width           =   4755
      End
      Begin VB.ListBox lbcResult 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   60
         TabIndex        =   23
         Top             =   1590
         Width           =   4755
      End
      Begin VB.Label lacResult 
         Appearance      =   0  'Flat
         Caption         =   "Results"
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
         Left            =   60
         TabIndex        =   20
         Top             =   75
         Width           =   705
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10455
      Top             =   5220
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
      Height          =   4860
      Left            =   30
      ScaleHeight     =   4800
      ScaleWidth      =   7485
      TabIndex        =   1
      Top             =   375
      Width           =   7545
      Begin VB.FileListBox lbcModFile 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   90
         MultiSelect     =   2  'Extended
         Pattern         =   "*.Exe"
         TabIndex        =   18
         Top             =   3270
         Visible         =   0   'False
         Width           =   7230
      End
      Begin VB.PictureBox pbcLbcModList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   90
         ScaleHeight     =   2100
         ScaleWidth      =   7230
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2610
         Width           =   7230
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
         Height          =   285
         Left            =   3885
         MaxLength       =   3
         TabIndex        =   14
         Top             =   855
         Width           =   540
      End
      Begin VB.PictureBox plcMod 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   795
         Left            =   90
         ScaleHeight     =   795
         ScaleWidth      =   5550
         TabIndex        =   2
         Top             =   60
         Width           =   5550
         Begin VB.OptionButton rbcMod 
            Caption         =   "Setup\Fonts"
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
            Height          =   240
            Index           =   9
            Left            =   3330
            TabIndex        =   12
            Top             =   525
            Width           =   1425
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Setup\ClientInstall\System32"
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
            Height          =   240
            Index           =   8
            Left            =   390
            TabIndex        =   11
            Top             =   525
            Width           =   2850
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Data"
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
            Height          =   240
            Index           =   7
            Left            =   4350
            TabIndex        =   10
            Top             =   255
            Width           =   735
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Generic Rpt"
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
            Height          =   240
            Index           =   6
            Left            =   2730
            TabIndex        =   9
            Top             =   255
            Width           =   1365
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Zero Btr"
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
            Height          =   240
            Index           =   5
            Left            =   1350
            TabIndex        =   8
            Top             =   255
            Width           =   1020
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "DDF"
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
            Height          =   240
            Index           =   4
            Left            =   390
            TabIndex        =   7
            Top             =   255
            Width           =   720
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Affiliate Rpt"
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
            Height          =   240
            Index           =   3
            Left            =   4575
            TabIndex        =   6
            Top             =   615
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Affiliate Exe"
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
            Height          =   240
            Index           =   2
            Left            =   3150
            TabIndex        =   5
            Top             =   645
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Reports (Rpt,Jpg,Gif,Bmp)"
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
            Height          =   240
            Index           =   1
            Left            =   3015
            TabIndex        =   4
            Top             =   -15
            Width           =   2625
         End
         Begin VB.OptionButton rbcMod 
            Caption         =   "Exe/DLL/OCX/PDB/XML"
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
            Height          =   240
            Index           =   0
            Left            =   390
            TabIndex        =   3
            Top             =   -15
            Width           =   2490
         End
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
         Left            =   105
         TabIndex        =   15
         Top             =   1260
         Width           =   2625
      End
      Begin VB.DirListBox lbcMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   2895
         TabIndex        =   16
         Top             =   1260
         Width           =   4425
      End
      Begin VB.ListBox lbcModList 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   75
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   2595
         Width           =   7230
      End
      Begin VB.Label lacPW 
         Appearance      =   0  'Flat
         Caption         =   "Zip Encrypt Password(Vnn; nn is version #)"
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
         Left            =   105
         TabIndex        =   13
         Top             =   885
         Width           =   3750
      End
   End
   Begin VB.CommandButton cmcZip 
      Appearance      =   0  'Flat
      Caption         =   "&Zip"
      Height          =   285
      Left            =   4890
      TabIndex        =   24
      Top             =   5325
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   6750
      TabIndex        =   25
      Top             =   5325
      Width           =   1140
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   2805
      TabIndex        =   0
      Top             =   30
      Width           =   2805
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   9720
      OleObjectBlob   =   "CSIZIP.frx":0000
      Top             =   5190
   End
   Begin VB.Label lacTotal 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5370
      TabIndex        =   28
      Top             =   4920
      Width           =   3045
   End
   Begin VB.Label lacResults 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2955
      TabIndex        =   27
      Top             =   45
      Width           =   5415
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5370
      TabIndex        =   26
      Top             =   4605
      Width           =   3045
   End
End
Attribute VB_Name = "CSIZip"
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
Dim tmZippedNames() As ZIPPEDNAMES
Dim smPassword As String
Dim smExport As String
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imFill As Integer   '0=Traffic Exe; 1=Traffic Rpt; 2=Affiliate Exe; 3=Affiliate Rpr; 4=DDF; 5=Zero_Btr; 6=Generic Rpt; 7=Data; 8=Setup\ClientInstall\System32; 9=Fonts
Dim smNowDate As String   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim smModFilePath As String
Dim smResultFilePath As String
Dim smTodaysDate As String
Dim imListField(1 To 4) As Integer
Dim bmAuto As Boolean
Dim bmSelectAll As Boolean

Private Sub cbcMod_Change()
    lbcMod.Path = cbcMod.Drive
    cbcResultDir.Drive = cbcMod.Drive
    mSetDirectories
End Sub
Private Sub cbcResultDir_Change()
    lbcResultDir.Path = cbcResultDir.Drive
End Sub
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
Private Sub cmcZip_Click()
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
    Dim slStr As String
    smPassword = UCase$(Trim$(edcPW.Text))
    If left$(smPassword, 1) <> "V" Then
        MsgBox " Password form not correct"
        Exit Sub
    End If
    slStr = Mid$(smPassword, 2)
    If Len(slStr) <> 2 Then
        MsgBox " Password form not correct"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    smTodaysDate = Format$(Now, "m/d/yy")
    smResultFilePath = lbcResultDir.Path & "\"
    ilPos = InStrRev(lbcMod.Path, "\")
    sgDBPath = left$(lbcMod.Path, ilPos) & "data\"
    smModFilePath = lbcMod.Path & "\"
    imCancelled = False
    ilRet = 0
    On Error GoTo cmcChkErr:
    slDateTime = FileDateTime(slMsgFile)
    If Not mOpenMsgFile(slMsgFile) Then
        Exit Sub
    End If
    lacResults.Caption = "Results Stored into: " & slMsgFile
    On Error GoTo 0
    Print #hmMsg, "Zipping Files " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmMsg, ""
    Select Case imFill
        Case 0
            Print #hmMsg, "Exe/DLL/OCX/PDB/XML Zipped"
        Case 1
            Print #hmMsg, "Reports (Rpt) Zipped"
        Case 2
            Print #hmMsg, "Affiliate Exe Zipped"
        Case 3
            Print #hmMsg, "Affiliate Rpt Zipped"
        Case 4
            Print #hmMsg, "DDF Zipped"
        Case 5
            Print #hmMsg, "Zero Btr Zipped"
        Case 6
            Print #hmMsg, "Generic Rpt Zipped"
        Case 7
            Print #hmMsg, "Data Zipped"
        Case 8
            Print #hmMsg, "Setup\ClientInstall\System32 Zipped"
        Case 9
            Print #hmMsg, "Fonts Zipped"
    End Select
    
    llCount = 0
    For ilLoop = 0 To lbcModList.ListCount - 1 Step 1
        If lbcModList.Selected(ilLoop) Then
            llCount = llCount + 1
        End If
    Next ilLoop
    ReDim tmZippedNames(0 To 0) As ZIPPEDNAMES
    lbcResult.Clear
    imProcessing = True
    plcGauge.Value = 0
    llProcessed = 0
    ilAddFile = False
    For ilLoop = 0 To lbcModList.ListCount - 1 Step 1
        If lbcModList.Selected(ilLoop) Then
            DoEvents
            slName = tmModName(ilLoop).sName
            lacProcessing.Caption = slName
            'If Not ilAddFile Then
                ilRet = mZipFiles(slName, tmModName(ilLoop).sDate, tmModName(ilLoop).sTime, tmModName(ilLoop).lSize)
            'Else
            '    ilRet = mAddZipFiles(slName, tmModName(ilLoop).sDate, tmModName(ilLoop).sTime)
            'End If
            'If (imFill = 4) Or (imFill = 5) Or (imFill = 7) Then
            '    ilAddFile = True
            'End If
            If imCancelled Then
                Print #hmMsg, ""
                Print #hmMsg, "Cancelled " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Close #hmMsg
                Exit Sub
            End If
            llProcessed = llProcessed + 1
            llPercent = (llProcessed * 100) \ llCount
            If llPercent >= 100 Then
                llPercent = 100
            End If
            plcGauge.Value = llPercent
        End If
    Next ilLoop
    Print #hmMsg, ""
    Print #hmMsg, "Completed " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close #hmMsg
    plcGauge.Value = 100
    lacProcessing.Caption = "Processing- Done"
    imProcessing = False
    cmcCancel.Caption = "&Done"
    'cmcZip.Enabled = False
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
cmcChkErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcPW_Change()
    tmcStart.Enabled = False
    
    mSetDirectories
    
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
    sgExePath = App.Path
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
    imTerminate = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass
    
    fgBoxGridH = 180      'Height of grid area (distance from bottom of form letter to bottom of form box)
    fgBoxInsetX = 30    'X Margin from box outline to text
    fgListHTArial825 = 210  '225  'Height of list item (arial 8.25 bold; DivideStyle=1-Black line; ListStyle=0-2D)
    fgListHtSerif825 = 195  '255  'Height of list item (san serif 8.25; Divide style=3=Raised; ListStyle=1-3D)
    igAlignCharWidth = 90
    igScrollBarWidth = 255
    
    imListField(1) = 15
    imListField(2) = 50 * igAlignCharWidth  '20 * igAlignCharWidth 'igAvgCharWidth * 5
    imListField(3) = 62 * igAlignCharWidth  '32 * igAlignCharWidth 'igAvgCharWidth * 12
    imListField(4) = 90 * igAlignCharWidth  '60 * igAlignCharWidth 'igAvgCharWidth * 12

    
    CSIZip.Height = cmcZip.top + 5 * cmcZip.Height / 3
    'gCenterStdAlone BtrMatch
    CSIZip.Move (Screen.Width - CSIZip.Width) \ 2, (Screen.Height - CSIZip.Height) \ 2 + 175 '+ Screen.Height \ 10
    CSIZip.Show
    
    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    smNowDate = Format$(Now, "m/d/yy")
    
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
    On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpStnFd.Txt"
    slToFile = smResultFilePath & "CSIZip.Txt"
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
                MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
                MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile, vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
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
    Dim slType As String
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
    
    '9/24/21 - to Automate Type (Exe, Report, etc)
    smExport = ""
    ilPos = InStr(1, slCommand, "/T:", 1)
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
        slType = UCase(Trim(smExport))
        Select Case slType
            Case "EXE": rbcMod(0).Value = True
            Case "REPORTS": rbcMod(1).Value = True
            Case "DDF": rbcMod(4).Value = True
            Case "ZERO": rbcMod(5).Value = True
            Case "GENERIC": rbcMod(6).Value = True
            Case "DATA": rbcMod(7).Value = True
            Case "SETUP": rbcMod(8).Value = True
            Case "FONTS": rbcMod(9).Value = True
        End Select
    End If
    
    
    '9/24/21 - to Automate Source path
    smExport = ""
    ilPos = InStr(1, slCommand, "/S:", 1)
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
        cbcMod.Drive = left$(smExport, 1)
        lbcMod.Path = smExport
    End If
    
    '9/24/21 - to Automate Selection of Source Files and to Quit After
    smExport = ""
    bmAuto = False
    ilPos = InStr(1, slCommand, "/AUTO", 1)
    If ilPos > 0 Then
        bmAuto = True
    End If
    
    
    '9/24/21 - to Automate Selection of Source Files and to Quit After
    smExport = ""
    bmSelectAll = False
    ilPos = InStr(1, slCommand, "/ALL", 1)
    If ilPos > 0 Then
        bmSelectAll = True
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
    Erase tmZippedNames
    Screen.MousePointer = vbDefault
    Unload CSIZip
    Set CSIZip = Nothing
    End
End Sub
Private Function mZipFiles(slInName As String, slInDate As String, slInTime As String, llSize As Long) As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim slName As String
    Dim slZipPathName As String
    Dim ilPos As Integer
    Dim ilNameFd As Integer
    Dim ilLoop As Integer
    Dim slExt As String
    Dim slMatchName As String
    initZIPCmdStruct zpcDZip
    On Error GoTo mZipFilesErr:
    ilPos = InStrRev(slInName, ".")
    If ilPos > 0 Then
        slName = left$(slInName, ilPos - 1)
        slExt = Mid(slInName, ilPos + 1)
        'slMatchName = slInName
        If Trim$(UCase$(slExt)) = "PDB" Then
            slMatchName = slName & ".exe"
        ElseIf Trim$(UCase$(slExt)) = "XML" Then
            slMatchName = slName & ".exe"
        Else
            slMatchName = slInName
        End If
    Else
        slName = slInName
        slExt = ""
        slMatchName = slName
    End If
    If (imFill = 4) Then
        slName = "DDF"
        slMatchName = slName
    ElseIf (imFill = 7) Then
        slName = "DATA"
        slMatchName = slName
    End If
    slZipPathName = LCase$(Trim$(mZipName(slName, slExt)))
    ilNameFd = False
    For ilLoop = LBound(tmZippedNames) To UBound(tmZippedNames) - 1 Step 1
        If StrComp(Trim$(slMatchName), Trim$(tmZippedNames(ilLoop).sName), 1) = 0 Then
            ilNameFd = True
            Exit For
        End If
    Next ilLoop
    If Not ilNameFd Then
        ilRet = 0
        slDateTime = FileDateTime(slZipPathName)
        If ilRet = 0 Then
            Kill slZipPathName
        End If
        tmZippedNames(UBound(tmZippedNames)).sName = slMatchName
        ReDim Preserve tmZippedNames(0 To UBound(tmZippedNames) + 1) As ZIPPEDNAMES
    End If
    On Error GoTo 0
    zpcDZip.ZIPFile = slZipPathName
    zpcDZip.ItemList = smModFilePath & Trim$(slInName)
    zpcDZip.OldAsLatestFlag = False 'False to use todays date; True to use date of latest module zipped
    zpcDZip.NoDirectoryNamesFlag = True
    zpcDZip.EncryptFlag = True
    zpcDZip.EncryptCode = smPassword
    If Not ilNameFd Then
        zpcDZip.ActionDZ = ZIP_ADD
    Else
        zpcDZip.ActionDZ = ZIP_UPDATE
    End If
    ilRet = zpcDZip.ErrorCode
    If ilRet = 0 Then
        Print #hmMsg, "    " & smModFilePath & Trim$(slInName) & " " & Trim$(slInDate) & " " & Trim$(slInTime) & " Size" & Str$(llSize) & " zipped into " & slZipPathName
'        Print #hmMsg, "        " & slInName & " " & slInDate & " " & slInTime
        lbcResult.AddItem Trim$(slInName) & " zipped into " & slName
    Else
        Print #hmMsg, "    " & smModFilePath & Trim$(slInName) & " zipped into " & slZipPathName & " Failed: " & Str$(ilRet)
        lbcResult.AddItem Trim$(slInName) & " not zipped"
    End If
    Exit Function
mZipFilesErr:
    ilRet = 1
    Resume Next
End Function
Private Function mZipName(slName As String, slExt As String) As String
    Select Case imFill
        Case 0
            mZipName = smResultFilePath & slName & ".E" & Mid$(smPassword, 2, 2)
        Case 1
            If StrComp(left$(slExt, 1), "R", vbTextCompare) <> 0 Then
                mZipName = smResultFilePath & slName & ".J" & Mid$(smPassword, 2, 2)
            Else
                mZipName = smResultFilePath & slName & ".R" & Mid$(smPassword, 2, 2)
            End If
        Case 2
            mZipName = smResultFilePath & slName & ".A" & Mid$(smPassword, 2, 2)
        Case 3
            mZipName = smResultFilePath & slName & ".B" & Mid$(smPassword, 2, 2)
        Case 4
            mZipName = smResultFilePath & slName & ".N" & Mid$(smPassword, 2, 2)
        Case 5
            mZipName = smResultFilePath & slName & ".Z" & Mid$(smPassword, 2, 2)
        Case 6
            mZipName = smResultFilePath & slName & ".G" & Mid$(smPassword, 2, 2)
        Case 7
            mZipName = smResultFilePath & slName & ".D" & Mid$(smPassword, 2, 2)
        Case 8
            mZipName = smResultFilePath & slName & ".S" & Mid$(smPassword, 2, 2)
        Case 9
            mZipName = smResultFilePath & slName & ".F" & Mid$(smPassword, 2, 2)
    End Select
End Function

Private Sub lbcModList_Click()
    pbcLbcModList_Paint
End Sub

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

Private Sub rbcMod_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcMod(Index).Value
    'End of coded added
    If Value Then
        imFill = Index
        lbcModList.Clear
        pbcLbcModList_Paint
        ReDim tmModName(0 To 0) As ZIPNAME
        Select Case imFill
            Case 0
                lbcModFile.Pattern = "*.Exe;*.vbx;*.dll;*.avi;*.tlb;*.lng;*.ini;*.ocx;*.oca;*.pdb;*.xml"
            Case 1
                lbcModFile.Pattern = "*.Rpt;*.gif;*jpg;*.bmp"
            Case 2
                lbcModFile.Pattern = "Af*.Exe;*.ocx;*.dll;*.ini;*.tlb;*.oca"
            Case 3
                lbcModFile.Pattern = "Af*.Rpt;C*.Rpt;L*.Rpt;*.gif;*jpg"
            Case 4
                lbcModFile.Pattern = "*.DDF;*.New"
            Case 5
                lbcModFile.Pattern = "*.Btr;*.Mkd"
            Case 6
                lbcModFile.Pattern = "*.Rpt"
            Case 7
                lbcModFile.Pattern = "*.*"
            Case 8
                '10241 added batch files
                lbcModFile.Pattern = "*.Exe;*.MSI;*.dll;*.ocx;*.oca;*.txt;*.bat"
            Case 9
                lbcModFile.Pattern = "*.Exe;*.TTF;*.fon"
        End Select
        tmcStart.Enabled = True
    End If
End Sub
Private Sub rbcMod_GotFocus(Index As Integer)
    tmcStart.Enabled = False
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
    Dim llSize As Long
    tmcStart.Enabled = False
    Screen.MousePointer = vbHourglass
    lbcModList.Clear
    pbcLbcModList_Paint
    ReDim tmModName(0 To 0) As ZIPNAME
    For ilLoop = 0 To lbcModFile.ListCount - 1 Step 1
        slName = lbcModFile.List(ilLoop)
        ilShow = True
        If imFill = 0 Then
            'If StrComp(left$(slName, 2), "AF", 1) = 0 Then
            '    ilShow = False
            'End If
        End If
        If imFill = 1 Then
            'If StrComp(left$(slName, 2), "AF", 1) = 0 Then
            '    ilShow = False
            'End If
        End If
        If imFill = 2 Then
            If StrComp(left$(slName, 2), "AF", 1) <> 0 Then
                ilShow = False
            End If
        End If
        If imFill = 3 Then
            If (StrComp(left$(slName, 2), "AF", 1) <> 0) And (StrComp(left$(slName, 1), "C", 1) <> 0) And (StrComp(left$(slName, 1), "L", 1) <> 0) Then
                ilShow = False
            End If
        End If
        If ilShow Then
            slDateTime = FileDateTime(smModFilePath & slName)
            slDate = DateValue(slDateTime)
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
            llSize = FileLen(smModFilePath & slName)
            ilUpper = UBound(tmModName)
            tmModName(ilUpper).sKey = slDate & slTime & UCase$(slName)
            tmModName(ilUpper).sName = LCase$(slName)
            tmModName(ilUpper).sDate = Format$(slDateTime, "m/d/yy")
            tmModName(ilUpper).sTime = Format$(slDateTime, "h:mm:ssam/pm")
            tmModName(ilUpper).lSize = llSize
            ReDim Preserve tmModName(0 To ilUpper + 1) As ZIPNAME
        End If
    Next ilLoop
    ilUpper = UBound(tmModName)
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tmModName(), 0), ilUpper, 0, LenB(tmModName(0)), 0, LenB(tmModName(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmModName) - 1 Step 1
        'ilPos = InStr(tmModName(ilLoop).sName, ".")
        'lbcModList.AddItem Trim$(Left$(tmModName(ilLoop).sName, ilPos - 1) & " " & tmModName(ilLoop).sDate & " " & tmModName(ilLoop).sTime)
        slName = Trim$(tmModName(ilLoop).sName)
        lbcModList.AddItem slName & "|" & Trim$(tmModName(ilLoop).sDate) & "|" & Trim$(tmModName(ilLoop).sTime)
    Next ilLoop
    pbcLbcModList_Paint
    Screen.MousePointer = vbDefault
    
    If bmAuto = True Then
        If bmSelectAll = True Then
            'Select all files
            For ilLoop = 0 To UBound(tmModName) - 1
                lbcModList.Selected(ilLoop) = True
            Next ilLoop
        Else
            'Select TODAYS files
            For ilLoop = 0 To UBound(tmModName) - 1
                If DateValue(tmModName(ilLoop).sDate) = DateValue(Date) Then
                    lbcModList.Selected(ilLoop) = True
                End If
            Next ilLoop
        End If
    End If
    
    pbcLbcModList_Paint
    Screen.MousePointer = vbDefault
    
    If bmAuto = True Then
        Screen.MousePointer = vbHourglass
        cmcZip_Click
        
        Screen.MousePointer = vbDefault
        cmcCancel_Click
    End If
End Sub
Private Sub plcMod_Paint()
    plcMod.CurrentX = 0
    plcMod.CurrentY = 0
    plcMod.Print "Zip"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Zip Files for Transfer to Internet"
End Sub
Private Sub zpcDZip_ZipRenameCallback(ItemName As String, iDate As Long, iTime As Long, lAttrib As Long, OrigItemName As String, rc As Long)
    
    Dim ilPos As Integer
    Dim slTempStr As String
    
    ilPos = InStrRev(ItemName, ".", Len(ItemName) - 1)
    slTempStr = UCase(Mid$(ItemName, ilPos + 1))
    
    If slTempStr = "EXE" Then
        ItemName = left$(ItemName, ilPos) & "fwe"
    End If
    
    If slTempStr = "DLL" Then
        ItemName = left$(ItemName, ilPos) & "fwl"
    End If
    
    If slTempStr = "OCA" Then
        ItemName = left$(ItemName, ilPos) & "fwa"
    End If
    
    If slTempStr = "OCX" Then
        ItemName = left$(ItemName, ilPos) & "fwx"
    End If
    
End Sub

Private Sub mSetDirectories()
    Dim ilRet As Integer
    Dim slDrive As String
    
    On Error GoTo TryProd:
    ilRet = 0
    slDrive = cbcMod.Drive
    If InStr(1, slDrive, "[", vbTextCompare) > 0 Then
        slDrive = left$(slDrive, 2)
    End If
    lbcMod.Path = slDrive & "\csi\" & Trim$(edcPW.Text)
    If ilRet = 1 Then
        ilRet = 0
        lbcMod.Path = slDrive & "\csi\prod"
        If ilRet = 1 Then
            ilRet = 0
            lbcMod.Path = slDrive & "\csi\prodengr"
        End If
    Else
        lbcResultDir.Path = lbcMod.Path
    End If
    If ilRet = 0 Then
        lbcResultDir.Path = lbcMod.Path & "\Updates"
    End If
    Exit Sub
TryProd:
    ilRet = 1
    Resume Next
End Sub
