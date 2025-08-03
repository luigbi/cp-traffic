VERSION 5.00
Begin VB.Form CSIDir 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   1005
   ClientTop       =   3045
   ClientWidth     =   6210
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
   ScaleHeight     =   3855
   ScaleWidth      =   6210
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
      Height          =   1905
      Left            =   255
      ScaleHeight     =   1845
      ScaleWidth      =   5460
      TabIndex        =   7
      Top             =   1290
      Width           =   5520
      Begin VB.ListBox lbcResult 
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
         Height          =   1785
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   5430
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5145
      Top             =   3315
   End
   Begin VB.PictureBox plcMod 
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
      Height          =   615
      Left            =   255
      ScaleHeight     =   555
      ScaleWidth      =   5460
      TabIndex        =   1
      Top             =   390
      Width           =   5520
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
         Left            =   3165
         TabIndex        =   2
         Top             =   60
         Width           =   2295
      End
      Begin VB.Label lacMod 
         Appearance      =   0  'Flat
         Caption         =   "Select Drive Letter For Counterpoint Folders"
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
         Height          =   420
         Left            =   90
         TabIndex        =   6
         Top             =   75
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmcCreate 
      Appearance      =   0  'Flat
      Caption         =   "C&reate Folders"
      Height          =   285
      Left            =   1395
      TabIndex        =   3
      Top             =   3435
      Width           =   1410
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3195
      TabIndex        =   4
      Top             =   3435
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
End
Attribute VB_Name = "CSIDir"
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
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim smNowDate As String   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim smTodaysDate As String
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
    Dim slDrive As String
    Dim ilRet As Integer
    
    imProcessing = True
    lbcResult.Clear
    slDrive = cbcMod.Drive
    ilRet = 0
    MkDir slDrive & "\CSI"
    ilRet = 0
    MkDir slDrive & "\CSI\PROD"
    'mMakeFolder slDrive & "\CSI\PROD\AFFEXE"
    'mMakeFolder slDrive & "\CSI\PROD\AFFREPORTS"
    'mMakeFolder slDrive & "\CSI\PROD\AFFSETUP"
    'mMakeFolder slDrive & "\CSI\PROD\AFFWORK"
    mMakeFolder slDrive & "\CSI\PROD\DATA"
    mMakeFolder slDrive & "\CSI\PROD\DATA\EXPORT"
    mMakeFolder slDrive & "\CSI\PROD\DATA\IMPORT"
    mMakeFolder slDrive & "\CSI\PROD\DATA\EXTRACT"
    mMakeFolder slDrive & "\CSI\PROD\DATA\EXTRACT\PE"
    mMakeFolder slDrive & "\CSI\PROD\DATADEMO"
    mMakeFolder slDrive & "\CSI\PROD\EXE"
    mMakeFolder slDrive & "\CSI\PROD\EXPORT"
    mMakeFolder slDrive & "\CSI\PROD\HELP"
    mMakeFolder slDrive & "\CSI\PROD\IMPORT"
    mMakeFolder slDrive & "\CSI\PROD\NEWDDF"
    mMakeFolder slDrive & "\CSI\PROD\REPORTS"
    mMakeFolder slDrive & "\CSI\PROD\REPORTS\GENERIC"
    mMakeFolder slDrive & "\CSI\PROD\SETUP"
    mMakeFolder slDrive & "\CSI\PROD\SETUP\BATCH"
    mMakeFolder slDrive & "\CSI\PROD\SETUP\FONTS"
    mMakeFolder slDrive & "\CSI\PROD\SETUP\TOPBOT"
    mMakeFolder slDrive & "\CSI\PROD\SETUP\WINZIP"
    mMakeFolder slDrive & "\CSI\PROD\UPDATES"
    mMakeFolder slDrive & "\CSI\PROD\UPDATES\DONE"
    mMakeFolder slDrive & "\CSI\PROD\UPDATES\TODO"
    mMakeFolder slDrive & "\CSI\PROD\ZERO_BTR"
    ilRet = 0
    MkDir slDrive & "\CSI\TEST"
    'mMakeFolder slDrive & "\CSI\TEST\AFFEXE"
    'mMakeFolder slDrive & "\CSI\TEST\AFFREPORTS"
    mMakeFolder slDrive & "\CSI\TEST\DATA"
    mMakeFolder slDrive & "\CSI\TEST\DATA\EXPORT"
    mMakeFolder slDrive & "\CSI\TEST\DATA\IMPORT"
    mMakeFolder slDrive & "\CSI\TEST\EXE"
    mMakeFolder slDrive & "\CSI\TEST\EXPORT"
    mMakeFolder slDrive & "\CSI\TEST\IMPORT"
    mMakeFolder slDrive & "\CSI\TEST\NEWDDF"
    mMakeFolder slDrive & "\CSI\TEST\REPORTS"
    mMakeFolder slDrive & "\CSI\TEST\REPORTS\GENERIC"
    'mMakeFolder slDrive & "\CSI\TEST\SETUP"
    'mMakeFolder slDrive & "\CSI\TEST\SETUP\BATCH"
    'mMakeFolder slDrive & "\CSI\TEST\SETUP\FONTS"
    'mMakeFolder slDrive & "\CSI\TEST\SETUP\TOPBOT"
    'mMakeFolder slDrive & "\CSI\TEST\SETUP\WINZIP"
    mMakeFolder slDrive & "\CSI\TEST\UPDATES"
    mMakeFolder slDrive & "\CSI\TEST\UPDATES\DONE"
    mMakeFolder slDrive & "\CSI\TEST\UPDATES\TODO"
    mMakeFolder slDrive & "\CSI\TEST\ZERO_BTR"
    
    imProcessing = False
    cmcCancel.Caption = "&Done"
    cmcCreate.Enabled = False
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub Form_Load()
    mInit
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
    CSIDir.Height = cmcCreate.top + 5 * cmcCreate.Height / 2
    'gCenterStdAlone BtrMatch
    CSIDir.Move (Screen.Width - CSIDir.Width) \ 2, (Screen.Height - CSIDir.Height) \ 2 + 175 '+ Screen.Height \ 10
    'CSIDir.Show
    
    Screen.MousePointer = vbHourglass
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
Private Sub mMakeFolder(slDrivePath As String)
    Dim ilRet As Integer
    'Move report logo to local C drice (c:\csi\rptlogo.bmp)
    On Error GoTo mMakeFolderErr:
    ilRet = 0
    MkDir slDrivePath
    If ilRet = 0 Then
        lbcResult.AddItem slDrivePath & " Created"
    ElseIf ilRet = 75 Then
        lbcResult.AddItem slDrivePath & " Previously Existed"
    Else
        lbcResult.AddItem slDrivePath & " Error #" & Str$(ilRet)
    End If
    DoEvents
    On Error GoTo 0
    Exit Sub
mMakeFolderErr:
    ilRet = Err
    Resume Next
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
    Screen.MousePointer = vbDefault
    Unload CSIDir
    Set CSIDir = Nothing
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Create Counterpoint Folders"
End Sub
