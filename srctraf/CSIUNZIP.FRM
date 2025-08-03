VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{0FB90DC1-97D1-11D1-87C0-444553540000}#1.0#0"; "duzactx.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CSIUnzip 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7200
   ClientLeft      =   915
   ClientTop       =   3225
   ClientWidth     =   10095
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
   ScaleHeight     =   7200
   ScaleWidth      =   10095
   Begin VB.ListBox lbcResult 
      Appearance      =   0  'Flat
      Height          =   1080
      ItemData        =   "CSIUNZIP.frx":0000
      Left            =   105
      List            =   "CSIUNZIP.frx":0002
      TabIndex        =   38
      Top             =   5010
      Width           =   9765
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   195
      Left            =   2565
      TabIndex        =   35
      Top             =   6195
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.FileListBox lbcAffReports 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   105
      MultiSelect     =   2  'Extended
      Pattern         =   "*.Exe"
      TabIndex        =   29
      Top             =   6435
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmcCreate 
      Appearance      =   0  'Flat
      Caption         =   "Create &Folder"
      Height          =   285
      Left            =   7275
      TabIndex        =   25
      Top             =   7005
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
      Height          =   4635
      Left            =   6375
      ScaleHeight     =   4575
      ScaleWidth      =   3435
      TabIndex        =   19
      Top             =   270
      Width           =   3495
      Begin VB.DirListBox lbcSaveDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   60
         TabIndex        =   34
         Top             =   2775
         Width           =   3315
      End
      Begin VB.DriveListBox cbcSaveDir 
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
         Left            =   1410
         TabIndex        =   33
         Top             =   2415
         Width           =   1965
      End
      Begin VB.PictureBox plcSave 
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
         Height          =   255
         Left            =   90
         ScaleHeight     =   255
         ScaleWidth      =   3195
         TabIndex        =   30
         Top             =   2175
         Width           =   3195
         Begin VB.OptionButton rbcSave 
            Caption         =   "No"
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
            Height          =   195
            Index           =   1
            Left            =   2370
            TabIndex        =   32
            Top             =   0
            Width           =   570
         End
         Begin VB.OptionButton rbcSave 
            Caption         =   "Yes"
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
            Height          =   195
            Index           =   0
            Left            =   1635
            TabIndex        =   31
            Top             =   0
            Width           =   645
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
         Left            =   1410
         TabIndex        =   21
         Top             =   45
         Width           =   1965
      End
      Begin VB.DirListBox lbcResultDir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   60
         TabIndex        =   22
         Top             =   435
         Width           =   3315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   75
         X2              =   3345
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Label lacSave 
         Appearance      =   0  'Flat
         Caption         =   "Save old into"
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
         Left            =   90
         TabIndex        =   28
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label lacResult 
         Appearance      =   0  'Flat
         Caption         =   "Unzip into subfolder of"
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
         Left            =   90
         TabIndex        =   20
         Top             =   45
         Width           =   1455
      End
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   6915
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
      Height          =   4635
      Left            =   105
      ScaleHeight     =   4575
      ScaleWidth      =   6045
      TabIndex        =   1
      Top             =   270
      Width           =   6105
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   39
         Top             =   4290
         Width           =   1350
      End
      Begin VB.FileListBox lbcModFile 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   60
         MultiSelect     =   2  'Extended
         Pattern         =   "*.Exe"
         TabIndex        =   18
         Top             =   3210
         Visible         =   0   'False
         Width           =   5235
      End
      Begin VB.PictureBox pbcLbcModList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   75
         ScaleHeight     =   1455
         ScaleWidth      =   4950
         TabIndex        =   36
         Top             =   2820
         Width           =   4950
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
         MaxLength       =   3
         TabIndex        =   14
         Top             =   930
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
         ScaleWidth      =   5895
         TabIndex        =   2
         Top             =   60
         Width           =   5895
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   9
            Left            =   3660
            TabIndex        =   12
            Top             =   540
            Value           =   1  'Checked
            Width           =   1410
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   8
            Left            =   690
            TabIndex        =   11
            Top             =   540
            Value           =   1  'Checked
            Width           =   2865
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   7
            Left            =   4050
            TabIndex        =   10
            Top             =   270
            Value           =   1  'Checked
            Width           =   750
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   6
            Left            =   2595
            TabIndex        =   9
            Top             =   270
            Value           =   1  'Checked
            Width           =   1380
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   5
            Left            =   1470
            TabIndex        =   8
            Top             =   270
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   4
            Left            =   690
            TabIndex        =   7
            Top             =   270
            Value           =   1  'Checked
            Width           =   705
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   3
            Left            =   4515
            TabIndex        =   6
            Top             =   675
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   2
            Left            =   3090
            TabIndex        =   5
            Top             =   675
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   1
            Left            =   3315
            TabIndex        =   4
            Top             =   0
            Value           =   1  'Checked
            Width           =   2610
         End
         Begin VB.CheckBox ckcMod 
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
            Height          =   210
            Index           =   0
            Left            =   690
            TabIndex        =   3
            Top             =   0
            Value           =   1  'Checked
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
         Left            =   2220
         TabIndex        =   15
         Top             =   1290
         Width           =   3090
      End
      Begin VB.DirListBox lbcMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1050
         Left            =   2235
         TabIndex        =   16
         Top             =   1650
         Width           =   3060
      End
      Begin VB.ListBox lbcModList 
         Appearance      =   0  'Flat
         Height          =   1500
         Left            =   60
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   2790
         Width           =   5235
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Unzip from subfolder of"
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
         TabIndex        =   37
         Top             =   1335
         Width           =   2085
      End
      Begin VB.Label lacPW 
         Appearance      =   0  'Flat
         Caption         =   "Zip Decrypt Password"
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
         TabIndex        =   13
         Top             =   960
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmcUnzip 
      Appearance      =   0  'Flat
      Caption         =   "&Unzip"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3525
      TabIndex        =   23
      Top             =   6810
      Width           =   1425
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5235
      TabIndex        =   24
      Top             =   6825
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
      Top             =   0
      Width           =   2160
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   8745
      OleObjectBlob   =   "CSIUNZIP.frx":0004
      Top             =   6705
   End
   Begin DUZACTXLibCtl.duzactxctrl zpcDUnZip 
      Left            =   7995
      OleObjectBlob   =   "CSIUNZIP.frx":00FA
      Top             =   6765
   End
   Begin VB.Label lacResults 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3210
      TabIndex        =   27
      Top             =   0
      Width           =   5805
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2160
      TabIndex        =   26
      Top             =   6540
      Width           =   5130
   End
End
Attribute VB_Name = "CSIUnzip"
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
Dim smPattern As String
Dim smPassword As String
Dim smExport As String
Dim smImport As String
Dim smSave As String
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
Dim smAffReports As String
Dim smSaveFilePath As String
Dim smTodaysDate As String
Dim imListField(1 To 4) As Integer
Dim smMessagesPath As String


Private Sub cbcMod_Change()
    
    lbcMod.Path = cbcMod.Drive
    cbcResultDir.Drive = cbcMod.Drive
    cbcSaveDir.Drive = cbcMod.Drive
    mSetDirectories
End Sub
Private Sub cbcResultDir_Change()
    lbcResultDir.Path = cbcResultDir.Drive
End Sub
Private Sub cbcSaveDir_Change()
    lbcSaveDir.Path = cbcSaveDir.Drive
End Sub
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilLoop As Integer
    ilValue = Value
    If (imSetAll) And (lbcModList.ListCount > 0) Then
        imAllClicked = True
        llRg = CLng(lbcModList.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcModList.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    pbcLbcModList_Paint
    cmcUnzip.Enabled = False
    For ilLoop = 0 To lbcModList.ListCount - 1 Step 1
        If lbcModList.Selected(ilLoop) Then
            cmcUnzip.Enabled = True
            Exit For
        End If
    Next ilLoop
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


Private Sub ckcMod_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcMod(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    lbcModList.Clear
    pbcLbcModList_Paint
    ReDim tmModName(0 To 0) As ZIPNAME
    smPattern = ""
    If ckcMod(0).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.E*"
        Else
            smPattern = smPattern & ";*.E*"
        End If
    End If
    If ckcMod(1).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.R*" & ";*.J*"
        Else
            smPattern = smPattern & ";*.R*" & ";*.J*"
        End If
    End If
    If ckcMod(2).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.A*"
        Else
            smPattern = smPattern & ";*.A*"
        End If
    End If
    If ckcMod(3).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.B*"
        Else
            smPattern = smPattern & ";*.B*"
        End If
    End If
    If ckcMod(4).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.N*"
        Else
            smPattern = smPattern & ";*.N*"
        End If
    End If
    If ckcMod(5).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.Z*"
        Else
            smPattern = smPattern & ";*.Z*"
        End If
    End If
    If ckcMod(6).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.G*"
        Else
            smPattern = smPattern & ";*.G*"
        End If
    End If
    If ckcMod(7).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.D*"
        Else
            smPattern = smPattern & ";*.D*"
        End If
    End If
    If ckcMod(8).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.S*"
        Else
            smPattern = smPattern & ";*.S*"
        End If
    End If
    If ckcMod(9).Value = vbChecked Then
        If Len(smPattern) = 0 Then
            smPattern = "*.F*"
        Else
            smPattern = smPattern & ";*.F*"
        End If
    End If
    lbcModFile.Pattern = smPattern
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
    Dim slStr As String
    smPassword = UCase$(Trim$(edcPW.Text))
    If left$(smPassword, 1) <> "V" Then
        MsgBox "Password form not correct"
        Exit Sub
    End If
    slStr = Mid$(smPassword, 2)
    If Len(slStr) <> 2 Then
        MsgBox "Password form not correct"
        Exit Sub
    End If
    smTodaysDate = Format$(Now, "m/d/yy")
    smResultFilePath = lbcResultDir.Path & "\"
    sgDBPath = smResultFilePath & "data\"
    If rbcSave(0).Value Then
        smSaveFilePath = lbcSaveDir.Path & "\"
        If StrComp(smResultFilePath, smSaveFilePath, 1) = 0 Then
            MsgBox "'Unzip into Subfolder' and 'Save Replaced Images' must be different folders"
            Exit Sub
        End If
    Else
        smSaveFilePath = ""
    End If
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
    Print #hmMsg, "CSI Unzipping Files " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
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
            ilPos = InStrRev(slName, ".")
            If ilPos > 0 Then
                slChar = Mid$(slName, ilPos + 1, 1)
                If StrComp(slChar, slPrevChar, 1) <> 0 Then
                    Select Case slChar
                        Case "E"
                            Print #hmMsg, "Exe/DLL/OCX/PDB/XML Unzipped"
                        Case "R", "J"
                            Print #hmMsg, "Reports (Rpt,bmp,jpg,gif) Unzipped"
                        Case "A"
                            Print #hmMsg, "Affiliate Exe Unzipped"
                        Case "B"
                            Print #hmMsg, "Affiliate Reports Unzipped"
                        Case "N"
                            Print #hmMsg, "DDFs Unzipped"
                        Case "Z"
                            Print #hmMsg, "Zero Btr Unzipped"
                        Case "G"
                            Print #hmMsg, "Generic Reports Unzipped"
                        Case "D"
                            Print #hmMsg, "Data Unzipped"
                    End Select
                    slPrevChar = slChar
                End If
                lacProcessing.Caption = slName
                ilRet = mUnZipFiles(slName)
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
        End If
    Next ilLoop
    Print #hmMsg, ""
    Print #hmMsg, "CSI Unzip Completed " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close #hmMsg
    plcGauge.Value = 100
    lacProcessing.Caption = "Processing- Done"
    imProcessing = False
    cmcCancel.Caption = "&Done"
    If ckcAll.Value = vbChecked Then
        ckcAll.Value = vbUnchecked
    End If
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
Private Sub lbcModList_Click()
    Dim ilLoop As Integer
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        pbcLbcModList_Paint
    End If
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
    imListField(2) = 20 * igAlignCharWidth 'igAvgCharWidth * 5
    imListField(3) = 32 * igAlignCharWidth 'igAvgCharWidth * 12
    imListField(4) = 60 * igAlignCharWidth 'igAvgCharWidth * 12

'    llValue = 1000
'    llRg = 0
'    llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    

    'ilRet = mGetIniInfo
    mParseCmmdLine
    CSIUnzip.Height = cmcUnzip.top + 5 * cmcUnzip.Height / 3
    'gCenterStdAlone BtrMatch
    CSIUnzip.Move (Screen.Width - CSIUnzip.Width) \ 2, (Screen.Height - CSIUnzip.Height) \ 2 ' + 175 '+ Screen.Height \ 10
    CSIUnzip.Show
    
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    smNowDate = Format$(Now, "m/d/yy")
    'smPattern = "*.E*;*.R*;*.A*;*.B*;*.N*;*.Z*;*.G*"
    smPattern = "*.E*;*.R*;*.J*;*.N*;*.Z*;*.G*;*.S*;*.F*"""
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
    
    
    slFilePath = smResultFilePath & "Data\Messages\"
    If fs.FolderExists(slFilePath) Then
        slToFile = slFilePath
    Else
        slToFile = smResultFilePath
    End If
    
    On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpStnFd.Txt"
    'slToFile = smResultFilePath & "CSIUnzip.Txt"
    'slToFile = smMessagesPath & "CSIUnzip.Txt"
    slToFile = slToFile & "CSIUnzip.Txt"
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
        'Open slToFile For Output As hmMsg
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
    smSave = ""
    ilPos = InStr(1, slCommand, "/O:", 1)
    If ilPos > 0 Then
        ilPos = ilPos + 3
        slChar = Trim$(Mid$(slCommand, ilPos, 1))
        Do While slChar = " "
            ilPos = ilPos + 1
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        Do While slChar <> " "
            smSave = smSave & slChar
            ilPos = ilPos + 1
            If ilPos > Len(slCommand) Then
                Exit Do
            End If
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        cbcSaveDir.Drive = left$(smSave, 1)
        lbcSaveDir.Path = smSave
        rbcSave(0).Value = True
    Else
        rbcSave(1).Value = True
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
    Unload CSIUnzip
    Set CSIUnzip = Nothing
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
    Dim slSubfolder As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slChar As String
    Dim ilCount As Integer
    Dim ilLoop As Integer
    Dim slFileName As String
    Dim slDate As String
    Dim slTime As String
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
    Dim ilSameTime As Integer
    'Build array of files to be unzipped
    ReDim tmUnZipName(1 To 1) As UNZIPNAME
    ilUpper = 0
    ilPos = InStrRev(slFromFile, ".")
    If ilPos > 0 Then
        slChar = Mid$(slFromFile, ilPos + 1, 1)
        Select Case slChar
            Case "E"
                slSubfolder = "exe"
            Case "R", "J"
                slSubfolder = "reports"
            Case "A"
                slSubfolder = "affexe"
            Case "B"
                ilRet = 0
                On Error GoTo mUnZipFilesErr:
                slSubfolder = "affrepor"
                lbcAffReports.Path = smResultFilePath & slSubfolder
                If ilRet <> 0 Then
                    slSubfolder = "affreports"
                End If
            Case "N"
                slSubfolder = "NewDDF"
            Case "Z"
                slSubfolder = "zero_btr"
            Case "G"
                slSubfolder = "reports\generic"
            Case "D"
                slSubfolder = "data"
            Case "S"
                slSubfolder = "Setup\ClientInstall\System32"
            Case "F"
                slSubfolder = "Setup\Fonts"
        End Select
        initUNZIPCmdStruct zpcDUnZip
        zpcDUnZip.ZIPFile = smModFilePath & slFromFile
        zpcDUnZip.QuietFlag = True
        zpcDUnZip.Filespec = "*.*"
        zpcDUnZip.NoDirectoryNamesFlag = True
        zpcDUnZip.ActionDZ = UNZIP_COUNTNAMEDZIPMEMBERS
        If zpcDUnZip.ErrorCode <> UE_OK Then
            lbcResult.AddItem "Not Valid Zip File " & Trim$(smModFilePath & slFromFile) & " Error #" & Str$(zpcDUnZip.ErrorCode)
            Print #hmMsg, "     " & "Not Valid Zip File " & Trim$(smModFilePath & slFromFile) & " Error #" & Str$(zpcDUnZip.ErrorCode)
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
                slDateTime = FileDateTime(smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName))
                If ilRet = 0 Then
                    ilFileExist = True
                    slDate = gAdjYear(Format$(slDateTime, "m/d/yy"))
                    slTime = Format$(slDateTime, "h:mm:ssAM/PM")
                    llSize = FileLen(smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName))
                    tmUnZipName(ilUpper).sOldDate = slDate
                    tmUnZipName(ilUpper).sOldTime = slTime
                    tmUnZipName(ilUpper).lOldSize = llSize
                    slFTime = Format$(slTime, "hh:mm:ss")
                    'ignore seconds as zip does not return the seconds
                    llFTime = 3600 * Hour(slFTime) + 60 * Minute(slFTime) '+ Second(slFTime)
                    slUTime = Format$(tmUnZipName(ilUpper).sTime, "hh:mm:ss")
                    llUTime = 3600 * Hour(slUTime) + 60 * Minute(slUTime) '+ Second(slUTime)
                    'If DateValue(tmUnZipName(ilUpper).sDate) > DateValue(slDate) Then
                    '    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '    ilAddFile = True
                    'ElseIf DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) Then
                    '    If llUTime >= llFTime Then
                    '        'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '        ilAddFile = True
                    '    Else
                    '        'ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
                    '        'If ilRet = 6 Then
                    '        '    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '        '    ilAddFile = True
                    '        'End If
                    '        If llSize <> tmUnZipName(ilUpper).lSize Then
                    '            ilOvMsg = True
                    '        Else
                    '            ilOvMsg = True
                    '        End If
                    '    End If
                    'Else
                    '    'ilRet = MsgBox("Overwrite: " & Trim$(tmUnZipName(ilUpper).sName) & Chr(10) & Chr(13) & "     " & slDate & " " & slTime & Str$(llSize) & " with " & Chr(10) & Chr(13) & "     " & Format$(tmUnZipName(ilUpper).sDate, "m/d/yy") & " " & Format$(tmUnZipName(ilUpper).sTime, "h:mm:ssAM/PM") & Str$(tmUnZipName(ilUpper).lSize), 4, "Overwrite")
                    '    'If ilRet = 6 Then
                    '    '    'ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                    '    '    ilAddFile = True
                    '    'End If
                    '    ilOvMsg = True
                    'End If
                    If DateValue(tmUnZipName(ilUpper).sDate) > DateValue(slDate) + 1 Then
                        ilAddFile = True
                    ElseIf DateValue(tmUnZipName(ilUpper).sDate) < DateValue(slDate) - 1 Then
                        ilOvMsg = True
                    Else
                        ilSameTime = False
                        If (llUTime = llFTime) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                            ilSameTime = True
                        Else
                            If Minute(slUTime) = Minute(slFTime) Then
                                If Hour(slUTime) = 0 Then
                                    If (Hour(slFTime) >= 21) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) + 1) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 3) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 1 Then
                                    If (Hour(slFTime) >= 22) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) + 1) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 4) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 2 Then
                                    If (Hour(slFTime) >= 23) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) + 1) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 5) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 23 Then
                                    If (Hour(slFTime) >= 20) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 2) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) - 1) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 22 Then
                                    If (Hour(slFTime) >= 19) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 1) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) - 1) Then
                                        ilSameTime = True
                                    End If
                                ElseIf Hour(slUTime) = 21 Then
                                    If (Hour(slFTime) >= 18) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                    If (Hour(slFTime) <= 0) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate) - 1) Then
                                        ilSameTime = True
                                    End If
                                Else
                                    If (Hour(slFTime) >= Hour(slUTime) - 3) And (Hour(slFTime) <= Hour(slUTime) + 3) And (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) Then
                                        ilSameTime = True
                                    End If
                                End If
                            End If
                        End If
                        If ilSameTime And (llSize = tmUnZipName(ilUpper).lSize) Then
                            ilAddFile = True
                        Else
                            If DateValue(tmUnZipName(ilUpper).sDate) > DateValue(slDate) And (Not ilSameTime) Then
                                ilAddFile = True
                            Else
                                If (DateValue(tmUnZipName(ilUpper).sDate) = DateValue(slDate)) And (llUTime > llFTime) And (Not ilSameTime) Then
                                    ilAddFile = True
                                Else
                                    ilOvMsg = True
                                End If
                            End If
                        End If
                    End If
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
                        ilRet = GetAttr(smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName)) And 1
                        If ilRet <> 0 Then
                            SetAttr smResultFilePath & slSubfolder & "\" & Trim$(tmUnZipName(ilUpper).sZipName), 0
                        End If
                    End If
                    ReDim Preserve tmUnZipName(1 To ilUpper + 1) As UNZIPNAME
                End If
            Next ilLoop
            
    
            For ilLoop = 1 To UBound(tmUnZipName) - 1 Step 1
                If (rbcSave(0).Value) And (tmUnZipName(ilLoop).lOldSize > 0) Then
                    ilRet = mZipFile(slFromFile, smResultFilePath & slSubfolder & "\" & tmUnZipName(ilLoop).sZipName, tmUnZipName(ilLoop).sOldDate, tmUnZipName(ilLoop).sOldTime, tmUnZipName(ilLoop).lOldSize)
                End If
                initUNZIPCmdStruct zpcDUnZip
                zpcDUnZip.ZIPFile = smModFilePath & slFromFile
                zpcDUnZip.Destination = smResultFilePath & slSubfolder
                zpcDUnZip.NoDirectoryNamesFlag = True
                zpcDUnZip.Filespec = Trim$(tmUnZipName(ilLoop).sName)
                zpcDUnZip.OverwriteFlag = True
                zpcDUnZip.DiagnosticFlag = True
                zpcDUnZip.DecryptFlag = True
                zpcDUnZip.DecryptCode = smPassword
                zpcDUnZip.ActionDZ = UNZIP_EXTRACT
                ilRet = zpcDUnZip.ErrorCode
                If ilRet = 0 Then
                    lbcResult.AddItem Trim$(tmUnZipName(ilLoop).sZipName) & " Unzipped into " & smResultFilePath & slSubfolder
                    If Trim$(tmUnZipName(ilLoop).sOldDate) = "" Then
                        Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sZipName) & " " & Trim$(tmUnZipName(ilLoop).sDate) & " " & Trim$(tmUnZipName(ilLoop).sTime) & " Size" & Str$(tmUnZipName(ilLoop).lSize) & " Unzipped into " & smResultFilePath & slSubfolder
                    Else
                        Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sZipName) & " " & Trim$(tmUnZipName(ilLoop).sDate) & " " & Trim$(tmUnZipName(ilLoop).sTime) & " Size" & Str$(tmUnZipName(ilLoop).lSize) & " Unzipped into " & smResultFilePath & slSubfolder & " Replaced: " & Trim$(tmUnZipName(ilLoop).sOldDate) & " " & Trim$(tmUnZipName(ilLoop).sOldTime) & " Size" & Str$(tmUnZipName(ilLoop).lOldSize)
                    End If
                Else
                    lbcResult.AddItem Trim$(tmUnZipName(ilLoop).sName) & " Error #" & Str$(ilRet)
                    Print #hmMsg, "     " & Trim$(tmUnZipName(ilLoop).sName) & " Error #" & Str$(ilRet)
                End If
            Next ilLoop
        End If
    End If
    Exit Function
mUnZipFilesErr:
    ilRet = Err
    Resume Next
End Function
Private Function mZipFile(slZipName As String, slInName As String, slInDate As String, slInTime As String, llSize As Long) As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim slName As String
    Dim slZipPathName As String
    Dim ilPos As Integer
    Dim ilNameFd As Integer
    Dim ilLoop As Integer
    initZIPCmdStruct zpcDZip
    On Error GoTo mZipFilesErr:
    slZipPathName = smSaveFilePath & Trim$(slZipName)
    ilNameFd = False
    For ilLoop = LBound(tmZippedNames) To UBound(tmZippedNames) - 1 Step 1
        If StrComp(Trim$(slZipName), Trim$(tmZippedNames(ilLoop).sName), 1) = 0 Then
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
        tmZippedNames(UBound(tmZippedNames)).sName = Trim$(slZipName)
        ReDim Preserve tmZippedNames(0 To UBound(tmZippedNames) + 1) As ZIPPEDNAMES
    End If
    On Error GoTo 0
    zpcDZip.ZIPFile = Trim$(slZipPathName)
    zpcDZip.ItemList = Trim$(slInName)
    zpcDZip.OldAsLatestFlag = True
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
        Print #hmMsg, "     " & Trim$(slInName) & " " & Trim$(slInDate) & " " & Trim$(slInTime) & " Size" & Str$(llSize) & " saved into " & Trim$(slZipPathName)
        'lbcResult.AddItem Trim$(slInName) & " zipped into " & slName
    Else
        Print #hmMsg, "     " & Trim$(slInName) & " saved into " & Trim$(slZipPathName) & " Failed: " & Str$(ilRet)
        'lbcResult.AddItem Trim$(slInName) & " not zipped"
    End If
    Exit Function
mZipFilesErr:
    ilRet = 1
    Resume Next
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

Private Sub rbcSave_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSave(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            cbcSaveDir.Enabled = True
            lbcSaveDir.Enabled = True
        Else
            cbcSaveDir.Enabled = False
            lbcSaveDir.Enabled = False
        End If
    End If
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
    smPassword = UCase$(Trim$(edcPW.Text))
    ReDim tmModName(0 To 0) As ZIPNAME
    For ilLoop = 0 To lbcModFile.ListCount - 1 Step 1
        slName = lbcModFile.List(ilLoop)
        ilShow = True
        ilPos = InStrRev(slName, ".")
        If ilPos > 0 Then
            slStr = Mid$(slName, ilPos + 1)
            If StrComp(slStr, "Exe", 1) = 0 Then
                ilShow = False
            End If
            If StrComp(slStr, "Zip", 1) = 0 Then
                ilShow = False
            End If
        Else
            ilShow = False
        End If
        If ilShow Then
            'initUNZIPCmdStruct zpcDUnZip
            'zpcDUnZip.ZIPFile = smModFilePath & slName
            'zpcDUnZip.QuietFlag = True
            'zpcDUnZip.Filespec = "*.*"
            'zpcDUnZip.Action = UNZIP_COUNTNAMEDZIPMEMBERS
            'If zpcDUnZip.ErrorCode <> UE_OK Then
            '    ilShow = False
            'Else
                If smPassword <> "" Then
                    ilPos = InStrRev(slName, ".")
                    If ilPos > 0 Then
                        slStr = Mid$(slName, ilPos + 2)
                        If InStr(1, smPassword, slStr, 1) = 0 Then
                            ilShow = False
                        End If
                    Else
                        ilShow = False
                    End If
                End If
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
            ilPos = InStrRev(slName, ".")
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
Private Sub plcSave_Paint()
    plcSave.CurrentX = 0
    plcSave.CurrentY = 0
    plcSave.Print "Save old images"
End Sub
Private Sub plcMod_Paint()
    plcMod.CurrentX = 0
    plcMod.CurrentY = 0
    plcMod.Print "Unzip"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Unzip Counterpoint Files"
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
        lbcSaveDir.Path = lbcMod.Path
    End If
    If ilRet = 0 Then
        lbcResultDir.Path = lbcMod.Path
        lbcSaveDir.Path = lbcMod.Path & "\updates\old"
        lbcMod.Path = lbcMod.Path & "\Updates"
    End If
    Exit Sub
TryProd:
    ilRet = 1
    Resume Next
End Sub
