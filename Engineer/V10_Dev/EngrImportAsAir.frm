VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form EngrImportAsAir 
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "EngrImportAsAir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckcAllLogs 
      Caption         =   "All Logs"
      Height          =   225
      Left            =   315
      TabIndex        =   7
      Top             =   2100
      Width           =   1050
   End
   Begin VB.DriveListBox cbcLogDrive 
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
      Left            =   4260
      TabIndex        =   6
      Top             =   405
      Width           =   3270
   End
   Begin VB.DirListBox lbcLogPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4260
      TabIndex        =   5
      Top             =   765
      Width           =   3240
   End
   Begin VB.FileListBox lbcLogFile 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   300
      MultiSelect     =   2  'Extended
      Pattern         =   "*.Log"
      TabIndex        =   4
      Top             =   405
      Width           =   3540
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   3165
      Width           =   1245
   End
   Begin VB.CommandButton cmcImport 
      Caption         =   "Import"
      Height          =   315
      Left            =   2205
      TabIndex        =   0
      Top             =   3165
      Width           =   1245
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7695
      Top             =   2355
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   3555
      FormDesignWidth =   7815
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   255
      Left            =   1020
      TabIndex        =   8
      Top             =   2850
      Visible         =   0   'False
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacScreen 
      Caption         =   "Import Engineering 'as Aired' Log"
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2595
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   5625
   End
End
Attribute VB_Name = "EngrImportAsAir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrImportAsAir - displays import csv information
'*
'*  Created Aug,1998 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private hmSEE As Integer

Private imTerminate As Integer
Private imImporting As Integer
Private hmFrom As Integer
Private hmMsg As Integer
Private smMsgFile As String
Private smNowDate As String
Private smNowTime As String
Private lmTotalNoBytes As Long
Private lmProcessedNoBytes As Long
Private smCurDir As String
Private lmFloodPercent As Long
Private imImportSelection As Integer
Private imFileNameListBoxIgnore As Integer
Private smFileNames As String
Private smRenameFile() As String

















Private Sub cbcLogDrive_Change()
    Screen.MousePointer = vbHourglass
    lbcLogPath.Path = cbcLogDrive.Drive
    Screen.MousePointer = vbDefault
End Sub

Private Sub ckcAllLogs_Click()
    Dim iValue As Integer
    Dim lRg As Long
    Dim lRet As Long
    'All check box has been selected or deselected
    If imFileNameListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllLogs.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    
    If lbcLogFile.ListCount > 0 Then        'if at least 1 audio type, set the off or on
        imFileNameListBoxIgnore = True
        lRg = CLng(lbcLogFile.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcLogFile.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imFileNameListBoxIgnore = False
    End If

End Sub

Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload EngrImportAsAir
End Sub

Private Sub cmcImport_Click()
    Dim ilLoop As Integer
    Dim slDrivePath As String
    Dim slName As String
    Dim llFileCount As Long
    Dim llTotalCount As Long
    Dim llPercent As Long
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slDateTime As String
    
    Screen.MousePointer = vbHourglass
    imImporting = False
    slDrivePath = lbcLogPath.Path
    llFileCount = 0
    plcGauge.Value = 0
    plcGauge.Visible = True
    If Right(slDrivePath, 1) <> "\" Then
        slDrivePath = slDrivePath & "\"
    End If
    For ilLoop = 0 To lbcLogFile.ListCount - 1 Step 1
        If lbcLogFile.Selected(ilLoop) Then
            llTotalCount = llTotalCount + 1
        End If
    Next ilLoop
    ReDim smRenameFile(0 To 0) As String
    For ilLoop = 0 To lbcLogFile.ListCount - 1 Step 1
        If lbcLogFile.Selected(ilLoop) Then
            slName = Trim$(lbcLogFile.List(ilLoop))
            lacMsg.Caption = "Processing " & slName
            ilRet = gLoadAsAirLog(slDrivePath & slName, sgAsAirLogDate, hmSEE)
            If ilRet Then
                smRenameFile(UBound(smRenameFile)) = slDrivePath & slName
                ReDim Preserve smRenameFile(0 To UBound(smRenameFile) + 1) As String
            End If
            llFileCount = llFileCount + 1
            llPercent = (llFileCount * CSng(100)) / llTotalCount
            If llPercent >= 100 Then
                llPercent = 100
            End If
            If plcGauge.Value <> llPercent Then
                plcGauge.Value = llPercent
            End If
            If imTerminate Then
                Exit For
            End If
        End If
    Next ilLoop
    On Error GoTo cmcImportErr:
    For ilLoop = 0 To UBound(smRenameFile) - 1 Step 1
        slName = smRenameFile(ilLoop)
        ilPos = InStr(1, slName, ".", vbTextCompare)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos) & "Old"
        End If
        ilRet = 0
        slDateTime = FileDateTime(slName)
        If ilRet = 0 Then
            Kill slName
        End If
        Name smRenameFile(ilLoop) As slName
    Next ilLoop
    On Error GoTo 0
    imImporting = False
    If Not imTerminate Then
        plcGauge.Value = 100
        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
    Else
        Unload EngrImportAsAir
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
cmcImportErr:
    ilRet = Err.Number
    Resume Next
End Sub



Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrImportAsAir
    gCenterFormModal EngrImportAsAir
End Sub

Private Sub Form_Load()
    Dim iUpper As Integer
    
    smCurDir = CurDir
    mInit
    igJobShowing(SCHEDULEJOB) = 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    btrDestroy hmSEE
    
    Erase smRenameFile
    
    If InStr(1, smCurDir, ":") > 0 Then
        ChDrive Left$(smCurDir, 1)
        ChDir smCurDir
    End If
    Set EngrImportAsAir = Nothing
    EngrSchd.Show vbModeless
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    Dim slFullPathName As String
    Dim slDrive As String
    Dim slPath As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slDate As String
    Dim ilPos As Integer
    
    Screen.MousePointer = vbHourglass
    imImporting = False
    imTerminate = False
    imFileNameListBoxIgnore = False
    smFileNames = ""
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ilRet = gGetTypeOfRecs_APE_AutoPath("C", sgCurrAPEStamp, "EngrImportAsAir-mInit", tgCurrAPE())
    For ilLoop = 0 To UBound(tgCurrAPE) - 1 Step 1
        If (tgCurrAPE(ilLoop).sType = "CI") Then
            smFileNames = Trim$(tgCurrAPE(ilLoop).sNewFileName) & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            ilPos = InStr(1, smFileNames, "Date", vbTextCompare)
            If ilPos > 0 Then
                If Trim$(tgCurrAPE(ilLoop).sDateFormat) <> "" Then
                    slDate = Format$(sgAsAirLogDate, Trim$(tgCurrAPE(ilLoop).sDateFormat))
                Else
                    slDate = Format$(sgAsAirLogDate, "yymmdd")
                End If
                smFileNames = Left$(smFileNames, ilPos - 1) & slDate & Mid(smFileNames, ilPos + 4)
            End If
            slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            If slPath <> "" Then
                If Right(slPath, 1) <> "\" Then
                    slPath = slPath & "\"
                End If
            End If
            lbcLogFile.Pattern = "*" & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            Exit For
        End If
    Next ilLoop
    ilPos = InStr(slPath, ":")
    If ilPos > 0 Then
        slDrive = Left$(slPath, ilPos)
        slPath = Mid$(slPath, ilPos + 1)
        If Right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        End If
        cbcLogDrive.Drive = slDrive
        lbcLogPath.Path = slPath
        slStr = lbcLogPath.Path
        If Right$(slStr, 1) <> "\" Then
            slStr = slStr & "\"
        End If
        lbcLogFile.fileName = slStr & smFileNames
    ElseIf Left(slPath, 2) = "\\" Then
        ilPos = InStr(3, slPath, "\", vbTextCompare)
        ilPos = InStr(ilPos + 1, slPath, "\", vbTextCompare)
        slDrive = Left$(slPath, ilPos - 1)
        'slPath = Mid$(slPath, ilPos + 1)
        If Right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        End If
        cbcLogDrive.Drive = slDrive
        lbcLogPath.Path = slPath
        slStr = lbcLogPath.Path
        If Right$(slStr, 1) <> "\" Then
            slStr = slStr & "\"
        End If
        lbcLogFile.fileName = slStr & smFileNames
    End If
    If lbcLogFile.ListCount > 0 Then
        ckcAllLogs.Value = vbChecked
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub lbcLogFile_Click()
    If imFileNameListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllLogs.Value = vbChecked Then
        imFileNameListBoxIgnore = True
        ckcAllLogs.Value = False
        imFileNameListBoxIgnore = False
    End If

End Sub

Private Sub lbcLogPath_Change()
    Dim slStr As String
    
    ckcAllLogs.Value = vbUnchecked
    slStr = lbcLogPath.Path
    If Right$(slStr, 1) <> "\" Then
        slStr = slStr & "\"
    End If
    lbcLogFile.fileName = slStr & smFileNames
End Sub
