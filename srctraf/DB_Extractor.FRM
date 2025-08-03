VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DB_Extractor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Extractor Utility"
   ClientHeight    =   6075
   ClientLeft      =   4245
   ClientTop       =   3630
   ClientWidth     =   7560
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
   Icon            =   "DB_Extractor.frx":0000
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
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
      Height          =   6075
      Left            =   0
      Picture         =   "DB_Extractor.frx":030A
      ScaleHeight     =   6015
      ScaleWidth      =   7485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   7545
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   150
         ScaleHeight     =   30
         ScaleWidth      =   45
         TabIndex        =   14
         Top             =   5745
         Width           =   45
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   105
         ScaleHeight     =   30
         ScaleWidth      =   15
         TabIndex        =   9
         Top             =   5475
         Width           =   15
      End
      Begin VB.ListBox lbcAirVehicle 
         Appearance      =   0  'Flat
         Height          =   240
         ItemData        =   "DB_Extractor.frx":9C57
         Left            =   2745
         List            =   "DB_Extractor.frx":9C59
         Sorted          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton cmcDropDown 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5400
         Picture         =   "DB_Extractor.frx":9C5B
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox edcDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4380
         MaxLength       =   20
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox pbcToggle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1440
         ScaleHeight     =   210
         ScaleWidth      =   795
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3225
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdVehicle 
         Height          =   2940
         Left            =   135
         TabIndex        =   8
         Top             =   1605
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   5186
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.PictureBox plcFrom 
         Height          =   375
         Left            =   1020
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   3
         Top             =   825
         Width           =   4305
         Begin VB.TextBox edcFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   15
            TabIndex        =   4
            Top             =   30
            Width           =   4230
         End
      End
      Begin VB.CommandButton cmcFrom 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Left            =   5715
         TabIndex        =   5
         Top             =   915
         Width           =   1485
      End
      Begin VB.Timer tmcClick 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1680
         Top             =   5085
      End
      Begin VB.CheckBox ckcDormant 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   1305
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   3495
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   105
      End
      Begin MSComDlg.CommonDialog CMDialogBox 
         Left            =   135
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   4100
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   4635
         Width           =   1875
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1950
         TabIndex        =   18
         Top             =   4575
         Width           =   3615
      End
      Begin VB.Label lacDormant 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Include Dormant"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   480
         TabIndex        =   7
         Top             =   1305
         Width           =   1635
      End
      Begin VB.Label lacFrom 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   2
         Top             =   915
         Width           =   810
      End
      Begin VB.Label lacStop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
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
         Left            =   3075
         TabIndex        =   17
         Top             =   5355
         Width           =   1335
      End
      Begin VB.Label lacMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize"
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
         Left            =   3075
         TabIndex        =   16
         Top             =   4845
         Width           =   1350
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   75
         Width           =   3210
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   4785
         TabIndex        =   15
         Top             =   4875
         Visible         =   0   'False
         Width           =   1410
      End
   End
End
Attribute VB_Name = "DB_Extractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: DB_Extractor.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text

Private Const LB_GETITEMHEIGHT = &H1A1

'Dim hmMsg As Integer

Dim imFirstActivate As Integer

Dim imDeleteForm As Integer '0=Use Extended calls; 1= Use GetFirst, Get Next

Dim imLastColSorted As Integer
Dim imLastSort As Integer

Dim hmFrom As Integer
Dim hmVehicles As Integer

Dim smPreselectNames() As String

Dim tmRecBuff As POPREC

Dim tmIntSrchKey As INTKEY0
Dim tmLongSrchKey As LONGKEY0

Dim tmVef As VEF                'Vehicle record image
Dim imVefRecLen As Integer      'Vehicle record length
Dim tmClf As CLF                'Vehicle record image
Dim imClfRecLen As Integer      'Vehicle record length
Dim hmClf As Integer
Dim tmClfSrchKey0 As CLFKEY0
Dim tmSbf As SBF                'Vehicle record image
Dim imSbfRecLen As Integer      'Vehicle record length
Dim hmSbf As Integer
Dim tmSbfSrchKey0 As SBFKEY0
Dim tmRif As RIF                'Vehicle record image
Dim tmDnf As DNF
Dim tmDrf As DRF                'Vehicle record image
Dim imDrfRecLen As Integer      'Vehicle record length
Dim hmDrf As Integer
Dim tmDrfSrchKey0 As DRFKEY0
Dim tmCrf As CRF                'Vehicle record image
Dim tmRvf As RVF                'Vehicle record image
Dim tmPhf As RVF                'Vehicle record image
Dim tmChf As CHF                'Vehicle record image
Dim tmSdf As SDF
Dim tmSmf As SMF                'Vehicle record image
Dim imSmfRecLen As Integer      'Vehicle record length
Dim hmSmf As Integer
Dim tmSmfSrchKey2 As LONGKEY0


Dim tmCff As CFF                'Vehicle record image
Dim imCffRecLen As Integer      'Vehicle record length
Dim hmCff As Integer
Dim tmCffSrchKey0 As CFFKEY0

Dim tmVpf As VPF                'Vehicle record image
Dim imVpfRecLen As Integer      'Vehicle record length
Dim hmVpf As Integer
Dim tmVpfSrchKey0 As INTKEY0

Dim tmVlf As VLF                'Vehicle record image
Dim imVlfRecLen As Integer      'Vehicle record length
Dim hmVlf As Integer
Dim tmVlfSrchKey0 As VLFKEY0

Dim tmSpf As SPF                'Vehicle record image
Dim imSpfRecLen As Integer      'Vehicle record length
Dim hmSpf As Integer

Dim hmRhf As Integer        'GIF Handle
Dim imRhfRecLen As Integer     'GIF record length
Dim tmRhf As RHF            'GIF record image

Dim hmAdf As Integer        'GIF Handle
Dim imAdfRecLen As Integer     'GIF record length
Dim tmAdf As ADF            'GIF record image

Dim hmAgf As Integer        'GIF Handle
Dim imAgfRecLen As Integer     'GIF record length
Dim tmAgf As AGF            'GIF record image

Dim smFromDBPath As String

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imConverting As Integer
Dim lmNowDate As Long
Dim imAutoRun As Integer            '1 = auto run flag, 0 = manual
Dim imClosed As Integer
Dim imFileNo As Integer
Dim imTotalFiles As Integer
Dim lmTotalRecords As Long
Dim lmRecordNo As Long
Dim lmAdjCount As Long

Dim tmRecreateSort() As RECREATESORT

Dim lmDeleteCode() As Long
Dim lmUpdateCode() As Long

Dim imLbcArrowSetting As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imChgMode As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim lmTopRow As Long
Dim bmIgnoreScroll As Boolean
Dim bmCtrlVisible As Boolean

Dim imPassNo As Integer

Const NAMEINDEX = 0
Const STATUSINDEX = 1
Const TYPEINDEX = 2
Const SORTINDEX = 3
Const STATEINDEX = 4
Const VEFCODEINDEX = 5



Private Sub cmcDropDown_Click()
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
    End Select

End Sub

Private Sub cmcFrom_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    igPathType = 0
    sgGetPath = edcFrom.Text
    lgCallTop = DB_Extractor.Top
    lgCallLeft = DB_Extractor.Left
    lgCallHeight = DB_Extractor.Height
    lgCallWidth = DB_Extractor.Width
    GetPath.Show vbModal
    If igGetPath = 0 Then
        edcFrom.Text = sgGetPath
    End If
    
    ChDir slCurDir
    
    Exit Sub
End Sub

Private Sub edcDropDown_Change()
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
    End Select

End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdVehicle.Col
            Case NAMEINDEX
            Case STATUSINDEX
            Case TYPEINDEX
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case grdVehicle.Col
            Case NAMEINDEX
            Case STATUSINDEX
            Case TYPEINDEX
        End Select
    End If
End Sub

Private Sub edcFrom_Change()
    mSetCommands
End Sub

Private Sub edcFrom_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub Form_Activate()
'    gShowBranner
End Sub

Private Sub Form_GotFocus()
    mSetShow
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        If imAutoRun = 0 Then
            MsgBox "Only one copy of DB_Extractor can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        End If
        gLogMsg "Second copy of DB_Extractor path: " & App.Path & " from " & Trim$(gGetComputerName()), "DB_Extractor_Log.txt", False
        End
    End If
    'gSetBkgdMode
    igBkgdProg = 11
    mInit
    If imAutoRun = 0 Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMinimized
    End If
    tmcClick.Interval = 2000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    If imClosed = True Then
        Exit Sub
    End If
    Cancel = 0
    tmcClick.Enabled = False
    
    'If imAutoRun = 0 Then
    '    ilRet = MsgBox("Stop the Export Projection", vbQuestion + vbYesNo, "Stop Service")
    '    If ilRet = vbNo Then
    '        Cancel = 1
    '        imTerminate = False
    '        Exit Sub
    '    End If
    'End If
    imClosed = True
    imTerminate = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    Erase tmRecreateSort
    Erase lmDeleteCode
    Erase lmUpdateCode
    
    btrStopAppl
    Set DB_Extractor = Nothing   'Remove data segment
    End
End Sub


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
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = "debug"
    igDirectCall = 0
    sgIniPath = ""
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 8))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    'imAutoRun = 1
    'ilPos = InStr(1, slCommand, "/UserInput", 1)
    'If ilPos > 0 Then
        imAutoRun = 0
        igBkgdProg = 0
        gLogMsg "Current Mode: UserInput.", "DB_Extractor_Log.txt", False
    'Else
    '    igBkgdProg = 8
    '    gLogMsg "Current Mode: Background.", "DB_Extractor.txt", False
    'End If
    gInitStdAlone DB_Extractor, slStr, ilTestSystem
End Sub



Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Vef.Btr", Len(tmVef)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

Private Sub grdVehicle_EnterCell()
    mSetShow
End Sub

Private Sub grdVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdVehicle.TopRow
    grdVehicle.Redraw = False
End Sub

Private Sub grdVehicle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    bmIgnoreScroll = False
    If Y < grdVehicle.RowHeight(0) Then
        grdVehicle.Col = grdVehicle.MouseCol
        mSortCol grdVehicle.Col
        Exit Sub
    End If
    ilCol = grdVehicle.MouseCol
    ilRow = grdVehicle.MouseRow
    If ilCol < grdVehicle.FixedCols Then
        grdVehicle.Redraw = True
        Exit Sub
    End If
    If ilRow < grdVehicle.FixedRows Then
        grdVehicle.Redraw = True
        Exit Sub
    End If
    If grdVehicle.TextMatrix(ilRow, NAMEINDEX) = "" Then
        grdVehicle.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdVehicle.TextMatrix(ilRow, NAMEINDEX) = ""
        grdVehicle.Row = ilRow + 1
        grdVehicle.Col = STATUSINDEX
        grdVehicle.Redraw = True
    Else
        grdVehicle.Row = ilRow
        grdVehicle.Col = ilCol
    End If
    grdVehicle.Redraw = True
    lmTopRow = grdVehicle.TopRow
    If Not mColOk() Then
        Exit Sub
    End If
    mEnableBox
End Sub

Private Sub grdVehicle_Scroll()
    If bmIgnoreScroll Then  'Or igGridIgnoreScroll Then
        bmIgnoreScroll = False
        Exit Sub
    End If
    If grdVehicle.Redraw = False Then
        grdVehicle.Redraw = True
        If lmTopRow < grdVehicle.FixedRows Then
            grdVehicle.TopRow = grdVehicle.FixedRows
        Else
            grdVehicle.TopRow = lmTopRow
        End If
        grdVehicle.Refresh
        grdVehicle.Redraw = False
    End If
    If (bmCtrlVisible) And (grdVehicle.Row >= grdVehicle.FixedRows) And (grdVehicle.Col >= 0) And (grdVehicle.Col < grdVehicle.Cols - 1) Then
        If grdVehicle.RowIsVisible(grdVehicle.Row) Then
            'pbcArrow.Move grdVehicle.Left - pbcArrow.Width - 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + (grdVehicle.RowHeight(grdVehicle.Row) - pbcArrow.Height) / 2
            'pbcArrow.Visible = True
            mSetFocus
        Else
            'pbcSetFocus.SetFocus
            pbcToggle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            'pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub lacDormant_Click()
    If ckcDormant.Value = vbChecked Then
        ckcDormant.Value = vbUnchecked
    Else
        ckcDormant.Value = vbChecked
    End If
End Sub

Private Sub lacMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacMin.Left - 30, lacMin.Top - 45
    imcOutline.Visible = True
End Sub


Private Sub lacStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacStop.Left - 30, lacStop.Top - 45
    imcOutline.Visible = True
End Sub


Private Sub pbcClickFocus_Click()
    imcOutline.Visible = False
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If bmCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdVehicle.Col = NAMEINDEX Then
                If grdVehicle.Row > grdVehicle.FixedRows Then
                    lmTopRow = -1
                    grdVehicle.Row = grdVehicle.Row - 1
                    If Not grdVehicle.RowIsVisible(grdVehicle.Row) Then
                        grdVehicle.TopRow = grdVehicle.TopRow - 1
                    End If
                    grdVehicle.Col = TYPEINDEX
                    mEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdVehicle.Col = grdVehicle.Col - 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdVehicle.TopRow = grdVehicle.FixedRows
        grdVehicle.Col = STATUSINDEX
        grdVehicle.Row = grdVehicle.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If bmCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdVehicle.Col = TYPEINDEX Then
                llRow = grdVehicle.Rows
                Do
                    llRow = llRow - 1
                Loop While grdVehicle.TextMatrix(llRow, NAMEINDEX) = ""
                llRow = llRow + 1
                If (grdVehicle.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdVehicle.Row = grdVehicle.Row + 1
                    If Not grdVehicle.RowIsVisible(grdVehicle.Row) Then 'Or (grdVehicle.Row - (grdVehicle.TopRow - grdVehicle.FixedRows) >= imInitNoRows) Then
                        bmIgnoreScroll = True
                        grdVehicle.TopRow = grdVehicle.TopRow + 1
                    End If
                    grdVehicle.Col = STATUSINDEX
                    'grdVehicle.TextMatrix(grdVehicle.Row, CODEINDEX) = 0
                    If Trim$(grdVehicle.TextMatrix(grdVehicle.Row, NAMEINDEX)) <> "" Then
                        If mColOk() Then
                            mEnableBox
                        Else
                            pbcClickFocus.SetFocus
                        End If
                    Else
                        pbcClickFocus.SetFocus
                    End If
                Else
                    If Trim$(grdVehicle.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdVehicle.Row + 1 >= grdVehicle.Rows Then
                            pbcClickFocus.SetFocus
                            Exit Sub
                        End If
                        grdVehicle.Row = grdVehicle.Row + 1
                        If (Not grdVehicle.RowIsVisible(grdVehicle.Row)) Then   'Or (grdVehicle.Row - (grdVehicle.TopRow - grdVehicle.FixedRows) >= imInitNoRows) Then
                            bmIgnoreScroll = True
                            grdVehicle.TopRow = grdVehicle.TopRow + 1
                        End If
                        grdVehicle.Col = STATUSINDEX
                        If mColOk() Then
                            mEnableBox
                        Else
                            pbcClickFocus.SetFocus
                        End If
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdVehicle.Col = grdVehicle.Col + 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdVehicle.TopRow = grdVehicle.FixedRows
        grdVehicle.Col = STATUSINDEX
        grdVehicle.Row = grdVehicle.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub pbcToggle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcToggle_KeyPress(KeyAscii As Integer)
    If lmEnableCol = STATUSINDEX Then
        If KeyAscii = Asc("I") Or (KeyAscii = Asc("i")) Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Include"
            pbcToggle_Paint
        ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Exclude"
            pbcToggle_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Include" Then
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Exclude"
                pbcToggle_Paint
            ElseIf grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Exclude" Then
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Include"
                pbcToggle_Paint
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcToggle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lmEnableCol = STATUSINDEX Then
        If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Include" Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Exclude"
            pbcToggle_Paint
        ElseIf grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Exclude" Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Include"
            pbcToggle_Paint
        End If
    End If
    pbcToggle_Paint
    mSetCommands
End Sub

Private Sub pbcToggle_Paint()
    pbcToggle.Cls
    pbcToggle.CurrentX = fgBoxInsetX
    pbcToggle.CurrentY = 0 'fgBoxInsetY
    pbcToggle.Print grdVehicle.TextMatrix(lmEnableRow, lmEnableCol)
End Sub

Private Sub plcSignon_Click()
    imcOutline.Visible = False
End Sub


Private Sub lacMin_Click()
    Dim ilRet As Integer
    If imAutoRun = 0 Then
        If Not imConverting Then
            mSetShow
            ilRet = mCheckSetting()
            If Not ilRet Then
                Exit Sub
            End If
            ilRet = MsgBox("WARNING! This will DESTROY much of the data in the database " & Trim(edcFrom.Text) & ", you must run this referencing a COPY of the orginal database", vbQuestion + vbOKCancel, "Extract")
            If ilRet = vbCancel Then
                Exit Sub
            End If
            imConverting = True
            Screen.MousePointer = vbHourglass
            gSetMousePointer grdVehicle, grdVehicle, vbHourglass
            pbcClickFocus.SetFocus
            gLogMsg "Extracting from database at " & Trim(edcFrom.Text), "DB_Extractor_Log.txt", False
            ilRet = mConvert()
            Screen.MousePointer = vbDefault
            gSetMousePointer grdVehicle, grdVehicle, vbDefault
            imConverting = False
            If ilRet Then
                mWritePreselectedVehicles
                gLogMsg "Extracting completed on database at " & Trim(edcFrom.Text), "DB_Extractor_Log.txt", False
                lacStop.Caption = "Done"
                lacMin.Enabled = False
            ElseIf imTerminate = True Then
                gLogMsg "Extracting terminated prior to completing on database at " & Trim(edcFrom.Text), "DB_Extractor_Log.txt", False
                lacStop.Caption = "Exit"
                lacMin.Enabled = False
            End If
        End If
    Else
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub lacStop_Click()
    If Not imConverting Then
        imTerminate = True
        mTerminate
    Else
        imTerminate = True
    End If
End Sub


Private Sub mBtrErrorMsg(slMsg As String, ilInRet As Integer)
    Dim ilRet As Integer
    
    ilRet = ilInRet
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    gLogMsg slMsg & Trim$(Str$(ilRet)), "DB_Extractor_Log.txt", False
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:8/17/05       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box for conventional           *
'*                      and selling vehicles           *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim llRow As Long
    Dim llCol As Long
    
    grdVehicle.Redraw = False
    grdVehicle.Rows = 2
    llRow = grdVehicle.FixedRows
    For llCol = NAMEINDEX To VEFCODEINDEX Step 1
        grdVehicle.TextMatrix(llRow, llCol) = ""
    Next llCol
    
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilVef).sType <> "P") And (tgMVef(ilVef).sType <> "A") Then
            If (tgMVef(ilVef).sState <> "D") Then
                mVehToGrid ilVef, llRow
            End If
        End If
    Next ilVef
    If ckcDormant.Value = vbChecked Then
        ckcDormant_Click
    Else
        mSortCol NAMEINDEX
        grdVehicle.Row = 0
        grdVehicle.Col = VEFCODEINDEX
    End If
    grdVehicle.Redraw = True
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



Private Sub ckcDormant_Click()
    Dim ilVef As Integer
    Dim llRow As Long
    Dim llLoop As Long
    
    grdVehicle.Redraw = False
    If ckcDormant.Value = vbChecked Then
        'Add Dormant
        llRow = grdVehicle.FixedRows
        For llLoop = grdVehicle.Rows - 1 To grdVehicle.FixedRows Step -1
            If grdVehicle.TextMatrix(llLoop, NAMEINDEX) <> "" Then
                llRow = llLoop + 1
                Exit For
            End If
        Next llLoop
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If ((tgMVef(ilVef).sType <> "P") And (tgMVef(ilVef).sType <> "A")) Or ((tgMVef(ilVef).sType = "P") And (tgMVef(ilVef).lPvfCode > 0)) Then
                If (tgMVef(ilVef).sState = "D") Then
                    mVehToGrid ilVef, llRow
                End If
            End If
        Next ilVef
        imLastColSorted = -1
        imLastSort = -1
        mSortCol NAMEINDEX
        grdVehicle.Row = 0
        grdVehicle.Col = VEFCODEINDEX
    Else
        'Remove dormant
        llRow = grdVehicle.Rows - 1
        Do While llRow >= grdVehicle.FixedRows
            If grdVehicle.TextMatrix(llRow, STATEINDEX) = "D" Then
                grdVehicle.RemoveItem llRow
            End If
            llRow = llRow - 1
        Loop
    End If
    grdVehicle.Redraw = True

    mSetCommands
End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slTodayDate As String
    Dim slLastYearDate As String
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVpf As Integer
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer

    mParseCmmdLine
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    If imAutoRun = 0 Then
        lacMin.Caption = "Extract"
    Else
        lacMin.Caption = "Minimize"
    End If
    CSIPass.Show vbModal
    If Not igPasswordOk Then
        lacMin.Enabled = False
    Else
        lacMin.Enabled = True
    End If
    pbcClickFocus.Top = -pbcClickFocus.Height - 100
    imDeleteForm = 1
    imClosed = False
    imTerminate = False
    imFirstActivate = True
    imChgMode = False
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdVehicle, grdVehicle, vbHourglass
    imLastColSorted = -1
    imLastSort = -1
    imConverting = False
    imBypassFocus = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    mReadPreselectedVehicles
    mSetGridColumns
    mSetGridTitles
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    
    mVehPop
        
    slTodayDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slTodayDate, True, slYear, slMonth, slDay
  
    edcFrom.Text = Left(sgDBPath, Len(sgDBPath) - 1)
    gSetMousePointer grdVehicle, grdVehicle, vbDefault

    Exit Sub
         
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Dim ilRet As Integer
  
    
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVehicle, grdVehicle, vbDefault
    'igParentRestarted = False
    'If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute "Done"
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    'End If
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVehicle, grdVehicle, vbDefault
    igManUnload = YES
    Unload Traffic
    Unload DB_Extractor
    Set DB_Extractor = Nothing   'Remve data segment
    igManUnload = NO
End Sub

Private Sub plcSignon_GotFocus()
    mSetShow
End Sub

Private Sub tmcClick_Timer()
    Dim slRepeat As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slDefaultFileName As String

    tmcClick.Enabled = False
    If imTerminate Then
        lacStop_Click
        Exit Sub
    End If
    
    Exit Sub
cmcExportDupNameErr:
    ilRet = 1
    Resume Next
End Sub

Public Sub mSetCommands()

'    If lbcVehicle.SelCount > 0 And edcTo.Text <> "" Then
'        'cmcExport.Enabled = True
'        lacMin.Enabled = True
'    Else
'        'cmcExport.Enabled = False
'        lacMin.Enabled = False
'    End If
    
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdVehicle.ColWidth(VEFCODEINDEX) = 0
    grdVehicle.ColWidth(SORTINDEX) = 0
    grdVehicle.ColWidth(STATEINDEX) = 0
    grdVehicle.ColWidth(STATUSINDEX) = grdVehicle.Width * 0.15
    grdVehicle.ColWidth(TYPEINDEX) = grdVehicle.Width * 0.2
    
    grdVehicle.ColWidth(NAMEINDEX) = grdVehicle.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To TYPEINDEX Step 1
        If ilCol <> NAMEINDEX Then
            grdVehicle.ColWidth(NAMEINDEX) = grdVehicle.ColWidth(NAMEINDEX) - grdVehicle.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdVehicle
    

End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    
    'Set column titles
    grdVehicle.TextMatrix(0, NAMEINDEX) = "Vehicle Name"
    grdVehicle.TextMatrix(0, STATUSINDEX) = "Status"
    grdVehicle.TextMatrix(0, TYPEINDEX) = "Type"
    grdVehicle.Row = 0
    For llCol = NAMEINDEX To TYPEINDEX Step 1
        grdVehicle.Col = llCol
        grdVehicle.CellBackColor = LIGHTBLUE
    Next llCol

End Sub

Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer


    For llRow = grdVehicle.FixedRows To grdVehicle.Rows - 1 Step 1
        slStr = Trim$(grdVehicle.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdVehicle.TextMatrix(llRow, ilCol)))
            slStr = grdVehicle.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                'Ascending
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVehicle.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                'Descending
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVehicle.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdVehicle, NAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Sub mVehToGrid(ilVef As Integer, llRow As Long)
    Dim ilLoop As Integer

    If llRow >= grdVehicle.Rows Then
        grdVehicle.AddItem ""
    End If
    grdVehicle.Row = llRow
    grdVehicle.TextMatrix(llRow, NAMEINDEX) = Trim$(tgMVef(ilVef).sName)
    grdVehicle.Col = NAMEINDEX
    grdVehicle.CellBackColor = LIGHTYELLOW
    
    If UBound(smPreselectNames) > LBound(smPreselectNames) Then
        grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Exclude"
        For ilLoop = LBound(smPreselectNames) To UBound(smPreselectNames) - 1 Step 1
            If Trim$(UCase(tgMVef(ilVef).sName)) = smPreselectNames(ilLoop) Then
                grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Include"
                Exit For
            End If
        Next ilLoop
    Else
        If tgMVef(ilVef).sState = "D" Then
            grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Exclude"
        Else
            grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Include"
        End If
    End If
    Select Case Trim$(tgMVef(ilVef).sType)
        Case "C"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Conventional"
        Case "S"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Selling"
        Case "A"
        Case "L"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Log"
        Case "V"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Virtual"
        Case "T"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Simul-cast"
        Case "P"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Package"
        Case "R"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Rep"
        Case "N"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "NTR"
        Case "G"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Sport"
    End Select
    grdVehicle.Col = TYPEINDEX
    grdVehicle.CellBackColor = LIGHTYELLOW
    grdVehicle.TextMatrix(llRow, SORTINDEX) = ""
    grdVehicle.TextMatrix(llRow, STATEINDEX) = tgMVef(ilVef).sState
    grdVehicle.TextMatrix(llRow, VEFCODEINDEX) = tgMVef(ilVef).iCode
    llRow = llRow + 1

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdVehicle.Row < grdVehicle.FixedRows) Or (grdVehicle.Row >= grdVehicle.Rows) Or (grdVehicle.Col < grdVehicle.FixedCols) Or (grdVehicle.Col >= grdVehicle.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdVehicle.Row
    lmEnableCol = grdVehicle.Col
    bmCtrlVisible = True
    Select Case grdVehicle.Col
        Case NAMEINDEX
            'ignore
        Case STATUSINDEX
        Case TYPEINDEX
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdVehicle.Row < grdVehicle.FixedRows) Or (grdVehicle.Row >= grdVehicle.Rows) Or (grdVehicle.Col < grdVehicle.FixedCols) Or (grdVehicle.Col >= grdVehicle.Cols - 1) Then
        Exit Sub
    End If
    bmCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdVehicle.Col - 1 Step 1
        llColPos = llColPos + grdVehicle.ColWidth(ilCol)
    Next ilCol
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
            pbcToggle.Move grdVehicle.Left + llColPos + 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + 30, grdVehicle.ColWidth(grdVehicle.Col) - 30, grdVehicle.RowHeight(grdVehicle.Row) - 15
            pbcToggle.Visible = True
            pbcToggle.SetFocus
        Case TYPEINDEX
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim llSvRow As Long
    Dim llSvCol As Long

    If (lmEnableRow >= grdVehicle.FixedRows) And (lmEnableRow < grdVehicle.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case STATUSINDEX
            Case TYPEINDEX
        End Select
    End If
    pbcToggle.Visible = False
    edcDropDown.Visible = False
    cmcDropDown.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    bmCtrlVisible = False
    mSetCommands
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdVehicle.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function



Sub mSetListBoxHeight(lbcCtrl As ListBox, llMaxHeight As Long)
'
'  flHeight = gListBoxHeight (ilNoRows, ilMaxRows)
'   Where:
'       ilNoRows (I) - current number of items within the list box
'       ilMaxRows (I) - max number of list box items to be displayed
'       flHeight (O) - height of list box in twips
'
    '+30 because of line above and below
    Dim llRowHeight As Long
    Dim slStr As String
    Dim ilMaxRow As Integer
    
    If lbcCtrl.ListCount > 0 Then
        'Determine standard height, set to small number so that only one row
        'height will be set (15 + 15 + RowHeight; 15 for size of boundaries)
        'Typical result is 300 (15 + 15 + 270)
'        lbcCtrl.Height = 10
        llRowHeight = 15 * SendMessageByString(lbcCtrl.hwnd, LB_GETITEMHEIGHT, 0, slStr)
        ilMaxRow = llMaxHeight / llRowHeight
        If lbcCtrl.ListCount <= ilMaxRow Then
'            lbcCtrl.Height = (lbcCtrl.Height - 30) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
            lbcCtrl.Height = (llRowHeight) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
        Else
'            lbcCtrl.Height = (lbcCtrl.Height - 30) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
            lbcCtrl.Height = (llRowHeight) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
        End If
    End If
End Sub



Sub mProcessArrowKey(ilShift As Integer, ilKeyCode As Integer, lbcCtrl As Control, ilLbcArrowSetting As Integer)
'
'   gProcessArrowKey Shift, KeyCode, lbcCtrl, imLbcArrowSetting
'   Where:
'       Shift (I)- Shift key state
'       KeyCode (I)- Key code
'       lbcCtrl (I)- list box control
'       ilLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilLoop As Integer

    If (ilShift And ALTMASK) > 0 Then
        lbcCtrl.Visible = Not lbcCtrl.Visible
    ElseIf (ilShift And SHIFTMASK) > 0 Then
    Else
        ilLbcArrowSetting = True
        If ilKeyCode = KEYUP Then    'Up arrow
            If lbcCtrl.ListIndex > 0 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex - 1
                If lbcCtrl.MultiSelect Then
                    For ilLoop = 0 To lbcCtrl.ListCount - 1 Step 1
                        If ilLoop = lbcCtrl.ListIndex Then
                            lbcCtrl.Selected(lbcCtrl.ListIndex) = True
                        Else
                            lbcCtrl.Selected(ilLoop) = False
                        End If
                    Next ilLoop
                End If
            End If
        Else
            If lbcCtrl.ListIndex < lbcCtrl.ListCount - 1 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex + 1
                If lbcCtrl.MultiSelect Then
                    For ilLoop = 0 To lbcCtrl.ListCount - 1 Step 1
                        If ilLoop = lbcCtrl.ListIndex Then
                            lbcCtrl.Selected(lbcCtrl.ListIndex) = True
                        Else
                            lbcCtrl.Selected(ilLoop) = False
                        End If
                    Next ilLoop
                End If
            End If
        End If
        ilLbcArrowSetting = False
    End If
End Sub
Sub mProcessLbcClick(lbcCtrl As Control, edcCtrl As Control, ilChgMode As Integer, ilLbcArrowSetting As Integer)
'
'   gProcessLbcClick lbcSPerson, edcDropDown, imChgMode,                                                imLbcArrowSetting
'   Where:
'       lbcSPerson (I)- list box control
'       edcDropDown (I)- User edit control
'       imChgMode (I/O)- Change mode flag
'       imLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilLbcArrowSet As Integer

    If ilChgMode = False Then
        ilChgMode = True
        ilLbcArrowSet = ilLbcArrowSetting
        edcCtrl.Text = lbcCtrl.List(lbcCtrl.ListIndex)
        If (edcCtrl.Visible) And (edcCtrl.Enabled) Then
            edcCtrl.SetFocus
        End If
        If Not ilLbcArrowSet Then
            lbcCtrl.Visible = False
        End If
        ilChgMode = False
    End If
    ilLbcArrowSetting = False
End Sub


Private Function mConvert() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilSellCode As Integer
    Dim ilPrevAirCode As Integer
    Dim ilVpfSell As Integer
    Dim ilVpfAir As Integer
    Dim llSvTotalRecords As Long
    Dim llSvRecordNo As Long
    Dim blFound As Boolean
    
    
    ReDim tmRecreateSort(0 To 0) As RECREATESORT
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVlf, "", smFromDBPath & "VLF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVlfRecLen = Len(tmVlf)
    For llRow = grdVehicle.FixedRows To grdVehicle.Rows - 1 Step 1
        If grdVehicle.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Include" Then
                slStr = grdVehicle.TextMatrix(llRow, VEFCODEINDEX)
                Do While Len(slStr) < 5
                    slStr = "0" & slStr
                Loop
                tmRecreateSort(UBound(tmRecreateSort)).sKey = slStr
                tmRecreateSort(UBound(tmRecreateSort)).iVefCode = Val(slStr)
                tmRecreateSort(UBound(tmRecreateSort)).bSellToConv = False
                tmRecreateSort(UBound(tmRecreateSort)).iAirAgreeVefCode = 0
                ReDim Preserve tmRecreateSort(0 To UBound(tmRecreateSort) + 1) As RECREATESORT
                If grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Selling" Then
                    ilPrevAirCode = -1
                    'Add Airing
                    ilSellCode = Val(grdVehicle.TextMatrix(llRow, VEFCODEINDEX))
                    'ilVpfSell = gBinarySearchVpf(ilSellCode)
                    'If ilVpfSell <> -1 Then
                    '    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '        If (tgMVef(ilLoop).sType = "A") And (tgMVef(ilLoop).sState <> "D") Then
                    '            ilVpfAir = gBinarySearchVpf(tgMVef(ilLoop).iCode)
                    '            If ilVpfAir <> -1 Then
                    '                If tgVpf(ilVpfSell).iSAGroupNo = tgVpf(ilVpfAir).iSAGroupNo Then
                    '                    slStr = Trim$(Str$(tgVpf(ilVpfAir).iVefKCode))
                    '                    Do While Len(slStr) < 5
                    '                        slStr = "0" & slStr
                    '                    Loop
                    '                    tmRecreateSort(UBound(tmRecreateSort)).sKey = slStr
                    '                    tmRecreateSort(UBound(tmRecreateSort)).iVefCode = Val(slStr)
                    '                    tmRecreateSort(UBound(tmRecreateSort)).bSellToConv = False
                    '                    tmRecreateSort(UBound(tmRecreateSort)).iAirAgreeVefCode = 0
                    '                    ReDim Preserve tmRecreateSort(0 To UBound(tmRecreateSort) + 1) As RECREATESORT
                    '                End If
                    '            End If
                    '        End If
                    '    Next ilLoop
                    'End If
                    tmVlfSrchKey0.iSellCode = ilSellCode
                    tmVlfSrchKey0.iSellDay = 0
                    tmVlfSrchKey0.iEffDate(0) = 0
                    tmVlfSrchKey0.iEffDate(1) = 0
                    tmVlfSrchKey0.iSellTime(0) = 0
                    tmVlfSrchKey0.iSellTime(1) = 0
                    tmVlfSrchKey0.iSellPosNo = 0
                    tmVlfSrchKey0.iSellSeq = 0
                    ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilSellCode)
                        If ilPrevAirCode <> tmVlf.iAirCode Then
                            blFound = False
                            For ilLoop = 0 To UBound(tmRecreateSort) - 1 Step 1
                                If tmRecreateSort(ilLoop).iVefCode = tmVlf.iAirCode Then
                                    blFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            ilPrevAirCode = tmVlf.iAirCode
                        Else
                            blFound = True
                        End If
                        If Not blFound Then
                            slStr = Trim$(Str$(tmVlf.iAirCode))
                            Do While Len(slStr) < 5
                                slStr = "0" & slStr
                            Loop
                            tmRecreateSort(UBound(tmRecreateSort)).sKey = slStr
                            tmRecreateSort(UBound(tmRecreateSort)).iVefCode = tmVlf.iAirCode
                            tmRecreateSort(UBound(tmRecreateSort)).bSellToConv = False
                            tmRecreateSort(UBound(tmRecreateSort)).iAirAgreeVefCode = 0
                            ReDim Preserve tmRecreateSort(0 To UBound(tmRecreateSort) + 1) As RECREATESORT
                        End If
                        ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            End If
        End If
    Next llRow
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    If UBound(tmRecreateSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmRecreateSort(), 0), UBound(tmRecreateSort), 0, LenB(tmRecreateSort(0)), 0, LenB(tmRecreateSort(0).sKey), 0
    End If
    imFileNo = 0
    imTotalFiles = 14
    imPassNo = 1
    lmAdjCount = 1
    
    hmVpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVpf, "", smFromDBPath & "VPF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVpfRecLen = Len(tmVpf)
    ilRet = mDeleteRecords("VEF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    
    ilRet = mDeleteRecords("ClF.BTR", INDEXKEY2)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mDeleteRecords("SBF.BTR", INDEXKEY1)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mDeleteRecords("RIF.BTR", INDEXKEY1)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", smFromDBPath & "DRF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imDrfRecLen = Len(tmDrf)
    ilRet = mDeleteRecords("DNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    
    ilRet = mDeleteRecords("CRF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mDeleteRecords("RVF.BTR", INDEXKEY2)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mDeleteRecords("PHF.BTR", INDEXKEY2)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", smFromDBPath & "CLF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", smFromDBPath & "CFF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCffRecLen = Len(tmCff)
    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", smFromDBPath & "SBF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSbfRecLen = Len(tmSbf)
    ilRet = mDeleteRecords("CHF.BTR", INDEXKEY0)
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", smFromDBPath & "SMF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSmfRecLen = Len(tmSmf)
    ilRet = mDeleteRecords("SDF.BTR", INDEXKEY3)
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    'Clear Revenue history
    imFileNo = imFileNo + 1
    lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
    lacInfo(1).Caption = "Clearing Revenue History"
    hmRhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRhf, "", smFromDBPath & "RHF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imRhfRecLen = Len(tmRhf)
    ilRet = btrGetFirst(hmRhf, tmRhf, imRhfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ilRet = btrDelete(hmRhf)
        ilRet = btrGetFirst(hmRhf, tmRhf, imRhfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Loop
    ilRet = btrClose(hmRhf)
    btrDestroy hmRhf
    
    'Set revenue balance to zero
    imFileNo = imFileNo + 1
    lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
    lacInfo(1).Caption = "Clearing Revenue Balance"
    hmSpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", smFromDBPath & "SPF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSpfRecLen = Len(tmSpf)
    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        gStrToPDN "0", 2, 6, tmSpf.sRB
        ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
    End If
    ilRet = btrClose(hmSpf)
    btrDestroy hmSpf
    
    imFileNo = imFileNo + 1
    lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
    lacInfo(1).Caption = "Clearing Advertiser Statistics"
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", smFromDBPath & "ADF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imAdfRecLen = Len(tmAdf)
    ilRet = btrGetFirst(hmAdf, tmAdf, imAdfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        tmAdf.iPct90 = 0
        gStrToPDN "", 2, 6, tmAdf.sCurrAR
        gStrToPDN "", 2, 6, tmAdf.sUnbilled
        gStrToPDN "", 2, 6, tmAdf.sHiCredit
        gStrToPDN "", 2, 6, tmAdf.sTotalGross
        ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)
        ilRet = btrGetNext(hmAdf, tmAdf, imAdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    
    imFileNo = imFileNo + 1
    lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
    lacInfo(1).Caption = "Clearing Agency Statistics"
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", smFromDBPath & "AGF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imAgfRecLen = Len(tmAgf)
    ilRet = btrGetFirst(hmAgf, tmAgf, imAgfRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        tmAgf.iPct90 = 0
        gStrToPDN "", 2, 6, tmAgf.sCurrAR
        gStrToPDN "", 2, 6, tmAgf.sUnbilled
        gStrToPDN "", 2, 6, tmAgf.sHiCredit
        gStrToPDN "", 2, 6, tmAgf.sTotalGross
        ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
        ilRet = btrGetNext(hmAgf, tmAgf, imAgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    
    
    lacInfo(1).Caption = "DB_Extractor Completed"
    lacInfo(0).Caption = ""
    mConvert = True
End Function

Private Function mCheckSetting() As Integer
    Dim slStr As String
    Dim ilRet As Integer
    mCheckSetting = True
    smFromDBPath = Trim$(edcFrom.Text)
    If smFromDBPath = "" Then
        mCheckSetting = False
        MsgBox "Please defined the Database location", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    smFromDBPath = gSetPathEndSlash(smFromDBPath, False)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFrom, "", smFromDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCheckSetting = False
        MsgBox "Check Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    ilRet = btrGetFirst(hmFrom, tmVef, imVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        mCheckSetting = False
        MsgBox "Check Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    ilRet = btrClose(hmFrom)
    btrDestroy hmFrom
    
End Function

Public Function mBinarySearchRecreateSort(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tmRecreateSort)
    ilMax = UBound(tmRecreateSort) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tmRecreateSort(ilMiddle).iVefCode Then
            'found the match
            mBinarySearchRecreateSort = ilMiddle
            Exit Function
        ElseIf ilCode < tmRecreateSort(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchRecreateSort = -1
End Function

Private Function mDeleteRecords(slFileName As String, ilAutoCodeKeyNo As Integer) As Integer
    Dim ilRecLen As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    Dim llNoRec As Long
    Dim ilPos As Integer
    Dim slName As String
    Dim ilInsertRecord As Integer
    
    If imPassNo = 1 Then
        imFileNo = imFileNo + 1
        If imDeleteForm = 0 Then
            lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles & " Gathering"
        Else
            lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
        End If
    End If
    lacInfo(1).Caption = slFileName
    DoEvents
    ilPos = InStr(1, slFileName, ".", vbTextCompare)
    If ilPos <= 0 Then
        mDeleteRecords = False
        MsgBox "File Name to convert in Error, Missing extension " & slFileName, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    slName = Left(slFileName, ilPos - 1)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFrom, "", smFromDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        mDeleteRecords = False
        MsgBox "Check Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    
    ilRecLen = btrRecordLength(hmFrom)  'Get and save record length
    If ilRecLen > Len(tmRecBuff) Then
        'Error message
        mDeleteRecords = False
        MsgBox "Increase General Buffer size to " & ilRecLen, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    If imPassNo = 1 Then
        lmTotalRecords = lmAdjCount * btrRecords(hmFrom)
        lmRecordNo = 0
    End If
    mClearCodes
    gLogMsg "Extracting from " & slFileName, "DB_Extractor_Log.txt", False

    If imDeleteForm = 0 Then
        
        llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
        btrExtClear hmFrom   'Clear any previous extend operation
        ilRet = btrGetFirst(hmFrom, tmRecBuff, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hmFrom)
            btrDestroy hmFrom
            Exit Function
        ElseIf ilRet <> BTRV_ERR_NONE Then
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            mDeleteRecords = False
            MsgBox "Get First Failed for " & slFileName & " with Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
            ilRet = btrClose(hmFrom)
            btrDestroy hmFrom
            Exit Function
        End If

        Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", slName, "") 'Set extract limits (all records)
    
        ilRet = btrExtAddField(hmFrom, 0, ilRecLen)  'Extract iCode field
        ''ilRet = btrExtGetNextExt(hlFile)  'Extract records
        'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)    'Extract recor
        ilRet = mExtGetNext(slFileName, ilRecLen)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilRecLen = btrRecordLength(hmFrom)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlFile, tlCodeNameBuff, ilLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)
                ilRet = mExtGetNext(slFileName, ilRecLen)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If imTerminate Then
                    gLogMsg "Extracting terminated", "DB_Extractor_Log.txt", False
                    mDeleteRecords = False
                    Exit Function
                End If
                lmRecordNo = lmRecordNo + 1
                lacInfo(1).Caption = slFileName & ": " & lmRecordNo & " of " & lmTotalRecords
                DoEvents
                mTestRecord slFileName
                'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)
                ilRet = mExtGetNext(slFileName, ilRecLen)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)
                    ilRet = mExtGetNext(slFileName, ilRecLen)
                Loop
            Loop
            ilRet = mDelete(slFileName, ilRecLen, ilAutoCodeKeyNo)
            ilRet = mUpdate(slFileName, ilRecLen, ilAutoCodeKeyNo)
        End If
        mClearCodes
    Else
        ilRet = mGetFirst(slFileName, ilRecLen, ilAutoCodeKeyNo)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If imTerminate Then
                gLogMsg "Extracting terminated", "DB_Extractor_Log.txt", False
                mDeleteRecords = False
                Exit Function
            End If
            lmRecordNo = lmRecordNo + 1
            lacInfo(1).Caption = slFileName & ": " & lmRecordNo & " of " & lmTotalRecords
            DoEvents
            mTestRecord slFileName
            ilRet = mDelete(slFileName, ilRecLen, ilAutoCodeKeyNo)
            ilRet = mUpdate(slFileName, ilRecLen, ilAutoCodeKeyNo)
            mClearCodes
            ilRet = mGetNext(slFileName, ilRecLen)
        Loop
        mTestRecord slFileName
    End If
    ilRet = btrClose(hmFrom)
    btrDestroy hmFrom
    gLogMsg "Extracting completed for " & slFileName, "DB_Extractor_Log.txt", False
    mDeleteRecords = True
End Function

Private Sub mTestRecord(slFileName As String)
    Dim ilVef As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llLogDate As Long
    Dim ilLoop As Integer
    Dim llTest As Long
    Dim blFound As Boolean
    Dim llIndex As Long
    Dim ilRet As Integer
    Dim llOffAirDate As Long
    Dim llDropDate As Long
    Dim llPvfCode As Long
    Dim llRow As Long
    Dim ilPos As Integer
    
    If UCase(slFileName) = "VEF.BTR" Then
        'tmVef = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVef.iCode)
        If ilVef <= -1 Then
            'mAddInteger tmVef.iCode, imVefCode()
            mAddLong CLng(tmVef.iCode), lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "CLF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmClf.iVefCode)
        If ilVef <= -1 Then
            mAddLong tmClf.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "SBF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmSbf.iBillVefCode)
        If ilVef <= -1 Then
            mAddLong tmSbf.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "RIF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmRif.iVefCode)
        If ilVef <= -1 Then
            mAddLong tmRif.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "DNF.BTR" Then
        'tmVbf = tmRecBuff
        ilPos = InStr(1, Trim$(UCase(tmDnf.sBookName)), "RADAR", vbTextCompare)
        If ilPos > 0 Then
            mAddLong CLng(tmDnf.iCode), lmDeleteCode()
        Else
            mAddLong CLng(tmDnf.iCode), lmUpdateCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "CRF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmCrf.iVefCode)
        If ilVef <= -1 Then
            mAddLong tmCrf.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "RVF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmRvf.iBillVefCode)
        If ilVef <= -1 Then
            mAddLong tmRvf.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "PHF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmPhf.iBillVefCode)
        If ilVef <= -1 Then
            mAddLong tmPhf.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "CHF.BTR" Then
        tmClfSrchKey0.lChfCode = tmChf.lCode
        tmClfSrchKey0.iLine = 0
        tmClfSrchKey0.iCntRevNo = 0
        tmClfSrchKey0.iPropVer = 0
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) Then
            mAddLong tmChf.lCode, lmUpdateCode()
        Else
            'Check if NTR only
            tmSbfSrchKey0.lChfCode = tmChf.lCode
            tmSbfSrchKey0.iDate(0) = 0
            tmSbfSrchKey0.iDate(1) = 0
            tmSbfSrchKey0.sTranType = ""
            ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tmChf.lCode) Then
                mAddLong tmChf.lCode, lmUpdateCode()
            Else
                mAddLong tmChf.lCode, lmDeleteCode()
            End If
        End If
        Exit Sub
    End If
    If UCase(slFileName) = "SDF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmSdf.iVefCode)
        If ilVef <= -1 Then
            mAddLong tmSdf.lCode, lmDeleteCode()
        End If
        Exit Sub
    End If

End Sub


Private Sub mAddLong(llCode As Long, llArray() As Long)
    Dim blFound As Boolean
    Dim llTest As Long
    
    If llCode > 0 Then
        blFound = False
        'For llTest = 0 To UBound(llArray) - 1 Step 1
        '    If llArray(llTest) = llCode Then
        '        blFound = True
        '        Exit For
        '    End If
        'Next llTest
        If Not blFound Then
            llArray(UBound(llArray)) = llCode
            ReDim Preserve llArray(0 To UBound(llArray) + 1) As Long
        End If
    End If
End Sub

Private Sub mAddInteger(ilCode As Integer, ilArray() As Integer)
    Dim blFound As Boolean
    Dim llTest As Long
    
    If ilCode > 0 Then
        blFound = False
        For llTest = 0 To UBound(ilArray) - 1 Step 1
            If ilArray(llTest) = ilCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            ilArray(UBound(ilArray)) = ilCode
            ReDim Preserve ilArray(0 To UBound(ilArray) + 1) As Integer
        End If
    End If
End Sub




Private Function mOpenFile(hlFrom As Integer, hlTo As Integer, slFileName As String) As Integer
    Dim ilRet As Integer
    
    hlFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlFrom, "", smFromDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenFile = False
        Exit Function
    End If
    mOpenFile = True
End Function

Private Function mCloseFile(hlFrom As Integer, hlTo As Integer) As Integer
    Dim ilRet As Integer
    
    ilRet = btrClose(hlFrom)
    btrDestroy hlFrom
    mCloseFile = True
End Function



Private Function mExtGetNext(slFileName As String, ilRecLen As Integer) As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    
    Select Case UCase(slFileName)
    
        Case "VEF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVef, ilRecLen, llRecPos)    'Extract record
        Case "CLF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmClf, ilRecLen, llRecPos)    'Extract record
        Case "SBF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSbf, ilRecLen, llRecPos)    'Extract record
        Case "RIF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmRif, ilRecLen, llRecPos)    'Extract record
        Case "DNF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmDnf, ilRecLen, llRecPos)    'Extract record
        Case "CRF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmCrf, ilRecLen, llRecPos)    'Extract record
        Case "RVF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmRvf, ilRecLen, llRecPos)    'Extract record
        Case "PHF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPhf, ilRecLen, llRecPos)    'Extract record
        Case "CHF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmChf, ilRecLen, llRecPos)    'Extract record
        Case "SDF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSdf, ilRecLen, llRecPos)    'Extract record
        Case Else
            ilRet = -1
    End Select
    mExtGetNext = ilRet
End Function

Private Function mDelete(slFileName As String, ilRecLen As Integer, ilAutoCodeKeyNo As Integer) As Integer
    Dim ilRet As Integer
    Dim llLoop As Long
    
    ilRet = -1
    If imDeleteForm = 0 Then
        lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles & " Deleting"
    End If
    Select Case UCase(slFileName)
    
        Case "VEF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmIntSrchKey.iCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmIntSrchKey, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmVpf)
                End If
                tmIntSrchKey.iCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmVef, ilRecLen, tmIntSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "CLF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmClf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "SBF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmSbf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "RIF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmRif, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "DNF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmDnf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "CRF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmCrf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "RVF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmRvf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "PHF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmPhf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "CHF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmChf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
        Case "SDF.BTR"
            For llLoop = 0 To UBound(lmDeleteCode) - 1 Step 1
                tmSmfSrchKey2.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, 2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmSmf)
                End If
                tmLongSrchKey.lCode = lmDeleteCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmSdf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmFrom)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmDeleteCode)
                End If
            Next llLoop
    End Select
    mDelete = True
End Function


Private Function mComputeContractAirTime(llChfCode As Long) As Long
    Dim ilRet As Integer
    Dim llTotalGross As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim ilDay As Integer
    
    llTotalGross = 0
    tmClfSrchKey0.lChfCode = llChfCode
    tmClfSrchKey0.iLine = 0
    tmClfSrchKey0.iCntRevNo = 0
    tmClfSrchKey0.iPropVer = 0
    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode)
        If (tmClf.sType <> "O") And (tmClf.sType <> "A") Then
            tmCffSrchKey0.lChfCode = llChfCode
            tmCffSrchKey0.iClfLine = tmClf.iLine
            tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
            tmCffSrchKey0.iPropVer = tmClf.iPropVer
            tmCffSrchKey0.iStartDate(0) = 0
            tmCffSrchKey0.iStartDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmCff.lChfCode = llChfCode) And (tmCff.iClfLine = tmClf.iLine)
                gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llStartDate
                gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llEndDate
                If llStartDate <= llEndDate Then
                    For llDate = llStartDate To llEndDate Step 7
                        If tmCff.sPriceType = "T" Then
                            If tmCff.sDyWk = "D" Then
                                For ilDay = 0 To 6 Step 1
                                    If (llDate + ilDay >= llStartDate) And (llDate + ilDay <= llEndDate) Then
                                        llTotalGross = llTotalGross + tmCff.iDay(ilDay) * tmCff.lActPrice
                                    End If
                                Next ilDay
                            Else
                                llTotalGross = llTotalGross + tmCff.iSpotsWk * tmCff.lActPrice
                            End If
                        End If
                    Next llDate
                End If
                ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mComputeContractAirTime = llTotalGross
End Function
Private Function mComputeContractNTR(llChfCode As Long) As Long
    Dim ilRet As Integer
    Dim llTotalGross As Long
    
    llTotalGross = 0
    tmSbfSrchKey0.lChfCode = llChfCode
    tmSbfSrchKey0.iDate(0) = 0
    tmSbfSrchKey0.iDate(1) = 0
    tmSbfSrchKey0.sTranType = "I"
    ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llChfCode)
        If (tmSbf.sTranType = "I") Then
            llTotalGross = llTotalGross + tmSbf.lGross * tmSbf.iNoItems
        End If
        ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mComputeContractNTR = llTotalGross
End Function
Private Function mUpdate(slFileName As String, ilRecLen As Integer, ilAutoCodeKeyNo As Integer)
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim ilVef As Integer
    
    ilRet = -1
    If imDeleteForm = 0 Then
        lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles & " Updating"
    End If
    Select Case UCase(slFileName)
        Case "DNF.BTR"
            For llLoop = 0 To UBound(lmUpdateCode) - 1 Step 1
                tmDrfSrchKey0.iDnfCode = lmUpdateCode(llLoop)
                tmDrfSrchKey0.sDemoDataType = ""
                tmDrfSrchKey0.iMnfSocEco = 0
                tmDrfSrchKey0.iVefCode = 0
                tmDrfSrchKey0.sInfoType = ""
                tmDrfSrchKey0.iRdfCode = 0
                ilRet = btrGetGreaterOrEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmDrf.iDnfCode = lmUpdateCode(llLoop))
                    '1-23-13 Always retain the population records
                    If tmDrf.sDemoDataType = "P" Then
                        ilRet = btrGetNext(hmDrf, tmDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Else
                    ilVef = mBinarySearchRecreateSort(tmDrf.iVefCode)
                    If ilVef <= -1 Then
                        ilRet = btrDelete(hmDrf)
                    End If
                    ilRet = btrGetNext(hmDrf, tmDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                Loop
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmUpdateCode)
                End If
            Next llLoop
        Case "CHF.BTR"
            For llLoop = 0 To UBound(lmUpdateCode) - 1 Step 1
                tmLongSrchKey.lCode = lmUpdateCode(llLoop)
                ilRet = btrGetEqual(hmFrom, tmChf, ilRecLen, tmLongSrchKey, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    tmChf.lAirTimeGross = mComputeContractAirTime(lmUpdateCode(llLoop))
                    tmChf.lNTRGross = mComputeContractNTR(lmUpdateCode(llLoop))
                    tmChf.lInputGross = tmChf.lAirTimeGross + tmChf.lNTRGross
                    ilRet = btrUpdate(hmFrom, tmChf, ilRecLen)
                End If
                If imDeleteForm = 0 Then
                    lacInfo(1).Caption = slFileName & ": " & llLoop + 1 & " of " & UBound(lmUpdateCode)
                End If
            Next llLoop
            ilRet = -1
    End Select
    mUpdate = True

End Function

Private Sub mReadPreselectedVehicles()
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim slStr(0 To 1) As String
    
    ReDim smPreselectNames(0 To 0)
    ilRet = 0
    On Error GoTo mReadPreselectedVehicles:
    hmVehicles = FreeFile
    Open Trim$(sgImportPath) & "DB_Extractor.Txt" For Input Access Read As hmVehicles
    If ilRet <> 0 Then
        Exit Sub
    End If
    Do
        On Error GoTo mReadPreselectedVehicles:
        Line Input #hmVehicles, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                gParseCDFields slLine, False, slStr()
                smPreselectNames(UBound(smPreselectNames)) = Trim(UCase(slStr(1)))
                ReDim Preserve smPreselectNames(0 To UBound(smPreselectNames) + 1) As String
            End If
        End If
    Loop Until ilEof
    Close hmVehicles
    Exit Sub
mReadPreselectedVehicles:
    ilRet = Err.Number
    Resume Next
End Sub
Private Sub mWritePreselectedVehicles()
    Dim ilRet As Integer
    Dim llRow As Long
    
    On Error Resume Next
    Kill Trim$(sgImportPath) & "DB_Extractor.Txt"
    On Error GoTo 0
    
    ilRet = 0
    On Error GoTo mWritePreselectedVehiclesErr:
    hmVehicles = FreeFile
    Open Trim$(sgImportPath) & "DB_Extractor.Txt" For Output As hmVehicles
    If ilRet = 0 Then
        For llRow = grdVehicle.FixedRows To grdVehicle.Rows - 1 Step 1
            If grdVehicle.TextMatrix(llRow, NAMEINDEX) <> "" Then
                If grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Include" Then
                    Print #hmVehicles, grdVehicle.TextMatrix(llRow, NAMEINDEX)
                End If
            End If
        Next llRow
        Close #hmVehicles
    End If
    Exit Sub
mWritePreselectedVehiclesErr:
    ilRet = Err.Number
    Resume Next
End Sub
Private Function mGetFirst(slFileName As String, ilRecLen As Integer, ilAutoCodeKeyNo As Integer) As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    
    Select Case UCase(slFileName)
    
        Case "VEF.BTR"
            ilRet = btrGetFirst(hmFrom, tmVef, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "CLF.BTR"
            ilRet = btrGetFirst(hmFrom, tmClf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "SBF.BTR"
            ilRet = btrGetFirst(hmFrom, tmSbf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "RIF.BTR"
            ilRet = btrGetFirst(hmFrom, tmRif, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "DNF.BTR"
            ilRet = btrGetFirst(hmFrom, tmDnf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "CRF.BTR"
            ilRet = btrGetFirst(hmFrom, tmCrf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "RVF.BTR"
            ilRet = btrGetFirst(hmFrom, tmRvf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "PHF.BTR"
            ilRet = btrGetFirst(hmFrom, tmPhf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "CHF.BTR"
            ilRet = btrGetFirst(hmFrom, tmChf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "SDF.BTR"
            ilRet = btrGetFirst(hmFrom, tmSdf, ilRecLen, ilAutoCodeKeyNo, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case Else
            ilRet = -1
    End Select
    mGetFirst = ilRet
End Function

Private Function mGetNext(slFileName As String, ilRecLen As Integer) As Integer
    Dim ilRet As Integer
    
    Select Case UCase(slFileName)
    
        Case "VEF.BTR"
            ilRet = btrGetNext(hmFrom, tmVef, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "CLF.BTR"
            ilRet = btrGetNext(hmFrom, tmClf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "SBF.BTR"
            ilRet = btrGetNext(hmFrom, tmSbf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "RIF.BTR"
            ilRet = btrGetNext(hmFrom, tmRif, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "DNF.BTR"
            ilRet = btrGetNext(hmFrom, tmDnf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "CRF.BTR"
            ilRet = btrGetNext(hmFrom, tmCrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "RVF.BTR"
            ilRet = btrGetNext(hmFrom, tmRvf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "PHF.BTR"
            ilRet = btrGetNext(hmFrom, tmPhf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "CHF.BTR"
            ilRet = btrGetNext(hmFrom, tmChf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case "SDF.BTR"
            ilRet = btrGetNext(hmFrom, tmSdf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)    'Extract record
        Case Else
            ilRet = -1
    End Select
    mGetNext = ilRet
End Function


Private Sub mClearCodes()
    ReDim lmDeleteCode(0 To 0) As Long
    ReDim lmUpdateCode(0 To 0) As Long
End Sub
