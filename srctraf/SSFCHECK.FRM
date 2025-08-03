VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SsfCheck 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   240
   ClientTop       =   1620
   ClientWidth     =   8580
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
   ScaleHeight     =   6030
   ScaleWidth      =   8580
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   240
      Left            =   165
      TabIndex        =   25
      Top             =   5265
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox plcPLMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   585
      ScaleHeight     =   255
      ScaleWidth      =   5850
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Width           =   5850
      Begin VB.OptionButton rbcPLMove 
         Caption         =   "Outsides"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4125
         TabIndex        =   7
         Top             =   0
         Width           =   1050
      End
      Begin VB.OptionButton rbcPLMove 
         Caption         =   "MG's"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   3405
         TabIndex        =   6
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.TextBox edcDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   4
      Top             =   390
      Width           =   1170
   End
   Begin VB.CheckBox ckcClone 
      Caption         =   "Recreate SSF"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5430
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton rbcCF 
      Caption         =   "Fix"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   15
      Width           =   675
   End
   Begin VB.OptionButton rbcCF 
      Caption         =   "Check"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   5430
      TabIndex        =   8
      Top             =   15
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.CheckBox ckcReschedule 
      Caption         =   "Reschedule Spots set to Missed"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5430
      TabIndex        =   10
      Top             =   225
      Width           =   3015
   End
   Begin VB.TextBox edcDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1605
      MaxLength       =   10
      TabIndex        =   2
      Top             =   390
      Width           =   1170
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6375
      Top             =   5340
   End
   Begin VB.PictureBox plcChkFix 
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
      Height          =   2790
      Left            =   1005
      ScaleHeight     =   2730
      ScaleWidth      =   6525
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1035
      Width           =   6585
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Vehicles"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   14
         Top             =   2055
         Width           =   1350
      End
      Begin VB.ListBox lbcVehs 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   150
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   60
         Width           =   6300
      End
      Begin VB.Label lacTest 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1545
         TabIndex        =   21
         Top             =   2265
         Width           =   4950
      End
      Begin VB.Label lacMissed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1545
         TabIndex        =   22
         Top             =   2460
         Width           =   4965
      End
      Begin VB.Label lacProcessing 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1545
         TabIndex        =   23
         Top             =   2070
         Width           =   4860
      End
   End
   Begin VB.CommandButton cmcChkFix 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   3060
      TabIndex        =   16
      Top             =   5400
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4935
      TabIndex        =   17
      Top             =   5400
      Width           =   1140
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   2175
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2175
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
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
      Left            =   7860
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5505
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Left            =   7245
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5430
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Left            =   6675
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5430
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcError 
      Appearance      =   0  'Flat
      Height          =   1290
      Left            =   180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3870
      Width           =   8160
   End
   Begin VB.Label lacDBPath 
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
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   45
      TabIndex        =   24
      Top             =   5760
      Width           =   6060
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3405
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "End Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   3045
      TabIndex        =   3
      Top             =   435
      Width           =   765
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Start Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   585
      TabIndex        =   1
      Top             =   435
      Width           =   885
   End
End
Attribute VB_Name = "SsfCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFCHECK.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Check/Fix MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmTo As Integer   'From file hanle
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF
Dim imVefRecLen As Integer     'VEF record length
Dim hmSsf As Integer        'Spot summary file handle
Dim hmSdf As Integer        'Spot file file handle
Dim hmSmf As Integer        'Spot move info file handle
Dim hmLcf As Integer        'Library calendar file handle
Dim hmCHF As Integer
Dim hmClf As Integer
Dim hmCff As Integer
Dim hmRdf As Integer
Dim hmCrf As Integer
Dim hmCgf As Integer
Dim hmGsf As Integer
Dim tmGsf As GSF
Dim tmGsfSrchKey3 As GSFKEY3
Dim imGsfRecLen As Integer
Dim hmGhf As Integer
'M for N Tracer record
Dim hmMtf As Integer        'Spot Detail
Dim hmSxf As Integer
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim lmNowDate As Long   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim imFirstActivate As Integer


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
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehs.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehs.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        If Value Then
            ckcClone.Enabled = True
        Else
            ckcClone.Value = vbUnchecked
            ckcClone.Enabled = False
        End If
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    If imProcessing Then
        imCancelled = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcChkFix_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilType As Integer
    Dim slType As String
    Dim sLCP As String
    Dim llEarliestLcfDate As Long
    Dim llLatestLcfDate As Long
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim llMissed As Long
    Dim llTotalVeh As Long
    Dim llVehProc As Long
    Dim llPercent As Long
    Dim slName As String
    Dim llTotalNoToProc As Long
    Dim llNoProc As Long
    Dim ilError As Integer
    Dim ilTestOnly As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilStartNoMissed As Integer
    Dim slMsg As String
    Dim llLen As Long
    Dim ilSeeMsg As Integer
    Dim ilRes As Integer
    Dim slMOrO As String
    Dim ilSvRet As Integer
    Dim ilVef As Integer

    If rbcCF(0).Value Then
        ilTestOnly = True   'Check Only
    Else
        If (Len(sgMDBPath) > 0) And (Len(sgSDBPath) > 0) Then
            ilRes = MsgBox("This will Fix only Data in " & sgDBPath, vbOKCancel + vbQuestion, "Fix")
            If ilRes = vbCancel Then
                Exit Sub
            End If
            ilRes = MsgBox("All Processing on " & sgMDBPath & " and " & sgSDBPath & " must be Stopped before Proceeding", vbOKCancel + vbQuestion, "Fix")
            If ilRes = vbCancel Then
                Exit Sub
            End If
        End If
        ilTestOnly = False  'Fix
    End If
    lbcError.Clear
    ilType = 0
    slType = "O"
    sLCP = "C"
    slStartDate = edcDate(0).Text
    If Not gValidDate(slStartDate) Then
        Beep
        edcDate(0).SetFocus
        Exit Sub
    End If
    slEndDate = UCase(edcDate(1).Text)
    If (Len(slEndDate) <> 0) And (slEndDate <> "TFN") Then
        If Not gValidDate(slEndDate) Then
            Beep
            edcDate(1).SetFocus
            Exit Sub
        End If
        If gDateValue(slEndDate) < gDateValue(slStartDate) Then
            Beep
            MsgBox "End Date must be after Start Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
            edcDate(1).SetFocus
            Exit Sub
        End If
    ElseIf slEndDate = "TFN" Then
        slEndDate = ""
    End If
    If Not ilTestOnly Then
        If (Not rbcPLMove(0).Value) And (Not rbcPLMove(1).Value) Then
            Beep
            MsgBox "Post Log Move Spot question must be defined ", vbOKOnly + vbInformation + vbApplicationModal, "Move define Error"
            rbcPLMove(0).SetFocus
            Exit Sub
        End If
    End If
    If rbcPLMove(0).Value Then
        slMOrO = "G"
    Else
        slMOrO = "O"
    End If
    Screen.MousePointer = vbHourglass
    ilRet = 0
    'On Error GoTo cmcChkFixErr:
    If ilTestOnly Then
        'slToFile = sgExportPath & "SSFCheck.Txt"
        slToFile = sgDBPath & "Messages\" & "SSFCheck.Txt"
    Else
        'slToFile = sgExportPath & "SSFFix.Txt"
        slToFile = sgDBPath & "Messages\" & "SSFFix.Txt"
    End If
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo cmcChkFixErr:
            'hmTo = FreeFile
            'Open slToFile For Append As hmTo
            ilRet = gFileOpen(slToFile, "Append", hmTo)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo cmcChkFixErr:
            'hmTo = FreeFile
            'Open slToFile For Output As hmTo
            ilRet = gFileOpen(slToFile, "Output", hmTo)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo cmcChkFixErr:
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo 0
    If ilTestOnly Then
        Print #hmTo, "SSF Check " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Else
        Print #hmTo, "SSF Fix " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    Print #hmTo, ""
    llTotalVeh = 0
    llVehProc = 0
    For ilLoop = 0 To lbcVehs.ListCount - 1 Step 1
        If lbcVehs.Selected(ilLoop) Then
            llTotalVeh = llTotalVeh + 1
        End If
    Next ilLoop
    llLen = 0
    ilSeeMsg = False
    imProcessing = True
    imCancelled = False
    plcGauge.Value = 0
    llTotalNoToProc = 0
    llNoProc = 0
    If Not ilTestOnly Then
        If (llTotalVeh = lbcVehs.ListCount) And (ckcClone.Visible = True) And (ckcClone.Value = vbChecked) And (ckcClone.Enabled) Then
            llTotalNoToProc = btrRecords(hmSsf)
        Else
            llTotalNoToProc = 0
        End If
        If Not gOpenSchFiles() Then
            MsgBox "Unable to Open Schedule Files", vbOKOnly + vbInformation + vbApplicationModal, "Open Error"
            Exit Sub
        End If
    Else
        llTotalNoToProc = 0
    End If
    For ilLoop = 0 To lbcVehs.ListCount - 1 Step 1
        If lbcVehs.Selected(ilLoop) Then
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get user name
            ilRet = gParseItem(slName, 3, "|", slName)    'Get user name
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get user name
            DoEvents
            If imCancelled Then
                Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Close #hmTo
                If Not ilTestOnly Then
                    gCloseSchFiles
                End If
                imProcessing = False
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            ilVefCode = Val(slCode)
            llEarliestLcfDate = gGetEarliestLCFDate(hmLcf, sLCP, ilVefCode)
            llLatestLcfDate = gGetLatestLCFDate(hmLcf, sLCP, ilVefCode)
            'Remove any SSF prior to llEarliestLcfDate
            'Not coded
            'Recreate SSF
            If gDateValue(slStartDate) > llEarliestLcfDate Then
                llStartDate = gDateValue(slStartDate)
            Else
                llStartDate = llEarliestLcfDate
            End If
            If Len(slEndDate) > 0 Then
                If gDateValue(slEndDate) < llLatestLcfDate Then
                    llEndDate = gDateValue(slEndDate)
                Else
                    llEndDate = llLatestLcfDate
                End If
            Else
                llEndDate = llLatestLcfDate
            End If
            If llEndDate > llStartDate Then
                llTotalNoToProc = llTotalNoToProc + (llEndDate - llStartDate) + 1
            Else
                llTotalNoToProc = llTotalNoToProc + 1
            End If
        End If
    Next ilLoop
    'Clone file and remake into clone, then rename at the end
    'This was removed because stepfirst not in SQL
'    If Not ilTestOnly Then
'        If (llTotalVeh = lbcVehs.ListCount) And (ckcClone.Visible = True) And (ckcClone.Value = vbChecked) And (ckcClone.Enabled) Then
'
'            lacProcessing.Caption = "Recreating Ssf in First to Next order"
'            DoEvents
'            llTotalRec = btrRecords(hmSsf)
'            'ReDim lmRecNoMoved(1 To llTotalRec)  As Long
'            ilNoRows = llTotalRec \ 32000 + 1
'            ReDim lmRecNoMoved(1 To ilNoRows, 1 To 32000) As Long
'            ilRet = btrClone(hmSsf, sgDBPath & "SsfNew.Btr", 0)
'            If ilRet <> BTRV_ERR_NONE Then
'                Screen.MousePointer = vbDefault
'                Print #hmTo, "Unable to clone Ssf- Error code " & Trim$(Str$(ilRet)) & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'                Close #hmTo
'                MsgBox "Unable to clone Ssf- Error code " & Trim$(Str$(ilRet))
'                imProcessing = False
'                Exit Sub
'            End If
'            ilRet = btrClose(hmSsf)
'            btrDestroy hmSsf
'            hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'            ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_VERIFY, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'            If ilRet <> BTRV_ERR_NONE Then
'                Screen.MousePointer = vbDefault
'                Print #hmTo, "Unable to open Ssf- Error code " & Trim$(Str$(ilRet)) & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'                Close #hmTo
'                MsgBox "Unable to open Ssf- Error code " & Trim$(Str$(ilRet))
'                imProcessing = False
'                Exit Sub
'            End If
'            'Move all records into clone
'            hmSsfNew = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'            ilRet = btrOpen(hmSsfNew, "", sgDBPath & "SsfNew.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'            If ilRet <> BTRV_ERR_NONE Then
'                Screen.MousePointer = vbDefault
'                Print #hmTo, "Unable to open SsfNew- Error code " & Trim$(Str$(ilRet)) & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
'                Close #hmTo
'                MsgBox "Unable to open SsfNew- Error code " & Trim$(Str$(ilRet))
'                imProcessing = False
'                Exit Sub
'            End If
'            llRecNo = 0
'            llNoProc = 0
'            imSsfRecLen = Len(tmSsf)
'            ilRet = gSSFStepFirst(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'            Do While ilRet = BTRV_ERR_NONE
'                llNoProc = llNoProc + 1
'                ilRet = gSSFGetPosition(hmSsf, llRecPos)
'                ilFound = False
'                'For llLoop = 1 To llRecNo Step 1
'                '    If lmRecNoMoved(llLoop) = llRecPos Then
'                '        ilFound = True
'                '        Exit For
'                '    End If
'                'Next llLoop
'                If Not ilFound Then
'                    llRecNo = llRecNo + 1
'                    llRecAdded = llRecAdded + 1
'                    ilRow = (llRecNo - 1) \ 32000 + 1
'                    ilCol = llRecNo - 32000 * CLng((ilRow - 1))
'                    lmRecNoMoved(ilRow, ilCol) = llRecPos
'                    imSsfRecLen = 17 + tmSsf.iCount * Len(tmAvail)
'                    ilRet = gSSFInsert(hmSsfNew, tmSsf, imSsfRecLen, INDEXKEY0)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        lmRecNoMoved(ilRow, ilCol) = 0
'                        llRecNo = llRecNo - 1
'                        llRecAdded = llRecAdded - 1
'                    End If
'                End If
'                lacTest.Caption = "Record Added" & Str$(llRecAdded) & " of" & Str(llTotalRec)
'                DoEvents
'                llPercent = (llNoProc * 100) \ llTotalNoToProc
'                If llPercent >= 99 Then
'                    llPercent = 99
'                ElseIf llPercent < 0 Then
'                    llPercent = 0
'                End If
'                plcGauge.Value = llPercent
'                DoEvents
'                imSsfRecLen = Len(tmSsf)
'                ilRet = gSSFStepNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'                Do While (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE)
'                    imSsfRecLen = Len(tmSsf)
'                    ilRet = gSSFStepNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'                Loop
'            Loop
'            If btrRecords(hmSsf) <> btrRecords(hmSsfNew) Then
'                lacProcessing.Caption = "Recreating Ssf in Last to Previous order"
'                DoEvents
'                'try creating once again, this time last, prev
'                imSsfRecLen = Len(tmSsf)
'                'llRecNo = 0
'                llNoProc = 0
'                llFirstRecNo = llRecNo
'                ilRet = gSSFStepLast(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'                Do While ilRet = BTRV_ERR_NONE
'                    llNoProc = llNoProc + 1
'                    ilRet = gSSFGetPosition(hmSsf, llRecPos)
'                    ilFound = False
'                    ilCol = 1
'                    ilRow = 1
'                    For llLoop = 1 To llFirstRecNo Step 1
'                        If lmRecNoMoved(ilRow, ilCol) = llRecPos Then
'                            ilFound = True
'                            Exit For
'                        End If
'                        ilCol = ilCol + 1
'                        If ilCol > 32000 Then
'                            ilCol = 1
'                            ilRow = ilRow + 1
'                        End If
'                    Next llLoop
'                    If Not ilFound Then
'                        llRecNo = llRecNo + 1
'                        llRecAdded = llRecAdded + 1
'                        ilRow = (llRecNo - 1) \ 32000 + 1
'                        ilCol = llRecNo - 32000 * CLng((ilRow - 1))
'                        lmRecNoMoved(ilRow, ilCol) = llRecPos
'                        imSsfRecLen = 17 + tmSsf.iCount * Len(tmAvail)
'                        ilRet = gSSFInsert(hmSsfNew, tmSsf, imSsfRecLen, INDEXKEY0)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            lmRecNoMoved(ilRow, ilCol) = 0
'                            llRecNo = llRecNo - 1
'                            llRecAdded = llRecAdded - 1
'                        End If
'                    End If
'                    lacTest.Caption = "Record Added" & Str$(llRecAdded) & " of" & Str(llTotalRec)
'                    DoEvents
'                    llPercent = (llNoProc * 100) \ llTotalNoToProc
'                    If llPercent >= 100 Then
'                        llPercent = 100
'                    ElseIf llPercent < 0 Then
'                        llPercent = 0
'                    End If
'                    plcGauge.Value = llPercent
'                    DoEvents
'                    imSsfRecLen = Len(tmSsf)
'                    ilRet = gSSFStepPrevious(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'                    Do While (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE)
'                        imSsfRecLen = Len(tmSsf)
'                        ilRet = gSSFStepPrevious(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'                    Loop
'                Loop
'    '            If btrRecords(hmSsf) <> btrRecords(hmSsfNew) Then
'    '                'try creating once again, this time by key
'    '                lacProcessing.Caption = "Recreating Ssf by Key"
'    '                DoEvents
'    '                ilRet = btrClose(hmSsfNew)
'    '                btrDestroy hmSsfNew
'    '                ilRet = btrClone(hmSsf, sgDBPath & "SsfNew.Btr", 0)
'    '                If ilRet <> BTRV_ERR_NONE Then
'    '                    MsgBox "Unable to clone Ssf- Error code " & Trim$(Str$(ilRet))
'    '                    Screen.MousePointer = vbDefault
'    '                    imProcessing = False
'    '                    Exit Sub
'    '                End If
'    '                'Move all records into clone
'    '                hmSsfNew = CBtrvTable() 'CBtrvObj()
'    '                ilRet = btrOpen(hmSsfNew, "", sgDBPath & "SsfNew.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    '                If ilRet <> BTRV_ERR_NONE Then
'    '                    MsgBox "Unable to open SsfNew- Error code " & Trim$(Str$(ilRet))
'    '                    Screen.MousePointer = vbDefault
'    '                    imProcessing = False
'    '                    Exit Sub
'    '                End If
'    '                llRecNo = 0
'    '                imSsfRecLen = Len(tmSsf)
'    '                ilRet = btrGetFirst(hmSsf, tmSsf, imSsfRecLen, INDEXKEY0, BTRV_LOCK_NONE)
'    '                Do While ilRet <> BTRV_ERR_END_OF_FILE
'    '                    If ilRet = BTRV_ERR_NONE Then
'    '                        llRecNo = llRecNo + 1
'    '                        lacTest.Caption = "Record" & Str$(llRecNo) & " of" & Str(llTotalRec)
'    '                        DoEvents
'    '                        imSsfRecLen = 17 + tmSsf.iCount * Len(tmAvail)
'    '                        ilRet = btrInsert(hmSsfNew, tmSsf, imSsfRecLen, INDEXKEY0)
'    '                    End If
'    '                    imSsfRecLen = Len(tmSsf)
'    '                    ilRet = btrGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE)
'    '                Loop
'    '            End If
'            End If
'            ilRet = btrClose(hmSsfNew)
'            btrDestroy hmSsfNew
'            ilRet = btrClose(hmSsf)
'            btrDestroy hmSsf
'            On Error Resume Next
'            Kill sgDBPath & "SsfOld.Btr"
'            On Error GoTo 0
'            Name sgDBPath & "Ssf.Btr" As sgDBPath & "SsfOld.Btr"
'            Name sgDBPath & "SsfNew.Btr" As sgDBPath & "Ssf.Btr"
'            hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'            ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        End If
'    End If
    For ilLoop = 0 To lbcVehs.ListCount - 1 Step 1
        If lbcVehs.Selected(ilLoop) Then
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get user name
            ilRet = gParseItem(slName, 3, "|", slName)    'Get user name
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get user name
            DoEvents
            If imCancelled Then
                Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Close #hmTo
                If Not ilTestOnly Then
                    gCloseSchFiles
                End If
                imProcessing = False
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            ilVefCode = Val(slCode)
            llEarliestLcfDate = gGetEarliestLCFDate(hmLcf, sLCP, ilVefCode)
            llLatestLcfDate = gGetLatestLCFDate(hmLcf, sLCP, ilVefCode)
            'Remove any SSF prior to llEarliestLcfDate
            'Not coded
            'Recreate SSF
            If gDateValue(slStartDate) > llEarliestLcfDate Then
                llStartDate = gDateValue(slStartDate)
            Else
                llStartDate = llEarliestLcfDate
            End If
            If Len(slEndDate) > 0 Then
                If gDateValue(slEndDate) < llLatestLcfDate Then
                    llEndDate = gDateValue(slEndDate)
                Else
                    llEndDate = llLatestLcfDate
                End If
            Else
                llEndDate = llLatestLcfDate
            End If
            If ilTestOnly Then
                slMsg = "Checking " & slName & " " & Format$(llStartDate, "m/d/yy") & "-" & Format$(llEndDate, "m/d/yy")
            Else
                slMsg = "Processing " & slName & " " & Format$(llStartDate, "m/d/yy") & "-" & Format$(llEndDate, "m/d/yy")
            End If
            lacProcessing.Caption = slMsg
            Print #hmTo, ""
            Print #hmTo, slMsg
            Print #hmTo, ""
            'ReDim lgReschSdfCode(1 To 1) As Long
            ReDim lgReschSdfCode(0 To 0) As Long
            Do  'Loop until no spots made missed
                ilStartNoMissed = UBound(lgReschSdfCode)
                For llDate = llStartDate To llEndDate Step 1
                    slDate = Format$(llDate, "m/d/yy")
                    If ilTestOnly Then
                        lacTest.Caption = slDate & "- Checking Ssf"
                    Else
                        lacTest.Caption = slDate & "- Recreating Ssf"
                    End If
                    gPackDate slDate, ilLogDate0, ilLogDate1
                    DoEvents
                    If imCancelled Then
                        Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                        Close #hmTo
                        If Not ilTestOnly Then
                            gCloseSchFiles
                        End If
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    'Test One- Check and fix Ssf as compared to LCF
                    ilType = 0
                    ilVef = gBinarySearchVef(ilVefCode)
                    If ilVef <> -1 Then
                        If tgMVef(ilVef).sType = "G" Then
                            tmGsfSrchKey3.iVefCode = ilVefCode
                            tmGsfSrchKey3.iGameNo = 0
                            ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get last current record to obtain date
                            Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode)
                                If (tmGsf.iAirDate(0) = ilLogDate0) And (tmGsf.iAirDate(1) = ilLogDate1) Then
                                    ilType = tmGsf.iGameNo
                                    ilRet = gMakeSSF(ilTestOnly, hmSsf, hmSdf, hmSmf, ilType, ilVefCode, ilLogDate0, ilLogDate1, 0)
                                    ilSvRet = ilRet
                                    For ilError = LBound(sgSSFErrorMsg) To UBound(sgSSFErrorMsg) - 1 Step 1
                                        Print #hmTo, sgSSFErrorMsg(ilError)
                                        If gOkAddStrToListBox(sgSSFErrorMsg(ilError), llLen, False) Then
                                            lbcError.AddItem sgSSFErrorMsg(ilError)
                                        Else
                                            If Not ilSeeMsg Then
                                                lbcError.AddItem "Out of room for messages, see output file"
                                                ilSeeMsg = True
                                            End If
                                        End If
                                    Next ilError
                                    DoEvents
                                    If Not ilSvRet Then
                                        Print #hmTo, "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                                        Close #hmTo
                                        If Not ilTestOnly Then
                                            gCloseSchFiles
                                        End If
                                        imProcessing = False
                                        Screen.MousePointer = vbDefault
                                        Exit Sub
                                    End If
                                End If
                                ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        Else
                            ilType = 0
                            ilRet = gMakeSSF(ilTestOnly, hmSsf, hmSdf, hmSmf, ilType, ilVefCode, ilLogDate0, ilLogDate1, 0)
                            ilSvRet = ilRet
                            For ilError = LBound(sgSSFErrorMsg) To UBound(sgSSFErrorMsg) - 1 Step 1
                                Print #hmTo, sgSSFErrorMsg(ilError)
                                If gOkAddStrToListBox(sgSSFErrorMsg(ilError), llLen, False) Then
                                    lbcError.AddItem sgSSFErrorMsg(ilError)
                                Else
                                    If Not ilSeeMsg Then
                                        lbcError.AddItem "Out of room for messages, see output file"
                                        ilSeeMsg = True
                                    End If
                                End If
                            Next ilError
                            DoEvents
                            If Not ilSvRet Then
                                Print #hmTo, "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                                Close #hmTo
                                If Not ilTestOnly Then
                                    gCloseSchFiles
                                End If
                                imProcessing = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                        End If
                    End If
                    '7/6/05- Move error messages
                    '7/6/05- Moved error messages before testing return
'                    For ilError = LBound(sgSSFErrorMsg) To UBound(sgSSFErrorMsg) - 1 Step 1
'                        Print #hmTo, sgSSFErrorMsg(ilError)
'                        If gOkAddStrToListBox(sgSSFErrorMsg(ilError), llLen, False) Then
'                            lbcError.AddItem sgSSFErrorMsg(ilError)
'                        Else
'                            If Not ilSeeMsg Then
'                                lbcError.AddItem "Out of room for messages, see output file"
'                                ilSeeMsg = True
'                            End If
'                        End If
'                    Next ilError
                    llMissed = UBound(lgReschSdfCode) - LBound(lgReschSdfCode)
                    If llMissed > 0 Then
                        If ilTestOnly Then
                            lacMissed.Caption = Trim$(str$(llMissed)) & " will be set to Missed"
                        Else
                            lacMissed.Caption = Trim$(str$(llMissed)) & " set to Missed"
                        End If
                    End If
                    DoEvents
                    If imCancelled Then
                        Print #hmTo, "Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                        Close #hmTo
                        If Not ilTestOnly Then
                            gCloseSchFiles
                        End If
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    'Test Two- Check that all Sdf spots are within Ssf as of date specified
                    If ilTestOnly Then
                        lacTest.Caption = slDate & "- Checking if Sdf and Ssf Match"
                    Else
                        lacTest.Caption = slDate & "- Synchronizing Sdf and Ssf"
                    End If
                    ilType = 0
                    ilRet = gSynchSdfSsf(ilTestOnly, hmSsf, hmSdf, hmSmf, hmVef, hmCHF, hmClf, hmCff, hmRdf, hmCrf, hmCgf, ilType, ilVefCode, ilLogDate0, ilLogDate1, slMOrO, hmSxf, hmGsf, hmGhf)
                    '7/6/05- moved output of messages prior to testing return status
                    ilSvRet = ilRet
                    For ilError = LBound(sgSSFErrorMsg) To UBound(sgSSFErrorMsg) - 1 Step 1
                        Print #hmTo, sgSSFErrorMsg(ilError)
                        If gOkAddStrToListBox(sgSSFErrorMsg(ilError), llLen, False) Then
                            lbcError.AddItem sgSSFErrorMsg(ilError)
                        Else
                            If Not ilSeeMsg Then
                                lbcError.AddItem "Out of room for messages, see output file"
                                ilSeeMsg = True
                            End If
                        End If
                    Next ilError
                    DoEvents
                    If Not ilSvRet Then
                        '7/6/05- Clear message and insert message into error list box
                        lacTest.Caption = ""
                        lbcError.AddItem "Error in gSynchSdfSsf"
                        Print #hmTo, "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                        Close #hmTo
                        If Not ilTestOnly Then
                            gCloseSchFiles
                        End If
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
'                    For ilError = LBound(sgSSFErrorMsg) To UBound(sgSSFErrorMsg) - 1 Step 1
'                        Print #hmTo, sgSSFErrorMsg(ilError)
'                        If gOkAddStrToListBox(sgSSFErrorMsg(ilError), llLen, False) Then
'                            lbcError.AddItem sgSSFErrorMsg(ilError)
'                        Else
'                            If Not ilSeeMsg Then
'                                lbcError.AddItem "Out of room for messages, see output file"
'                                ilSeeMsg = True
'                            End If
'                        End If
'                    Next ilError
                    llMissed = UBound(lgReschSdfCode) - LBound(lgReschSdfCode)
                    If llMissed > 0 Then
                        If ilTestOnly Then
                            lacMissed.Caption = Trim$(str$(llMissed)) & " will be set to Missed"
                        Else
                            lacMissed.Caption = Trim$(str$(llMissed)) & " set to Missed"
                        End If
                    End If
                    DoEvents
                    If imCancelled Then
                        Print #hmTo, "Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                        Close #hmTo
                        If Not ilTestOnly Then
                            gCloseSchFiles
                        End If
                        imProcessing = False
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    'If ilTestOnly Then
                        llNoProc = llNoProc + 1
                        'llPercent = (llVehProc * 100) \ llTotalVeh
                        llPercent = (llNoProc * CSng(100)) / llTotalNoToProc
                        If llPercent >= 99 Then
                            llPercent = 99
                        ElseIf llPercent < 0 Then
                            llPercent = 0
                        End If
                        plcGauge.Value = llPercent
                        DoEvents
                    'End If
                Next llDate
            Loop While (ilStartNoMissed < UBound(lgReschSdfCode)) And (Not ilTestOnly)
            'If Not ilTestOnly Then
            '    llNoProc = llNoProc + llEndDate - llStartDate + 1
            '    'llPercent = (llVehProc * 100) \ llTotalVeh
                llPercent = (llNoProc * CSng(100)) / llTotalNoToProc
                If llPercent >= 99 Then
                    llPercent = 99
                ElseIf llPercent < 0 Then
                    llPercent = 0
                End If
                plcGauge.Value = llPercent
                DoEvents
            'End If
            llMissed = UBound(lgReschSdfCode) - LBound(lgReschSdfCode)
            If ilTestOnly Then
                lbcVehs.List(ilLoop) = lbcVehs.List(ilLoop) & ", " & Trim$(str$(llMissed)) & " Will be Missed "
            Else
                lbcVehs.List(ilLoop) = lbcVehs.List(ilLoop) & ", Missed " & Trim$(str$(llMissed))
            End If
            If (ckcReschedule.Value = vbChecked) And (llMissed > 0) Then
                lacTest.Caption = "Rescheduling spots set to missed "
                ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
                If Not ilRet Then
                    Print #hmTo, "Error in Processing, Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                    Close #hmTo
                    If Not ilTestOnly Then
                        gCloseSchFiles
                    End If
                    imProcessing = False
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            DoEvents
            If imCancelled Then
                Print #hmTo, "Cancelled" & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Close #hmTo
                If Not ilTestOnly Then
                    gCloseSchFiles
                End If
                imProcessing = False
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            lacProcessing.Caption = ""
            lacMissed.Caption = ""
            lacTest.Caption = ""
            llVehProc = llVehProc + 1
            'llPercent = (llVehProc * 100) \ llTotalVeh
            'If llPercent >= 100 Then
            '    llPercent = 100
            'End If
            'plcGauge.Value = llPercent
        End If
    Next ilLoop
    Print #hmTo, ""
    Print #hmTo, "Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close #hmTo
    plcGauge.Value = 100
    DoEvents
    lacProcessing.Caption = "Processing- Done "
    Screen.MousePointer = vbDefault
    If Not ilTestOnly Then
        If (Len(sgMDBPath) > 0) And (Len(sgSDBPath) > 0) Then
            If sgMDBPath = sgDBPath Then
                ilRes = MsgBox("The Data Must be Copied from " & sgDBPath & " to " & sgSDBPath, vbOKOnly + vbQuestion, "Fix")
            Else
                ilRes = MsgBox("The Data Must be Copied from " & sgDBPath & " to " & sgMDBPath, vbOKOnly + vbQuestion, "Fix")
            End If
        End If
        gCloseSchFiles
    End If
    imProcessing = False
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Exit Sub
'cmcChkFixErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
Private Sub edcDate_GotFocus(Index As Integer)
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        tmcStart.Enabled = True
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub edcDate_KeyPress(Index As Integer, KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    'igParentRestarted = True
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase sgSSFErrorMsg
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    btrDestroy hmGhf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmMtf)
    btrDestroy hmMtf
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    
    Set SsfCheck = Nothing
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehs_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        mSetCommands
    End If
End Sub
Private Sub lbcVehs_GotFocus()
'    Screen.MousePointer = vbHourGlass
'    lbcVehs.Clear
'    mPopulate

'    mSetCommands      'Disable LinksDef if No Vehicles are selected
'    Screen.MousePointer = vbDefault
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
    imFirstActivate = True
    imTerminate = False
    imSetAll = True
    imAllClicked = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass

    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    SsfCheck.Height = lacDBPath.Top + 2 * lacDBPath.Height
    gCenterStdAlone SsfCheck
    SsfCheck.Show

'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gGetSchParameters
    Screen.MousePointer = vbHourglass
    lacDBPath.Caption = "D.B. Path=> " & sgDBPath
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    slDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slDate)
    edcDate(0).Text = Format$(gDateValue(slDate) + 1, "m/d/yy")
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    imVefRecLen = Len(tmVef)  'btrRecordLength(hlVef)  'Get and save record length
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
    hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    hmMtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMtf, "", sgDBPath & "Mtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        lgMtfNoRecs = btrRecords(hmMtf)
        btrDestroy hmMtf
    Else
        lgMtfNoRecs = 0
    End If
    hmSxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", SsfCheck
    On Error GoTo 0
    gObtainPostLogAvailCode 'Set igPLAnfCode
    If tgSpf.sPLMove = "M" Then
        rbcPLMove(0).Value = True
        plcPLMove.Visible = False
    ElseIf tgSpf.sPLMove = "O" Then
        rbcPLMove(1).Value = True
        plcPLMove.Visible = False
    End If
    'Moved to lbcVehs GotFocus so hourglass will show up will building list box
'    Screen.MousePointer = vbHourGlass
'    lbcVehs.Clear
'    mPopulate

'    mSetCommands      'Disable LinksDef if No Vehicles are selected
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
    slCommand = sgCommandStr    'Command$
    'ilPos = InStr(1, slCommand, "/CS", 1)
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (ilPos > 0) Then
    '    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
     '   ilTestSystem = False
     '   ilPos = InStr(1, slCommand, "/S:Test", 1)
    '    If ilPos > 0 Then
    '        ilTestSystem = True
    '    End If
    '    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    '    If ilPos > 0 Then
    '        ilTestSystem = False
    '    End If
    '    slCommand = "Traffic\Guide"
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get user name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone SsfCheck, slStr, ilTestSystem
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/25/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box              *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilVehCode As Integer
    Dim ilVehIndex As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slLogDate As String
    'ilRet = gPopUserVehicleBox(SsfCheck, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, lbcVehs, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(SsfCheck, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, lbcVehs, tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopUserVehicleBox: Vehicle)", SsfCheck
        On Error GoTo 0
        'Add last log date
        For ilLoop = 0 To lbcVehs.ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVehCode = Val(slCode)
            ilVehIndex = gVpfFind(SsfCheck, ilVehCode)
            gUnpackDate tgVpf(ilVehIndex).iLLD(0), tgVpf(ilVehIndex).iLLD(1), slLogDate
            If slLogDate <> "" Then
                lbcVehs.List(ilLoop) = lbcVehs.List(ilLoop) & " " & slLogDate
            End If
        Next ilLoop
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    Exit Sub
End Sub
'************************************************************
'          Procedure Name : mSetCommands
'
'    Created : 4/17/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments:  Set Control properties
'
'
'************************************************************
'
Private Sub mSetCommands()
    If lbcVehs.ListCount <= 0 Then
        cmcChkFix.Enabled = False
    Else
        If lbcVehs.SelCount > 0 Then
            cmcChkFix.Enabled = True
        Else
            cmcChkFix.Enabled = False
        End If
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

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload SsfCheck
    'End
    igManUnload = NO
End Sub
Private Sub rbcCF_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcCF(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then   'Check
            ckcReschedule.Value = vbUnchecked
            ckcReschedule.Enabled = False
            ckcClone.Value = vbUnchecked
            ckcClone.Enabled = False
            cmcChkFix.Caption = "C&heck"
        Else
            ckcReschedule.Enabled = False   'True
            If ckcAll.Value = vbChecked Then
                ckcClone.Enabled = True
            Else
                ckcClone.Value = vbUnchecked
                ckcClone.Enabled = False
            End If
            cmcChkFix.Caption = "&Fix"
        End If
    End If
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imFirstTime Then
        'mInitDDE
        imFirstTime = False
    End If
    Screen.MousePointer = vbHourglass
    lbcVehs.Clear
    mPopulate

    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
End Sub
Private Sub plcPLMove_Paint()
    plcPLMove.CurrentX = 0
    plcPLMove.CurrentY = 0
    plcPLMove.Print "Set Discrepant Post Log Moved Spots as"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Check/Fix Schedule Files"
End Sub
