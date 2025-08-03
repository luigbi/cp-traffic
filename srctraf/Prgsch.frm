VERSION 5.00
Begin VB.Form PrgSch 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5640
   ClientLeft      =   870
   ClientTop       =   1215
   ClientWidth     =   7215
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
   ScaleHeight     =   5640
   ScaleWidth      =   7215
   Begin VB.ListBox lbcNotSchd 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4290
      Width           =   7035
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2970
      TabIndex        =   9
      Top             =   3990
      Width           =   1350
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1170
      Top             =   5085
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
      Left            =   6030
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   5355
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   525
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
      Left            =   6690
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox plcPrgSch 
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
      Height          =   3390
      Left            =   255
      ScaleHeight     =   3330
      ScaleWidth      =   6585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   6645
      Begin VB.ListBox lbcVehs 
         Appearance      =   0  'Flat
         Height          =   2970
         Left            =   135
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   195
         Width           =   6300
      End
   End
   Begin VB.CommandButton cmcSchedule 
      Appearance      =   0  'Flat
      Caption         =   "&Schedule"
      Height          =   285
      Left            =   2115
      TabIndex        =   4
      Top             =   5280
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4005
      TabIndex        =   5
      Top             =   5280
      Width           =   1140
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1875
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1875
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   165
      Top             =   5190
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacChange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Changes to Events, Require Changes to Links"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2070
      TabIndex        =   3
      Top             =   3780
      Width           =   3090
   End
End
Attribute VB_Name = "PrgSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Prgsch.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Program Scheduling MODULE DEFINITIONS
'
'   Created : 4/25/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF
Dim imVefRecLen As Integer     'VEF record length
Dim tmVefSrchKey As INTKEY0  'Vef key record image
Dim tmAtt As ATT                'ATT record image
Dim tmATTSrchKey1 As INTKEY0     'ATT key 1 image
Dim imAttRecLen As Integer      'ATT record length
Dim hmAtt As Integer            'Agreement file handle

Dim hmMtf As Integer        'M for N Tracking
Dim tmVefSch() As VEFSCH
'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imFirstActivate As Integer
Dim imFirstTime As Integer
Dim imScheduling As Integer
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
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
    If lbcVehs.ListCount <= 0 Then
        Exit Sub
    End If
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehs.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehs.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    If imScheduling Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSchedule_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilVehWithoutLinks As Integer
    Dim ilVehWithAffiliate As Integer
    Dim ilVef As Integer
    Dim slType As String
    
    If imScheduling Then
        Exit Sub
    End If
    lbcNotSchd.Clear
    ilVehWithoutLinks = False
    For ilIndex = 0 To lbcVehs.ListCount - 1 Step 1
        If lbcVehs.Selected(ilIndex) Then
            For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
                If ilIndex = tmVefSch(ilLoop).iGroup Then
                    If tmVefSch(ilLoop).sType = "S" Then
                        'If all terminate dates = 0, then no links defined
                        If (tmVefSch(ilLoop).iTermDate0(0) = 0) And (tmVefSch(ilLoop).iTermDate0(1) = 0) And (tmVefSch(ilLoop).iTermDate6(0) = 0) And (tmVefSch(ilLoop).iTermDate6(1) = 0) And (tmVefSch(ilLoop).iTermDate7(0) = 0) And (tmVefSch(ilLoop).iTermDate7(1) = 0) Then
                            ilVehWithoutLinks = True
                        Else
                            tmVefSch(ilLoop).sOnAirSchStatus = "S"
                            tmVefSch(ilLoop).sAltSchStatus = "S"
                        End If
                    ElseIf tmVefSch(ilLoop).sType = "A" Then
                        'If all terminate dates = 0, then no links defined
                        If (tmVefSch(ilLoop).iTermDate0(0) = 0) And (tmVefSch(ilLoop).iTermDate0(1) = 0) And (tmVefSch(ilLoop).iTermDate6(0) = 0) And (tmVefSch(ilLoop).iTermDate6(1) = 0) And (tmVefSch(ilLoop).iTermDate7(0) = 0) And (tmVefSch(ilLoop).iTermDate7(1) = 0) Then
                            ilVehWithoutLinks = True
                        Else
                            tmVefSch(ilLoop).sOnAirSchStatus = "S"
                            tmVefSch(ilLoop).sAltSchStatus = "S"
                        End If
                    End If
                End If
            Next ilLoop
        End If
    Next ilIndex
'    If ilVehWithoutLinks Then
'        ilRet = MsgBox("Changes To Vehicle Events Without Link Changes, Schedule Anyway", vbYesNo + vbQuestion, "Links Missing")
'        If ilRet = vbNo Then
'            If ckcAll.Value = vbChecked Then
'                ckcAll.Value = vbUnchecked
'            Else
'                imAllClicked = True
'                ilValue = False
'                llRg = CLng(lbcVehs.ListCount - 1) * &H10000 Or 0
'                llRet = SendMessageByNum(lbcVehs.hwnd, LB_SELITEMRANGE, ilValue, llRg)
'                imAllClicked = False
'            End If
'            cmcCancel.SetFocus
'            imScheduling = False
'            Exit Sub
'        End If
'    End If
    Screen.MousePointer = vbHourglass
    gObtainMissedReasonCode
    If gOpenSchFiles() Then
        'Process selling and airing link changes with a terminate date defined in tmVefSch
        ilRet = gBuildVCF(tmVefSch(), lbcNotSchd)
        If ilRet Then
            For ilIndex = 0 To lbcVehs.ListCount - 1 Step 1
                If lbcVehs.Selected(ilIndex) Then
                    For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
                        If ilIndex = tmVefSch(ilLoop).iGroup Then
                            tmVefSch(ilLoop).sOnAirSchStatus = "S"
                            tmVefSch(ilLoop).sAltSchStatus = "S"
                        End If
                    Next ilLoop
                End If
            Next ilIndex
            'gBuildLCF checks for pending events for vehicle- if none the vehicle is bypassed
            ilRet = gBuildLCF(tmVefSch(), lbcNotSchd)
        End If
        gCloseSchFiles
    End If
    ilVehWithAffiliate = False
    If tgSpf.sGUseAffSys = "Y" Then
        For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
            'Exclude Sports as done in GameSchd
            ilVef = gBinarySearchVef(tmVefSch(ilLoop).iVefCode)
            If ilVef <> -1 Then
                slType = tgMVef(ilVef).sType
            Else
                slType = "G"
            End If
            If (tmVefSch(ilLoop).sOnAirSchStatus = "S") And (slType <> "G") Then
                tmATTSrchKey1.iCode = tmVefSch(ilLoop).iVefCode
                ilRet = btrGetEqual(hmAtt, tmAtt, imAttRecLen, tmATTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilVehWithAffiliate = True
                    Exit For
                End If
            End If
        Next ilLoop
    End If
    If ilVehWithoutLinks Then
        ilRet = MsgBox("Selling/Airing Vehicle Program Changes Scheduled, Check Links", vbOKOnly + vbExclamation, "Check Links")
    End If
    If ilVehWithAffiliate Then
        ilRet = MsgBox("Program Changes Scheduled, Advise Affiliate Department to Adjust Agreements", vbOKOnly + vbExclamation, "Check Links")
    End If
    lbcVehs.Clear
    ckcAll.Value = vbUnchecked
    DoEvents
    If imTerminate Then
        Screen.MousePointer = vbDefault
        mTerminate
        Exit Sub
    End If
    mPopulate

    mSetCommands      'Disable LinksDef if No Vehicles are selected
    cmcCancel.Caption = "&Done"
    imScheduling = False
    Screen.MousePointer = vbDefault
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    gGetSchParameters
    Dim ilRet As Integer
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmAtt)
    btrDestroy hmAtt

    Set PrgSch = Nothing

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
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        tmcStart.Enabled = True
        Screen.MousePointer = vbDefault
    End If
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
    imTerminate = False
    imFirstActivate = True
    imSetAll = True
    imAllClicked = False
    imScheduling = False
    Screen.MousePointer = vbHourglass

    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    'If tgSpf.sSSellNet = "Y" Then
    '    lacChange.Visible = True
    'Else
        lacChange.Visible = False
    'End If
    PrgSch.Height = cmcSchedule.Top + 5 * cmcSchedule.Height / 3
    gCenterStdAlone PrgSch
    'PrgSch.Show

'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gGetSchParameters
    ReDim tmVefSch(0 To 0) As VEFSCH
    sgMovePass = "N"
    sgCompPass = "N"
    sgPreemptPass = "N"
    Screen.MousePointer = vbHourglass
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PrgSch
    On Error GoTo 0
    imVefRecLen = Len(tmVef)  'btrRecordLength(hlVef)  'Get and save record length
    
    hmAtt = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAtt, "", sgDBPath & "Att.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", PrgSch
    On Error GoTo 0
    imAttRecLen = Len(tmAtt)
    
    hmMtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMtf, "", sgDBPath & "Mtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        lgMtfNoRecs = btrRecords(hmMtf)
    Else
        lgMtfNoRecs = 0
    End If
    btrDestroy hmMtf

    'Moved to lbcVehs GotFocus so hourglass will show up will building list box
'    Screen.MousePointer = vbHourGlass
'    lbcVehs.Clear
'    mPopulate

'    mSetCommands      'Disable LinksDef if No Vehicles are selected
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lbcNotSchd.Clear
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
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
    Dim slHelpSystem As String
    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    '    slCommand = "Traffic\Guide"
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get user name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
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
        imShowHelpMsg = True
        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
            imShowHelpMsg = False
        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone PrgSch, slStr, ilTestSystem
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
    Dim ilLoop As Integer
    Dim ilLargestGrpNo As Integer
    Dim slName As String
    Dim ilGroup As Integer
    Dim ilPending As Integer
    ilRet = gBuildVehSchInfo(True, True, tmVefSch())
    'Show all selling/airing first
    ilLargestGrpNo = -1
    For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
        If (tmVefSch(ilLoop).sOnAirSchStatus = "P") And (tmVefSch(ilLoop).sType = "S") And (tmVefSch(ilLoop).iGroup <> -1) Then
            If tmVefSch(ilLoop).iGroup > ilLargestGrpNo Then
                ilLargestGrpNo = tmVefSch(ilLoop).iGroup
            End If
        End If
    Next ilLoop
    For ilGroup = 1 To ilLargestGrpNo Step 1
        slName = ""
        'Show selling vehicles then airing
        ilPending = False
        For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
            If (tmVefSch(ilLoop).sType = "S") And (tmVefSch(ilLoop).iGroup = ilGroup) Then
                If tmVefSch(ilLoop).sOnAirSchStatus = "P" Then
                    ilPending = True
                End If
                tmVefSch(ilLoop).iGroup = lbcVehs.ListCount
                tmVefSrchKey.iCode = tmVefSch(ilLoop).iVefCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If slName = "" Then
                        slName = Trim$(tmVef.sName)
                    Else
                        slName = slName & ";" & Trim$(tmVef.sName)
                    End If
                End If
            End If
        Next ilLoop
        For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
            If (tmVefSch(ilLoop).sType = "A") And (tmVefSch(ilLoop).iGroup = ilGroup) Then
                If tmVefSch(ilLoop).sOnAirSchStatus = "P" Then
                    ilPending = True
                End If
                tmVefSch(ilLoop).iGroup = lbcVehs.ListCount
                tmVefSrchKey.iCode = tmVefSch(ilLoop).iVefCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If InStr(slName, "/") = 0 Then
                        slName = slName & "/" & Trim$(tmVef.sName)
                    Else
                        slName = slName & ";" & Trim$(tmVef.sName)
                    End If
                End If
            End If
        Next ilLoop
        If (slName <> "") And ilPending Then
            lbcVehs.AddItem slName
        End If
    Next ilGroup
    For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
        'If (tmVefSch(ilLoop).sOnAirSchStatus = "P") And (tmVefSch(ilLoop).sType <> "S") And (tmVefSch(ilLoop).sType <> "A") Then
        If (tmVefSch(ilLoop).sOnAirSchStatus = "P") And (tmVefSch(ilLoop).iGroup = -1) Then
            tmVefSch(ilLoop).iGroup = lbcVehs.ListCount
            tmVefSrchKey.iCode = tmVefSch(ilLoop).iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                lbcVehs.AddItem Trim$(tmVef.sName)
            End If
        End If
    Next ilLoop
    For ilLoop = LBound(tmVefSch) To UBound(tmVefSch) - 1 Step 1
        If (tmVefSch(ilLoop).sAltSchStatus = "P") And (tmVefSch(ilLoop).sType <> "S") And (tmVefSch(ilLoop).sType <> "A") Then
            tmVefSch(ilLoop).iGroup = lbcVehs.ListCount
            tmVefSrchKey.iCode = tmVefSch(ilLoop).iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                lbcVehs.AddItem Trim$(tmVef.sName)
            End If
        End If
    Next ilLoop
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
        cmcSchedule.Enabled = False
    Else
        If lbcVehs.SelCount > 0 Then
            cmcSchedule.Enabled = True
        Else
            cmcSchedule.Enabled = False
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
    Unload PrgSch
    'End
    igManUnload = NO
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imFirstTime Then
        imFirstTime = False
    End If
    Screen.MousePointer = vbHourglass
    lbcVehs.Clear
    mPopulate

    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print " Program Scheduling"
End Sub
