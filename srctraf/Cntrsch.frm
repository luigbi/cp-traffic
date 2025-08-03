VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CntrSch 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5940
   ClientLeft      =   1020
   ClientTop       =   885
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
   Icon            =   "Cntrsch.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   7215
   Begin VB.Timer tmcAPI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   5520
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStatus 
      Height          =   3735
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6588
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   705
      Top             =   5250
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   195
      Left            =   2235
      TabIndex        =   14
      Top             =   5145
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcBkgd 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5310
      Top             =   5100
   End
   Begin VB.ListBox lbcNotSchd 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4050
      Width           =   7035
   End
   Begin VB.CheckBox ckcShow 
      Caption         =   "Show All"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3660
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Contracts"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2805
      TabIndex        =   9
      Top             =   3660
      Width           =   1425
   End
   Begin VB.ListBox lbcCntrCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5790
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   -15
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1125
      Top             =   5340
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
      Left            =   6330
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5445
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
      Left            =   5940
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
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
      Left            =   6630
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5220
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
      Left            =   240
      ScaleHeight     =   3330
      ScaleWidth      =   6585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   6645
      Begin VB.ListBox lbcCntr 
         Appearance      =   0  'Flat
         Height          =   2970
         Left            =   195
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
      Left            =   1995
      TabIndex        =   4
      Top             =   5565
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3975
      TabIndex        =   3
      Top             =   5565
      Width           =   1140
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1875
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1875
   End
   Begin VB.Label lacCntr 
      Alignment       =   2  'Center
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
      Height          =   240
      Left            =   435
      TabIndex        =   11
      Top             =   4830
      Width           =   6345
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   225
      Top             =   5430
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacCount 
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
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   5070
      Width           =   2070
   End
End
Attribute VB_Name = "CntrSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Cntrsch.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Contract Scheduling MODULE DEFINITIONS
'
'   Created : 5/02/94       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
'Dim hmChf As Integer        'Vehicle file handle
Dim tmChf As CHF
Dim imCHFRecLen As Integer     'CHF record length
'Module Status Flags
Dim imFirstActivate As Integer
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim imSchType As Integer        '0=Hold; 1=Selected contract; 2=Import; 3=Background
Dim lmCntrCode As Long          'Contract code number
Dim imFirstTime As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imScheduling As Integer
Dim imShowHelpMsg As Integer    'True=Show help messages; False=Ignore help message system
Dim lmSleepTime As Long
Dim smCntrNotSchd() As String
Dim smCntrSchdErrors() As String

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
        llRg = CLng(lbcCntr.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcCntr.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcShow_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcShow.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If Value Then
        mPopulate
        ckcAll.Visible = True
        ckcShow.Visible = False
        mSetCommands
    End If
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slMsg                                                   *
'******************************************************************************************
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim slCode As String
    Dim llCount As Long
    Dim llTotalCount As Long
    Dim llPercent As Long
    Dim slName As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim slItem() As String
    Dim slRes As String
    
    If imScheduling Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    '------------------------------------------
    'Megaphone API
    SetupStatusGrid
    grdStatus.Visible = True
    grdStatus.Enabled = False
    'Populate Status Grid
    llTotalCount = 0
    For ilIndex = 0 To lbcCntr.ListCount - 1 Step 1
        If lbcCntr.Selected(ilIndex) Then
            If grdStatus.TextMatrix(grdStatus.rows - 1, grdStatus_CntrNo_Col) <> "" Then
                grdStatus.AddItem ""
            End If
            llTotalCount = llTotalCount + 1
            ilRet = gParseItem(lbcCntrCode.List(ilIndex), 1, "\", slName)    'Get user name
            ilRet = gParseItem(lbcCntrCode.List(ilIndex), 2, "\", slCode)    'Get user name
            slItem = Split(slName, "|")
            grdStatus.TextMatrix(llTotalCount, grdStatus_DigitalCntr_Col) = "" 'Will be set later
            grdStatus.TextMatrix(llTotalCount, grdStatus_LineDetail_Col) = "" 'Will be set later
            grdStatus.TextMatrix(llTotalCount, grdStatus_CHFCode_Col) = slCode
            grdStatus.TextMatrix(llTotalCount, grdStatus_Advertiser_Col) = slItem(0)
            grdStatus.TextMatrix(llTotalCount, grdStatus_CntrNo_Col) = Val(slItem(2))
        End If
    Next ilIndex
    cmcSchedule.Enabled = False
    bgCntrSchError = False
    mGetMegaphoneAvfInfo
    sgAPIActivityLog = sgAPIActivityLog & "Megaphone AVF Code: " & igMegaphoneAvfCode & vbCrLf

    'tmcBkgd.Enabled = False
    imScheduling = True
    plcGauge.Visible = True
    plcGauge.Value = 0
    lacCount.Caption = ""
    lacCntr.Caption = ""
    gObtainMissedReasonCode
    'llTotalCount = 0
    lbcNotSchd.Clear
    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    If gOpenSchFiles() Then
        'For ilIndex = 0 To lbcCntr.ListCount - 1 Step 1
        '    If lbcCntr.Selected(ilIndex) Then
        '        llTotalCount = llTotalCount + 1
        '    End If
        'Next ilIndex
        llCount = 0
        
        'Megaphone API
        lacCntr.Caption = "Scheduling Contracts..."
        'For ilIndex = 0 To lbcCntr.ListCount - 1 Step 1
        '    If lbcCntr.Selected(ilIndex) Then
        
        'Megaphone API
        For ilIndex = 1 To grdStatus.rows - 1
            lgCntrSchGridRow = CLng(ilIndex)
            sgCntrSchStatus = ""
            mColorGridRow CLng(lgCntrSchGridRow), GridColor_Blue
            If ilIndex > 5 Then
                grdStatus.TopRow = ilIndex - 5
            End If
            sgCntrSchStatus = "Scheduling " & grdStatus.TextMatrix(ilIndex, grdStatus_Advertiser_Col) & "/" & grdStatus.TextMatrix(ilIndex, grdStatus_CntrNo_Col)
            sgAPIActivityLog = "--------------------------------------------------------" & vbCrLf
            sgAPIActivityLog = sgAPIActivityLog & sgCntrSchStatus & vbCrLf
            sgAPIActivityLog = sgAPIActivityLog & Format(Now, "MM/dd/yyyy hh:mm:ssAM/PM") & vbCrLf
            tmcAPI_Timer 'Updates the status grid
            'Add call to scheduler
            'ilRet = gParseItem(lbcCntrCode.List(ilIndex), 1, "\", slName)    'Get user name
            'ilRet = gParseItem(lbcCntrCode.List(ilIndex), 2, "\", slCode)    'Get user name
            'slStr = Trim$(lbcCntr.List(ilIndex))
            'ilPos = InStr(slStr, "&")
            'If ilPos > 0 Then
            '    slStr = Left$(slStr, ilPos - 1) & "&&" & Mid$(slStr, ilPos + 1)
            'End If
            'lacCntr.Caption = "Processing: " & slStr
            DoEvents
            If imTerminate Then
                gCloseSchFiles
                Screen.MousePointer = vbDefault
                mTerminate
                lacCntr.Caption = "Scheduling Canceled!"
                sgAPIActivityLog = sgAPIActivityLog & "Scheduling Canceled!" & vbCrLf
                'Exit Sub
                GoTo ExitScheduling
            End If
            sgApplyVehPreemptRule = "Y"
            sgVehPreemptRule = "N"
            
            'Megaphone API
            'ilRet = gSchCntr(Val(slCode), lbcNotSchd, slStr)
            ilRet = gSchCntr(Val(grdStatus.TextMatrix(ilIndex, grdStatus_CHFCode_Col)), lbcNotSchd, grdStatus.TextMatrix(ilIndex, grdStatus_Advertiser_Col) & "/" & grdStatus.TextMatrix(ilIndex, grdStatus_CntrNo_Col))
            
            'If this was a Megaphone Contract, Flag the row so we can email results
            grdStatus.TextMatrix(ilIndex, grdStatus_DigitalCntr_Col) = "0"
            If bgMegaphoneContract = True Then
                grdStatus.TextMatrix(ilIndex, grdStatus_DigitalCntr_Col) = igMegaphoneAvfCode
            Else
                'Not a API contract so, let's clear the logs
                sgAPIActivityLog = ""
            End If
            
            tmcAPI_Timer 'updates the status
            sgApplyVehPreemptRule = "N"
            sgVehPreemptRule = "N"
            
            If bgCntrSchError Then
                mColorGridRow lgCntrSchGridRow, GridColor_Red
                CNTRSCHD.mErrorMsg sgCntrSchStatus, lbcNotSchd
            Else
                mColorGridRow lgCntrSchGridRow, GridColor_Green
            End If
            
            If igSchdInBkgd = False Then
                If Not ilRet Then
                    'If bgCntrSchError Then
                    '    mColorGridRow lgCntrSchGridRow, GridColor_Red
                    '    CNTRSCHD.mErrorMsg sgCntrSchStatus, lbcNotSchd
                    '    mLogErrorMsg sgAPIActivityLog
                    'End If
                    tmcAPI_Timer
                    grdStatus.Enabled = True
                    gCloseSchFiles
                    Screen.MousePointer = vbDefault
                    
                    If bgCntrSchError Then
                        'if any API contracts failed
                        mColorGridRow lgCntrSchGridRow, GridColor_Red
                        CNTRSCHD.mErrorMsg sgCntrSchStatus, lbcNotSchd
                        sgAPIActivityLog = sgAPIActivityLog & "Scheduling Not Completed!" & vbCrLf
                        ilRet = MsgBox("Scheduling Not Completed-see log ScheduleAPI.Txt, Try Later", vbOKOnly + vbExclamation, "Scheduling")
                    Else
                        ilRet = MsgBox("Scheduling Not Completed-see Message Area for Error Messages, Try Later", vbOKOnly + vbExclamation, "Scheduling")
                    End If
                    
                    'mTerminate
                    lbcCntr.Clear
                    cmcCancel.SetFocus
                    imScheduling = False
                    lacCntr.Caption = "Scheduling not complete!"
                    CntrSch.mColorGridRow CLng(ilIndex), GridColor_Red 'problem
                    'Exit Sub
                    GoTo ExitScheduling
                End If
            Else
                mWriteErrors smCntrSchdErrors()
            End If
            DoEvents
            If imTerminate Then
                gCloseSchFiles
                Screen.MousePointer = vbDefault
                mTerminate
                lacCntr.Caption = "Scheduling Canceled!"
                sgAPIActivityLog = sgAPIActivityLog & "Scheduling Canceled!" & vbCrLf
                CntrSch.mColorGridRow CLng(ilIndex), GridColor_Red
                grdStatus.Enabled = True
                'Exit Sub
                GoTo ExitScheduling
            End If
            llCount = llCount + 1
            llPercent = (llCount * CSng(100)) / llTotalCount
            If llPercent >= 100 Then
                If llCount = llTotalCount Then
                    plcGauge.Value = 100
                Else
                    plcGauge.Value = 99
                End If
            Else
                plcGauge.Value = llPercent
            End If
            lacCount.Caption = Trim$(str$(llCount)) & " of" & str$(llTotalCount)
            'End If
        Next ilIndex
        plcGauge.Visible = False
        
        'Process preempted spots
        lacCntr.Caption = "Processing Preempted Spots..."
        'sgAPIActivityLog = sgAPIActivityLog & lacCntr.Caption & vbCrLf
        lacCount.Caption = ""
        DoEvents
        If imTerminate Then
            gCloseSchFiles
            Screen.MousePointer = vbDefault
            mTerminate
            lacCntr.Caption = "Scheduling Canceled!"
            sgAPIActivityLog = sgAPIActivityLog & "Scheduling Canceled!" & vbCrLf
            CntrSch.mColorGridRow CLng(ilIndex), GridColor_Red
            grdStatus.Enabled = True
            'Exit Sub
            GoTo ExitScheduling
        End If
        sgApplyVehPreemptRule = "Y"
        sgVehPreemptRule = "N"
        ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
        sgApplyVehPreemptRule = "N"
        sgVehPreemptRule = "N"
        If igSchdInBkgd = False Then
            If Not ilRet Then
                gCloseSchFiles
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Scheduling Not Completed-see Message Area for Error Messages, Try Later", vbOKOnly + vbExclamation, "Scheduling")
                'mTerminate
                lbcCntr.Clear
                cmcCancel.SetFocus
                imScheduling = False
                lacCntr.Caption = "Scheduling not complete!"
                sgAPIActivityLog = sgAPIActivityLog & "Scheduling not complete!" & vbCrLf
                CntrSch.mColorGridRow CLng(ilIndex), GridColor_Red
                grdStatus.Enabled = True
                'Exit Sub
                GoTo ExitScheduling
            End If
        End If
        lacCntr.Caption = ""
        gCloseSchFiles
        plcGauge.Visible = True
        'lbcCntr.Clear
        If imSchType = 0 Then   'Hold
            'mPopulate
            For ilIndex = lbcCntr.ListCount - 1 To 0 Step -1
                If lbcCntr.Selected(ilIndex) Then
                    lbcCntr.RemoveItem ilIndex
                    lbcCntrCode.RemoveItem ilIndex
                End If
            Next ilIndex
            mSetCommands      'Disable LinksDef if No Vehicles are selected
            cmcCancel.Caption = "&Done"
        ElseIf imSchType = 1 Then   'Selected
            Screen.MousePointer = vbDefault
            'mTerminate
            'Exit Sub
            GoTo ExitScheduling
        ElseIf imSchType = 3 Then   'Background
            'tmcBkgd.Enabled = True
        End If
    End If
    
ExitScheduling:
    imScheduling = False
    
    'Send an Email if Digital Contracts processed, to users setup to receive digital contract emails
    slRes = mSendAPIEmail
    If InStr(1, slRes, "Error:") > 0 Then
        'Try Again
        slRes = mSendAPIEmail
        If InStr(1, slRes, "Error:") > 0 Then
            MsgBox "Error Sending Email!" & slRes & vbCrLf & "Use the content in ScheduleAPI.Txt to reconstruct the message and email manually.", vbExclamation, "Email Digital API Summary Error"
        End If
    End If
    
    If lbcNotSchd.ListCount > 0 Then
        lacCntr.Caption = "Scheduling incomplete!  Check list for errors."
    Else
        lacCntr.Caption = "Scheduling complete!"
    End If
    If cmcCancel.Enabled = True And cmcCancel.Visible = True Then cmcCancel.SetFocus
    
    If sgAPIActivityLog <> "" Then
        mLogAPIMsg sgAPIActivityLog
    End If
    
    grdStatus.Enabled = True
    Screen.MousePointer = vbDefault
    
    'Self Terminate only when No Errors and Scheduling one contract (Called by another screen)
    If imSchType = 1 And lbcNotSchd.ListCount = 0 And bgCntrSchError = False Then  'Selected
        mTerminate
    End If
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
    If imSchType = 1 Then   'Selected
        'ckcShow.Visible = True
        ckcAll.Visible = True   'False
    Else
        'ckcShow.Visible = False
        ckcAll.Visible = True
    End If
    Me.KeyPreview = True
    CntrSch.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub

Private Sub Form_Load()
    gSetBkgdMode
    If igBkgdProg = 10 Then
        igBkgdProg = 1
        igSchdInBkgd = True
    Else
        igSchdInBkgd = False
    End If
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase smCntrNotSchd
    Erase smCntrSchdErrors
    gRemoveLockCntr
    If igSchdInBkgd Then
        tmcSetTime.Enabled = False
        gCloseTmf
        gCloseVpf
        
        Set CntrSch = Nothing

        btrStopAppl
        End
    End If
    Set CntrSch = Nothing
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcCntr_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        mSetCommands
    End If
End Sub

Private Sub lbcCntr_GotFocus()
    If imFirstTime Then
        Screen.MousePointer = vbHourglass
        tmcStart.Enabled = True
        Screen.MousePointer = vbDefault
    End If
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
    lmSleepTime = 10000
    Screen.MousePointer = vbHourglass
    ReDim smCntrNotSchd(0 To 0) As String
    ReDim smCntrSchdErrors(0 To 0) As String
    grdStatus.Visible = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    If igSchdInBkgd Then
        CntrSch.height = cmcSchedule.Top + 3 * cmcSchedule.height
    Else
        CntrSch.height = cmcSchedule.Top + 5 * cmcSchedule.height / 3
    End If
    gCenterStdAlone CntrSch
    'CntrSch.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    'hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    'ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'On Error GoTo mInitErr
    'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CntrSch
    'On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'btrRecordLength(hlChf)  'Get and save record length
        
    lbcNotSchd.Clear
    'Obtain all vehicle options
    ilRet = gVpfRead()

    gGetSchParameters

    Screen.MousePointer = vbDefault
    If imTerminate Then
        Exit Sub
    End If
    
    Exit Sub

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
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim ilRet As Integer
'    igSchdInBkgd = False
    igDoEvent = False
    igDirectCall = 0
    sgIniPath = ""
'    slCommand = Command$
'    If StrComp(slCommand, "Debug", 1) = 0 Then
'        igStdAloneMode = True 'False  'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = False
'        imShowHelpMsg = False
'        slCommand = "Traffic\Guide\Hold"
'    ElseIf InStr(1, slCommand, "/CS", 1) > 0 Then
    If igSchdInBkgd = True Then
        igStdAloneMode = True 'False  'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
        imShowHelpMsg = False
        'If InStr(1, slCommand, "/Background", 1) > 0 Then
        '    igSchdInBkgd = True
        'End If
        'If InStr(1, slCommand, "/TimeSlice", 1) > 0 Then
            igDoEvent = True
        'End If
        slCommand = "Traffic\Guide\Hold"
    Else
        slCommand = sgCommandStr
        igStdAloneMode = False  'Switch from/to stand alone mode
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
    End If
    If igSchdInBkgd Then
        gInitStdAlone CntrSch, slStr, ilTestSystem
    End If
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source(Hold or C then contract # or Import)
    If igSchdInBkgd Then
        imSchType = 3
        CntrSch.Caption = "Scheduling in Background"
        'CntrSch.MinButton = True
        'CntrSch.MaxButton = True
        'CntrSch.WindowState = 1
    ElseIf Left$(slStr, 1) = "#" Then
        slStr = right$(slStr, Len(slStr) - 1)   'Remove C
        lmCntrCode = Val(slStr)
        imSchType = 1
    ElseIf StrComp(slStr, "Hold", 1) = 0 Then
        imSchType = 0
    ElseIf StrComp(slStr, "Import", 1) = 0 Then
        imSchType = 2
    Else
        imSchType = -1
    End If
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
    Dim ilRet As Integer 'btrieve status

    ilRet = gPopCntrSchBox(CntrSch, 0, lbcCntr, lbcCntrCode, lbcNotSchd)
    If ilRet <> CP_MSG_NOPOPREQ Then
        If imSchType <> 3 Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gPopCntrSchBox)", CntrSch
            On Error GoTo 0
        End If
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
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
Private Sub mSetCommands()
    If lbcCntr.ListCount <= 0 Then
        cmcSchedule.Enabled = False
    Else
        If lbcCntr.SelCount > 0 Then
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
'*            Comments: terminate Contract scheduling  *
'*                                                     *
'*******************************************************
Private Sub mTerminate()

    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload CntrSch
    'End
    igManUnload = NO
End Sub

Private Sub lbcNotSchd_DblClick()
    MsgBox lbcNotSchd.List(lbcNotSchd.ListIndex)
End Sub

Private Sub tmcAPI_Timer()
    If bgApiDelay > 0 Then
        bgApiDelay = bgApiDelay - 100
    End If
    If sgCntrSchStatus <> "" And grdStatus.cols >= grdStatus_Status_Col Then
        grdStatus.TextMatrix(lgCntrSchGridRow, grdStatus_Status_Col) = sgCntrSchStatus
        grdStatus.TextMatrix(lgCntrSchGridRow, grdStatus_LineDetail_Col) = sgCntrLineDetail
    End If
End Sub

Private Sub tmcBkgd_Timer()
'    Dim ilLoop As Integer

    tmcBkgd.Enabled = False
'    mPopulate
'    For ilLoop = lbcCntr.ListCount - 1 To 0 Step -1
'        lbcCntr.Selected(ilLoop) = True
'    Next ilLoop
'    mWriteErrors smCntrNotSchd()
''        llRg = CLng(lbcCntr.ListCount - 1) * &H10000 + 0
''        llRet = SendMessage(lbcCntr.Hwnd, &H400 + 28, True, llRg)
'    mSetCommands      'Disable LinksDef if No Vehicles are selected
'    If cmcSchedule.Enabled Then
'        cmcSchedule_Click
'    Else
'        tmcBkgd.Enabled = True
'    End If
    gOpenTmf
    gOpenVpf
    tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
    tmcSetTime.Enabled = True
    mBkgdTask
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "CSS"
End Sub

Private Sub tmcStart_Timer()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    tmcStart.Enabled = False
    If imFirstTime Then
        imFirstTime = False
    Else
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lbcCntrCode.Clear
    lbcCntr.Clear
    If imSchType = 0 Then   'Hold
        mPopulate
        mSetCommands      'Disable LinksDef if No Vehicles are selected
    ElseIf imSchType = 1 Then   'Select contract
        mPopulate
        lbcCntr.Visible = False
        'Debug.Print "looking for " & lmCntrCode
        For ilLoop = 0 To lbcCntrCode.ListCount - 1 Step 1
            slNameCode = lbcCntrCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'Debug.Print slNameCode
            If Val(slCode) = lmCntrCode Then
                lbcCntr.Selected(ilLoop) = True
                Exit For
            End If
        Next ilLoop
        lbcCntr.Visible = True
        'tmChfSrchKey.lCode = lmCntrCode
        'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        'If ilRet <> BTRV_ERR_NONE Then
        '    cmcCancel.SetFocus
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If
        'lbcCntr.AddItem Trim$(Str$(tmChf.lCntrNo))
        'lbcCntrCode.AddItem Trim$(Str$(tmChf.lCntrNo)) & "\" & Trim$(Str$(tmChf.lCode))
        'lbcCntr.ListIndex = 0
        ''Wait until list box value shown
        'For ilLoop = 1 To 20 Step 1
        '    DoEvents
        'Next ilLoop
        mSetCommands      'Disable LinksDef if No Vehicles are selected
        'cmcSchedule.SetFocus
        'SendKeys "{Enter}"
    ElseIf imSchType = 2 Then   'Import
        mPopulate
        For ilLoop = lbcCntr.ListCount - 1 To 0 Step -1
            lbcCntr.Selected(ilLoop) = True
        Next ilLoop
'        llRg = CLng(lbcCntr.ListCount - 1) * &H10000 + 0
'        llRet = SendMessage(lbcCntr.Hwnd, &H400 + 28, True, llRg)
        mSetCommands      'Disable LinksDef if No Vehicles are selected
        If cmcSchedule.Enabled Then
            cmcSchedule.SetFocus
        Else
            cmcCancel.SetFocus
        End If
    ElseIf imSchType = 3 Then   'Background
        tmcBkgd.Enabled = True
        CntrSch.WindowState = 1
    Else
        cmcCancel.SetFocus
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print " Contract Scheduling"
End Sub

Private Sub mWriteErrors(slLastErrors() As String)
    Dim ilLoop As Integer
    Dim slMsg As String
    Dim ilFound As Integer
    Dim ilTest As Integer
    For ilLoop = 0 To lbcNotSchd.ListCount - 1 Step 1
        slMsg = lbcNotSchd.List(ilLoop)
        ilFound = False
        For ilTest = 0 To UBound(slLastErrors) - 1 Step 1
            If StrComp(slLastErrors(ilTest), slMsg, vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilTest
        If Not ilFound Then
            sgTmfStatus = "E"
            gLogMsg slMsg, "BkgdSchdErrors.Txt", False
        End If
    Next ilLoop
    ReDim slLastErrors(0 To 0) As String
    For ilLoop = 0 To lbcNotSchd.ListCount - 1 Step 1
        slMsg = lbcNotSchd.List(ilLoop)
        slLastErrors(UBound(slLastErrors)) = slMsg
        ReDim Preserve slLastErrors(0 To UBound(slLastErrors) + 1) As String
    Next ilLoop
End Sub

Private Sub mBkgdTask()
    Dim ilLoop As Integer
    Do
        Sleep lmSleepTime
        If imTerminate Then
            mTerminate
            Exit Sub
        End If
        For ilLoop = 0 To 100 Step 1
            DoEvents
        Next ilLoop
        mPopulate
        For ilLoop = lbcCntr.ListCount - 1 To 0 Step -1
            lbcCntr.Selected(ilLoop) = True
        Next ilLoop
        mWriteErrors smCntrNotSchd()
    '        llRg = CLng(lbcCntr.ListCount - 1) * &H10000 + 0
    '        llRet = SendMessage(lbcCntr.Hwnd, &H400 + 28, True, llRg)
        mSetCommands      'Disable LinksDef if No Vehicles are selected
        If cmcSchedule.Enabled Then
            gUpdateTaskMonitor 1, "CSS"
            cmcSchedule_Click
            gRemoveLockCntr
            gUpdateTaskMonitor 2, "CSS"
        End If
    Loop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       SetupStatusGrid
' Description:       Clears and Sets up grdStatus with headers and column widths
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       3/8/2024-10:46:50
' Parameters :
'--------------------------------------------------------------------------------
Sub SetupStatusGrid()
    grdStatus.cols = 6
    grdStatus.FixedCols = 0
    grdStatus.rows = 2
    
    'Make headers
    grdStatus.TextMatrix(0, grdStatus_DigitalCntr_Col) = "Vendor"
    grdStatus.TextMatrix(0, grdStatus_LineDetail_Col) = "Line Details"
    grdStatus.TextMatrix(0, grdStatus_CHFCode_Col) = "CHFCode"
    grdStatus.TextMatrix(0, grdStatus_Advertiser_Col) = "Advertiser"
    grdStatus.TextMatrix(0, grdStatus_CntrNo_Col) = "Contract"
    grdStatus.TextMatrix(0, grdStatus_Status_Col) = "Status"
    
    'Col Width
    grdStatus.ColWidth(grdStatus_DigitalCntr_Col) = 0
    grdStatus.ColWidth(grdStatus_LineDetail_Col) = 0
    grdStatus.ColWidth(grdStatus_CHFCode_Col) = 0
    grdStatus.ColWidth(grdStatus_Advertiser_Col) = 1240
    grdStatus.ColWidth(grdStatus_CntrNo_Col) = 1110
    grdStatus.ColWidth(grdStatus_Status_Col) = 4200
    
    'Alignment
    grdStatus.ColAlignment(grdStatus_CntrNo_Col) = flexAlignLeftCenter
    grdStatus.ColAlignment(grdStatus_Status_Col) = flexAlignLeftCenter
    
    'Clear 1st row
    grdStatus.TextMatrix(1, grdStatus_DigitalCntr_Col) = ""
    grdStatus.TextMatrix(1, grdStatus_LineDetail_Col) = ""
    grdStatus.TextMatrix(1, grdStatus_CHFCode_Col) = ""
    grdStatus.TextMatrix(1, grdStatus_Advertiser_Col) = ""
    grdStatus.TextMatrix(1, grdStatus_CntrNo_Col) = ""
    grdStatus.TextMatrix(1, grdStatus_Status_Col) = ""
    
    'Dont allow focus
    grdStatus.Enabled = False
    tmcAPI.Enabled = True
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mColorGridRow
' Description:       Color a specified row with a certain color in grdStatus
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       3/8/2024-10:45:50
' Parameters :       llRow (Long)
'                    iGridColor (Integer)
'--------------------------------------------------------------------------------
Sub mColorGridRow(llRow As Long, iGridColor As Integer)
    Dim ilLoop As Integer
    grdStatus.Row = llRow
    If grdStatus.cols < 4 Then Exit Sub
    For ilLoop = grdStatus_Advertiser_Col To grdStatus_Status_Col
        grdStatus.Col = ilLoop
        If iGridColor = 0 Then grdStatus.CellBackColor = RGB(255, 255, 255) 'white (UnSelected)
        If iGridColor = 1 Then grdStatus.CellBackColor = RGB(209, 248, 255) 'blue (Selected)
        If iGridColor = 2 Then grdStatus.CellBackColor = RGB(161, 255, 110) 'green (Scheduled)
        If iGridColor = 3 Then grdStatus.CellBackColor = RGB(250, 202, 191) 'red (Problem)
    Next ilLoop
End Sub

'--------------------------------------------------------------------------------
' Procedure  :       mSendAPIEmail
' Description:       email results of the contract push to Vendor API
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       3/8/2024-10:44:28
' Parameters :
'--------------------------------------------------------------------------------
Function mSendAPIEmail() As String
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    Dim slEMail  As String
    Dim slEmailAddresses As String
    Dim slErrorMessage As String
    Dim slFromAddress As String
    Dim slClientName As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim tlSrchKey0 As INTKEY0
    Dim ilRet As Integer
    Dim hlSite As Integer
    Dim tlSite As SITE
    Dim blFound As Boolean 'At least one recipient found
    Dim olEmailer As CEmail
    Dim slBody As String
    
    If grdStatus.rows < 2 Then
        Exit Function
    End If
    
    slBody = mGenerateEmailBody
    If slBody = "" Then
        mSendAPIEmail = "N/A"
        Exit Function
    End If
    
    lacCntr.Caption = "Sending Email..."
    lacCntr.Refresh
    
    slClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
        Else
            tlSrchKey0.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hlMnf, tlMnf, Len(tlMnf), tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slClientName = Trim$(tlMnf.sName)
            End If
        End If
    End If
    
    Set olEmailer = New CEmail
    
    If Not olEmailer Is Nothing Then
        With olEmailer
            'Get EmailAddresses
            slSQLQuery = "SELECT cefComment FROM CEF_Comments_Events WHERE cefCode in ( SELECT urfEMailCefCode FROM URF_User_Options WHERE urfDigitalCntrAlert='Y' AND urfEMailCefCode <> 0 AND urfDelete <>'Y' )"
            Set rst_Temp = gSQLSelectCall(slSQLQuery)
            While Not rst_Temp.EOF
                slEMail = Trim(rst_Temp!cefComment)
                If slEMail <> "" Then
                    blFound = True
                    .AddTOAddress slEMail
                    If slEmailAddresses <> "" Then slEmailAddresses = slEmailAddresses & ","
                    slEmailAddresses = slEmailAddresses & slEMail
                    If Len(.ErrorMessage) > 0 Then
                        slErrorMessage = .ErrorMessage
                        GoTo ExitSendEmail
                    End If
                End If
                rst_Temp.MoveNext
            Wend
            rst_Temp.Close
            
            'Found at least one recipeint?
            If blFound = True Then
                sgAPIActivityLog = sgAPIActivityLog & "Emailing: " & slEmailAddresses & vbCrLf
                sgAPIActivityLog = sgAPIActivityLog & slBody
                
                .FromAddress = Trim(tgSite.sEmailAcctName)
                '.FromName = Trim(tgSite.sEmailFromName)
                .FromName = "Counterpoint"
                .Subject = "Digital Contracts schedule summary " & Format(Now, "MM/DD/YY hh:mmAMPM")
                .Message = slBody
                
                If Not .Send() Then
                    slErrorMessage = "Send failed.  " & .ErrorMessage
                    sgAPIActivityLog = sgAPIActivityLog & "Error: " & slErrorMessage & vbCrLf
                    mSendAPIEmail = "Error: Send Failed"
                End If
            Else
                slErrorMessage = "No Email addresses to send to"
                sgAPIActivityLog = sgAPIActivityLog & "Error: " & slErrorMessage & vbCrLf
                mSendAPIEmail = "N/A"
                Exit Function
            End If
        End With
    Else
        sgAPIActivityLog = sgAPIActivityLog & "Emailer not installed!" & vbCrLf
        mSendAPIEmail = "Error: Emailer not installed!"
    End If
    
ExitSendEmail:
    If slErrorMessage <> "" Then
        mSendAPIEmail = "Error:" & slErrorMessage
    Else
        mSendAPIEmail = "Success"
    End If
End Function

'--------------------------------------------------------------------------------
' Procedure  :       mGenerateEmailBody
' Description:       Generate Email Body containing summary of all Digital Line that are associated with a Vendor API from grdStatus
' Created by :       J. White
' Machine    :       CSI-FM
' Date-Time  :       3/8/2024-10:43:06
' Parameters :
'--------------------------------------------------------------------------------
Function mGenerateEmailBody() As String
    Dim slString As String
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim blMadeHeader As Boolean
    slString = ""
    
    'Email header
    slString = slString & "Summary of digital contracts sent via API:" & vbCrLf
    
    'grdStatus rows - Successful
    blMadeHeader = False
    For ilLoop = 1 To grdStatus.rows - 1
        If Val(grdStatus.TextMatrix(ilLoop, grdStatus_DigitalCntr_Col)) <> 0 Then
            If InStr(1, grdStatus.TextMatrix(ilLoop, grdStatus_Status_Col), "Scheduled:") > 0 Then
                If blMadeHeader = False Then
                    slString = slString & "----Successful---" & vbCrLf
                    blMadeHeader = True
                End If
                blFound = True
                slString = slString & grdStatus.TextMatrix(ilLoop, grdStatus_Advertiser_Col) & "/" & grdStatus.TextMatrix(ilLoop, grdStatus_CntrNo_Col)
                slString = slString & ": " & grdStatus.TextMatrix(ilLoop, grdStatus_Status_Col) & vbCrLf
                If grdStatus.TextMatrix(ilLoop, grdStatus_LineDetail_Col) <> "" Then
                    slString = slString & " -> " & grdStatus.TextMatrix(ilLoop, grdStatus_LineDetail_Col) & vbCrLf
                End If
            End If
        End If
    Next ilLoop
    If blMadeHeader Then slString = slString & vbCrLf
    
    'grdStatus rows - unsuccessful
    blMadeHeader = False
    For ilLoop = 1 To grdStatus.rows - 1
        If Val(grdStatus.TextMatrix(ilLoop, grdStatus_DigitalCntr_Col)) <> 0 Then
            If InStr(1, grdStatus.TextMatrix(ilLoop, grdStatus_Status_Col), "Scheduled:") = 0 Then
                If blMadeHeader = False Then
                    slString = slString & "----Unsuccessful---" & vbCrLf
                    blMadeHeader = True
                End If
                blFound = True
                slString = slString & grdStatus.TextMatrix(ilLoop, grdStatus_Advertiser_Col) & "/" & grdStatus.TextMatrix(ilLoop, grdStatus_CntrNo_Col)
                slString = slString & ": " & grdStatus.TextMatrix(ilLoop, grdStatus_Status_Col) & vbCrLf
                If grdStatus.TextMatrix(ilLoop, grdStatus_LineDetail_Col) <> "" Then
                    If grdStatus.TextMatrix(ilLoop, grdStatus_LineDetail_Col) <> "" Then slString = slString & " -> " & grdStatus.TextMatrix(ilLoop, grdStatus_LineDetail_Col) & vbCrLf
                End If
            End If
        End If
    Next ilLoop
    
    If blFound = True Then
        mGenerateEmailBody = slString
    End If
End Function
