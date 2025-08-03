VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CffCheck 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3615
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
   ScaleHeight     =   3615
   ScaleWidth      =   8580
   Begin VB.OptionButton rbcFix 
      Caption         =   "Fix"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   45
      Width           =   960
   End
   Begin VB.OptionButton rbcCheck 
      Caption         =   "Check"
      Height          =   255
      Left            =   6540
      TabIndex        =   10
      Top             =   45
      Value           =   -1  'True
      Width           =   960
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   240
      Left            =   165
      TabIndex        =   8
      Top             =   2685
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
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
      Left            =   5325
      MaxLength       =   10
      TabIndex        =   3
      Top             =   390
      Width           =   1170
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
      Left            =   2190
      MaxLength       =   10
      TabIndex        =   1
      Top             =   390
      Width           =   1170
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6375
      Top             =   2760
   End
   Begin VB.CommandButton cmcChkFix 
      Appearance      =   0  'Flat
      Caption         =   "C&heck"
      Height          =   285
      Left            =   3060
      TabIndex        =   5
      Top             =   2820
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4935
      TabIndex        =   6
      Top             =   2820
      Width           =   1140
   End
   Begin VB.ListBox lbcError 
      Appearance      =   0  'Flat
      Height          =   1290
      Left            =   180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1290
      Width           =   8160
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   180
      TabIndex        =   12
      Top             =   1005
      Width           =   8160
   End
   Begin VB.Label lacScreen 
      Caption         =   "Contract Flight Check"
      Height          =   225
      Left            =   105
      TabIndex        =   9
      Top             =   60
      Width           =   2115
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
      TabIndex        =   7
      Top             =   3225
      Width           =   8190
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Sunday End Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   3750
      TabIndex        =   2
      Top             =   435
      Width           =   1410
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Monday Start Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   435
      Width           =   1575
   End
End
Attribute VB_Name = "CffCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CffCheck.Frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Type Defs (Removed)                                                            *
'*  CgfChanged                                                                            *
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
Dim hmCHF As Integer
Dim hmClf As Integer
Dim hmCff As Integer

Dim tmChf As CHF            'CHF record image
'Contract record information
Dim tmClfList() As CLFLIST            'CLF record image
'Contract record information
Dim tmCfflist() As CFFLIST            'CFF record image
Dim tmCffSrchKey As CFFKEY0 'CFF key record image

Dim tmCgfList() As CGFLIST

Dim tmChfAdvtExt() As CHFADVTEXT


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
Dim bmChooseFix As Boolean

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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llDate                        llStartDate                   llEndDate                 *
'*  slDate                        llTotalVeh                    llVehProc                 *
'*  llPercent                     llTotalNoToProc               llNoProc                  *
'*  ilError                       slMsg                         llLen                     *
'*  ilRes                         ilSvRet                       ilVef                     *
'*                                                                                        *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slName As String
    Dim ilTestOnly As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilSeeMsg As Integer

    ilTestOnly = True   'Check Only
    lbcError.Clear
    slStartDate = edcDate(0).Text
    If Not gValidDate(slStartDate) Then
        Beep
        MsgBox "Invalid Start Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
        edcDate(0).SetFocus
        Exit Sub
    End If
    If gWeekDayStr(slStartDate) <> 0 Then
        Beep
        MsgBox "Start Date must be at Monday ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
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
        If gWeekDayStr(slEndDate) <> 6 Then
            Beep
            MsgBox "End Date must be at Sunday ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
            edcDate(1).SetFocus
            Exit Sub
        End If
    Else
        slEndDate = "12/31/2069"
    End If
    Screen.MousePointer = vbHourglass
    ilRet = 0
    'On Error GoTo cmcChkFixErr:
    slToFile = sgDBPath & "Messages\" & "CffCheck.csv"
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
    Print #hmTo, "Contract Flight Check " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    Print #hmTo, "Contract,Line #,Vehicle,Week Start Date,Week End Date,Week Start Date,Week End Date"
    ilSeeMsg = False
    imProcessing = True
    imCancelled = False
    plcGauge.Value = 0
    ilRet = mCffCheck(slStartDate, slEndDate)
    If imCancelled Then
        Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close #hmTo
        lacProcessing.Caption = "Processing- Cancelled "
    Else
        Print #hmTo, ""
        Print #hmTo, "Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close #hmTo
        plcGauge.Value = 100
        DoEvents
        lacProcessing.Caption = "Processing- Done "
    End If
    cmcCancel.Caption = "Done"
    If Not ilRet Then
        ilRet = MsgBox("Results stored into " & slToFile, vbOKOnly + vbInformation, "Error")
    End If
    Screen.MousePointer = vbDefault
    imProcessing = False
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Exit Sub
'cmcChkFixErr:
'    ilRet = Err.Number
'    Resume Next
End Sub

Private Sub edcDate_Change(Index As Integer)
    mSetCommands
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
    sgDoneMsg = ""
    
    Erase tmClfList
    Erase tmCfflist
    Erase tmCgfList
    Erase tmChfAdvtExt
    
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    
    Set CffCheck = Nothing

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
    CffCheck.Height = lacDBPath.Top + 2 * lacDBPath.Height
    gCenterStdAlone CffCheck
    'CffCheck.Show

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
    bmChooseFix = False
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CffCheck
    On Error GoTo 0
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CffCheck
    On Error GoTo 0
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CffCheck
    On Error GoTo 0
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                         slTestSystem              *
'*  ilTestSystem                                                                          *
'******************************************************************************************

    Dim slCommand As String
    slCommand = sgCommandStr    'Command$
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
    If Not gValidDate(edcDate(0).Text) Then
        cmcChkFix.Enabled = False
    Else
        cmcChkFix.Enabled = True
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
    Unload CffCheck
    'End
    igManUnload = NO
End Sub



Private Sub rbcCheck_Click()
    cmcChkFix.Caption = "Check"
    bmChooseFix = False
End Sub

Private Sub rbcFix_Click()
    cmcChkFix.Caption = "Fix"
    bmChooseFix = True
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imFirstTime Then
        'mInitDDE
        imFirstTime = False
    End If
    Screen.MousePointer = vbHourglass
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Screen.MousePointer = vbDefault
End Sub

Function mCffCheck(slStartDate As String, slEndDate As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  blFoundACgf                                                                           *
'******************************************************************************************

    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilChf As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilTest As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llTStartDate As Long
    Dim llTEndDate As Long
    Dim ilRet As Integer
    Dim ilFlightRemoved As Integer
    Dim ilVefIndex As Integer
    Dim slVehicleName As String
    Dim llNoProc As Long
    Dim llTotalNoToProc As Long
    Dim llPercent As Long
    Dim blFixOk As Boolean
    Dim slCffStartDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    
    slCntrStatus = "HO"
    slCntrType = "C"
    ilHOType = 1
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(CffCheck, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Function
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    llNoProc = 0
    llTotalNoToProc = UBound(tmChfAdvtExt)
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChfAdvtExt(ilChf).lCode, False, tmChf, tmClfList(), tmCfflist())
        If ilRet Then
            lacProcessing.Caption = "Processing contract " & tmChfAdvtExt(ilChf).lCntrNo
            For ilClf = LBound(tmClfList) To UBound(tmClfList) - 1 Step 1
                ilVefIndex = gBinarySearchVef(tmClfList(ilClf).ClfRec.iVefCode)
                If ilVefIndex <> -1 Then
                    slVehicleName = Trim$(tgMVef(ilVefIndex).sName)
                Else
                    slVehicleName = ""
                End If
                gUnpackDateLong tmClfList(ilClf).ClfRec.iStartDate(0), tmClfList(ilClf).ClfRec.iStartDate(1), llSDate
                gUnpackDateLong tmClfList(ilClf).ClfRec.iEndDate(0), tmClfList(ilClf).ClfRec.iEndDate(1), llEDate
                If llSDate <= llEDate Then
                    ilFlightRemoved = False
                    For ilCff = LBound(tmCfflist) To UBound(tmCfflist) - 1 Step 1
                        If tmClfList(ilClf).ClfRec.iLine = tmCfflist(ilCff).CffRec.iClfLine Then
                            If (tmCfflist(ilCff).CffRec.iStartDate(0) <> -1) Or (tmCfflist(ilCff).CffRec.iStartDate(1) <> -1) Then
                                gUnpackDateLong tmCfflist(ilCff).CffRec.iStartDate(0), tmCfflist(ilCff).CffRec.iStartDate(1), llStartDate
                                gUnpackDateLong tmCfflist(ilCff).CffRec.iEndDate(0), tmCfflist(ilCff).CffRec.iEndDate(1), llEndDate
                                For ilTest = ilCff + 1 To UBound(tmCfflist) - 1 Step 1
                                    If tmClfList(ilClf).ClfRec.iLine = tmCfflist(ilTest).CffRec.iClfLine Then
                                        gUnpackDateLong tmCfflist(ilTest).CffRec.iStartDate(0), tmCfflist(ilTest).CffRec.iStartDate(1), llTStartDate
                                        gUnpackDateLong tmCfflist(ilTest).CffRec.iEndDate(0), tmCfflist(ilTest).CffRec.iEndDate(1), llTEndDate
                                        If (llStartDate <= llTEndDate) And (llTStartDate <= llEndDate) Then
                                            Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llStartDate, "m/d/yy") & "," & Format$(llEndDate, "m/d/yy") & "," & Format(llTStartDate, "m/d/yy") & "," & Format$(llTEndDate, "m/d/yy")
                                            lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & slVehicleName & "," & Format(llStartDate, "m/d/yy") & "," & Format$(llEndDate, "m/d/yy") & "," & Format(llTStartDate, "m/d/yy") & "," & Format$(llTEndDate, "m/d/yy")
                                            If bmChooseFix Then
                                                ilRet = btrBeginTrans(hmCff, 1000)
                                                If (llStartDate <= llTStartDate) And (llEndDate >= llTEndDate) Then
                                                    ilRet = mFixCffByDelete(tmCfflist(ilTest))
                                                    If ilRet <> BTRV_ERR_NONE Then
                                                        ilRet = btrAbortTrans(hmCff)
                                                        Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed & "," & slVehicleName & "," & "Error=" & ilRet
                                                        lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed
                                                    Else
                                                        ilRet = btrEndTrans(hmCff)
                                                        tmCfflist(ilTest).CffRec.iStartDate(0) = -1
                                                        tmCfflist(ilTest).CffRec.iStartDate(1) = -1
                                                        ilFlightRemoved = True
                                                    End If
                                                ElseIf (llTStartDate <= llStartDate) And (llTEndDate >= llEndDate) Then
                                                    ilRet = mFixCffByDelete(tmCfflist(ilCff))
                                                    If ilRet <> BTRV_ERR_NONE Then
                                                        ilRet = btrAbortTrans(hmCff)
                                                        Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed & "," & slVehicleName & "," & "Error=" & ilRet
                                                        lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed
                                                    Else
                                                        ilRet = btrEndTrans(hmCff)
                                                        tmCfflist(ilCff).CffRec.iStartDate(0) = -1
                                                        tmCfflist(ilCff).CffRec.iStartDate(1) = -1
                                                        ilFlightRemoved = True
                                                    End If
                                                    Exit For
                                                Else
                                                    If llStartDate < llTStartDate Then
                                                        slCffStartDate = Format$(llEndDate, "m/d/yy")
                                                        slCffStartDate = gObtainNextMonday(slCffStartDate)
                                                        ilRet = mModifyCffDates(ilTest, slCffStartDate)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            ilRet = btrAbortTrans(hmCff)
                                                            Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed & "," & slVehicleName & "," & "Error=" & ilRet
                                                            lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed
                                                        Else
                                                            ilRet = btrEndTrans(hmCff)
                                                        End If
                                                    Else
                                                        slCffStartDate = Format$(llTEndDate, "m/d/yy")
                                                        slCffStartDate = gObtainNextMonday(slCffStartDate)
                                                        ilRet = mModifyCffDates(ilCff, slCffStartDate)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            ilRet = btrAbortTrans(hmCff)
                                                            Print #hmTo, tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed & "," & slVehicleName & "," & "Error=" & ilRet
                                                            lbcError.AddItem tmChfAdvtExt(ilChf).lCntrNo & "," & tmClfList(ilClf).ClfRec.iLine & "," & Not Fixed
                                                        Else
                                                            ilRet = btrEndTrans(hmCff)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilTest
                            End If
                        End If
                    Next ilCff
                    If ilFlightRemoved Then
                        Do While ilCff < UBound(tmCfflist)
                            If (tmCfflist(ilCff).CffRec.iStartDate(0) = -1) And (tmCfflist(ilCff).CffRec.iStartDate(1) = -1) Then
                                For ilTest = ilCff To UBound(tmCfflist) - 1 Step 1
                                    tmCfflist(ilTest) = tmCfflist(ilTest + 1)
                                Next ilTest
                                ReDim Preserve tmCfflist(LBound(tmCfflist) To UBound(tmCfflist) - 1) As CFFLIST
                            Else
                                ilCff = ilCff + 1
                            End If
                        Loop
                    End If
                    
                End If
                If imCancelled Then
                    Exit Function
                End If
                'If ilCntrError Then
                '    Exit For
                'End If
            Next ilClf
        End If
        If imCancelled Then
            Exit Function
        End If
        llNoProc = llNoProc + 1
        llPercent = (llNoProc * CSng(100)) / llTotalNoToProc
        If llPercent >= 99 Then
            llPercent = 99
        ElseIf llPercent < 0 Then
            llPercent = 0
        End If
        plcGauge.Value = llPercent
        DoEvents
    Next ilChf
End Function




Private Function mFixCffByDelete(tlCffList As CFFLIST) As Integer
    Dim ilRet As Integer
    Dim tlCff As CFF
    Dim ilRecLen As Integer
    
    tmCffSrchKey.lChfCode = tlCffList.CffRec.lChfCode
    tmCffSrchKey.iClfLine = tlCffList.CffRec.iClfLine
    tmCffSrchKey.iCntRevNo = tlCffList.CffRec.iCntRevNo
    tmCffSrchKey.iPropVer = tlCffList.CffRec.iPropVer
    tmCffSrchKey.iStartDate(0) = tlCffList.CffRec.iStartDate(0)
    tmCffSrchKey.iStartDate(1) = tlCffList.CffRec.iStartDate(1)
    
    ilRecLen = Len(tlCff)
    ilRet = btrGetEqual(hmCff, tlCff, ilRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        ilRet = btrDelete(hmCff)
    End If
    mFixCffByDelete = ilRet
End Function



Private Function mModifyCffDates(ilCffIndex As Integer, slMonday As String) As Integer
    ' change cff record's dates to match new gsf dates
    ' input ilcffindex- Record to change. Start date to change to. index of clf.
    'output.  Change cff file   Also changes tmcffSrchKey.
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim tlCff As CFF
    Dim llNewSunday As Long
    
    ilRecLen = Len(tlCff)
    Do
        tmCffSrchKey.lChfCode = tmCfflist(ilCffIndex).CffRec.lChfCode
        tmCffSrchKey.iClfLine = tmCfflist(ilCffIndex).CffRec.iClfLine
        tmCffSrchKey.iCntRevNo = tmCfflist(ilCffIndex).CffRec.iCntRevNo
        tmCffSrchKey.iPropVer = tmCfflist(ilCffIndex).CffRec.iPropVer
        tmCffSrchKey.iStartDate(0) = tmCfflist(ilCffIndex).CffRec.iStartDate(0)
        tmCffSrchKey.iStartDate(1) = tmCfflist(ilCffIndex).CffRec.iStartDate(1)
        ilRet = btrGetEqual(hmCff, tlCff, ilRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            'Can not modify key: Therefore, delete and insert record
            ilRet = btrDelete(hmCff)
            If ilRet <> BTRV_ERR_NONE Then
                mModifyCffDates = ilRet
                Exit Function
            End If
            gPackDate slMonday, tlCff.iStartDate(0), tlCff.iStartDate(1)
            tlCff.lCode = 0
            ilRet = btrInsert(hmCff, tlCff, ilRecLen, INDEXKEY1)
            If ilRet = BTRV_ERR_NONE Then
                tmCfflist(ilCffIndex).CffRec = tlCff
                mModifyCffDates = ilRet
                Exit Function
            End If
        Else
            mModifyCffDates = ilRet
            Exit Function
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    mModifyCffDates = ilRet
End Function



