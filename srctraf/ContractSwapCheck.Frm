VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ContractSwapCheck 
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
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   11
      Top             =   390
      Width           =   1170
   End
   Begin VB.OptionButton rbcFix 
      Caption         =   "Fix"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   45
      Width           =   960
   End
   Begin VB.OptionButton rbcCheck 
      Caption         =   "Check"
      Height          =   255
      Left            =   6540
      TabIndex        =   8
      Top             =   45
      Value           =   -1  'True
      Width           =   960
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   240
      Left            =   165
      TabIndex        =   6
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
      Index           =   0
      Left            =   2775
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
      TabIndex        =   3
      Top             =   2820
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4935
      TabIndex        =   4
      Top             =   2820
      Width           =   1140
   End
   Begin VB.ListBox lbcError 
      Appearance      =   0  'Flat
      Height          =   1290
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1290
      Width           =   8160
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "or Contract # "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   4380
      TabIndex        =   12
      Top             =   435
      Width           =   1200
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   180
      TabIndex        =   10
      Top             =   1005
      Width           =   8160
   End
   Begin VB.Label lacScreen 
      Caption         =   "Contract Swap Check"
      Height          =   225
      Left            =   105
      TabIndex        =   7
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
      TabIndex        =   5
      Top             =   3225
      Width           =   8190
   End
   Begin VB.Label lacDate 
      Appearance      =   0  'Flat
      Caption         =   "Contract Active Start Date "
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   435
      Width           =   2100
   End
End
Attribute VB_Name = "ContractSwapCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey0 As LONGKEY0  'CHF key record image (contract #)
Dim tmChfSrchKey1 As CHFKEY1
Dim imCHFRecLen As Integer        'RVF record length

Dim hmClf As Integer
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey0 As CLFKEY0
Dim tmClfSrchKey2 As LONGKEY0  'CHF key record image (contract #)
Dim imClfRecLen As Integer        'RVF record length

Dim hmCff As Integer
Dim tmCff As CFF            'CFF record image
Dim tmCffSrchKey1 As LONGKEY0  'CHF key record image (contract #)
Dim imCffRecLen As Integer        'RVF record length

Private rst_Chf As ADODB.Recordset
Private rst_Clf As ADODB.Recordset
Private rst_Cff As ADODB.Recordset


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

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim slCntrNo As String
    Dim slName As String
    Dim ilTestOnly As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilSeeMsg As Integer

    ilTestOnly = True   'Check Only
    lbcError.Clear
    slStartDate = Trim$(edcDate(0).Text)
    If slStartDate <> "" Then
        If Not gValidDate(slStartDate) Then
            Beep
            MsgBox "Invalid Start Date ", vbOKOnly + vbInformation + vbApplicationModal, "Date Error"
            edcDate(0).SetFocus
            Exit Sub
        End If
    End If
    slCntrNo = Trim$(edcDate(1).Text)
    Screen.MousePointer = vbHourglass
    ilRet = 0
    'On Error GoTo cmcChkFixErr:
    slToFile = sgDBPath & "Messages\" & "ContractSwapCheck.csv"
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
    If bmChooseFix Then
        Print #hmTo, "Contract Swap Fix moded " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Else
        Print #hmTo, "Contract Swap Check moded " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    Print #hmTo, ""
    Print #hmTo, "Contract #,Start Date,End Date,Revision #"
    ilSeeMsg = False
    imProcessing = True
    imCancelled = False
    plcGauge.Value = 0
    ilRet = mContractSwapCheck(slStartDate, slCntrNo)
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
    If Index = 0 Then
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
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

Private Sub Form_Terminate()
    Dim ilRet As Integer
    
    On Error Resume Next
    
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    
    rst_Chf.Close
    rst_Clf.Close
    rst_Cff.Close

    igManUnload = YES
    Set ContractSwapCheck = Nothing   'Remove data segment
    igManUnload = NO

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
    ContractSwapCheck.Height = lacDBPath.Top + 2 * lacDBPath.Height
    gCenterStdAlone ContractSwapCheck
    'ContractSwapCheck.Show

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
    edcDate(0).Text = ""
    bmChooseFix = False
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ContractSwapCheck
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ContractSwapCheck
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ContractSwapCheck
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
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
    If edcDate(0).Text <> "" Then
        If Not gValidDate(edcDate(0).Text) Then
            cmcChkFix.Enabled = False
        Else
            cmcChkFix.Enabled = True
        End If
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
    Dim ilRet As Integer
    
    Screen.MousePointer = vbDefault
    Unload ContractSwapCheck
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

Function mContractSwapCheck(slStartDate As String, slCntrNo As String) As Integer
    Dim slSQLQuery As String
    Dim llTotalNoToProc As Long
    Dim llTotalNoChf As Long
    Dim llTotalNoClf As Long
    Dim llTotalNoCff As Long
    Dim llNoProc As Long
    Dim llPercent As Long
    Dim llPrevCntrNo As Long
    Dim ilLastSchCntRevNo As Integer
    Dim slDelete As String
    Dim ilRet As Integer
    Dim llSvChfCode As Long
    Dim slSvSchStatus As String
    Dim slSvDelete As String
    Dim slFilterCntrNo As String
    
    mContractSwapCheck = True
    slSQLQuery = "Select Count(*) From ("
    slSQLQuery = slSQLQuery & "Select distinct ChfCntrNo from CLF_Contract_Line "
    slSQLQuery = slSQLQuery & " Left Outer Join chf_contract_header On clfChfCode = chfCode "
    slSQLQuery = slSQLQuery & " where clfCntRevNo <> chfCntRevNo"
    If Trim$(slStartDate) <> "" Then
        slSQLQuery = slSQLQuery & " and chfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    End If
    If Trim$(slCntrNo) <> "" Then
        slSQLQuery = slSQLQuery & " and chfCntrNo = " & slCntrNo
    End If
    slSQLQuery = slSQLQuery & ")X"
    'Set rst_Chf = cnn.Execute(slSQLQuery)
    Set rst_Chf = gSQLSelectCall(slSQLQuery)
    llTotalNoChf = rst_Chf(0).Value
    
    slSQLQuery = "Select Count(*) From ("
    slSQLQuery = slSQLQuery & "Select distinct ChfCntrNo, chfCode, chfCntRevNo, chfStartDate, chfEndDate, clfCode from CLF_Contract_Line "
    slSQLQuery = slSQLQuery & " Left Outer Join chf_contract_header On clfChfCode = chfCode "
    slSQLQuery = slSQLQuery & " where clfCntRevNo <> chfCntRevNo"
    If Trim$(slStartDate) <> "" Then
        slSQLQuery = slSQLQuery & " and chfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    End If
    If Trim$(slCntrNo) <> "" Then
        slSQLQuery = slSQLQuery & " and chfCntrNo = " & slCntrNo
    End If
    slSQLQuery = slSQLQuery & ")X"
    'Set rst_Clf = cnn.Execute(slSQLQuery)
    Set rst_Clf = gSQLSelectCall(slSQLQuery)
    llTotalNoClf = rst_Clf(0).Value
    
    slSQLQuery = "Select Count(*) From ("
    slSQLQuery = slSQLQuery & "Select distinct chfCntrNo, chfCode, chfCntRevNo, chfStartDate, chfEndDate, cffCode from CFF_Contract_Flight "
    slSQLQuery = slSQLQuery & " Left Outer Join chf_contract_header On cffChfCode = chfCode "
    slSQLQuery = slSQLQuery & " where cffCntRevNo <> chfCntRevNo"
    If Trim$(slStartDate) <> "" Then
        slSQLQuery = slSQLQuery & " and chfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    End If
    If Trim$(slCntrNo) <> "" Then
        slSQLQuery = slSQLQuery & " and chfCntrNo = " & slCntrNo
    End If
    slSQLQuery = slSQLQuery & ")X"
    'Set rst_Cff = cnn.Execute(slSQLQuery)
    Set rst_Cff = gSQLSelectCall(slSQLQuery)
    llTotalNoCff = rst_Cff(0).Value
    llTotalNoToProc = llTotalNoChf + llTotalNoClf + llTotalNoCff
    llNoProc = 0
    
    llPrevCntrNo = -1

    slSQLQuery = "Select distinct chfCntrNo, chfCode, chfCntRevNo, chfStartDate, chfEndDate, clfCode from CLF_Contract_Line "
    slSQLQuery = slSQLQuery & " Left Outer Join chf_contract_header On clfChfCode = chfCode "
    slSQLQuery = slSQLQuery & " where clfCntRevNo <> chfCntRevNo"
    If Trim$(slStartDate) <> "" Then
        slSQLQuery = slSQLQuery & " and chfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    End If
    If Trim$(slCntrNo) <> "" Then
        slSQLQuery = slSQLQuery & " and chfCntrNo = " & slCntrNo
    End If
    slSQLQuery = slSQLQuery & " Order by chfCntrNo, ChfCntRevNo desc"
    'Set rst_Clf = cnn.Execute(slSQLQuery)
    Set rst_Clf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Clf.EOF
        If llPrevCntrNo <> rst_Clf!chfCntrNo Then
            ilLastSchCntRevNo = -1
            Print #hmTo, rst_Clf!chfCntrNo & "," & Format(rst_Clf!ChfStartDate, sgShowDateForm) & "," & Format(rst_Clf!ChfEndDate, sgShowDateForm) & "," & rst_Clf!chfCntRevNo
            tmChfSrchKey1.lCntrNo = rst_Clf!chfCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = rst_Clf!chfCntrNo) And (tmChf.sSchStatus <> "F")
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (tmChf.lCntrNo = rst_Clf!chfCntrNo) And (tmChf.sSchStatus = "F") Then
                ilLastSchCntRevNo = tmChf.iCntRevNo
            End If
            llPrevCntrNo = rst_Clf!chfCntrNo
            If slFilterCntrNo <> "" Then
                slFilterCntrNo = slFilterCntrNo & "," & rst_Clf!chfCntrNo
            Else
                slFilterCntrNo = rst_Clf!chfCntrNo
            End If
        End If
        If bmChooseFix Then
            'Read Line that matched
            tmClfSrchKey2.lCode = rst_Clf!clfCode
            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                'Read Contract matching the line with the wrong CntRevNo
                tmChfSrchKey1.lCntrNo = rst_Clf!chfCntrNo
                tmChfSrchKey1.iCntRevNo = tmClf.iCntRevNo
                tmChfSrchKey1.iPropVer = tmClf.iPropVer
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    'Delete the Line
                    ilRet = btrDelete(hmClf)
                    'Add the line
                    tmClf.lChfCode = tmChf.lCode
                    ilRet = btrInsert(hmClf, tmClf, imClfRecLen, INDEXKEY0)
                End If
            End If
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
        rst_Clf.MoveNext
    Loop
    
    llPrevCntrNo = -1
    slSQLQuery = "Select distinct chfCntrNo, chfCode, chfCntRevNo, chfStartDate, chfEndDate, cffCode from CFF_Contract_Flight "
    slSQLQuery = slSQLQuery & " Left Outer Join chf_contract_header On cffChfCode = chfCode "
    slSQLQuery = slSQLQuery & " where cffCntRevNo <> chfCntRevNo"
    If Trim$(slStartDate) <> "" Then
        slSQLQuery = slSQLQuery & " and chfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    End If
    If Trim$(slCntrNo) <> "" Then
        slSQLQuery = slSQLQuery & " and chfCntrNo = " & slCntrNo
    End If
    slSQLQuery = slSQLQuery & " Order by chfCntrNo, ChfCntRevNo desc"
    'Set rst_Cff = cnn.Execute(slSQLQuery)
    Set rst_Cff = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Cff.EOF
        If llPrevCntrNo <> rst_Cff!chfCntrNo Then
            ilLastSchCntRevNo = -1
            tmChfSrchKey1.lCntrNo = rst_Cff!chfCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = rst_Cff!chfCntrNo) And (tmChf.sSchStatus <> "F")
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (tmChf.lCntrNo = rst_Cff!chfCntrNo) And (tmChf.sSchStatus = "F") Then
                ilLastSchCntRevNo = tmChf.iCntRevNo
            End If
            llPrevCntrNo = rst_Cff!chfCntrNo
            If InStr(1, slFilterCntrNo, rst_Cff!chfCntrNo, vbTextCompare) <= 0 Then
                Print #hmTo, rst_Cff!chfCntrNo & "," & Format(rst_Cff!ChfStartDate, sgShowDateForm) & "," & Format(rst_Cff!ChfEndDate, sgShowDateForm) & "," & rst_Cff!chfCntRevNo
                If slFilterCntrNo <> "" Then
                    slFilterCntrNo = slFilterCntrNo & "," & rst_Clf!chfCntrNo
                Else
                    slFilterCntrNo = rst_Clf!chfCntrNo
                End If
            End If
        End If
        If bmChooseFix Then
            'Read Line that matched
            tmCffSrchKey1.lCode = rst_Cff!cffCode
            ilRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                'Read Contract matching the line with the wrong CntRevNo
                tmChfSrchKey1.lCntrNo = rst_Cff!chfCntrNo
                tmChfSrchKey1.iCntRevNo = tmCff.iCntRevNo
                tmChfSrchKey1.iPropVer = tmCff.iPropVer
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    'Delete the Line
                    ilRet = btrDelete(hmCff)
                    'Add the line
                    tmCff.lChfCode = tmChf.lCode
                    ilRet = btrInsert(hmCff, tmCff, imCffRecLen, INDEXKEY0)
                End If
            End If
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
        rst_Cff.MoveNext
    Loop
    
    If slFilterCntrNo = "" Then
        plcGauge.Value = 100
        DoEvents
        Exit Function
    End If
    If Not bmChooseFix Then
        plcGauge.Value = 100
        DoEvents
        Exit Function
    End If
    llPrevCntrNo = -1
    slSQLQuery = "Select distinct ChfCntrNo, chfCode, chfCntRevNo from "
    slSQLQuery = slSQLQuery & " chf_contract_header "
    If Trim$(slStartDate) <> "" Then
        slSQLQuery = slSQLQuery & " where chfEndDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
        'If Trim$(slCntrNo) <> "" Then
            slSQLQuery = slSQLQuery & " and chfCntrNo in (" & slFilterCntrNo & ")"
        'End If
    Else
        'If Trim$(slCntrNo) <> "" Then
            slSQLQuery = slSQLQuery & " where chfCntrNo  in (" & slFilterCntrNo & ")"
        'End If
    End If
    slSQLQuery = slSQLQuery & " Order by chfCntrNo, ChfCntRevNo desc"
    'Set rst_Chf = cnn.Execute(slSQLQuery)
    Set rst_Chf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Chf.EOF
        tmChfSrchKey0.lCode = rst_Chf!chfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmChf.sStatus = "O" Or tmChf.sStatus = "H" Then
                If (llPrevCntrNo = -1) Or (llPrevCntrNo <> tmChf.lCntrNo) Then
                    'Update Lines
                    mResetClfSchStatusAndDelete tmChf.lCode, "F", "N"
                    'Update flights
                    mResetCffSchStatusAndDelete tmChf.lCode, "N"
                    tmChf.sSchStatus = "F"
                    tmChf.sDelete = "N"
                    llPrevCntrNo = tmChf.lCntrNo
                Else
                    'Update Lines
                    mResetClfSchStatusAndDelete tmChf.lCode, "F", "Y"
                    'Update flights
                    mResetCffSchStatusAndDelete tmChf.lCode, "Y"
                    tmChf.sSchStatus = "F"
                    tmChf.sDelete = "Y"
                End If
                ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
            End If
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
        rst_Chf.MoveNext
    Loop
  
End Function

Private Sub mResetClfSchStatusAndDelete(llChfCode As Long, slSchStatus As String, slDelete As String)
    Dim slSQLQuery As String
    'Dim rst As ADODB.Recordset
    slSQLQuery = "Update clf_Contract_Line Set clfDelete = '" & slDelete & "'"
    If slSchStatus <> "" Then
        slSQLQuery = slSQLQuery & ", clfSchStatus = '" & slSchStatus & "'"
    End If
    slSQLQuery = slSQLQuery & " Where clfChfCode = " & llChfCode
    'Set rst = cnn.Execute(slSQLQuery)
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gHandleError "TrafficErrors.txt", "ContractSwapCheck: mResetClfSchStatusAndDelete"
    End If
End Sub

Private Sub mResetCffSchStatusAndDelete(llChfCode As Long, slDelete As String)
    Dim slSQLQuery As String
    'Dim rst As ADODB.Recordset
    slSQLQuery = "Update cff_Contract_Flight Set cffDelete = '" & slDelete & "'"
    slSQLQuery = slSQLQuery & " Where cffChfCode = " & llChfCode
    'Set rst = cnn.Execute(slSQLQuery)
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gHandleError "TrafficErrors.txt", "ContractSwapCheck: mResetCffSchStatusAndDelete"
    End If
End Sub
