VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form AvgToPay 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5565
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
   ScaleHeight     =   5565
   ScaleWidth      =   8580
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   240
      Left            =   3000
      TabIndex        =   4
      Top             =   4485
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6375
      Top             =   4800
   End
   Begin VB.CommandButton cmcChkFix 
      Appearance      =   0  'Flat
      Caption         =   "&Fix"
      Height          =   285
      Left            =   2835
      TabIndex        =   1
      Top             =   4860
      Width           =   1140
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4710
      TabIndex        =   2
      Top             =   4860
      Width           =   1140
   End
   Begin VB.ListBox lbcMsg 
      Appearance      =   0  'Flat
      Height          =   3390
      ItemData        =   "AvgToPay.frx":0000
      Left            =   180
      List            =   "AvgToPay.frx":0002
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   705
      Width           =   8160
   End
   Begin VB.Label lacZeroBalance 
      Caption         =   "Prior to running this fix, run Collection Zero-Purge"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   195
      TabIndex        =   7
      Top             =   360
      Width           =   4410
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   1005
      Width           =   8160
   End
   Begin VB.Label lacScreen 
      Caption         =   "Days to Pay"
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   60
      Width           =   1305
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
      TabIndex        =   3
      Top             =   5265
      Width           =   8190
   End
End
Attribute VB_Name = "AvgToPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'**********************************************************
'                Day to Pay MODULE DEFINITIONS
'
'   Created : 10/7/19       By : D. LeVine
'   Modified :              By :
'
'**********************************************************
Option Explicit
Option Compare Text
Dim hmTo As Integer   'From file hanle

Private rst_Rvf As ADODB.Recordset
Private rst_Phf As ADODB.Recordset
Private rst_adf As ADODB.Recordset
Private rst_Agf As ADODB.Recordset
Private rst_Tmp As ADODB.Recordset

'Module Status Flags
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imChgMode As Integer        'Change mode status (so change not entered when in change)
Dim imBSMode As Integer         'Backspace flag
Dim lmNowDate As Long   'Todays date
Dim imFirstTime As Integer
Dim imCancelled As Integer
Dim imProcessing As Integer
Dim imFirstActivate As Integer
Private Type AVGTOPAYINFO
    iCode As Integer
    iCurLstToPay As Integer
    iCurAvgToPay As Integer
    iCurNoInvPd As Integer
    iNewLstToPay As Integer
    iNewAvgToPay As Integer
    iNewNoInvPd As Integer
End Type
Dim tmAdfAvgToPayInfo() As AVGTOPAYINFO
Dim tmAgfAvgToPayInfo() As AVGTOPAYINFO


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
    Dim slEndDate As String
    Dim slName As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilSeeMsg As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim ilMonth As Integer

    lbcMsg.Clear
    Screen.MousePointer = vbHourglass
    ilRet = 0
    'On Error GoTo cmcChkFixErr:
    slToFile = sgDBPath & "Messages\" & "AvgToPay.csv"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo cmcChkFixErr:
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            cmcCancel.SetFocus
            Exit Sub
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
            MsgBox "Open " & slToFile & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Print #hmTo, "Days to Pay " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    Print #hmTo, "Name,Days to Pay,Days to Pay,Difference,# of Invoices,# of Invoices,Difference,Last # Days to Pay,Last # Days to Pay,Difference"
    Print #hmTo, ",Original,Fixed,,Original,Fixed,,Original,Fixed,"
    ilSeeMsg = False
    imProcessing = True
    imCancelled = False
    plcGauge.Value = 0
    ilRet = mAvgToPay()
    If imCancelled Then
        Print #hmTo, "Cancelled " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        lacProcessing.Caption = "Processing- Cancelled "
    Else
        plcGauge.Value = 100
        DoEvents
        lacProcessing.Caption = "Processing- Done "
        Print #hmTo, "Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    End If
    Close #hmTo
    cmcChkFix.Enabled = False
    cmcCancel.Caption = "Done"
    Screen.MousePointer = vbDefault
    If Not ilRet Then
        ilRet = MsgBox("Results stored into " & slToFile, vbOKOnly + vbInformation, "Information")
    End If
    imProcessing = False
    mSetCommands      'Disable LinksDef if No Vehicles are selected
    Exit Sub
'cmcChkFixErr:
'    ilRet = Err.Number
'    Resume Next
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
    rst_Rvf.Close
    rst_Phf.Close
    rst_adf.Close
    rst_Agf.Close
    rst_Tmp.Close

    Set AvgToPay = Nothing

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
    Dim slStr As String
    
    imFirstActivate = True
    imTerminate = False
    imCancelled = False
    imProcessing = False
    Screen.MousePointer = vbHourglass

    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    AvgToPay.Height = lacDBPath.Top + 2 * lacDBPath.Height
    gCenterStdAlone AvgToPay
    'AvgToPay.Show

'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Screen.MousePointer = vbHourglass
    lacDBPath.Caption = "D.B. Path=> " & sgDBPath
    imChgMode = False
    imBSMode = False
    imFirstTime = True
    slDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slDate)
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
    cmcChkFix.Enabled = True
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
    Unload AvgToPay
    'End
    igManUnload = NO
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

Function mAvgToPay() As Integer
    Dim slSQLQuery As String
    Dim llInvDate As Long
    Dim llUpper As Long
    Dim ilAdf As Integer
    Dim ilAgf As Integer
    Dim ilLoop As Integer
    Dim llTotalNoToProc As Long
    Dim llNoProc As Long
    Dim ilIndex As Integer
    
    ReDim tmAdfAvgToPayInfo(0 To 0) As AVGTOPAYINFO
    ReDim tmAgfAvgToPayInfo(0 To 0) As AVGTOPAYINFO
    llTotalNoToProc = 0
    llNoProc = 0
    slSQLQuery = "Select Count(1) from ADF_Advertisers Where adfCode In (Select distinct phfAdfCode from PHF_Payment_History Where phfTranType = 'PI' And phfCashTrade = 'C')"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    If Not rst_adf.EOF Then
        llTotalNoToProc = 2 * rst_adf(0).Value
    End If
    slSQLQuery = "Select Count(1) from AGF_Agencies Where agfCode In (Select distinct phfAgfCode from PHF_Payment_History Where phfTranType = 'PI' And phfCashTrade = 'C' And phfAgfCode > 0)"
    Set rst_Agf = gSQLSelectCall(slSQLQuery)
    If Not rst_Agf.EOF Then
        llTotalNoToProc = llTotalNoToProc + 2 * rst_Agf(0).Value
    End If
    slSQLQuery = "Select distinct rvfCntrNo,rvfInvNo, sum (rvfNet) as sumnet from RVF_Receivables group by rvfCntrNo, rvfInvNo"
    Set rst_Rvf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Rvf.EOF
        llTotalNoToProc = llTotalNoToProc + 1
        rst_Rvf.MoveNext
    Loop
    
    slSQLQuery = "Select distinct phfAdfCode, phfagfCode, phfInvNo from phf_Payment_History Where phfCashTrade = 'C' And phfTranType = 'PI'"
    Set rst_Phf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Phf.EOF
        llTotalNoToProc = llTotalNoToProc + 1
        rst_Phf.MoveNext
    Loop
    lbcMsg.AddItem "Getting Advertiser Values"
    'Get current values (adfLstToPay, adfAvgToPay adfNoInvPd) and save in array
    slSQLQuery = "Select distinct adfCode, adfLstToPay, adfAvgToPay, adfNoInvPd from ADF_Advertisers Where (adfCode In (Select distinct phfAdfCode from PHF_Payment_History Where phfTranType = 'PI' And phfCashTrade = 'C') Or adfCode In (Select distinct rvfAdfCode from RVF_Receivables Where rvfTranType = 'PI' And rvfCashTrade = 'C')) Order by UCase(adfName)"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_adf.EOF
        llUpper = UBound(tmAdfAvgToPayInfo)
        tmAdfAvgToPayInfo(llUpper).iCode = rst_adf!adfCode
        tmAdfAvgToPayInfo(llUpper).iCurLstToPay = rst_adf!adfLstToPay
        tmAdfAvgToPayInfo(llUpper).iCurAvgToPay = rst_adf!adfAvgToPay
        tmAdfAvgToPayInfo(llUpper).iCurNoInvPd = rst_adf!adfNoInvPd
        tmAdfAvgToPayInfo(llUpper).iNewLstToPay = 0
        tmAdfAvgToPayInfo(llUpper).iNewAvgToPay = 0
        tmAdfAvgToPayInfo(llUpper).iNewNoInvPd = 0
        ReDim Preserve tmAdfAvgToPayInfo(0 To llUpper + 1) As AVGTOPAYINFO
        mSetGauge llNoProc, llTotalNoToProc
        rst_adf.MoveNext
    Loop
    lbcMsg.List(0) = "Getting Advertiser Values->Done"
    lbcMsg.AddItem "Getting Agency Values"
    slSQLQuery = "Select distinct agfCode, agfLstToPay, agfAvgToPay, agfNoInvPd from AGF_Agencies Where (agfCode In (Select distinct phfAgfCode from PHF_Payment_History Where phfTranType = 'PI' And phfCashTrade = 'C' And phfAgfCode > 0)) or (agfCode In (Select distinct rvfAgfCode from RVF_Receivables Where rvfTranType = 'PI' And rvfCashTrade = 'C' And rvfAgfCode > 0)) Order by UCase(agfName)"
    Set rst_Agf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Agf.EOF
        llUpper = UBound(tmAgfAvgToPayInfo)
        tmAgfAvgToPayInfo(llUpper).iCode = rst_Agf!agfCode
        tmAgfAvgToPayInfo(llUpper).iCurLstToPay = rst_Agf!agfLstToPay
        tmAgfAvgToPayInfo(llUpper).iCurAvgToPay = rst_Agf!agfAvgToPay
        tmAgfAvgToPayInfo(llUpper).iCurNoInvPd = rst_Agf!agfNoInvPd
        tmAgfAvgToPayInfo(llUpper).iNewLstToPay = 0
        tmAgfAvgToPayInfo(llUpper).iNewAvgToPay = 0
        tmAgfAvgToPayInfo(llUpper).iNewNoInvPd = 0
        ReDim Preserve tmAgfAvgToPayInfo(0 To llUpper + 1) As AVGTOPAYINFO
        mSetGauge llNoProc, llTotalNoToProc
        rst_Agf.MoveNext
    Loop
    lbcMsg.List(1) = "Getting Agency Values->Done"
    
    'Get PI Records
    lbcMsg.AddItem "Retrieving Payment History"
    slSQLQuery = "Select distinct phfAdfCode, phfagfCode, phfInvNo from phf_Payment_History Where phfCashTrade = 'C' And phfTranType = 'PI' Order by phfadfCode"
    Set rst_Phf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Phf.EOF
        'Get Invoice date
        slSQLQuery = "Select phfTranDate from phf_Payment_History Where phfCashTrade = 'C' And phfTranType = 'IN' And phfInvNo = " & rst_Phf!phfInvNo & " Order by phfTranDate"
        Set rst_Tmp = gSQLSelectCall(slSQLQuery)
        If Not rst_Tmp.EOF Then
            'Get Invoice date
            llInvDate = gDateValue(Format(rst_Tmp!phfTranDate, "ddddd"))
            'Get Advertiser values (AvgToPay)
            slSQLQuery = "Select Top 1 phfAdfCode, phfDateEntrd from phf_Payment_History Where phfCashTrade = 'C' And phfTranType = 'PI' And phfInvNo = " & rst_Phf!phfInvNo & " Order by phfDateEntrd Desc"
            Set rst_adf = gSQLSelectCall(slSQLQuery)
            If Not rst_adf.EOF Then
                'Find advteriser in array
                ilAdf = -1
                For ilLoop = 0 To UBound(tmAdfAvgToPayInfo) - 1 Step 1
                    If tmAdfAvgToPayInfo(ilLoop).iCode = rst_adf!phfAdfCode Then
                        ilAdf = ilLoop
                        Exit For
                    End If
                Next ilLoop
                If ilAdf <> -1 Then
                    tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay = ((10 * ((CLng(tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd) * tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay _
                    + gDateValue(Format(rst_adf!phfDateEntrd, "ddddd")) - llInvDate)) \ (tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd + 1)) + 5) \ 10
                    tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd = tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd + 1
                    tmAdfAvgToPayInfo(ilAdf).iNewLstToPay = gDateValue(Format(rst_adf!phfDateEntrd, "ddddd")) - llInvDate
                End If
            End If
            'Get Agency value (AvgToPay)
            slSQLQuery = "Select Top 1 phfAgfCode, phfDateEntrd from phf_Payment_History Where phfCashTrade = 'C' And phfTranType = 'PI' And phfAgfCode > 0 And phfInvNo = " & rst_Phf!phfInvNo & " Order by phfDateEntrd Desc"
            Set rst_Agf = gSQLSelectCall(slSQLQuery)
            If Not rst_Agf.EOF Then
                'Find advteriser in array
                ilAgf = -1
                For ilLoop = 0 To UBound(tmAgfAvgToPayInfo) - 1 Step 1
                    If tmAgfAvgToPayInfo(ilLoop).iCode = rst_Agf!phfAgfCode Then
                        ilAgf = ilLoop
                        Exit For
                    End If
                Next ilLoop
                If ilAgf <> -1 Then
                    tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay = ((10 * ((CLng(tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd) * tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay _
                    + gDateValue(Format(rst_Agf!phfDateEntrd, "ddddd")) - llInvDate)) \ (tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd + 1) + 5)) \ 10
                    tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd = tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd + 1
                    tmAgfAvgToPayInfo(ilAgf).iNewLstToPay = gDateValue(Format(rst_Agf!phfDateEntrd, "ddddd")) - llInvDate
                End If
            End If
        End If
        mSetGauge llNoProc, llTotalNoToProc
        rst_Phf.MoveNext
    Loop
    lbcMsg.List(2) = "Retrieving Payment History->Done"
    'Get PI Records
    lbcMsg.AddItem "Retrieving Receivables"
    slSQLQuery = "Select distinct rvfCntrNo, rvfInvNo, sum(rvfNet) as sumnet from RVF_Receivables group by rvfCntrNo, rvfInvNo"
    Set rst_Rvf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Rvf.EOF
        'Get Invoice date
        If rst_Rvf!Sumnet = "0" And rst_Rvf!rvfCntrNo > 0 And rst_Rvf!rvfInvNo > 0 Then
            slSQLQuery = "Select rvfTranDate from RVF_Receivables Where rvfCashTrade = 'C' And rvfTranType = 'IN' And rvfInvNo = " & rst_Rvf!rvfInvNo & " And rvfCntrNo = " & rst_Rvf!rvfCntrNo & " Order by rvfTranDate"
            Set rst_Tmp = gSQLSelectCall(slSQLQuery)
            If Not rst_Tmp.EOF Then
                'Get Invoice date
                llInvDate = gDateValue(Format(rst_Tmp!rvfTranDate, "ddddd"))
                'Get Advertiser values (AvgToPay)
                slSQLQuery = "Select Top 1 rvfAdfCode, rvfDateEntrd from RVF_Receivables Where rvfCashTrade = 'C' And rvfTranType = 'PI' And rvfInvNo = " & rst_Rvf!rvfInvNo & " And rvfCntrNo = " & rst_Rvf!rvfCntrNo & " Order by rvfDateEntrd Desc"
                Set rst_adf = gSQLSelectCall(slSQLQuery)
                If Not rst_adf.EOF Then
                    'Find advteriser in array
                    ilAdf = -1
                    For ilLoop = 0 To UBound(tmAdfAvgToPayInfo) - 1 Step 1
                        If tmAdfAvgToPayInfo(ilLoop).iCode = rst_adf!rvfAdfCode Then
                            ilAdf = ilLoop
                            Exit For
                        End If
                    Next ilLoop
                    If ilAdf <> -1 Then
                        tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay = ((10 * ((CLng(tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd) * tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay _
                        + gDateValue(Format(rst_adf!rvfDateEntrd, "ddddd")) - llInvDate)) \ (tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd + 1)) + 5) \ 10
                        tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd = tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd + 1
                        tmAdfAvgToPayInfo(ilAdf).iNewLstToPay = gDateValue(Format(rst_adf!rvfDateEntrd, "ddddd")) - llInvDate
                    End If
                End If
                'Get Agency value (AvgToPay)
                slSQLQuery = "Select Top 1 rvfAgfCode, rvfDateEntrd from RVF_Receivables Where rvfCashTrade = 'C' And rvfTranType = 'PI' And rvfAgfCode > 0 And rvfInvNo = " & rst_Rvf!rvfInvNo & " And rvfCntrNo = " & rst_Rvf!rvfCntrNo & " Order by rvfDateEntrd Desc"
                Set rst_Agf = gSQLSelectCall(slSQLQuery)
                If Not rst_Agf.EOF Then
                    'Find advteriser in array
                    ilAgf = -1
                    For ilLoop = 0 To UBound(tmAgfAvgToPayInfo) - 1 Step 1
                        If tmAgfAvgToPayInfo(ilLoop).iCode = rst_Agf!rvfAgfCode Then
                            ilAgf = ilLoop
                            Exit For
                        End If
                    Next ilLoop
                    If ilAgf <> -1 Then
                        tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay = ((10 * ((CLng(tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd) * tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay _
                        + gDateValue(Format(rst_Agf!rvfDateEntrd, "ddddd")) - llInvDate)) \ (tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd + 1)) + 5) \ 10
                        tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd = tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd + 1
                        tmAgfAvgToPayInfo(ilAgf).iNewLstToPay = gDateValue(Format(rst_Agf!rvfDateEntrd, "ddddd")) - llInvDate
                    End If
                End If
            End If
        End If
        mSetGauge llNoProc, llTotalNoToProc
        rst_Rvf.MoveNext
    Loop
    lbcMsg.List(3) = "Retrieving Receivable->Done"
    
    lbcMsg.AddItem "Reset Advertiser Values"
    'Reset Fields
    Print #hmTo, "*** Advertisers ***"
    For ilAdf = 0 To UBound(tmAdfAvgToPayInfo) - 1 Step 1
        slSQLQuery = "Update ADF_Advertisers Set "
        slSQLQuery = slSQLQuery & "adfLstToPay = " & tmAdfAvgToPayInfo(ilAdf).iNewLstToPay & ", "
        slSQLQuery = slSQLQuery & "adfAvgToPay = " & tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay & ", "
        slSQLQuery = slSQLQuery & "adfNoInvPd = " & tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd
        slSQLQuery = slSQLQuery & " Where adfCode = " & tmAdfAvgToPayInfo(ilAdf).iCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "AvgToPayErrors.txt", "AvgToPay: mAvgToPay"
        End If
        mSetGauge llNoProc, llTotalNoToProc
        ilIndex = gBinarySearchAdf(tmAdfAvgToPayInfo(ilAdf).iCode)
        If ilIndex <> -1 Then
            Print #hmTo, """" & Trim$(tgCommAdf(ilIndex).sName) & """" & "," & tmAdfAvgToPayInfo(ilAdf).iCurAvgToPay & "," & tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay & "," & tmAdfAvgToPayInfo(ilAdf).iCurAvgToPay - tmAdfAvgToPayInfo(ilAdf).iNewAvgToPay & "," & _
            tmAdfAvgToPayInfo(ilAdf).iCurNoInvPd & "," & tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd & "," & tmAdfAvgToPayInfo(ilAdf).iCurNoInvPd - tmAdfAvgToPayInfo(ilAdf).iNewNoInvPd & "," & _
            tmAdfAvgToPayInfo(ilAdf).iCurLstToPay & "," & tmAdfAvgToPayInfo(ilAdf).iNewLstToPay & "," & tmAdfAvgToPayInfo(ilAdf).iCurLstToPay - tmAdfAvgToPayInfo(ilAdf).iNewLstToPay
        End If
    Next ilAdf
    lbcMsg.List(4) = "Reset Advertiser Values->Done"
    lbcMsg.AddItem "Reset Agency Values"
    Print #hmTo, ""
    Print #hmTo, ""
    Print #hmTo, "*** Agencies ***"
    For ilAgf = 0 To UBound(tmAgfAvgToPayInfo) - 1 Step 1
        slSQLQuery = "Update AGF_Agencies Set "
        slSQLQuery = slSQLQuery & "agfLstToPay = " & tmAgfAvgToPayInfo(ilAgf).iNewLstToPay & ", "
        slSQLQuery = slSQLQuery & "agfAvgToPay = " & tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay & ", "
        slSQLQuery = slSQLQuery & "agfNoInvPd = " & tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd
        slSQLQuery = slSQLQuery & " Where agfCode = " & tmAgfAvgToPayInfo(ilAgf).iCode
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "AvgToPayErrors.txt", "AvgToPay: mAvgToPay"
        End If
        mSetGauge llNoProc, llTotalNoToProc
        ilIndex = gBinarySearchAgf(tmAgfAvgToPayInfo(ilAgf).iCode)
        If ilIndex <> -1 Then
            Print #hmTo, """" & Trim$(tgCommAgf(ilIndex).sName) & """" & "," & tmAgfAvgToPayInfo(ilAgf).iCurAvgToPay & "," & tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay & "," & tmAgfAvgToPayInfo(ilAgf).iCurAvgToPay - tmAgfAvgToPayInfo(ilAgf).iNewAvgToPay & "," & _
            tmAgfAvgToPayInfo(ilAgf).iCurNoInvPd & "," & tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd & "," & tmAgfAvgToPayInfo(ilAgf).iCurNoInvPd - tmAgfAvgToPayInfo(ilAgf).iNewNoInvPd & "," & _
            tmAgfAvgToPayInfo(ilAgf).iCurLstToPay & "," & tmAgfAvgToPayInfo(ilAgf).iNewLstToPay & "," & tmAgfAvgToPayInfo(ilAgf).iCurLstToPay - tmAgfAvgToPayInfo(ilAgf).iNewLstToPay
        End If
    Next ilAgf
    lbcMsg.List(5) = "Reset Agency Values->Done"
End Function


Private Sub mSetGauge(llNoProc As Long, llTotalNoToProc As Long)
    Dim llPercent As Long
    llNoProc = llNoProc + 1
    llPercent = (llNoProc * CSng(100)) / llTotalNoToProc
    If llPercent >= 99 Then
        llPercent = 99
    ElseIf llPercent < 0 Then
        llPercent = 0
    End If
    plcGauge.Value = llPercent
    DoEvents
End Sub
