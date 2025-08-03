VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmCPTTAgree 
   Caption         =   "Check CPTT between Affiliate and Web"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "AffCPTTAgree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9885
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1275
      TabIndex        =   1
      Top             =   60
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "11/8/2010"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   0   'False
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   585
      Width           =   3825
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Results"
      Top             =   585
      Width           =   3825
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9675
      Top             =   3300
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   4605
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3180
      ItemData        =   "AffCPTTAgree.frx":08CA
      Left            =   5070
      List            =   "AffCPTTAgree.frx":08CC
      TabIndex        =   6
      Top             =   1035
      Width           =   4455
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "AffCPTTAgree.frx":08CE
      Left            =   135
      List            =   "AffCPTTAgree.frx":08D0
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1050
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9480
      Top             =   4740
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5610
      FormDesignWidth =   9885
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   5820
      TabIndex        =   4
      Top             =   5115
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   5
      Top             =   5100
      Width           =   1575
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   7
      Top             =   5055
      Width           =   5490
   End
   Begin VB.Label Label1 
      Caption         =   "Check Week"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1395
   End
End
Attribute VB_Name = "frmCPTTAgree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private smDate As String     'Export Date
Private imVefCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imChecking As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmTo As Integer
Private hmVehicles As Integer






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
Private Function mOpenMsgFile(sMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgExportDirectory & "CPTTAgree.Txt"
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Checking CPTT between Affiliate and Web: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    mReadPreselectedVehicles
End Sub




Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub cmdCheck_Click()
    Dim iLoop As Integer
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    Dim sFileName As String
    Dim sLetter As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    lbcMsg.Clear
    If lbcVehicles.SelCount <= 0 Then
        gMsgBox "Vehicle must be specified.", vbOKOnly
        Exit Sub
    End If
    If edcDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        edcDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        edcDate.SetFocus
    Else
        smDate = Format(edcDate.Text, sgShowDateForm)
    End If
    smDate = gObtainPrevMonday(smDate)
    Screen.MousePointer = vbHourglass
    
    If Not mOpenMsgFile(sMsgFileName) Then
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    imChecking = True
    lacResult.Caption = ""
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            ilRet = mCPTTAgree()
            If imTerminate Then
                Exit For
            End If
        End If
    Next iLoop
    If Not imTerminate Then
        mWritePreselectedVehicles
    End If
    imChecking = False
    If imTerminate Then
        Print #hmMsg, "** Completed Checking CPTT: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Else
        Print #hmMsg, "** Terminated Checking CPTT: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    End If
    Close #hmMsg
    lacResult.Caption = "See: " & sMsgFileName & " for Result Summary"
    cmdCheck.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdCheckErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTT Agree: mCheck_Click"
End Sub

Private Sub cmdCancel_Click()
    If imChecking Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmCPTTAgree
End Sub


Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmCPTTAgree
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    smDate = ""
    imAllClick = False
    imTerminate = False
    imChecking = False
    imFirstTime = True
    
    mFillVehicle
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If imChecking Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    Set frmCPTTAgree = Nothing
End Sub


Private Sub lbcVehicles_Click()
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
End Sub

Private Sub edcDate_Change()
    lbcMsg.Clear
End Sub


Private Function mCPTTAgree()
    Dim ilRet As Integer
    Dim slResult As String
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT cpttAtfCode, cpttShfCode, cpttStartDate FROM cptt "
    SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & imVefCode
    SQLQuery = SQLQuery + " AND cpttPostingStatus = 0"
    SQLQuery = SQLQuery + " AND cpttStartDate = '" & Format$(smDate, sgSQLDateForm) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        'Doug: add your call here to get counts from Web
        
        'Add saving error results here
        'ilRet = 0
        'On Error GoTo mCPTTAgreeErr:
        'Print #hmTo, slResult
        'If ilRet <> 0 Then
        '    mCPTTAgree = False
        '    Exit Function
        'End If
        'lbcMsg.AddItem slRecord
        'On Error GoTo ErrHand
        rst.MoveNext
    Wend
    mCPTTAgree = True
    Exit Function
mCPTTAgreeErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTT Agree-mCPTTAgree"
    mCPTTAgree = False
    Exit Function
    
End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmCPTTAgree
End Sub

Private Sub mReadPreselectedVehicles()
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim llRow As Long
    
    ReDim smPreselectNames(0 To 0)
    ilRet = 0
    On Error GoTo mReadPreselectedVehicles:
    'hmVehicles = FreeFile
    'Open Trim$(sgImportDirectory) & "CPTTAgree.Txt" For Input Access Read As hmVehicles
    ilRet = gFileOpen(Trim$(sgImportDirectory) & "CPTTAgree.Txt", "Input Access Read", hmVehicles)
    If ilRet <> 0 Then
        Close hmVehicles
        Exit Sub
    End If
    Do
        'On Error GoTo mReadPreselectedVehicles:
        If EOF(hmVehicles) Then
            Exit Do
        End If
        Line Input #hmVehicles, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                For llRow = 0 To lbcVehicles.ListCount - 1 Step 1
                    If StrComp(lbcVehicles.List(llRow), slLine, vbTextCompare) = 0 Then
                        lbcVehicles.Selected(llRow) = True
                        Exit For
                    End If
                Next llRow
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
    Kill Trim$(sgImportDirectory) & "CPTTAgree.Txt"
    On Error GoTo 0
    
    ilRet = 0
    On Error GoTo mWritePreselectedVehiclesErr:
    'hmVehicles = FreeFile
    'Open Trim$(sgImportDirectory) & "CPTTAgree.Txt" For Output As hmVehicles
    ilRet = gFileOpen(Trim$(sgImportDirectory) & "CPTTAgree.Txt", "Output", hmVehicles)
    If ilRet = 0 Then
        For llRow = 0 To lbcVehicles.ListCount - 1 Step 1
            If lbcVehicles.Selected(llRow) Then
                Print #hmVehicles, lbcVehicles.List(llRow)
            End If
        Next llRow
        Close #hmVehicles
    End If
    Exit Sub
mWritePreselectedVehiclesErr:
    ilRet = Err.Number
    Resume Next
End Sub
