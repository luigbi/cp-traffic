VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmCheckBkout 
   Caption         =   "Duplicate Blackouts Fix"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "AffCheckBkout.frx":0000
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
      ItemData        =   "AffCheckBkout.frx":08CA
      Left            =   5070
      List            =   "AffCheckBkout.frx":08CC
      TabIndex        =   6
      Top             =   1035
      Width           =   4455
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "AffCheckBkout.frx":08CE
      Left            =   135
      List            =   "AffCheckBkout.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
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
      Caption         =   "Fix"
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
      Caption         =   "Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1395
   End
End
Attribute VB_Name = "frmCheckBkout"
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
Private lmTotalRemoved As Long
Private Type BKOUTINFO
    lAstCode As Long
    lLstCode As Long
    iCPStatus As Integer
    bRemove As Boolean
    sFeedDate As String * 10
    sFeedTime As String * 11
    sISCI As String * 30
End Type
Private tmBkoutInfo() As BKOUTINFO
Private rst_att As ADODB.Recordset
Private rst_Ast As ADODB.Recordset
Private rst_Chk As ADODB.Recordset



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
    slToFile = sgExportDirectory & "CheckBkout.Txt"
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
    Print #hmMsg, "** Checking and fixing Duplicate Blackouts: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
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
    'mReadPreselectedVehicles
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
    lmTotalRemoved = 0
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            mCheckBkout
            If imTerminate Then
                Exit For
            End If
        End If
    Next iLoop
    Print #hmMsg, "** Total Removed: " & lmTotalRemoved & " **"
    If Not imTerminate Then
        'mWritePreselectedVehicles
    End If
    imChecking = False
    If imTerminate Then
        Print #hmMsg, "** Completed Check and Fix Duplicate Blackouts: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Else
        Print #hmMsg, "** Terminated Check and Fix Duplicate Blackouts: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
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
    Unload frmCheckBkout
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
    gSetFonts frmCheckBkout
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
    Erase tmBkoutInfo
    rst_att.Close
    rst_Ast.Close
    rst_Chk.Close
    Set frmCheckBkout = Nothing
End Sub


Private Sub lbcVehicles_Click()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCancel.Caption = "&Cancel"

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
    cmdCheck.Enabled = True
    cmdCancel.Caption = "&Cancel"
End Sub


Private Sub mCheckBkout()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilShtt As Integer
    Dim slResult As String
    Dim slSQLQuery As String
    Dim blAnyCPStatusZero As Boolean
    Dim blAnyCPStatusOne As Boolean
    Dim ilLoop As Integer
    Dim ilRetain As Integer
    Dim slStr As String
    Dim llIndex As Long
    
    On Error GoTo ErrHand
    ilVef = gBinarySearchVef(CLng(imVefCode))
    If ilVef = -1 Then
        Exit Sub
    End If
    slSQLQuery = "SELECT distinct attCode FROM att "
    slSQLQuery = slSQLQuery + " WHERE attVefCode = " & imVefCode
    slSQLQuery = slSQLQuery & " AND attLoad <= 1"
    'slSQLQuery = slSQLQuery & " AND attNoAirPlays <= 1"
    slSQLQuery = slSQLQuery + " AND attOffAir >= '" & Format$(smDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery + " AND attDropDate >= '" & Format$(smDate, sgSQLDateForm) & "'"
    Set rst_att = gSQLSelectCall(slSQLQuery)
    While Not rst_att.EOF
        slSQLQuery = "Select astSdfCode, astVefCode, astShfCode, astDatCode, lstBkoutLstCode, astFeedDate, Count(1) from ast "
        slSQLQuery = slSQLQuery & " Left Outer Join lst On astLsfCode = lstCode"
        slSQLQuery = slSQLQuery & " Where astAtfCode =" & rst_att!attCode
        slSQLQuery = slSQLQuery & " And lstBkoutLstCode > 0"
        slSQLQuery = slSQLQuery & " And astFeedDate >= '" & Format$(smDate, sgSQLDateForm) & "'"
        slSQLQuery = slSQLQuery & " Group By astSdfCode, astVefCode, astShfCode, astDatCode, lstBkoutLstCode, astFeedDate Having Count(1) > 1"
        Set rst_Ast = gSQLSelectCall(slSQLQuery)
        Do While Not rst_Ast.EOF
            ilShtt = gBinarySearchStationInfoByCode(rst_Ast!astShfCode)
            ReDim tmBkoutInfo(0 To 0) As BKOUTINFO
            blAnyCPStatusZero = False
            blAnyCPStatusOne = False
            slSQLQuery = "Select astCode, astLsfCode, astCPStatus, astFeedDate, astFeedTime, lstISCI from ast "
            slSQLQuery = slSQLQuery & " Left Outer Join lst On astLsfCode = lstCode"
            slSQLQuery = slSQLQuery & " Where astAtfCode =" & rst_att!attCode
            slSQLQuery = slSQLQuery & " And astSdfCode =" & rst_Ast!astSdfCode
            slSQLQuery = slSQLQuery & " And astDatCode =" & rst_Ast!astDatCode
            slSQLQuery = slSQLQuery & " And lstBkoutLstCode =" & rst_Ast!lstBkoutLstCode
            slSQLQuery = slSQLQuery & " And astFeedDate = '" & Format$(rst_Ast!astFeedDate, sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " Order By astSdfCode, astDatCode, astFeedDate, astFeedTime"
            Set rst_Chk = gSQLSelectCall(slSQLQuery)
            Do While Not rst_Chk.EOF
                'save spots
                tmBkoutInfo(UBound(tmBkoutInfo)).lAstCode = rst_Chk!astCode
                tmBkoutInfo(UBound(tmBkoutInfo)).lLstCode = rst_Chk!astLsfCode
                tmBkoutInfo(UBound(tmBkoutInfo)).iCPStatus = rst_Chk!astCPStatus
                If rst_Chk!astCPStatus = 0 Then blAnyCPStatusZero = True
                If rst_Chk!astCPStatus = 1 Then blAnyCPStatusOne = True
                tmBkoutInfo(UBound(tmBkoutInfo)).bRemove = False
                tmBkoutInfo(UBound(tmBkoutInfo)).sFeedDate = Format(rst_Chk!astFeedDate, sgShowDateForm)
                tmBkoutInfo(UBound(tmBkoutInfo)).sFeedTime = Format(rst_Chk!astFeedTime, sgShowTimeWSecForm)
                tmBkoutInfo(UBound(tmBkoutInfo)).sISCI = rst_Chk!lstISCI
                ReDim Preserve tmBkoutInfo(0 To UBound(tmBkoutInfo) + 1) As BKOUTINFO
                rst_Chk.MoveNext
            Loop
            'determine which to remove
            'Rule 1: check if some posted and others are not. Remove the not posted
            If (blAnyCPStatusZero And blAnyCPStatusOne) Then
                For ilLoop = 0 To UBound(tmBkoutInfo) - 1 Step 1
                    If tmBkoutInfo(ilLoop).iCPStatus = 0 Then tmBkoutInfo(ilLoop).bRemove = True
                Next ilLoop
            End If
            'Rule 2: retain the oldest lstCode if more then one posted
            ilRetain = -1
            For ilLoop = 0 To UBound(tmBkoutInfo) - 1 Step 1
                If tmBkoutInfo(ilLoop).bRemove = False Then
                    If ilRetain = -1 Then
                        ilRetain = ilLoop
                    Else
                        If tmBkoutInfo(ilLoop).lLstCode < tmBkoutInfo(ilRetain).lLstCode Then
                            tmBkoutInfo(ilRetain).bRemove = True
                            ilRetain = ilLoop
                        Else
                            tmBkoutInfo(ilLoop).bRemove = True
                        End If
                    End If
                End If
            Next ilLoop
            'Alternate rule: Retain the one that is on the web
            
            'Remove lst and ast
            For ilLoop = 0 To UBound(tmBkoutInfo) - 1 Step 1
                If tmBkoutInfo(ilLoop).bRemove = True Then
                    slSQLQuery = "Delete From lst Where lstCode = " & tmBkoutInfo(ilLoop).lLstCode
                    If gSQLWaitNoMsgBox(slSQLQuery, False) = 0 Then
                        slSQLQuery = "Delete From ast Where astCode = " & tmBkoutInfo(ilLoop).lAstCode
                        If gSQLWaitNoMsgBox(slSQLQuery, False) = 0 Then
                            lmTotalRemoved = lmTotalRemoved + 1
                            If ilShtt <> -1 Then
                                slStr = Trim$(tgVehicleInfo(ilVef).sVehicle) & "/" & Trim$(tgStationInfoByCode(ilShtt).sCallLetters) & ": " & Trim$(tmBkoutInfo(ilLoop).sFeedDate) & " " & Trim$(tmBkoutInfo(ilLoop).sFeedTime) & " " & Trim$(tmBkoutInfo(ilLoop).sISCI)
                            Else
                                slStr = Trim$(tgVehicleInfo(ilVef).sVehicle) & ": " & Trim$(tmBkoutInfo(ilLoop).sFeedDate) & " " & Trim$(tmBkoutInfo(ilLoop).sFeedTime) & " " & Trim$(tmBkoutInfo(ilLoop).sISCI)
                            End If
                            Print #hmMsg, "  Removed:" & slStr
                        End If
                    End If
                Else
                    slStr = Trim$(tgVehicleInfo(ilVef).sVehicle) & ": " & Format$(rst_Ast!astFeedDate, sgShowDateForm)
                    llIndex = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slStr)
                    If llIndex < 0 Then
                        lbcMsg.AddItem slStr
                    End If
                End If
            Next ilLoop
            
            rst_Ast.MoveNext
        Loop
                
        rst_att.MoveNext
    Wend
    Exit Sub
mCheckBkoutErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTT Agree-mCheckBkout"
    Exit Sub
    
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmCheckBkout
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
    'Open Trim$(sgImportDirectory) & "CheckBkout.Txt" For Input Access Read As hmVehicles
    ilRet = gFileOpen(Trim$(sgImportDirectory) & "CheckBkout.Txt", "Input Access Read", hmVehicles)
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
    Kill Trim$(sgImportDirectory) & "CheckBkout.Txt"
    On Error GoTo 0
    
    ilRet = 0
    On Error GoTo mWritePreselectedVehiclesErr:
    'hmVehicles = FreeFile
    'Open Trim$(sgImportDirectory) & "CheckBkout.Txt" For Output As hmVehicles
    ilRet = gFileOpen(Trim$(sgImportDirectory) & "CheckBkout.Txt", "Output", hmVehicles)
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
