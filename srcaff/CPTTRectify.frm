VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form CPTTRectify 
   Caption         =   "CPTT Rectify"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "CPTTRectify.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9885
   Begin V81CPTTRectify.CSI_Calendar edcStartDate 
      Height          =   255
      Left            =   1605
      TabIndex        =   1
      Top             =   45
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   450
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin V81CPTTRectify.CSI_Calendar edcEndDate 
      Height          =   255
      Left            =   5340
      TabIndex        =   3
      Top             =   45
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   450
      Text            =   "9/22/2012"
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
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin VB.CheckBox ckcNotAst 
      Caption         =   "Output Detail (Date, Vehicle, Station) for 'Not Created' count"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3885
      TabIndex        =   15
      Top             =   450
      Width           =   5115
   End
   Begin VB.CheckBox ckcAST 
      Caption         =   "Create missing Affiliate Spots(ast records)"
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   450
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9690
      Top             =   3990
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   870
      Width           =   3825
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Results"
      Top             =   870
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
      TabIndex        =   6
      Top             =   4890
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3180
      ItemData        =   "CPTTRectify.frx":08CA
      Left            =   5070
      List            =   "CPTTRectify.frx":08CC
      TabIndex        =   9
      Top             =   1320
      Width           =   4455
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "CPTTRectify.frx":08CE
      Left            =   135
      List            =   "CPTTRectify.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1335
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
      FormDesignHeight=   5925
      FormDesignWidth =   9885
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Rectify"
      Height          =   375
      Left            =   5820
      TabIndex        =   7
      Top             =   5400
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   8
      Top             =   5385
      Width           =   1575
   End
   Begin VB.Label lacVehicle 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1185
      TabIndex        =   13
      Top             =   4650
      Width           =   6810
   End
   Begin VB.Label lacProgress 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1185
      TabIndex        =   12
      Top             =   4980
      Width           =   6810
   End
   Begin VB.Label lacWeekEndDate 
      Caption         =   "Week End Date"
      Height          =   255
      Left            =   3885
      TabIndex        =   2
      Top             =   75
      Width           =   1920
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   10
      Top             =   5340
      Width           =   5490
   End
   Begin VB.Label lacWeekStartDate 
      Caption         =   "Week Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1920
   End
End
Attribute VB_Name = "CPTTRectify"
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

Private imGuideUstCode As Integer
Private bmNoPervasive As Boolean

Private smStartDate As String     'Export Date
Private smEndDate As String
Private lmStartDate As Long     'Export Date
Private lmEndDate As Long
Private imVefCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imChecking As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmTo As Integer
Private lmTotal As Long
Private lmRunningCount As Long
Private lmCreated As Long
Private lmNotPosted As Long
Private lmPartiallyPosted As Long
Private lmCompletedPosted As Long
Private lmAstMissing As Long
Private lmAstCreated As Long
Private lmAstNotCreated As Long
Private lmExtra As Long
Private hmVehicles As Integer
Private bmErrorMsgLogged As Boolean
Private Const FORMNAME As String = "CPTTRectify"
Private cptt_rst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private vpf_rst As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private hmAst As Integer
Private tmAstInfo() As ASTINFO

Private Type CPTTSTATIC
    sKey As String * 50
    lDate As Long
    iVefCode As Integer
    lCreated As Long
    lNotPosted As Long
    lPartiallyPosted As Long
    lCompletedPosted As Long
    lAstMissing As Long
    lAstCreated As Long
    lAstNotCreated As Long
    lExtra As Long
End Type
Dim tmCpttStatic() As CPTTSTATIC
Private Type NOTCREATED
    sKey As String * 60
    lDate As Long
    iVefCode As Integer
    iShttCode As Integer
End Type
Dim tmNotCreated() As NOTCREATED








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
    slToFile = sgMsgDirectory & "CPTTRectify_" & Format$(Now, "mmddyy") & ".Csv"
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
            On Error GoTo 0
            ilRet = 0
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
            On Error Resume Next
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
    Print #hmMsg, " "
    Print #hmMsg, "** Rectify Post CP: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, "Start Date: " & smStartDate & " End Date: " & smEndDate
    If ckcAST.Value = vbChecked Then
        Print #hmMsg, "Week Date-# Weeks:,Vehicle,Created,Not Posted,Partially Posted,Completely Posted,Missing Spots,Extra,Created Spots,Spots Not Created"
    Else
        Print #hmMsg, "Week Date-# Weeks:,Vehicle,Created,Not Posted,Partially Posted,Completely Posted,Missing Spots,Extra"
    End If
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
    mSetControl
End Sub

Private Sub ckcAST_Click()
    mSetControl
End Sub

Private Sub cmdCheck_Click()
    Dim ilLoop As Integer
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
    Dim ilTotal As Integer
    Dim ilCount As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llTotalTime As Long
    Dim llLoop As Long
    Dim llVef As Long
    Dim slVehicleName As String
    Dim slStr As String
    Dim ilShtt As Integer
    Dim slStation As String

    On Error GoTo ErrHand
    If imChecking = True Then
        Exit Sub
    End If
    imTerminate = False
    lacProgress.Caption = ""
    lacResult.Caption = ""
    lbcMsg.Clear
    If lbcVehicles.SelCount <= 0 Then
        gMsgBox "Vehicle must be specified.", vbOKOnly
        Exit Sub
    End If
    If edcStartDate.Text = "" Then
        gMsgBox "Start Date must be specified.", vbOKOnly
        edcStartDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcStartDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid start date (m/d/yy).", vbCritical
        edcStartDate.SetFocus
    Else
        smStartDate = Format(edcStartDate.Text, sgShowDateForm)
    End If
    smStartDate = gObtainPrevMonday(smStartDate)
    If (edcEndDate.Text <> "") And (edcEndDate.Text <> "TFN") Then
        If gIsDate(edcEndDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid end date (m/d/yy).", vbCritical
            edcEndDate.SetFocus
        Else
            smEndDate = Format(edcEndDate.Text, sgShowDateForm)
        End If
        smEndDate = gObtainNextSunday(smEndDate)
    Else
        smEndDate = ""
    End If
    If smEndDate = "" Then
        smEndDate = gObtainNextSunday(Format(gNow(), "m/d/yy"))
    End If
    Screen.MousePointer = vbHourglass
    bmErrorMsgLogged = False
    If Not mOpenMsgFile(sMsgFileName) Then
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If

    lacProgress.Caption = ""
    lacVehicle.Caption = ""
    DoEvents
    
    imChecking = True
    lacResult.Caption = ""
    llStartTime = timeGetTime
    lmTotal = 0
    lmRunningCount = 0
    ReDim tmCpttStatic(0 To 0) As CPTTSTATIC
    ReDim tmNotCreated(0 To 0) As NOTCREATED
    For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilLoop) Then
            lmTotal = lmTotal + 1
        End If
    Next ilLoop
    lmStartDate = gDateValue(smStartDate)
    lmEndDate = gDateValue(smEndDate)
    lacProgress.Caption = ""
    'Print #hmMsg, "Type,attCode,Feed Date,Vehicle,Station"
    For ilLoop = 0 To lbcVehicles.ListCount - 1
        DoEvents
        If lbcVehicles.Selected(ilLoop) Then
            lmRunningCount = lmRunningCount + 1
            lacProgress.Caption = "Processed " & lmRunningCount & " of " & lmTotal
            'Get hmTo handle
            lacVehicle.Caption = "Checking: " & Trim$(lbcVehicles.List(ilLoop))
            DoEvents
            imVefCode = lbcVehicles.ItemData(ilLoop)
            cmdCancel.Caption = "&Cancel"
            ilRet = mAddCptt(imVefCode)
            If imTerminate Then
                Exit For
            End If
        End If
    Next ilLoop
    lmCreated = 0
    lmNotPosted = 0
    lmPartiallyPosted = 0
    lmCompletedPosted = 0
    lmAstMissing = 0
    lmAstCreated = 0
    lmAstNotCreated = 0
    lmExtra = 0
    For llLoop = 0 To UBound(tmCpttStatic) - 1 Step 1
        slStr = gDateValue(Format(tmCpttStatic(llLoop).lDate, "m/d/yy"))
        Do While Len(slStr) < 6
            slStr = "0" & slStr
        Loop
        llVef = gBinarySearchVef(CLng(tmCpttStatic(llLoop).iVefCode))
        If llVef <> -1 Then
            slVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
        Else
            slVehicleName = tmCpttStatic(llLoop).iVefCode
            Do While Len(slVehicleName) < 6
                slVehicleName = "0" & slVehicleName
            Loop
        End If
        tmCpttStatic(llLoop).sKey = slStr & slVehicleName
    Next llLoop
    If UBound(tmCpttStatic) - 1 > 0 Then
        ArraySortTyp fnAV(tmCpttStatic(), 0), UBound(tmCpttStatic), 0, LenB(tmCpttStatic(0)), 0, LenB(tmCpttStatic(0).sKey), 0
    End If
    For llLoop = 0 To UBound(tmCpttStatic) - 1 Step 1
        lmCreated = lmCreated + tmCpttStatic(llLoop).lCreated
        lmNotPosted = lmNotPosted + tmCpttStatic(llLoop).lNotPosted
        lmPartiallyPosted = lmPartiallyPosted + tmCpttStatic(llLoop).lPartiallyPosted
        lmCompletedPosted = lmCompletedPosted + tmCpttStatic(llLoop).lCompletedPosted
        lmAstMissing = lmAstMissing + tmCpttStatic(llLoop).lAstMissing
        lmExtra = lmExtra + tmCpttStatic(llLoop).lExtra
        lmAstCreated = lmAstCreated + tmCpttStatic(llLoop).lAstCreated
        lmAstNotCreated = lmAstNotCreated + tmCpttStatic(llLoop).lAstNotCreated
        llVef = gBinarySearchVef(CLng(tmCpttStatic(llLoop).iVefCode))
        If ckcAST.Value = vbChecked Then
            If llVef <> -1 Then
                Print #hmMsg, Format(tmCpttStatic(llLoop).lDate, "m/d/yy") & "," & Trim$(tgVehicleInfo(llVef).sVehicle) & "," & tmCpttStatic(llLoop).lCreated & "," & tmCpttStatic(llLoop).lNotPosted & "," & tmCpttStatic(llLoop).lPartiallyPosted & "," & tmCpttStatic(llLoop).lCompletedPosted & "," & tmCpttStatic(llLoop).lAstMissing & "," & tmCpttStatic(llLoop).lExtra & "," & tmCpttStatic(llLoop).lAstCreated & "," & tmCpttStatic(llLoop).lAstNotCreated
            Else
                Print #hmMsg, Format(tmCpttStatic(llLoop).lDate, "m/d/yy") & "," & tmCpttStatic(llLoop).iVefCode & "," & tmCpttStatic(llLoop).lCreated & "," & tmCpttStatic(llLoop).lNotPosted & "," & tmCpttStatic(llLoop).lPartiallyPosted & "," & tmCpttStatic(llLoop).lCompletedPosted & "," & tmCpttStatic(llLoop).lAstMissing & "," & tmCpttStatic(llLoop).lExtra & "," & tmCpttStatic(llLoop).lAstCreated & "," & tmCpttStatic(llLoop).lAstNotCreated
            End If
        Else
            If llVef <> -1 Then
                Print #hmMsg, Format(tmCpttStatic(llLoop).lDate, "m/d/yy") & "," & Trim$(tgVehicleInfo(llVef).sVehicle) & "," & tmCpttStatic(llLoop).lCreated & "," & tmCpttStatic(llLoop).lNotPosted & "," & tmCpttStatic(llLoop).lPartiallyPosted & "," & tmCpttStatic(llLoop).lCompletedPosted & "," & tmCpttStatic(llLoop).lAstMissing & "," & tmCpttStatic(llLoop).lExtra
            Else
                Print #hmMsg, Format(tmCpttStatic(llLoop).lDate, "m/d/yy") & "," & tmCpttStatic(llLoop).iVefCode & "," & tmCpttStatic(llLoop).lCreated & "," & tmCpttStatic(llLoop).lNotPosted & "," & tmCpttStatic(llLoop).lPartiallyPosted & "," & tmCpttStatic(llLoop).lCompletedPosted & "," & tmCpttStatic(llLoop).lAstMissing & "," & tmCpttStatic(llLoop).lExtra
            End If
        End If
    Next llLoop
    If ckcNotAst.Value = vbChecked Then
        Print #hmMsg, ""
        Print #hmMsg, "Date,Vehicle,Station"
        If UBound(tmNotCreated) - 1 > 0 Then
            ArraySortTyp fnAV(tmNotCreated(), 0), UBound(tmNotCreated), 0, LenB(tmNotCreated(0)), 0, LenB(tmNotCreated(0).sKey), 0
        End If
        For llLoop = 0 To UBound(tmNotCreated) - 1 Step 1
            llVef = gBinarySearchVef(CLng(tmNotCreated(llLoop).iVefCode))
            If llVef <> -1 Then
                slVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
            Else
                slVehicleName = tmNotCreated(llLoop).iVefCode
            End If
            ilShtt = gBinarySearchStationInfoByCode(tmNotCreated(llLoop).iShttCode)
            If ilShtt <> -1 Then
                slStation = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
            Else
                slStation = tmNotCreated(llLoop).iShttCode
            End If
            Print #hmMsg, Format(tmNotCreated(llLoop).lDate, "m/d/yy") & "," & slVehicleName & "," & slStation
        Next llLoop
        Print #hmMsg, ""
    End If
    lbcMsg.AddItem "Total Posted Weeks(cptt) Created " & lmCreated
    Print #hmMsg, "Total Posted Weeks(cptt) Created " & lmCreated
    lbcMsg.AddItem "Number of Weeks Not Posted " & lmNotPosted
    Print #hmMsg, "Number of Weeks  Not Posted " & lmNotPosted
    lbcMsg.AddItem "Number of Weeks Partially Posted " & lmPartiallyPosted
    Print #hmMsg, "Number of Weeks Partially Posted " & lmPartiallyPosted
    lbcMsg.AddItem "Number of Weeks Completely Posted " & lmCompletedPosted
    Print #hmMsg, "Number of Weeks Completely Posted " & lmCompletedPosted
    lbcMsg.AddItem "Number of Weeks Affiliate Spots(ast) Missing " & lmAstMissing
    Print #hmMsg, "Number of Weeks Affiliate Spots(ast) Missing " & lmAstMissing
    lbcMsg.AddItem "Number of Extra Posted Weeks(cptt) Exist " & lmExtra
    Print #hmMsg, "Number of Extra Posted Weeks(cptt) Exist " & lmExtra
    If ckcAST.Value = vbChecked Then
        lbcMsg.AddItem "Number of Weeks Affiliate Spots(ast) Created " & lmAstCreated
        Print #hmMsg, "Number of Weeks Affiliate Spots(ast) Created " & lmAstCreated
        lbcMsg.AddItem "Number of Weeks Affiliate Spots(ast) Not Created " & lmAstNotCreated
        Print #hmMsg, "Number of Weeks Affiliate Spots(ast) Not Created " & lmAstNotCreated
    End If
    lacProgress.Caption = ""
    lacVehicle.Caption = ""
    imChecking = False
    If Not imTerminate Then
        Print #hmMsg, "** Completed CPTT Rectify: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Else
        Print #hmMsg, "** Terminated CPTT Rectify: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    End If
    Print #hmMsg, " "
    Close #hmMsg
    lacResult.Caption = "See: " & sMsgFileName & " for Result Summary"
    llEndTime = timeGetTime
    llTotalTime = llEndTime - llStartTime
    lbcMsg.AddItem "Run Time = " & gTimeString(llTotalTime / 1000, True)
    cmdCheck.Enabled = False    'True
    cmdCancel.SetFocus
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdCheckErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTTRectify: mCheck_Click"
End Sub

Private Sub cmdCancel_Click()
    If imChecking Then
        imTerminate = True
        Exit Sub
    End If
    edcStartDate.Text = ""
    Unload CPTTRectify
End Sub


Private Sub edcEndDate_Change()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Rectify"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub edcEndDate_GotFocus()
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Rectify"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub edcStartDate_GotFocus()
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Rectify"
    cmdCancel.Caption = "&Cancel"
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

Private Sub Form_GotFocus()
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Rectify"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts CPTTRectify
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    Dim slAffToWebDate As String
    Dim slWebToAffDate As String
        
    Screen.MousePointer = vbHourglass
    smStartDate = ""
    imAllClick = False
    imTerminate = False
    imChecking = False
    imFirstTime = True
    
    mInit
    
    gOpenMKDFile hmAst, "Ast.Mkd"
    
    mFillVehicle
    chkAll.Value = vbChecked
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If imChecking Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    gCloseMKDFile hmAst, "Ast.Mkd"
    
    cptt_rst.Close
    att_rst.Close
    ast_rst.Close
    vpf_rst.Close
    lst_rst.Close
    
    Erase tmCpttStatic
    Erase tmNotCreated
    Erase tmAstInfo
    
    Erase tgCifCpfInfo1
    Erase tgCrfInfo1
    Erase lgUserLogUlfCode
    Erase tgCopyRotInfo
    Erase tgGameInfo
    Erase tgStationInfoByCode
    Erase tgCpfInfo
    Erase tgMarketInfo
    Erase tgMSAMarketInfo
    Erase tgTerritoryInfo
    Erase tgCityInfo
    Erase tgCountyInfo
    Erase tgAreaInfo
    Erase tgMonikerInfo
    Erase tgOperatorInfo
    Erase tgMarketRepInfo
    Erase tgServiceRepInfo
    Erase tgAffAEInfo
    Erase tgSellingVehicleInfo
    Erase tgVpfOptions
    Erase tgLstInfo
    Erase tgAttInfo1
    Erase tgShttInfo1
    Erase tgCpttInfo
    Erase sgAufsKey
    Erase tgRBofRec
    Erase tgSplitNetLastFill
    Erase tgAvailNamesInfo
    Erase tgMediaCodesInfo
    Erase tgTitleInfo
    Erase tgOwnerInfo
    Erase tgFormatInfo
    Erase tgVffInfo
    Erase tgTeamInfo
    Erase tgLangInfo
    Erase tgTimeZoneInfo
    Erase tgStateInfo
    Erase tgSubtotalGroupInfo
    Erase tgAttExpMon
    Erase tgReportNames
    Erase tgRff
    Erase tgRffExtended
    Erase tgUstInfo
    Erase tgDeptInfo
    
    
    Erase tgStationInfo
    Erase tgVehicleInfo
    Erase tgRnfInfo
    Erase tgAdvtInfo
    Erase sgStationImportTitles
    
    '9/11/06: Split Network stuff
    Erase tgRBofRec
    Erase tgSplitNetLastFill
    
    On Error Resume Next
    rstAlertUlf.Close
    On Error Resume Next
    rstAlert.Close
    'gLogMsg "Closing Pervasive API Engine. User: " & gGetComputerName(), "WebExportLog.Txt", False
    mClosePervasiveAPI
    cnn.Close
    
    Set CPTTRectify = Nothing
End Sub


Private Sub lbcVehicles_Click()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Rectify"
    cmdCancel.Caption = "&Cancel"
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
    mSetControl
End Sub

Private Sub edcStartDate_Change()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Rectify"
    cmdCancel.Caption = "&Cancel"
End Sub




Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload CPTTRectify
End Sub


Private Sub mInit()
    Dim sBuffer As String
    Dim lSize As Long
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim ilValue As Integer
    Dim ilValue8 As Integer
    Dim slDate As String
    Dim ilDatabase As Integer
    Dim ilLocation As Integer
    Dim ilSQL As Integer
    Dim ilForm As Integer
    Dim sMsg As String
    Dim iLoop As Integer
    Dim sCurDate As String
    Dim sAutoLogin As String
    Dim slTimeOut As String
    Dim slDSN As String
    Dim slStartIn As String
    Dim slStartStdMo As String
    Dim slTemp As String
    ReDim sWin(0 To 13) As String * 1
    Dim ilIsTntEmpty As Integer
    Dim ilIsShttEmpty As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    Dim EmailExists_rst As ADODB.Recordset
    '5/11/11
    Dim blAddGuide As Boolean
    'dan 2/23/12 can't have error handler in error handler
    Dim blNeedToCloseCnn As Boolean
    Dim slXMLINIInputFile As String
    
    
    sgCommand = Command$
    blNeedToCloseCnn = False
    'Display gMsgBox
    'igShowMsgBox = True shows the gMsgBox.
    'igShowMsgBox = False does not show any gMsgBox
    
    'Warning: One thing to remember is that if you are expecting a return value from a gMsgBox
    'and you turn gMsgBox off then you need to make sure that you handle that case.
    'example:   ilRet = gMsgBox "xxxx"
    igShowMsgBox = True
    
    'igDemoMode = False
    'If InStr(sgCommand, "Demo") Then
        igDemoMode = True
    'End If
    
    'Used to speed-up testing exports with multiple files reduce record count needed to create a new file
    igSmallFiles = False
    If InStr(sgCommand, "SmallFiles") Then
        igSmallFiles = True
    End If
    
    igAutoImport = False
    slStartIn = CurDir$
    sgCurDir = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommand, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
        
    sgBS = Chr$(8)  'Backspace
    sgTB = Chr$(9)  'Tab
    sgLF = Chr$(10) 'Line Feed (New Line)
    sgCR = Chr$(13) 'Carriage Return
    sgCRLF = sgCR + sgLF
   
   
    ilRet = 0
    ilLocation = False
    ilDatabase = False
    sgDatabaseName = ""
    sgReportDirectory = ""
    sgExportDirectory = ""
    sgImportDirectory = ""
    sgExeDirectory = ""
    sgLogoDirectory = ""
    sgPasswordAddition = ""
    sgSQLDateForm = "yyyy-mm-dd"
    sgCrystalDateForm = "yyyy,mm,dd"
    sgSQLTimeForm = "hh:mm:ss"
    igSQLSpec = 1               'Pervasive 2000
    sgShowDateForm = "m/d/yyyy"
    sgShowTimeWOSecForm = "h:mma/p"
    sgShowTimeWSecForm = "h:mm:ssa/p"
    igWaitCount = 10
    igTimeOut = -1
    sgWallpaper = ""
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    sgLogoName = "rptlogo.bmp"
    sgNowDate = ""
    ilPos = InStr(1, sgCommand, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommand, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommand, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gIsDate(slDate) Then
            sgNowDate = slDate
        End If
    End If
    
    If Not gLoadOption("Locations", "Logo", sgLogoPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    
    
    If Not gLoadOption("Database", "Name", sgDatabaseName) Then
        gMsgBox "Affiliat.Ini [Database] 'Name' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Reports", sgReportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Reports' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Export' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Exe", sgExeDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Exe' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    End If
    
        
    'Import is optional
    If gLoadOption("Locations", "Import", sgImportDirectory) Then
        sgImportDirectory = gSetPathEndSlash(sgImportDirectory, True)
    Else
        sgImportDirectory = ""
    End If
    
    If gLoadOption("Locations", "ContractPDF", sgContractPDFPath) Then
        sgContractPDFPath = gSetPathEndSlash(sgContractPDFPath, True)
    Else
        sgContractPDFPath = ""
    End If
    
    
    'Commented out below because I can't see why you would need a backslash
    'on the end of a DSN name
    'sgDatabaseName = gSetPathEndSlash(sgDatabaseName)
    sgReportDirectory = gSetPathEndSlash(sgReportDirectory, True)
    sgExportDirectory = gSetPathEndSlash(sgExportDirectory, True)
    sgExeDirectory = gSetPathEndSlash(sgExeDirectory, True)
    sgLogoDirectory = gSetPathEndSlash(sgLogoDirectory, True)
    
    Call gLoadOption("SQLSpec", "Date", sgSQLDateForm)
    Call gLoadOption("SQLSpec", "Time", sgSQLTimeForm)
    If gLoadOption("SQLSpec", "System", sBuffer) Then
        If sBuffer = "P7" Then
            igSQLSpec = 0
        End If
    End If
    If gLoadOption("Locations", "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption("Locations", "Wallpaper", sgWallpaper)
    
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    
    If Not gLoadOption("Locations", "DBPath", sgDBPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    
    'Set Message folder
    If Not gLoadOption("Locations", "DBPath", sgMsgDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload CPTTRectify
        Exit Sub
    Else
        sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory, True) & "Messages\"
'        sgMsgDirectory = CurDir
'        If InStr(1, sgMsgDirectory, "Data", vbTextCompare) Then
'            sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory) & "Messages\"
'        Else
'            sgMsgDirectory = sgExportDirectory
'        End If
    End If
    
    ' Not sure what section this next item is coming from. The original code did not specify.
    'Call gLoadOption("SQLSpec", "WaitCount", sBuffer)
    'igWaitCount = Val(sBuffer)
    
    On Error GoTo ErrHand
    Set cnn = New ADODB.Connection
   
    'Set env = rdoEnvironments(0)
    'cnn.CursorDriver = rdUseOdbc
    
    'Set cnn = cnn.OpenConnection(dsName:="Affiliate", Prompt:=rdDriverCompleteRequired)
    ' The default timeout is 15 seconds. This always fails on my PC the first time I run this program.


    slDSN = sgDatabaseName
    'ttp 4905.  Need to try connection. If it fails, try one more time, after sleeping.
    'cnn.Open "DSN=" & slDSN
    
    On Error GoTo ERRNOPERVASIVE
    ilRet = 0
    cnn.Open "DSN=" & slDSN
    
    On Error GoTo ErrHand
    If ilRet = 1 Then
        Sleep 2000
        cnn.Open "DSN=" & slDSN
    End If

    
    
    'Example of using a user name and password
    'cnn.Open "DSN=" & slDSN, "Master", "doug"
    Set rst = New ADODB.Recordset

    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
 
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(slDSN, Len(slDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        slDSN = Left(slDSN, Len(slDSN) - 1)
    End If
    'Set cnn = cnn.OpenConnection(dsName:=slDSN, Prompt:=rdDriverCompleteRequired)
    'If igTimeOut >= 0 Then
    '    cnn.QueryTimeout = igTimeOut
    'End If
    'Code modified for testing
    
    
    If Not mOpenPervasiveAPI Then
        Unload CPTTRectify
        Exit Sub
    End If
    
    
    'Test for Guide- if not added- add
    'SQLQuery = "Select MAX(ustCode) from ust"
    'Set rst = cnn.Execute(SQLQuery)
    ''If rst(0).Value = 0 Then
    'If IsNull(rst(0).Value) Then
    ''5/11/11
    '    blAddGuide = True
    'Else
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = cnn.Execute(SQLQuery)
        If rst.EOF Then
            blAddGuide = True
        Else
            blAddGuide = False
            imGuideUstCode = rst!ustCode
        End If
    'End If
    If blAddGuide Then
    '5/11/11
        'SQLQuery = "INSERT INTO ust(ustName, ustPassword, ustState)"
        'SQLQuery = SQLQuery & "VALUES ('Guide', 'Guide', 0)"
        sCurDate = Format(Now, sgShowDateForm)
        For iLoop = 0 To 13 Step 1
            sWin(iLoop) = "I"
        Next iLoop
        '5/11/11
        'mResetGuideGlobals
        SQLQuery = "INSERT INTO ust(ustName, ustReportName, ustPassword, "
        SQLQuery = SQLQuery & "ustState, ustPassDate, ustActivityLog, ustWin1, "
        SQLQuery = SQLQuery & "ustWin2, ustWin3, ustWin4, "
        SQLQuery = SQLQuery & "ustWin5, ustWin6, ustWin7, "
        SQLQuery = SQLQuery & "ustWin8, ustWin9, ustPledge, "
        SQLQuery = SQLQuery & "ustExptSpotAlert, ustExptISCIAlert, ustTrafLogAlert, "
        SQLQuery = SQLQuery & "ustWin10, ustWin11, ustWin12, ustWin13, "
        SQLQuery = SQLQuery & "ustWin14, ustWin15, ustPhoneNo, ustCity, ustEMailCefCode, ustAllowedToBlock, "
        SQLQuery = SQLQuery & "ustWin16, "
        SQLQuery = SQLQuery & "ustUserInitials, "
        SQLQuery = SQLQuery & "ustDntCode, "
        SQLQuery = SQLQuery & "ustAllowCmmtChg, "
        SQLQuery = SQLQuery & "ustAllowCmmtDelete, "
        SQLQuery = SQLQuery & "ustUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "VALUES ('" & "Guide" & "', "
        SQLQuery = SQLQuery & "'" & "System" & "', '" & "Guide" & "', "
        SQLQuery = SQLQuery & 0 & ", '" & Format$(sCurDate, sgSQLDateForm) & "', '" & "V" & "', '" & sgUstWin(1) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(2) & "', '" & sgUstWin(3) & "', '" & sgUstWin(4) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(5) & "', '" & sgUstWin(6) & "', '" & sgUstWin(7) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(8) & "', '" & sgUstWin(9) & "', '" & sgUstPledge & "', "
        SQLQuery = SQLQuery & "'" & sgExptSpotAlert & "', '" & sgExptISCIAlert & "', '" & sgTrafLogAlert & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(10) & "', '" & sgUstWin(11) & "', '" & sgUstWin(12) & "', '" & sgUstWin(13) & "', "
        SQLQuery = SQLQuery & "'" & sgUstClear & "', '" & sgUstDelete & "', "
        SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', " & 0 & ", '" & "Y" & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(sgUstWin(0)) & "', "
        SQLQuery = SQLQuery & "'" & "G" & "', "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & sgUstAllowCmmtChg & "', "
        SQLQuery = SQLQuery & "'" & sgUstAllowCmmtDelete & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        cnn.BeginTrans
        blNeedToCloseCnn = True
        'cnn.ConnectionTimeout = 30  ' Increase from the default of 15 to 30 seconds.
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHand:
        End If
        cnn.CommitTrans
        blNeedToCloseCnn = False
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = cnn.Execute(SQLQuery)
        If Not rst.EOF Then
            imGuideUstCode = rst!ustCode
        Else
            imGuideUstCode = 0
        End If
    End If
    
    gUsingCSIBackup = False
    gUsingXDigital = False
    gWegenerExport = False
    gOLAExport = False
    ' Dan M added spfusingFeatures2
    SQLQuery = "SELECT spfGClient, spfGAlertInterval, spfGUseAffSys, spfUsingFeatures7, spfUsingFeatures2, spfUsingFeatures8"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = cnn.Execute(SQLQuery)
    
    If Not rst.EOF Then
        If UCase(rst!spfGUseAffSys) <> "Y" Then
            gMsgBox "The Affiliate system has not been activated.  Please call Counterpoint.", vbCritical
            Unload CPTTRectify
            Exit Sub
        End If
        ilValue8 = Asc(rst!spfUsingFeatures8)
        If (ilValue8 And ALLOWMSASPLITCOPY) <> ALLOWMSASPLITCOPY Then
            gUsingMSARegions = False
        Else
            gUsingMSARegions = True
        End If
        If (ilValue8 And ISCIEXPORT) <> ISCIEXPORT Then
            gISCIExport = False
        Else
            gISCIExport = True
        End If
        ilValue = Asc(rst!spfUsingFeatures7)
        If (ilValue And CSIBACKUP) <> CSIBACKUP Then
            gUsingCSIBackup = False
        Else
            gUsingCSIBackup = True
        End If
        
        If ((ilValue And XDIGITALISCIEXPORT) <> XDIGITALISCIEXPORT) And ((ilValue8 And XDIGITALBREAKEXPORT) <> XDIGITALBREAKEXPORT) Then
            gUsingXDigital = False
        Else
            gUsingXDigital = True
        End If
        If (ilValue And WEGENEREXPORT) <> WEGENEREXPORT Then
            gWegenerExport = False
        Else
            gWegenerExport = True
        End If
        If (ilValue And OLAEXPORT) <> OLAEXPORT Then
            gOLAExport = False
        Else
            gOLAExport = True
        End If
        ilValue = Asc(rst!spfusingfeatures2)
        If (ilValue And STRONGPASSWORD) <> STRONGPASSWORD Then
            bgStrongPassword = False
        Else
            bgStrongPassword = True
        End If
    End If
    
    If Not rst.EOF Then
        sgClientName = Trim$(rst!spfGClient)
        igAlertInterval = rst!spfGAlertInterval
    Else
        sgClientName = "Unknown"
        gMsgBox "Client name is not defined in Site Options"
        igAlertInterval = 0
    End If
    
    If InStr(1, sgCommand, "NoAlerts", vbTextCompare) > 0 Then
        'For Debug ONLY
        igAlertInterval = 0
    End If
    
    If Trim$(sgNowDate) = "" Then
        If InStr(1, sgClientName, "XYZ Broadcasting", vbTextCompare) > 0 Then
            sgNowDate = "12/15/1999"
        End If
    End If


    ilRet = gInitGlobals()
    If ilRet = 0 Then
        'While Not gVerifyWebIniSettings()
        '    frmWebIniOptions.Show vbModal
        '    If Not igWebIniOptionsOK Then
        '        Unload CPTTRectify
        '        Exit Sub
        '    End If
        'Wend
    End If
    
    Call gLoadOption("Database", "AutoLogin", sAutoLogin)
    
    
    On Error GoTo ErrHand
    'If Not igAutoImport Then
    '    ilRet = mInitAPIReport()      '4-19-04
    'End If
    
    
    ilRet = gTestWebVersion()
    'Move report logo to local C drice (c:\csi\rptlogo.bmp)
    ilRet = 0
    On Error GoTo mStartUpErr:
    'slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
    'If ilRet <> 0 Then
    '    ilRet = 0
    '    MkDir "C:\CSI"
    '    If ilRet = 0 Then
    '        FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
    '    Else
    '        FileCopy sgDBPath & "RptLogo.Bmp", sgLogoPath & "RptLogo.Bmp"
    '    End If
    'Else
    '    ilRet = 0
    '    slDateTime2 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
    '    If ilRet = 0 Then
    '        If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
    '            FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
    '        End If
    '    End If
    'End If
     'ttp 5260
    'If Dir(sgLogoPath & "RptLogo.jpg") > "" Then
    '    If Dir("c:\csi\RptLogo.jpg") = "" Then
    '        FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
    '    'ok, both exist.  is logopath's more recent?
    '    Else
    '        slDateTime1 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
    '        slDateTime2 = FileDateTime("C:\CSI\RptLogo.jpg")
    '        If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
     '           FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
    '        End If
    '    End If
    'End If
    'Determine number if X-Digital HeadEnds
    ReDim sgXDSSection(0 To 0) As String
    'slXMLINIInputFile = gXmlIniPath(True)
    'If LenB(slXMLINIInputFile) <> 0 Then
    '    ilRet = gSearchFile(slXMLINIInputFile, "[XDigital", True, 1, sgXDSSection())
    'End If
    'Test to see if this function has been ran before, if so don't run it again
    igEmailNeedsConv = False
    mCreateStatustype
    ilRet = gPopMarkets()
    ilRet = gPopMSAMarkets()         'MSA markets
    ilRet = gPopMntInfo("T", tgTerritoryInfo())
    ilRet = gPopMntInfo("C", tgCityInfo())
    ilRet = gPopOwnerNames()
    ilRet = gPopStations()
    ilRet = gPopVehicleOptions()
    ilRet = gPopVehicles()
    ilRet = gPopSellingVehicles()
    ilRet = gPopAdvertisers()
    ilRet = gPopReportNames()
    ilRet = gGetLatestRatecard()
    ilRet = gPopTimeZones()
    ilRet = gPopStates()
    ilRet = gPopFormats()
    ilRet = gPopAvailNames()
    ilRet = gPopMediaCodes()
    
    Exit Sub

mStartUpErr:
    ilRet = Err.Number
    Resume Next
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
'    gMsg = ""
'    For Each gErrSQL In cnn.Errors
'        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
'            gMsg = "A SQL error has occured: "
'            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
'        End If
'    Next gErrSQL
'    On Error Resume Next
'    cnn.RollbackTrans
'    On Error GoTo 0
'    If gMsg = "" Then
'        gMsgBox "Error at Start-up " & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
'    End If
    'ttp 5217
    gHandleError "", FORMNAME & "-Form_Load"
    'ttp 4905 need to quit app
    bmNoPervasive = True
    If blNeedToCloseCnn Then
        cnn.RollbackTrans
    End If
    'unload affiliate  ttp 4905
    tmcTerminate.Enabled = True
End Sub
Private Sub mCreateStatustype()
    'Agreement only shows status- 1:; 2:; 5: and 9:
    'All other screens show all the status
    tgStatusTypes(0).sName = "1-Aired Live"        'In Agreement and Pre_Log use 'Air Live'
    tgStatusTypes(0).iPledged = 0
    tgStatusTypes(0).iStatus = 0
    tgStatusTypes(1).sName = "2-Aired Delay B'cast" '"2-Aired In Daypart"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(1).iPledged = 1
    tgStatusTypes(1).iStatus = 1
    tgStatusTypes(2).sName = "3-Not Aired Tech Diff"
    tgStatusTypes(2).iPledged = 2
    tgStatusTypes(2).iStatus = 2
    tgStatusTypes(3).sName = "4-Not Aired Blackout"
    tgStatusTypes(3).iPledged = 2
    tgStatusTypes(3).iStatus = 3
    tgStatusTypes(4).sName = "5-Not Aired Other"
    tgStatusTypes(4).iPledged = 2
    tgStatusTypes(4).iStatus = 4
    tgStatusTypes(5).sName = "6-Not Aired Product"
    tgStatusTypes(5).iPledged = 2
    tgStatusTypes(5).iStatus = 5
    tgStatusTypes(6).sName = "7-Aired Outside Pledge"  'In Pre-Log use 'Air-Outside Pledge'
    tgStatusTypes(6).iPledged = 3
    tgStatusTypes(6).iStatus = 6
    tgStatusTypes(7).sName = "8-Aired Not Pledged"  'in Pre-Log use 'Air-Not Pledged'
    tgStatusTypes(7).iPledged = 3
    tgStatusTypes(7).iStatus = 7
    'D.S. 11/6/08 remove the "or Aired" from the status 9 description
    'Affiliate Meeting Decisions item 5) f-iv
    'tgStatusTypes(8).sName = "9-Not Carried or Aired"
    tgStatusTypes(8).sName = "9-Not Carried"
    tgStatusTypes(8).iPledged = 2
    tgStatusTypes(8).iStatus = 8
    tgStatusTypes(9).sName = "10-Delay Cmml/Prg"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(9).iPledged = 1
    tgStatusTypes(9).iStatus = 9
    tgStatusTypes(10).sName = "11-Air Cmml Only"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(10).iPledged = 1
    tgStatusTypes(10).iStatus = 10
    tgStatusTypes(ASTEXTENDED_MG).sName = "MG"
    tgStatusTypes(ASTEXTENDED_MG).iPledged = 3
    tgStatusTypes(ASTEXTENDED_MG).iStatus = ASTEXTENDED_MG
    tgStatusTypes(ASTEXTENDED_BONUS).sName = "Bonus"
    tgStatusTypes(ASTEXTENDED_BONUS).iPledged = 3
    tgStatusTypes(ASTEXTENDED_BONUS).iStatus = ASTEXTENDED_BONUS
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).sName = "Replacement"
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iPledged = 3
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iStatus = ASTEXTENDED_REPLACEMENT
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).sName = "15-Missed MG Bypassed"
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iPledged = 2
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iStatus = ASTAIR_MISSED_MG_BYPASS
End Sub



Private Function mGetCpttComplete(ilVefCode As Integer, ilShttCode As Integer, llAttCode As Long, llDate As Long, ilCpttStatus As Integer, ilCpttPostingStatus As Integer, slAstStatus As String) As Boolean
    'Created by D.S. June 2007  Modified Dan M 11/02/10 V81 new values in cptt added 2/25/2011
    'Set the CPTT week's value

    Dim ilStatus As Integer
    Dim llVeh As Long
    'new values in cptt
    Dim slMondayFeedDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slSuDate As String
    Dim llAstCount As Long
    Dim llNotAiredCount As Long
    Dim blNoSpotsAired As Boolean
    Dim slMsg As String
    Dim ilShtt As Integer
    
    On Error GoTo ErrHand
    
    mGetCpttComplete = False
    
    ilCpttStatus = 0
    ilCpttPostingStatus = 0
    slAstStatus = "N"
    
    If imTerminate Then
        Exit Function
    End If
    slMondayFeedDate = Format(llDate, "m/d/yy")
    slSuDate = DateAdd("d", 6, slMondayFeedDate)
    
    SQLQuery = "Select Count(*) FROM ast WHERE astAtfCode = " & llAttCode
    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
    Set ast_rst = cnn.Execute(SQLQuery)
    If Not ast_rst.EOF Then
        llAstCount = ast_rst(0).Value
        If llAstCount > 0 Then
            'Set any Not Aired to received as they are not exported
            For ilStatus = 0 To UBound(tgStatusTypes) Step 1
                DoEvents
                If imTerminate Then
                    Exit Function
                End If
                If (tgStatusTypes(ilStatus).iPledged = 2) Then
                    SQLQuery = "UPDATE ast SET "
                    SQLQuery = SQLQuery & "astCPStatus = " & "1"    'Received
                    SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAttCode
                    SQLQuery = SQLQuery & " AND astCPStatus = 0"
                    SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
                    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                End If
            Next ilStatus
            llNotAiredCount = 0
            For ilStatus = 0 To UBound(tgStatusTypes) Step 1
                DoEvents
                If imTerminate Then
                    Exit Function
                End If
                If (tgStatusTypes(ilStatus).iPledged = 2) Then
                    SQLQuery = "Select Count(*) FROM ast WHERE astAtfCode = " & llAttCode
                    SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
                    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    Set ast_rst = cnn.Execute(SQLQuery)
                    If Not ast_rst.EOF Then
                        llNotAiredCount = llNotAiredCount + ast_rst(0).Value
                    End If
                End If
            Next ilStatus
            If llAstCount <> llNotAiredCount Then
                blNoSpotsAired = False
            Else
                blNoSpotsAired = True
            End If
            If imTerminate Then
                Exit Function
            End If
    
            'Determine if CPTTStatus should to set to 0=Partial or 1=Completed:  because of above code, will always be complete
            SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
            SQLQuery = SQLQuery & " AND astAtfCode = " & llAttCode
            SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            Set ast_rst = cnn.Execute(SQLQuery)
            DoEvents
            If imTerminate Then
                Exit Function
            End If
            If ast_rst.EOF Then
                'Set CPTT as complete
                SQLQuery = "UPDATE cptt SET "
                llVeh = gBinarySearchVef(CLng(ilVefCode))
                If llVeh <> -1 Then
                    If (tgVehicleInfo(llVeh).sVehType = "G") And (DateValue(slSuDate) > DateValue(Format$(gNow(), "m/d/yy"))) Then
                        ilCpttStatus = 0 'Partial
                        ilCpttPostingStatus = 1 'Partial
                        mAddStatic llDate, 1, 0, 1, 0, 0, 0, 0, 0
                        mAddNotCreated llDate, ilVefCode, ilShttCode
                    Else
                        slAstStatus = "C"
                        If blNoSpotsAired Then
                            ilCpttStatus = 2 'Complete
                        Else
                            ilCpttStatus = 1 'Complete
                        End If
                        ilCpttPostingStatus = 2  'Complete
                        mAddStatic llDate, 1, 0, 0, 1, 0, 0, 0, 0
                    End If
                Else
                    slAstStatus = "C"
                    If blNoSpotsAired Then
                        ilCpttStatus = 2 'Complete
                    Else
                        ilCpttStatus = 1 'Complete
                    End If
                    ilCpttPostingStatus = 2  'Complete
                    mAddStatic llDate, 1, 0, 0, 1, 0, 0, 0, 0
                End If
            Else
                SQLQuery = "Select count(*) FROM ast WHERE astCPStatus = 1"
                SQLQuery = SQLQuery & " AND astAtfCode = " & llAttCode
                SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                Set ast_rst = cnn.Execute(SQLQuery)
                If Not ast_rst.EOF Then
                    If (ast_rst(0).Value > 0) And (ast_rst(0).Value <> llNotAiredCount) Then
                '        lmNumberPartials = lmNumberPartials + 1
                '        slMsg = Trim$(Str$(llAttCode))
                '        slMsg = slMsg & "," & slMondayFeedDate
                '        llVeh = gBinarySearchVef(CLng(ilVefCode))
                '        If llVeh <> -1 Then
                '            slMsg = slMsg & "," & Trim$(tgVehicleInfo(llVeh).sVehicle)
                '        End If
                '        ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
                '        If ilShtt <> -1 Then
                '            slMsg = slMsg & "," & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                '        End If
                '        Print #hmMsg, "Partially Posted:," & slMsg
                        ilCpttStatus = 0
                        ilCpttPostingStatus = 1 'Partial
                        mAddStatic llDate, 1, 0, 1, 0, 0, 0, 0, 0
                        mAddNotCreated llDate, ilVefCode, ilShttCode
                    Else
                        mAddStatic llDate, 1, 1, 0, 0, 0, 0, 0, 0
                    End If
                Else
                    mAddStatic llDate, 1, 1, 0, 0, 0, 0, 0, 0
                End If
            End If
        Else
            mAddStatic llDate, 1, 1, 0, 0, 1, 0, 0, 0
        End If
    Else
        mAddStatic llDate, 1, 1, 0, 0, 1, 0, 0, 0
    End If
    mGetCpttComplete = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTTRectify: mSetCpttComplete"
End Function


Private Function mAddCptt(ilVefCode As Integer)
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilCycle As Integer
    Dim slCurDate As String
    Dim ilStatus As Integer
    Dim ilPostingStatus As Integer
    Dim slAstStatus As String
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slTime As String
    Dim ilRet As Integer
    Dim llCpttCode As Long
    Dim slDate As String
    Dim slLLD As String
    Dim llLLD As Long
    Dim blCheckAst As Boolean
    
    On Error GoTo ErrHand
    mAddCptt = False
    
    'Test if CPTT exist
    SQLQuery = "SELECT vpfLLD, vpfGenLog"
    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & ilVefCode & ")"
    slLLD = ""
    Set vpf_rst = cnn.Execute(SQLQuery)
    If Not vpf_rst.EOF Then
        If Not IsNull(vpf_rst!vpfLLD) Then
            If gIsDate(vpf_rst!vpfLLD) Then
                'set sLLD to last log date
                slLLD = Format$(vpf_rst!vpfLLD, sgShowDateForm)
            End If
        End If
        If vpf_rst!vpfGenLog = "N" Then
            'slLLD = ""
            If Not gIsDate(slLLD) Then
                slLLD = Format(gNow(), "m/d/yy")
            Else
                If gDateValue(Format(gNow(), "m/d/yy")) > gDateValue(slLLD) Then
                    slLLD = Format(gNow(), "m/d/yy")
                End If
            End If
        End If
    End If
    If slLLD = "" Then
        mAddCptt = True
        Exit Function
    End If
    llLLD = gDateValue(gAdjYear(slLLD))
    ilCycle = 7  'rst_Att!vpfLNoDaysCycle
    slCurDate = Format(gNow(), sgShowDateForm)
    slTime = Format("12:00AM", "hh:mm:ss")
    
    SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
    SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(smEndDate), sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(smStartDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(smStartDate), sgSQLDateForm) & "')" & ")"
    Set att_rst = cnn.Execute(SQLQuery)
    Do While Not att_rst.EOF
        DoEvents
        If imTerminate Then
            Exit Function
        End If
        llStartDate = gDateValue(Format(att_rst!attOnAir, "m/d/yy"))
        If gDateValue(Format(att_rst!attOffAir, "m/d/yy")) < gDateValue(Format(att_rst!attDropDate, "m/d/yy")) Then
            llEndDate = gDateValue(Format(att_rst!attOffAir, "m/d/yy"))
        Else
            llEndDate = gDateValue(Format(att_rst!attDropDate, "m/d/yy"))
        End If
        If lmStartDate > llStartDate Then
            llStartDate = lmStartDate
        End If
        If lmEndDate < llEndDate Then
            llEndDate = lmEndDate
        End If
        If llLLD < llEndDate Then
            llEndDate = llLLD
        End If
        If att_rst!attCarryCmml = 0 Then
            If llLLD >= llStartDate Then
                For llDate = llStartDate To llEndDate Step 7
                    DoEvents
                    If imTerminate Then
                        Exit Function
                    End If
                    If ckcAST.Value = vbChecked Then
                        blCheckAst = True
                    Else
                        blCheckAst = False
                    End If
                    SQLQuery = "Select * FROM cptt WHERE cpttVefCode = " & ilVefCode
                    SQLQuery = SQLQuery & " AND cpttShfCode = " & att_rst!attshfCode
                    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(llDate, sgSQLDateForm) & "'"
                    Set cptt_rst = cnn.Execute(SQLQuery)
                    If Not cptt_rst.EOF Then
                        If cptt_rst!cpttatfCode <> att_rst!attCode Then
                            SQLQuery = "Select * FROM cptt WHERE cpttatfCode = " & att_rst!attCode
                            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(llDate, sgSQLDateForm) & "'"
                            Set cptt_rst = cnn.Execute(SQLQuery)
                        End If
                    End If
                    If cptt_rst.EOF Then
                        SQLQuery = "Select COUNT(lstCode) from LST"
                        SQLQuery = SQLQuery + " WHERE"
                        SQLQuery = SQLQuery + " lstLogVefCode = " & ilVefCode
                        SQLQuery = SQLQuery + " AND lstLogDate >= " & "'" & Format(llDate, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery + " AND lstLogDate <= " & "'" & Format(llDate + 6, sgSQLDateForm) & "'"
                        Set lst_rst = cnn.Execute(SQLQuery)
                        'If rst.EOF Then
                        If lst_rst(0).Value <> 0 Then
                            'Add CPTT
                            ilRet = mGetCpttComplete(ilVefCode, att_rst!attshfCode, att_rst!attCode, llDate, ilStatus, ilPostingStatus, slAstStatus)
                            DoEvents
                            If imTerminate Then
                                Exit Function
                            End If
                            If att_rst!attServiceAgreement = "Y" Then
                                ilStatus = 1
                            End If
                            SQLQuery = "INSERT INTO cptt(cpttCode, cpttAtfCode, cpttShfCode, cpttVefCode, "
                            SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, cpttCycle, "
                            SQLQuery = SQLQuery & "cpttAirTime, cpttStatus, cpttUsfCode, "
                            SQLQuery = SQLQuery & "cpttPrintStatus, cpttPostingStatus, cpttAstStatus)"
                            SQLQuery = SQLQuery & " VALUES "
                            SQLQuery = SQLQuery & "(Replace, " & att_rst!attCode & ", " & att_rst!attshfCode & ", " & att_rst!attvefCode & ", "
                            SQLQuery = SQLQuery & "'" & Format$(gAdjYear(slCurDate), sgSQLDateForm) & "', '" & Format(gAdjYear(Format$(llDate, "m/d/yy")), sgSQLDateForm) & "', " & ilCycle & ", "
                            SQLQuery = SQLQuery & "'" & Format$(slTime, sgSQLTimeForm) & "', " & ilStatus & ", " & igUstCode & ","
                            SQLQuery = SQLQuery & 0 & ", " & ilPostingStatus & ", '" & slAstStatus & "'" & ")"
                            'cnn.Execute SQLQuery, rdExecDirect
                            llCpttCode = gInsertAndReturnCode(SQLQuery, "cptt", "cpttCode", "Replace")
                            If llCpttCode <> -1 Then
                                'Don't get compliant count as AST created
                                'slDate = Format(llDate, "m/d/yy")
                                'ilRet = gSetCpttCount(att_rst!attCode, slDate, slDate)
                                gFileChgdUpdate "cptt.mkd", False
                            Else
                                blCheckAst = False
                            End If
                        Else
                            blCheckAst = False
                        End If
                    Else
                        SQLQuery = "Select COUNT(lstCode) from LST"
                        SQLQuery = SQLQuery + " WHERE"
                        SQLQuery = SQLQuery + " lstLogVefCode = " & ilVefCode
                        SQLQuery = SQLQuery + " AND lstLogDate >= " & "'" & Format(llDate, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery + " AND lstLogDate <= " & "'" & Format(llDate + 6, sgSQLDateForm) & "'"
                        Set lst_rst = cnn.Execute(SQLQuery)
                        'If rst.EOF Then
                        If lst_rst(0).Value = 0 Then
                            blCheckAst = False
                            mAddStatic llDate, 0, 0, 0, 0, 0, 0, 0, 1
                            mAddNotCreated llDate, att_rst!attvefCode, att_rst!attshfCode
                        End If
                    End If
                    If (blCheckAst) And (att_rst!attPostingType > 1) Then
                        SQLQuery = "Select COUNT(astCode) from AST"
                        SQLQuery = SQLQuery + " WHERE"
                        SQLQuery = SQLQuery + " astAtfCode = " & att_rst!attCode
                        SQLQuery = SQLQuery + " AND astFeedDate >= " & "'" & Format(llDate, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery + " AND astFeedDate <= " & "'" & Format(llDate + 6, sgSQLDateForm) & "'"
                        Set ast_rst = cnn.Execute(SQLQuery)
                        'If rst.EOF Then
                        If ast_rst(0).Value = 0 Then
                            ilRet = gBuildAst(hmAst, att_rst!attCode, Format(llDate, "m/d/yy"), tmAstInfo())
                            If ilRet Then
                                mAddStatic llDate, 0, 0, 0, 0, 0, 1, 0, 0
                            Else
                                mAddStatic llDate, 0, 0, 0, 0, 0, 0, 1, 0
                                mAddNotCreated llDate, att_rst!attvefCode, att_rst!attshfCode
                            End If
                        End If
                    End If
                Next llDate
            End If
        End If
        att_rst.MoveNext
    Loop
    mAddCptt = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTTRectify: mAddCptt"
End Function

Private Sub mAddStatic(llDate As Long, llCreated As Long, llNotPosted As Long, llPartiallyPosted As Long, llCompletedPosted As Long, llASTMissing As Long, llAstCreated As Long, llAstNotCreated As Long, llExtra As Long)
    Dim blFound As Boolean
    Dim llLoop As Long
    Dim llIndex As Long
    
    blFound = False
    For llLoop = 0 To UBound(tmCpttStatic) - 1 Step 1
        If (llDate = tmCpttStatic(llLoop).lDate) And (imVefCode = tmCpttStatic(llLoop).iVefCode) Then
            blFound = True
            llIndex = llLoop
        End If
    Next llLoop
    If Not blFound Then
        ReDim Preserve tmCpttStatic(0 To UBound(tmCpttStatic) + 1) As CPTTSTATIC
        llIndex = UBound(tmCpttStatic) - 1
        tmCpttStatic(llIndex).lDate = llDate
        tmCpttStatic(llIndex).lCreated = 0
        tmCpttStatic(llIndex).lNotPosted = 0
        tmCpttStatic(llIndex).lPartiallyPosted = 0
        tmCpttStatic(llIndex).lCompletedPosted = 0
        tmCpttStatic(llIndex).lAstMissing = 0
        tmCpttStatic(llIndex).lAstCreated = 0
        tmCpttStatic(llIndex).lAstNotCreated = 0
        tmCpttStatic(llIndex).lExtra = 0
    End If
    tmCpttStatic(llIndex).lDate = llDate
    tmCpttStatic(llIndex).iVefCode = imVefCode
    tmCpttStatic(llIndex).lCreated = tmCpttStatic(llIndex).lCreated + llCreated
    tmCpttStatic(llIndex).lNotPosted = tmCpttStatic(llIndex).lNotPosted + llNotPosted
    tmCpttStatic(llIndex).lPartiallyPosted = tmCpttStatic(llIndex).lPartiallyPosted + llPartiallyPosted
    tmCpttStatic(llIndex).lCompletedPosted = tmCpttStatic(llIndex).lCompletedPosted + llCompletedPosted
    tmCpttStatic(llIndex).lAstMissing = tmCpttStatic(llIndex).lAstMissing + llASTMissing
    tmCpttStatic(llIndex).lAstCreated = tmCpttStatic(llIndex).lAstCreated + llAstCreated
    tmCpttStatic(llIndex).lAstNotCreated = tmCpttStatic(llIndex).lAstNotCreated + llAstNotCreated
    tmCpttStatic(llIndex).lExtra = tmCpttStatic(llIndex).lExtra + llExtra
End Sub

Private Sub mAddNotCreated(llDate As Long, ilVefCode As Integer, ilShttCode As Integer)
    Dim llIndex As Long
    Dim llVef As Long
    Dim ilShtt As Integer
    Dim slVehicleName As String
    Dim slStation As String
    Dim slStr As String
    
    If ckcNotAst.Value = vbUnchecked Then
        Exit Sub
    End If
    llIndex = UBound(tmNotCreated)
    slStr = llDate
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    llVef = gBinarySearchVef(CLng(ilVefCode))
    If llVef <> -1 Then
        slVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
    Else
        slVehicleName = ilVefCode
        Do While Len(slVehicleName) < 6
            slVehicleName = "0" & slVehicleName
        Loop
    End If
    ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
    If ilShtt <> -1 Then
        slStation = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
    Else
        slStation = ilShttCode
        Do While Len(slStation) < 5
            slStation = "0" & slStation
        Loop
    End If
    tmNotCreated(llIndex).sKey = slStr & slVehicleName & slStation
    tmNotCreated(llIndex).lDate = llDate
    tmNotCreated(llIndex).iVefCode = ilVefCode
    tmNotCreated(llIndex).iShttCode = ilShttCode
    ReDim Preserve tmNotCreated(0 To llIndex + 1) As NOTCREATED
End Sub

Private Sub mSetControl()
    If ckcAST.Value = vbChecked Then
        If (lbcVehicles.SelCount >= 1) And (lbcVehicles.SelCount <= 2) Then
            ckcNotAst.Enabled = True
        Else
            ckcNotAst.Value = vbUnchecked
            ckcNotAst.Enabled = False
        End If
    Else
        ckcNotAst.Value = vbUnchecked
        ckcNotAst.Enabled = False
    End If
End Sub
