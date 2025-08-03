VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCPTTCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CPTT Check"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "AffCPTTCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Check Posted Only Dates"
      Height          =   855
      Left            =   1320
      TabIndex        =   11
      Top             =   4560
      Width           =   7095
      Begin VB.TextBox txtEnd 
         Height          =   285
         Left            =   5040
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblEnd 
         Alignment       =   2  'Center
         Caption         =   "End Date:"
         Height          =   210
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblStart 
         Alignment       =   2  'Center
         Caption         =   "Start Date:"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtWeeksChanged 
      Height          =   285
      Left            =   8040
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTtlWeeks 
      Height          =   285
      Left            =   5250
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txtWeekNum 
      Height          =   285
      Left            =   3870
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.CommandButton cmdPostedWeeks 
      Caption         =   "Check &Posted Only"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   5625
      Width           =   2295
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "&Fix All"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   5625
      Width           =   1215
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4560
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6270
      FormDesignWidth =   9225
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7245
      TabIndex        =   3
      Top             =   5625
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "C&heck All"
      Height          =   375
      Left            =   1305
      TabIndex        =   0
      Top             =   5625
      Width           =   1215
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   3840
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6773
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Hidden"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vehicle"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Station"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "# Prior"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Agreement"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "# After"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "# to Add"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "# to Delete"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Of"
      Height          =   210
      Left            =   4650
      TabIndex        =   9
      Top             =   4110
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Weeks Updated:"
      Height          =   210
      Left            =   6120
      TabIndex        =   8
      Top             =   4110
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lacProgress 
      Alignment       =   2  'Center
      Caption         =   "Processing Week:"
      Height          =   210
      Left            =   90
      TabIndex        =   5
      Top             =   4110
      Visible         =   0   'False
      Width           =   3585
   End
End
Attribute VB_Name = "frmCPTTCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAffDP - Affiliate Daypart Information
'*
'*  Created August,2001 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc.
'*****************************************************
Option Explicit
Dim imCheckCleared As Integer
Dim imFixedCleared As Integer
Dim smCPTTCheck As String
Dim imCPTTCheckOK As Integer
Dim tmCPTTCheck() As CPTTCHECK
Dim tmWebSpotsInfo() As WEBSPOTSINFO
Dim imTerminate As Integer
Dim imDone As Integer
Dim smLogFileName As String
Private smFileName As String
Private smStatus As String
Private smMsg1 As String
Private smMsg2 As String
Private smDTStamp As String
Private lmAstNotFoundOnWeb As Long
Private lmAstWebExported_NotPostedLocal As Long
Private lmAstWebNotExported_PostedLocal As Long
Private smWebWorkStatus As String



Const cmOneSecond As Long = 1000


Private Sub cmdCheck_Click()

    smCPTTCheck = ""
    Screen.MousePointer = vbHourglass
    smLogFileName = "CpttCheckLog.txt"
    imTerminate = False
    If Not imCheckCleared Then
        gLogMsg "", smLogFileName, True
        imCheckCleared = True
    Else
        gLogMsg "", smLogFileName, False
    End If
    gLogMsg "**** Starting Program ****", smLogFileName, False
    mCheckAndFixCPTTs "C", smLogFileName
    gLogMsg "**** Ending Program ****", smLogFileName, False
    Screen.MousePointer = vbDefault
    If imTerminate Then
        imDone = True
        Unload frmCPTTCheck
    Else
        cmdCheck.Enabled = True
        cmdCancel.Caption = "&Done"
    End If
    DoEvents
End Sub

Private Sub cmdCancel_Click()
    
    Dim ilRet As Integer
    
    If Not imDone Then
        If Not gTerminate(smLogFileName) Then
        
            Exit Sub
        Else
            imDone = True
        End If
    Else
        Unload frmCPTTCheck
    End If
    imTerminate = True

End Sub

Private Sub cmdFix_Click()
    
    Dim ilRet As Integer
    
    smCPTTCheck = ""
    imTerminate = False
    Screen.MousePointer = vbHourglass
        smLogFileName = "CpttFixLog.txt"
    If Not imFixedCleared Then
        gLogMsg "", smLogFileName, True
        imFixedCleared = True
    Else
        gLogMsg "", smLogFileName, False
    End If
    gLogMsg "**** Starting Program ****", smLogFileName, False
    mCheckAndFixCPTTs "F", smLogFileName
    gLogMsg "**** Ending Program ****", smLogFileName, False
    Screen.MousePointer = vbDefault
    If imTerminate Then
        imDone = True
        Unload frmCPTTCheck
    Else
        cmdCheck.Enabled = True
        cmdCancel.Caption = "&Done"
    End If
    DoEvents
End Sub



Private Sub cmdPostedWeeks_Click()
    
    smCPTTCheck = ""
    Screen.MousePointer = vbHourglass
    smLogFileName = "CpttCheckPostedOnly.txt"
    txtWeekNum.Visible = True
    txtTtlWeeks.Visible = True
    txtWeeksChanged.Visible = True
    lacProgress.Visible = True
    Label2.Visible = True
    Label1.Visible = True
    txtStart.Visible = True
    txtEnd.Visible = True
    
    imTerminate = False
    If Not imCheckCleared Then
        gLogMsg "", smLogFileName, False
        imCheckCleared = True
    Else
        gLogMsg "", smLogFileName, False
    End If
    gLogMsg "**** Starting Program ****", smLogFileName, False
    mCheckPostedWeeks
    gLogMsg "**** Ending Program ****", smLogFileName, False
    Screen.MousePointer = vbDefault
    If imTerminate Then
        imDone = True
        Unload frmCPTTCheck
    Else
        cmdCheck.Enabled = True
        cmdCancel.SetFocus
        cmdCancel.Caption = "&Done"
        Screen.MousePointer = vbDefault

    End If
    MsgBox "To review the results please: " & sgMsgDirectory & smLogFileName
    DoEvents

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    imCheckCleared = False
    imFixedCleared = False
    Me.Width = Screen.Width / 1.1
    Me.Height = Screen.Height / 1.35
    Me.Top = (Screen.Height - Me.Height) / 1.5
    Me.Left = (Screen.Width - Me.Width) / 2
    Screen.MousePointer = vbDefault
    imTerminate = False
    imDone = True
End Sub
Private Sub mCheckPostedWeeks()

' D.S. 04/2010 Purpose: Look for all cptt weeks that are posted as complete. Look at each cptt weeks ast records and then
' and then determine if the entire week was posted by the station as not aired.  If so, then set the did not air
' flag.  The result is that in Post CP the week shows in blue under the None Aired radio button and the new missing
' weeks report can be ran with accurate listings.  This is important so that the networks can follow-up on stations
' that are no longer airing spots.

        
    Dim ilRet As Integer
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    Dim rst_Cptt As ADODB.Recordset
    Dim rst_att As ADODB.Recordset
    Dim rst_Ast As ADODB.Recordset
    Dim llCpttTotalCount As Long
    Dim llCount As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilSpotAired As Integer
    Dim slVefName As String
    Dim slStaName As String
    Dim sLWeek As String
    Dim slTemp As String
    Dim llChgCount As Long
    Dim slMode As String
    Dim slStartDateRange As String
    Dim slEndDateRange As String
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    
    On Error GoTo ErrHand:
        
    slStartDateRange = Format(txtStart.Text, "yyyy-mm-dd")
    slEndDateRange = Format(txtEnd.Text, "yyyy-mm-dd")
    
    If slStartDateRange <> "" Then
        ilRet = gIsDate(slStartDateRange)
        If Not ilRet Then
            MsgBox ("The Start Date is Not Valid")
            txtStart.SetFocus
            Exit Sub
        End If
    End If
    
    If slStartDateRange = "" Then
        slStartDateRange = "1969-12-31"
        txtStart.Text = "12/31/1969"
    Else
        txtStart.Text = Format(slStartDateRange, "mm/dd/yyyy")
    End If
    
    If slEndDateRange <> "" Then
        ilRet = gIsDate(slEndDateRange)
        If Not ilRet Then
            MsgBox ("The End Date is Not Valid")
            txtEnd.SetFocus
            Exit Sub
        End If
    End If
        
    If slEndDateRange = "" Then
        slEndDateRange = "2069-12-31"
        txtEnd.Text = "TFN"
    Else
        txtEnd.Text = Format(slEndDateRange, "mm/dd/yyyy")
    End If
    
    gLogMsg "Scanning Date Range of " & Format(slStartDateRange, "mm/dd/yyyy") & " to " & Format(slEndDateRange, "mm/dd/yyyy"), smLogFileName, False
    
    If cmdPostedWeeks.Value = True Then
        slMode = "C"
    End If
        
    llChgCount = 0
    llCount = 0
    txtWeekNum.Text = ""
    txtTtlWeeks.Text = ""
    txtWeeksChanged.Text = ""

    'SQLQuery = "SELECT count(cpttCode) FROM cptt where cpttPostingStatus = 2 And cpttstatus = 1"
    SQLQuery = "SELECT count(cpttCode) FROM cptt where cpttPostingStatus = 2 And cpttstatus = 1 And cpttStartDate >= " & "'" & slStartDateRange & "'" & " And cpttStartDate <= " & "'" & slEndDateRange & "'"
    
    
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If rst_Cptt(0).Value > 0 Then
        llCpttTotalCount = rst_Cptt(0).Value
    Else
        llCpttTotalCount = 0
    End If
    txtTtlWeeks.Text = llCpttTotalCount
    txtWeekNum.Text = llCount
    
    'Check cpttPostingStatus = 2 (Complete) And cpttstatus = 1 (Posting Complete and Some Spots Aired)
    'SQLQuery = "SELECT cpttCode, cpttAtfCode, cpttStartDate, cpttVefCode, cpttShfCode FROM cptt where cpttPostingStatus = 2 And cpttstatus = 1"
    SQLQuery = "SELECT cpttCode, cpttAtfCode, cpttStartDate, cpttVefCode, cpttShfCode FROM cptt where cpttPostingStatus = 2 And cpttstatus = 1 And cpttStartDate >= " & "'" & slStartDateRange & "'" & " And cpttStartDate <= " & "'" & slEndDateRange & "'"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    ilRet = ilRet
    
    While Not rst_Cptt.EOF
        slStartDate = rst_Cptt!CpttStartDate
        slEndDate = DateAdd("d", 6, slStartDate)
        slStartDate = gAdjYear(Format$(slStartDate, "m/d/yy"))
        slEndDate = gAdjYear(DateAdd("d", 6, slStartDate))
        'Get all of the Ast record for the given Att code and Start and End dates
        '12/13/13
        'SQLQuery = "Select astCode, astPledgeStartTime, astStatus, astPledgeStatus, astAirTime, astCPStatus FROM ast WHERE "
        SQLQuery = "Select * FROM ast WHERE "
        SQLQuery = SQLQuery + " astAtfCode = " & rst_Cptt!cpttatfCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')"
        Set rst_Ast = gSQLSelectCall(SQLQuery)
        llCount = llCount + 1
        If llCount Mod 10 = 0 Then
            txtWeekNum.Text = llCount
            DoEvents
        End If
        
        'debug
        'If rst_Cptt!cpttCode = 4947967 Then
        '    ilRet = ilRet
        'End If

        ilSpotAired = False
        While Not rst_Ast.EOF And Not ilSpotAired
            DoEvents
            
            '12/13/13: Obtain Pledge information from Dat
            tlDatPledgeInfo.lAttCode = rst_Ast!astAtfCode
            tlDatPledgeInfo.lDatCode = rst_Ast!astDatCode
            tlDatPledgeInfo.iVefCode = rst_Ast!astVefCode
            tlDatPledgeInfo.sFeedDate = Format(rst_Ast!astFeedDate, "m/d/yy")
            tlDatPledgeInfo.sFeedTime = Format(rst_Ast!astFeedTime, "hh:mm:ssam/pm")
            ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
            
            ''Was it Pledged to Air? astPledgeStatus <= 1 indicates Yes, >= 2 indicates No
            ''tgStatusTypes(rst_Ast!astPledgeStatus).iPledged
            '12/13/13
            'If rst_Ast!astPledgeStatus <= 1 Or rst_Ast!astPledgeStatus = 8 Or rst_Ast!astPledgeStatus = 9 Then
            If tlDatPledgeInfo.iPledgeStatus <= 1 Or tlDatPledgeInfo.iPledgeStatus = 8 Or tlDatPledgeInfo.iPledgeStatus = 9 Then
            ''If tgStatusTypes(rst_Ast!astPledgeStatus).iPledged <= 1 Then
                'We now know that it was Plegded to Air
                'If rst_Ast!astCPStatus = 1 Then
                    'We now know that it was Recieved (astCPStatus = 1 = Recv.)
                    If gGetAirStatus(rst_Ast!astStatus) <= 1 Or gGetAirStatus(rst_Ast!astStatus) = 9 Then
                        'We now know that it did Air
                        ilSpotAired = True
                    'Else
                    '    'We now know that it did Air
                    '    ilRet = ilRet
                    End If
                'End If
            End If
            rst_Ast.MoveNext
        Wend

        If Not ilSpotAired Then
            slVefName = gGetVehNameByVefCode(rst_Cptt!cpttvefcode)
            slStaName = gGetCallLettersByShttCode(rst_Cptt!cpttshfcode)
            sLWeek = Format(rst_Cptt!CpttStartDate, "m/d/yy")
            slTemp = "Updating Cptt Number: " & rst_Cptt!cpttCode & ", " & slStaName & ", " & slVefName & ", " & sLWeek
            gLogMsg slTemp, smLogFileName, False
            
            SQLQuery = "UPDATE cptt SET "
            SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete
            SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
            SQLQuery = SQLQuery + " WHERE cpttCode = " & rst_Cptt!cpttCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError smLogFileName, "CPTTCheck-mCheckPostedWeeks"
                rst_Cptt.Close
                rst_Ast.Close
                Exit Sub
            End If
            llChgCount = llChgCount + 1
            txtWeeksChanged.Text = llChgCount
        Else
            'debug break stop only
            ilSpotAired = ilSpotAired
        End If
        
        rst_Cptt.MoveNext
        
        If imTerminate Then
            rst_Cptt.Close
            rst_Ast.Close
            Exit Sub
        End If
    Wend
    gFileChgdUpdate "cptt.mkd", True
    txtWeekNum.Text = llCount
    gLogMsg "", smLogFileName, False
    gLogMsg "*** Total Number of Weeks Updated: " & llChgCount & " ***", smLogFileName, False
    rst_Cptt.Close
    rst_Ast.Close
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTTCheck-mCheckPostedWeeks"
End Sub

Private Sub mCheckAndFixCPTTs(slMode As String, smLogFileName As String)
    'slMode (I)- C=Check; F=Fix
    Dim slLLD As String
    Dim llLLD As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llMonCheckDate As Long
    Dim llOffAirDate As Long
    Dim llDropDate As Long
    Dim slOnAirDate As String
    Dim llMonOnAirDate As Long
    Dim llDate As Long
    Dim ilCycle As Integer
    Dim slTime As String
    Dim slCurDate As String
    Dim slCurMoDate As String
    Dim llCurMoDate As Long
    Dim llAttCount As Long
    Dim llCpttCount As Long
    Dim llAttTotalCount As Long
    Dim llCpttTotalCount As Long
    Dim ilStatus As Integer
    Dim ilPostingStatus As Integer
    Dim slAstStatus As String
    Dim slStartDate As String
    Dim slMoDate As String
    Dim slSuDate As String
    Dim llLoop As Long
    Dim llSubTotalRecordsDeleted As Long
    Dim llTotalRecordsDeleted As Long
    Dim llDuplicates As Long
    Dim llPostComplete As Long
    Dim llPostPartial As Long
    Dim llPostCleared As Long
    Dim llPrevCpttAtt As Long
    Dim llPrevCpttStartDate As Long
    Dim llCpttStartDate As Long
    Dim llAstLatestDate As Long
    Dim ilCreateCPTT As Integer
    Dim llTotalNoAdd As Long
    Dim llTotalNoDelete As Long
    Dim llTotalNoPrior As Long
    Dim llTotalNoAfter As Long
    Dim llChangedAstCPStatus As Long
    Dim ilChangedAstCPStatusFlag As Integer
    Dim llChangeAstCPStatusForCPTT As Long
    Dim ilFound As Integer
    Dim llIndex As Long
    Dim llPrevAttCode As Long
    Dim ilAccessToWeb As Integer
    Dim ilResetAst As Integer
    Dim ilBypassTimeCheck As Integer
    Dim ilZoneAdj As Integer
    Dim llLogDate As Long
    Dim llLogTime As Long
    Dim ilDayOk As Integer
    Dim ilVef As Integer
    Dim ilLocalAdj As Integer
    Dim slZone As String
    Dim ilZone As Integer
    Dim llFdStTime As Long
    Dim llFdEdTime As Long
    Dim ilrst_AST1 As Integer
    Dim ilGetDAT As Integer
    Dim ilDat As Integer
    Dim llCount As Long
    Dim llMatchAst As Long
    Dim slDataArray() As String
    Dim llExported As Long
    Dim tlDatPledgeInfo1 As DATPLEDGEINFO
    Dim tlDatPledgeInfo2 As DATPLEDGEINFO
    Dim ilRet As Integer
    Dim ilAstDescrepant As Integer
    Dim llMoDate As Long
    Dim slTempMoDate As String
    Dim llLstRet As Long
    Dim ilSpotsAired As Integer
    
    Dim ilPartialPost As Integer
    Dim llReset_0_Cnt As Long
    Dim llReset_1_Cnt As Long
    Dim llTtlAstCnt As Long
    Dim llIdx As Long
    Dim slVefStr As String
    Dim llAufsKey As Long
    
    Dim llAst_0() As Long
    Dim llAst_1() As Long
    
    Dim llPartialSet_0 As Long
    Dim llPartialSet_1 As Long
    Dim llNotFoundOnWeb As Long
    Dim llTemp As Long
    Dim ilResult As Integer
    Dim ilOk As Integer
    
    Dim llIdx1 As Long
    Dim llIdx2 As Long
    Dim llIdx3 As Long
    Dim llIdx4 As Long
    Dim llIdx5 As Long
    Dim llIdx6 As Long
    Dim llIdx7 As Long
    Dim llIdx8 As Long
    Dim llIdx9 As Long
    Dim llIdx10 As Long
    Dim llIdx11 As Long
    Dim llIdx12 As Long
    Dim llIdx13 As Long
    Dim llIdx14 As Long
    Dim llIdx15 As Long
    
    Dim llMinCptt As Long
    Dim llMaxCptt As Long
    Dim llCpttIdx As Long
    Dim llDat As Long
    Dim llAttRet As Long
    Dim llShttRet As Long
    
    Dim rst_att As ADODB.Recordset
    Dim rst_Vpf As ADODB.Recordset
    Dim rst_Cptt As ADODB.Recordset
    Dim rst_Ast1 As ADODB.Recordset
    Dim rst_Ast2 As ADODB.Recordset
    Dim rst_DAT As ADODB.Recordset
    
    'Debug information
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    
    
    On Error GoTo ErrHand
    
    'Debug Counts
    llIdx1 = 0
    llIdx2 = 0
    llIdx3 = 0
    llIdx4 = 0
    llIdx5 = 0
    llIdx6 = 0
    llIdx7 = 0
    llIdx8 = 0
    llIdx9 = 0
    llIdx10 = 0
    llIdx11 = 0
    llIdx12 = 0
    llIdx13 = 0
    llIdx14 = 0

   
    
    '********************** Pop All of the data that we can **********************

    lacProgress.Caption = "Standby...Gathering cptt records."
    ilRet = gPopCpttInfo
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopCpttInfo Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If
    llCpttTotalCount = lgCpttCount

    lacProgress.Caption = "Standby...Gathering Auf records."
    ilRet = gPopAuf
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopAuf Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If

    lacProgress.Caption = "Standby...Gathering spot records."
    ilRet = gPopLstInfo
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopLstInfo Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If

    lacProgress.Caption = "Standby...Gathering agreement records."
    ilRet = gPopAttInfo
    llAttTotalCount = lgAttCount
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopAttInfo Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If

    lacProgress.Caption = "Standby...Gathering station records."
    ilRet = gPopShttInfo
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopShttInfo Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If

    imDone = False


    lbcView.ListItems.Clear
    mSetColumnWidth
    llTotalNoAdd = 0
    llTotalNoDelete = 0
    llTotalNoPrior = 0
    llTotalNoAfter = 0
    llChangedAstCPStatus = 0
    llChangeAstCPStatusForCPTT = 0


    ilCycle = 7  'rst_Att!vpfLNoDaysCycle
    slTime = Format("12:00AM", "hh:mm:ss")
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurMoDate = gAdjYear(gObtainPrevMonday(slCurDate))
    llCurMoDate = DateValue(slCurMoDate)

    llAttCount = 0
    gLogMsg "**** Processing " & CStr(llAttTotalCount) & " Agreements ****", smLogFileName, False

    If imTerminate Then
        imDone = True
        Exit Sub
    End If

    ReDim tmCPTTCheck(0 To llAttTotalCount) As CPTTCHECK
    llIndex = LBound(tmCPTTCheck)

    llCpttCount = 0
    If gUsingWeb Then
        DoEvents

        SQLQuery = "SELECT Count(cpttCode) FROM cptt"
        Set rst_Cptt = gSQLSelectCall(SQLQuery)
        If rst_Cptt(0).Value > 0 Then
            llCpttTotalCount = rst_Cptt(0).Value
        Else
            llCpttTotalCount = 0
        End If
    End If

    llTotalRecordsDeleted = 0
    If (gUsingWeb) And (slMode = "F") Then
        ilAccessToWeb = gHasWebAccess()
        If Not ilAccessToWeb Then
           gLogMsg "**** Unable to Update Web ****", smLogFileName, False
        End If
    Else
        ilAccessToWeb = False
    End If

    'Doug- on 11/17/06 Move code here so that user will see the error sooner
    ReDim tmWebSpotsInfo(0 To 0) As WEBSPOTSINFO
    'Add code here to set the Complete posting flag
'    'If gUsingWeb And (slMode = "F") And ilAccessToWeb Then
    If gUsingWeb And ilAccessToWeb Then
        'Doug-Fix this call please.
        'Move slDataArry into tmWebSpotsInfo
        'Note:  two fields in tmWebSpotsInfo:  astCode and flag where flag = 0 if Not Posted, 1=Posted but not sent back to affilate; 2= Posted and sent back to affiliate
        If imTerminate Then
            imDone = True
            Exit Sub
        End If

        ilRet = mGetWebAstRecs()
        If Not ilRet Then
            gLogMsg "Error: No WebAstRecs.txt file was returned form the web site.  Shutting down. ", smLogFileName, False
            gMsgBox "Error: No WebAstRecs.txt file was returned form the web site.  Shutting down. "
            imDone = True
            Exit Sub
        End If

        If imTerminate Then
            imDone = True
            Exit Sub
        End If
    End If
    DoEvents


    lacProgress.Caption = ""
    llPrevAttCode = -1

    SQLQuery = "Select * from att"
    If smCPTTCheck <> "" Then
        llMonCheckDate = DateValue(gAdjYear(gObtainPrevMonday(smCPTTCheck)))
        SQLQuery = SQLQuery & " WHERE " & "(attOffAir >= '" & Format$(gAdjYear(smCPTTCheck), sgSQLDateForm) & "') And (attDropDate >= '" & Format$(gAdjYear(smCPTTCheck), sgSQLDateForm) & "')"
    Else
        llMonCheckDate = 0
    End If
    Set rst_att = gSQLSelectCall(SQLQuery)

    While Not rst_att.EOF
        If imTerminate Then
            imDone = True
            Exit Sub
        End If
        llAttCount = llAttCount + 1
        lacProgress.Caption = "Agreements being Checked: " & Trim$(Str$(llAttCount)) & " of " & Trim$(Str$(llAttTotalCount)) & ", Post Complete being Checked: "  '& Trim$(Str$(llCpttCount)) & " of " & Trim$(Str$(llCpttTotalCount))
        txtWeekNum.Text = Trim$(Str$(llCpttCount))
        txtTtlWeeks.Text = Trim$(Str$(llCpttTotalCount))
        DoEvents
        'Obtain latest AST date too determine how CPTT should be created
        'If cptt is prior to latest ast date, then only create is ast exist
        'If cptt is after latest ast date, then always create
        If (rst_att!attExportType <> 0) Then
            SQLQuery = "Select max(astFeedDate) FROM ast WHERE "
            SQLQuery = SQLQuery + " astAtfCode = " & rst_att!attCode
            Set rst_Ast1 = gSQLSelectCall(SQLQuery)
            'If rst_Ast.EOF Then
            If IsNull(rst_Ast1(0).Value) Then
                llAstLatestDate = 0
            Else
                llAstLatestDate = DateValue(gAdjYear(Format$(rst_Ast1(0).Value, sgShowDateForm)))
            End If
        Else
            llAstLatestDate = 0
        End If
        'Test if CPTT exist
        SQLQuery = "SELECT vpfLLD, vpfLNoDaysCycle, vpfGenLog"
        SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
        SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & rst_att!attvefCode & ")"
        slLLD = ""
        Set rst_Vpf = gSQLSelectCall(SQLQuery)
        If Not rst_Vpf.EOF Then
            If Not IsNull(rst_Vpf!vpfLLD) Then
                If gIsDate(rst_Vpf!vpfLLD) Then
                    'set sLLD to last log date
                    slLLD = Format$(rst_Vpf!vpfLLD, sgShowDateForm)
                End If
            End If
            If rst_Vpf!vpfGenLog = "N" Then
                slLLD = ""
            End If
        End If
        If slLLD <> "" Then
            ilFound = False
            tmCPTTCheck(llIndex).iShfCode = rst_att!attshfCode
            tmCPTTCheck(llIndex).iVefCode = rst_att!attvefCode
            tmCPTTCheck(llIndex).lAttCode = rst_att!attCode
            tmCPTTCheck(llIndex).iNoAdd = 0
            tmCPTTCheck(llIndex).iNoDelete = 0
            tmCPTTCheck(llIndex).iNoPrior = 0
            tmCPTTCheck(llIndex).iNoAfter = 0
            tmCPTTCheck(llIndex).lAgreementStart = DateValue(gAdjYear(rst_att!attOnAir))
            tmCPTTCheck(llIndex).lAgreementMoStart = DateValue(gAdjYear(gObtainPrevMonday(rst_att!attOnAir)))

            'Set only if required to check by call to mGetAirWeekAdj
            tmCPTTCheck(llIndex).iAirWeekAdj = 0

            llLLD = DateValue(gAdjYear(slLLD))
            llOffAirDate = DateValue(gAdjYear(rst_att!attOffAir))
            llDropDate = DateValue(gAdjYear(rst_att!attDropDate))
            If llDropDate < llOffAirDate Then
                tmCPTTCheck(llIndex).lAgreementEnd = llDropDate
                If llLLD < llDropDate Then
                    llEndDate = llLLD
                Else
                    llEndDate = llDropDate
                End If
            Else
                tmCPTTCheck(llIndex).lAgreementEnd = llOffAirDate
                If llLLD < llOffAirDate Then
                    llEndDate = llLLD
                Else
                    llEndDate = llOffAirDate
                End If
            End If
            slOnAirDate = rst_att!attOnAir
            llMonOnAirDate = DateValue(gAdjYear(gObtainPrevMonday(slOnAirDate)))
            If smCPTTCheck <> "" Then
                If llMonCheckDate > llMonOnAirDate Then
                    llStartDate = llMonCheckDate
                Else
                    llStartDate = llMonOnAirDate
                End If
            Else
                llStartDate = llMonOnAirDate
            End If
            If tmCPTTCheck(llIndex).lAgreementStart <= tmCPTTCheck(llIndex).lAgreementEnd Then
                For llDate = llStartDate To llEndDate Step 7
                    If imTerminate Then
                        Exit Sub
                    End If
                    SQLQuery = "SELECT * FROM cptt WHERE"
                    SQLQuery = SQLQuery & " cpttAtfCode = " & rst_att!attCode
                    SQLQuery = SQLQuery & " And cpttStartDate = " & "'" & Format$(gAdjYear(Format$(llDate, "m/d/yy")), sgSQLDateForm) & "'"
                    Set rst_Cptt = gSQLSelectCall(SQLQuery)
                    If rst_Cptt.EOF Then
                        'Check if should be created
                        'Check if cptt should be created:
                        'if exporting then see if ast exist and if so then add cptt.  ast will not exist after ast latest date.
                        'For manual posting, create cptt as testing ast will yield no information about deleting cptt
                        If (rst_att!attExportType <> 0) And (llDate <= llAstLatestDate) Then
                            slMoDate = gAdjYear(Format$(llDate, "m/d/yy"))
                            slSuDate = gAdjYear(DateAdd("d", 6, slMoDate))
                            SQLQuery = "Select astCode FROM ast WHERE "
                            SQLQuery = SQLQuery + " astAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                            Set rst_Ast1 = gSQLSelectCall(SQLQuery)
                            If rst_Ast1.EOF Then
                                ilCreateCPTT = False
                            Else
                                ilCreateCPTT = True
                            End If
                        Else
                            'If in past, ignore creation of CPTT
                            'This is done to avoid added in one pass, then deleting weeks in other pass
                            If llDate < llCurMoDate Then
                                ilCreateCPTT = False
                            Else
                                ilCreateCPTT = True
                            End If
                        End If
                        If ilCreateCPTT Then
                            ilFound = True
                            tmCPTTCheck(llIndex).iNoAdd = tmCPTTCheck(llIndex).iNoAdd + 1
                            llTotalNoAdd = llTotalNoAdd + 1
                            If slMode = "F" Then
                                ilStatus = 0
                                ilPostingStatus = 0
                                slAstStatus = "N"
                                If rst_att!attExportType <> 0 Then
                                    slMoDate = gAdjYear(Format$(llDate, "m/d/yy"))
                                    slSuDate = gAdjYear(DateAdd("d", 6, slMoDate))
                                    SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
                                    SQLQuery = SQLQuery + " AND astAtfCode = " & rst_att!attCode
                                    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                                    Set rst_Ast1 = gSQLSelectCall(SQLQuery)
                                    If rst_Ast1.EOF Then
                                        ilStatus = 1
                                        ilPostingStatus = 2
                                        slAstStatus = "C"
                                    End If
                                End If
                                '6/8/19
                                If rst_att!attServiceAgreement = "Y" Then
                                    ilStatus = 1
                                End If
                                SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                                SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, "
                                SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode, cpttPostingStatus, cpttAstStatus)"
                                SQLQuery = SQLQuery & " VALUES "
                                SQLQuery = SQLQuery & "(" & rst_att!attCode & ", " & rst_att!attshfCode & ", " & rst_att!attvefCode & ", "
                                SQLQuery = SQLQuery & "'" & Format$(gAdjYear(slCurDate), sgSQLDateForm) & "', '" & Format(gAdjYear(Format$(llDate, "m/d/yy")), sgSQLDateForm) & "', "
                                SQLQuery = SQLQuery & "" & ilStatus & ", " & igUstCode & ", " & ilPostingStatus & ", '" & slAstStatus & "'" & ")"
                                'cnn.Execute SQLQuery, rdExecDirect
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                    Exit Sub
                                End If
                                gFileChgdUpdate "cptt.mkd", True
                                gLogMsg "AttCode = " & rst_att!attCode & ", Vehicle: " & gGetVehNameByVefCode(rst_att!attvefCode) & ", Station: " & gGetCallLettersByShttCode(rst_att!attshfCode) & " Date Created: " & gAdjYear(Format(llDate, "m/d/yy")), smLogFileName, False
                            End If
                        End If
                    Else
                        'Check if cptt should be removed:
                        'if exporting then see if ast exist and if not remove cptt
                        'For manual posting, retain cptt as testing ast will yield no information about deleting cptt

                        If (rst_att!attExportType <> 0) And (llDate <= llAstLatestDate) Then
                            slMoDate = gAdjYear(Format$(llDate, "m/d/yy"))
                            slSuDate = gAdjYear(DateAdd("d", 6, slMoDate))
                            SQLQuery = "Select astCode FROM ast WHERE "
                            SQLQuery = SQLQuery + " astAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                            Set rst_Ast1 = gSQLSelectCall(SQLQuery)
                            If rst_Ast1.EOF Then
                                llPrevAttCode = rst_att!attCode
                                While Not rst_Cptt.EOF
                                    ilFound = True
                                    tmCPTTCheck(llIndex).iNoDelete = tmCPTTCheck(llIndex).iNoDelete + 1
                                    If slMode = "F" Then
                                        'To cover the case where an agreement is changed from a manual type
                                        'to a web type.  We don't want to delete the manual cptt recs if they
                                        'had any posting activity. i.e. cpttStatus <> 0 or cpttPostingStatus <> 0
                                        If rst_Cptt!cpttStatus = 0 And rst_Cptt!cpttPostingStatus = 0 Then
                                            SQLQuery = "DELETE FROM cptt WHERE cpttCode = " & rst_Cptt!cpttCode
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                                Exit Sub
                                            End If
                                            llTotalNoDelete = llTotalNoDelete + 1
                                        End If
                                    End If
                                    rst_Cptt.MoveNext
                                Wend
                            End If
                        End If
                    End If
                Next llDate
                SQLQuery = "SELECT Count(cpttCode) FROM cptt WHERE"
                SQLQuery = SQLQuery & " cpttAtfCode = " & rst_att!attCode
                SQLQuery = SQLQuery & " And cpttStartDate < " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementMoStart, "m/d/yy")), sgSQLDateForm) & "'"
                Set rst_Cptt = gSQLSelectCall(SQLQuery)
                If Not rst_Cptt.EOF Then
                    'ExportType: 0=Manual; 1=Export
                    If (rst_Cptt(0).Value > 0) And ((rst_att!attExportType = 0) Or (rst_att!attExportType = 1)) Then
                        ilFound = True
                        tmCPTTCheck(llIndex).iNoPrior = rst_Cptt(0).Value
                        llTotalNoPrior = llTotalNoPrior + tmCPTTCheck(llIndex).iNoPrior

                        If slMode = "F" Then
                            'Delete from CPTT
                            SQLQuery = "DELETE FROM cptt"
                            SQLQuery = SQLQuery & " WHERE "
                            SQLQuery = SQLQuery + " cpttAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery & " And cpttStartDate < " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementMoStart, "m/d/yy")), sgSQLDateForm) & "'"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            'Delete from AST
                            SQLQuery = "DELETE FROM ast"
                            SQLQuery = SQLQuery & " WHERE "
                            SQLQuery = SQLQuery & " astAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery & " And astAirDate < " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementMoStart, "m/d/yy")), sgSQLDateForm) & "'"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            'Send command to Web to remove spots
                            '7701
                            If (rst_att!attExportType = 1) And gUsingWeb And ilAccessToWeb Then
                           ' If (rst_att!attExportType = 1) And ((rst_att!attExportToWeb = "Y") Or (rst_att!attWebInterface = "C")) And gUsingWeb And ilAccessToWeb Then
                                gLogMsg "Delete Web Spots: AttCode = " & CLng(rst_att!attCode) & ", Vehicle = " & gGetVehNameByVefCode(rst_att!attvefCode) & ", Station = " & gGetCallLettersByShttCode(rst_att!attshfCode) & " And PledgeStartDate < " & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementMoStart, "m/d/yy")), sgSQLDateForm), smLogFileName, False
                                SQLQuery = "Delete From Spots Where attCode = " & CLng(rst_att!attCode) & " And PledgeStartDate < '" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementMoStart, "m/d/yy")), sgSQLDateForm) & "'"
                                llSubTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
                                If llSubTotalRecordsDeleted = -1 Then
                                    gLogMsg "Error trying to delete AST from the web.", smLogFileName, False
                                Else
                                    llTotalRecordsDeleted = llTotalRecordsDeleted + llSubTotalRecordsDeleted
                                    gLogMsg "    " & CStr(llSubTotalRecordsDeleted) & " Web Ast Records were found to Delete.", smLogFileName, False
                                    SQLQuery = "Delete From SpotRevisions Where attCode = " & CLng(rst_att!attCode) & " And PledgeStartDate < '" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementMoStart, "m/d/yy")), sgSQLDateForm) & "'"
                                    llSubTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
                                    gLogMsg "    " & CStr(llSubTotalRecordsDeleted) & " Web Ast Spot Revision Records were found to Delete.", smLogFileName, False
                                End If
                            End If
                        End If
                    End If
                End If
                If imTerminate Then
                    imDone = True
                    Exit Sub
                End If
                SQLQuery = "SELECT Count(cpttCode) FROM cptt WHERE"
                SQLQuery = SQLQuery & " cpttAtfCode = " & rst_att!attCode
                SQLQuery = SQLQuery & " And cpttStartDate > " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm) & "'"
                Set rst_Cptt = gSQLSelectCall(SQLQuery)
                If Not rst_Cptt.EOF Then
                    If (rst_Cptt(0).Value > 0) And ((rst_att!attExportType = 0) Or (rst_att!attExportType = 1)) Then
                        ilFound = True
                        tmCPTTCheck(llIndex).iNoAfter = rst_Cptt(0).Value
                        llTotalNoAfter = llTotalNoAfter + tmCPTTCheck(llIndex).iNoAfter

                        If slMode = "F" Then
                            'Delete from CPTT where weeks are past the end date of the agreement
                            SQLQuery = "DELETE FROM cptt"
                            SQLQuery = SQLQuery & " WHERE "
                            SQLQuery = SQLQuery & " cpttAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery & " And cpttStartDate > " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm) & "'"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            'Delete from AST where spots are past the end date of the agreement
                            SQLQuery = "DELETE FROM ast"
                            SQLQuery = SQLQuery & " WHERE "
                            SQLQuery = SQLQuery & " astAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery & " And astAirDate > " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm) & "'"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            'Send command to Web to remove spots
                            'D.S. 11/11/13
                            'If (rst_att!attExportType = 1) And ((rst_att!attExportToWeb = "Y") Or (rst!att!attWebInterface = "C")) And gUsingWeb And ilAccessToWeb Then
                            '7701
                            If (rst_att!attExportType = 1) And gUsingWeb And ilAccessToWeb Then
                           ' If (rst_att!attExportType = 1) And ((rst_att!attExportToWeb = "Y") Or (rst_att!attWebInterface = "C")) And gUsingWeb And ilAccessToWeb Then
                                gLogMsg "Delete Web Spots: AttCode = " & CLng(rst_att!attCode) & ", Vehicle = " & gGetVehNameByVefCode(rst_att!attvefCode) & ", Station = " & gGetCallLettersByShttCode(rst_att!attshfCode) & " And PledgeStartDate > " & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm), smLogFileName, False
                                SQLQuery = "Delete From Spots Where attCode = " & CLng(rst_att!attCode) & " And PledgeStartDate > '" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm) & "'"
                                llSubTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
                                If llSubTotalRecordsDeleted = -1 Then
                                    gLogMsg "Error trying to delete AST from the web.", smLogFileName, False
                                Else
                                    llTotalRecordsDeleted = llTotalRecordsDeleted + llSubTotalRecordsDeleted
                                    gLogMsg "    " & CStr(llSubTotalRecordsDeleted) & " Web Spot With Dates Beyond " & Format$(tmCPTTCheck(llIndex).lAgreementEnd, "mm/dd/yy") & ", the Agreements End date, were Deleted.", smLogFileName, False
                                    SQLQuery = "Delete From SpotRevisions Where attCode = " & CLng(rst_att!attCode) & " And PledgeStartDate > '" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm) & "'"
                                    llSubTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
                                    gLogMsg "    " & CStr(llSubTotalRecordsDeleted) & " Web Ast Spot Revision Records were found to Delete.", smLogFileName, False
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'Remove all
                If imTerminate Then
                    imDone = True
                    Exit Sub
                End If
                SQLQuery = "SELECT Count(cpttCode) FROM cptt WHERE"
                SQLQuery = SQLQuery & " cpttAtfCode = " & rst_att!attCode
                Set rst_Cptt = gSQLSelectCall(SQLQuery)
                If Not rst_Cptt.EOF Then
                    If (rst_Cptt(0).Value > 0) And ((rst_att!attExportType = 0) Or (rst_att!attExportType = 1)) Then
                        ilFound = True
                        tmCPTTCheck(llIndex).iNoAfter = rst_Cptt(0).Value
                        llTotalNoAfter = llTotalNoAfter + tmCPTTCheck(llIndex).iNoAfter

                        If slMode = "F" Then
                            'Delete from CPTT
                            SQLQuery = "DELETE FROM cptt"
                            SQLQuery = SQLQuery & " WHERE "
                            SQLQuery = SQLQuery & " cpttAtfCode = " & rst_att!attCode
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            'Delete from AST
                            SQLQuery = "DELETE FROM ast"
                            SQLQuery = SQLQuery & " WHERE "
                            SQLQuery = SQLQuery & " astAtfCode = " & rst_att!attCode
                            SQLQuery = SQLQuery & " And astAirDate > " & "'" & Format$(gAdjYear(Format$(tmCPTTCheck(llIndex).lAgreementEnd, "m/d/yy")), sgSQLDateForm) & "'"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            'Send command to Web to remove spots
                            'D.S. 11/11/13
                            'If (rst_att!attExportType = 1) And ((rst_att!attExportToWeb = "Y") Or (rst!att!attWebInterface = "C")) And gUsingWeb And ilAccessToWeb Then
                            '7701
                            If (rst_att!attExportType = 1) And gUsingWeb And ilAccessToWeb Then
                           ' If (rst_att!attExportType = 1) And ((rst_att!attExportToWeb = "Y") Or (rst_att!attWebInterface = "C")) And gUsingWeb And ilAccessToWeb Then
                                gLogMsg "Delete Web Spots: AttCode = " & CLng(rst_att!attCode) & ", Vehicle = " & gGetVehNameByVefCode(rst_att!attvefCode) & ", Station = " & gGetCallLettersByShttCode(rst_att!attshfCode) & " And Agreement End Prior to Start", smLogFileName, False
                                SQLQuery = "Delete From Spots Where attCode = " & CLng(rst_att!attCode)
                                llSubTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
                                If llSubTotalRecordsDeleted = -1 Then
                                    gLogMsg "Error trying to delete AST from the web.", smLogFileName, False
                                Else
                                    llTotalRecordsDeleted = llTotalRecordsDeleted + llSubTotalRecordsDeleted
                                    gLogMsg "    " & CStr(llSubTotalRecordsDeleted) & " Web Ast Records were found to Delete.", smLogFileName, False
                                    SQLQuery = "Delete From SpotRevisions Where attCode = " & CLng(rst_att!attCode)
                                    llSubTotalRecordsDeleted = gExecWebSQLWithRowsEffected(SQLQuery)
                                    gLogMsg "    " & CStr(llSubTotalRecordsDeleted) & " Web Ast Spot Revision Records were found to Delete.", smLogFileName, False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If ilFound Then
            llIndex = llIndex + 1
            If llIndex = UBound(tmCPTTCheck) Then
                ReDim Preserve tmCPTTCheck(0 To llIndex + 100) As CPTTCHECK
            End If
        End If
        rst_att.MoveNext
    Wend
    If imTerminate Then
        imDone = True
        Exit Sub
    End If

    If imTerminate Then
        imDone = True
        Exit Sub
    End If
    ReDim Preserve tmCPTTCheck(0 To llIndex) As CPTTCHECK
   
    If slMode = "F" Then
        gLogMsg "**** A total of " & CStr(llTotalRecordsDeleted) & " Web Ast Records were Deleted. ****", smLogFileName, False
    End If
    'Re-get cptt count as some might have been deleted above
    llCpttCount = 0
    llPostComplete = 0
    llPostPartial = 0
    llPostCleared = 0
    llDuplicates = 0
    llPartialSet_0 = 0
    llPartialSet_1 = 0
    llNotFoundOnWeb = 0
    llPrevCpttAtt = -1
    llPrevCpttStartDate = -1
    lmAstWebNotExported_PostedLocal = 0
    lmAstWebExported_NotPostedLocal = 0
    lmAstNotFoundOnWeb = 0
    
    DoEvents
    If gUsingWeb Then
        gLogMsg "**** Processing " & CStr(llCpttTotalCount) & " for Posting Complete flags and Duplicate CPTTs ****", smLogFileName, False
    Else
        gLogMsg "**** Processing " & CStr(llCpttTotalCount) & " for Duplicate CPTTs ****", smLogFileName, False
    End If
    'SQLQuery = "SELECT * FROM cptt ORDER BY cpttAtfCode, cpttStartDate, cpttPostingStatus  DESC"
    'SQLQuery = "SELECT * FROM cptt"
    
    llMinCptt = 1
    llMaxCptt = 5000
    
    ilRet = gPopCpttInfo
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopCpttInfo Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If
    
    ilRet = gPopAuf
    If Not ilRet Then
        gLogMsg "**** ERROR: gPopAuf Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If
    
    llCpttTotalCount = lgCpttCount
    
    For llCpttIdx = 0 To UBound(tgCpttInfo) - 1 Step 1
        llIdx1 = llIdx1 + 1
        If imTerminate Then
            imDone = True
            Exit Sub
        End If
        DoEvents
        ilChangedAstCPStatusFlag = False
        llCpttCount = llCpttCount + 1
        
        ilResult = llCpttCount Mod 100
        If ilResult = 0 Then
            gLogMsg "Checking CPTT Number: " & Trim$(Str$(llIdx1)) & " of " & Trim$(Str$(lgCpttCount)), smLogFileName, False
            lacProgress.Caption = "Agreements Checked: " & Trim$(Str$(llAttCount)) & " of " & Trim$(Str$(lgCpttCount)) & ", Post Complete being Checked: " '& Trim$(Str$(llCpttCount)) & " of " & Trim$(Str$(llCpttTotalCount))
            txtWeekNum.Text = Trim$(Str$(llCpttCount))
            txtTtlWeeks.Text = Trim$(Str$(llCpttTotalCount))
        End If


        'DoEvents
         llAttRet = gBinarySearchAtt(tgCpttInfo(llCpttIdx).cpttatfCode)
         If llAttRet <> -1 Then
            'cpttStatus = 2 indicates Not Aired
            If (tgAttInfo1(llAttRet).attExportType <> 0) And (tgCpttInfo(llCpttIdx).cpttStatus <> 2) Then
                'DoEvents
                slStartDate = Format$(Trim$(tgCpttInfo(llCpttIdx).CpttStartDate), sgShowDateForm)
                slMoDate = gAdjYear(gObtainPrevMonday(slStartDate))
                slSuDate = gAdjYear(DateAdd("d", 6, slMoDate))
                ilGetDAT = True
                ilZoneAdj = False
                ilBypassTimeCheck = False
                slZone = ""
                ReDim tlDat(0 To 30) As DATRST
                
                On Error GoTo BadAstRec
                
                ilOk = True
                'SQLQuery = "Select * FROM ast WHERE "
                '12/13/13: Pledge obtained from DAT
                'SQLQuery = "Select astCode, astStatus, astCPStatus, astPledgeStatus, astPledgeStartTime, astAirTime, astlsfCode, astShfCode FROM ast WHERE "
                SQLQuery = "Select * FROM ast WHERE "
                SQLQuery = SQLQuery + " astAtfCode = " & tgCpttInfo(llCpttIdx).cpttatfCode
                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                Set rst_Ast1 = gSQLSelectCall(SQLQuery)
                If Not rst_Ast1.EOF And ilOk Then
                    'DoEvents
                    'Check if astCPstatus set wrong
                    If (tgAttInfo1(llAttRet).attExportType = 1) And gUsingWeb And ilAccessToWeb Then
                        'llExported = gCheckIfSpotsHaveBeenExported(tgAttInfo1(llAttRet).attvefCode, slMoDate, tgAttInfo1(llAttRet).attExportType)
                        'pad out the vefcode with spaces
                        slVefStr = Trim$(Str$(tgAttInfo1(llAttRet).attvefCode))
                        Do While Len(slVefStr) < 5
                             slVefStr = "0" & slVefStr
                        Loop
                        'pad out the date with spaces
                        llMoDate = DateValue(slMoDate)
                        slTempMoDate = CStr(llMoDate)
                        Do While Len(slTempMoDate) < 6
                          slTempMoDate = "0" & slTempMoDate
                        Loop
                
                        'concatenate the vefcode and date to form the key value
                        slVefStr = slVefStr & slTempMoDate
                        llExported = gBinarySearchAuf(slVefStr)
                        If llExported = -1 Then
                            llNotFoundOnWeb = llNotFoundOnWeb + 1
                        End If
                        If llExported <> -1 Then
                            ilBypassTimeCheck = True
                            llReset_0_Cnt = 0
                            llReset_1_Cnt = 0
                            llTtlAstCnt = 0
                            ilAstDescrepant = False
                            ReDim llAst_0(0 To 5000) As Long
                            ReDim llAst_1(0 To 5000) As Long
                            
                            'Gather up all of the spots and put them in the appropriate array.
                            'We may or may not need the array later.
                            Do While Not rst_Ast1.EOF
                                llIdx2 = llIdx2 + 1
                                If imTerminate Then
                                    imDone = True
                                    Exit Sub
                                End If
                                
                                'Don't blow out the arrays
                                If llReset_1_Cnt >= UBound(llAst_1) Then
                                   ReDim Preserve llAst_1(0 To UBound(llAst_1) + 5000) As Long
                                End If
                                
                                If llReset_0_Cnt >= UBound(llAst_0) Then
                                   ReDim Preserve llAst_0(0 To UBound(llAst_0) + 5000) As Long
                                End If
                                
                                '12/13/13: Obtain Pledge information from Dat
                                tlDatPledgeInfo1.lAttCode = rst_Ast1!astAtfCode
                                tlDatPledgeInfo1.lDatCode = rst_Ast1!astDatCode
                                tlDatPledgeInfo1.iVefCode = rst_Ast1!astVefCode
                                tlDatPledgeInfo1.sFeedDate = Format(rst_Ast1!astFeedDate, "m/d/yy")
                                tlDatPledgeInfo1.sFeedTime = Format(rst_Ast1!astFeedTime, "hh:mm:ssam/pm")
                                ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo1)

                                
                                llMatchAst = mBinarySearchWebAstInfo(rst_Ast1!astCode)
                                'Can we find a match in the web ast info table?
                                If llMatchAst <> -1 Then
                                    If tmWebSpotsInfo(llMatchAst).iFlag = 2 Then
                                        'Web has Posted and Exported the spot
                                        If rst_Ast1!astCPStatus = 0 Then  'not recv.
                                            ilAstDescrepant = True
                                            lmAstWebExported_NotPostedLocal = lmAstWebExported_NotPostedLocal + 1
                                        End If
                                        llAst_1(llReset_1_Cnt) = rst_Ast1!astCode
                                        llTtlAstCnt = llTtlAstCnt + 1
                                        llReset_1_Cnt = llReset_1_Cnt + 1
                                    Else
                                        'Web may have or may not Posted, but it has NOT Exported the spot
                                        If rst_Ast1!astCPStatus = 1 Then  'recv.
                                            ilAstDescrepant = True
                                            lmAstWebNotExported_PostedLocal = lmAstWebNotExported_PostedLocal + 1
                                        End If
                                        llAst_0(llReset_0_Cnt) = rst_Ast1!astCode
                                        llTtlAstCnt = llTtlAstCnt + 1
                                        llReset_0_Cnt = llReset_0_Cnt + 1
                                    End If
                                Else
                                    'D.S. 7/28/08  Avoid Not Carried spots from creating a partail week
                                    'This is where the spot was not found on the web because it was not
                                    'supposed to carried in the first place.  I added the outside if statement
                                    '12/13/13
                                    If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo1.iPledgeStatus)).iStatus <> 2 Then
                                        If rst_Ast1!astCPStatus = 0 Then
                                            ilAstDescrepant = True
                                            lmAstNotFoundOnWeb = lmAstNotFoundOnWeb + 1
                                        End If
                                        llAst_1(llReset_1_Cnt) = rst_Ast1!astCode
                                        llTtlAstCnt = llTtlAstCnt + 1
                                        llReset_1_Cnt = llReset_1_Cnt + 1
                                    End If
                                End If
                                rst_Ast1.MoveNext
                            Loop
                            
                            If ilAstDescrepant Then
                                'UnPost or Post ALL of the spots
                                If (llReset_0_Cnt = llTtlAstCnt) Or (llReset_0_Cnt >= llReset_1_Cnt) Then
                                    ilResetAst = 0
                                Else
                                    ilResetAst = 1
                                End If
                                
                                'Update all of the spots, partial or not
                                If (slMode = "F") Then
                                    llIdx3 = llIdx3 + 1
                                    llChangedAstCPStatus = llChangedAstCPStatus + llTtlAstCnt
                                    SQLQuery = "UPDATE ast SET "
                                    SQLQuery = SQLQuery + "astCPStatus = " & ilResetAst
                                    SQLQuery = SQLQuery + " WHERE"
                                    SQLQuery = SQLQuery + " astAtfCode = " & tgCpttInfo(llCpttIdx).cpttatfCode
                                    SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/10/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                        Exit Sub
                                    End If
                                End If
                                
                                'Do we have a partail post?
                                ilPartialPost = False
                                If (llReset_0_Cnt <> llTtlAstCnt) And (ilResetAst = 0) Then
                                    ilPartialPost = True
                                    'We have Partial, Post all spots to whatever has the most
                                    'then individually post the rest
                                    ilResetAst = 1
                                Else
                                    If (llReset_1_Cnt <> llTtlAstCnt) And (ilResetAst = 1) Then
                                        ilPartialPost = True
                                        'We have Partial, Post all spots to whatever has the most
                                        'then individually post the rest
                                        ilResetAst = 0
                                    End If
                                End If
                                
                                If ilPartialPost And ilResetAst = 0 Then
                                    If (slMode = "F") Then
                                        llPartialSet_0 = llPartialSet_0 + (llReset_0_Cnt - 1)
                                        For llIdx = 0 To llReset_0_Cnt - 1 Step 1
                                            If imTerminate Then
                                                imDone = True
                                                Exit Sub
                                            End If
                                            llIdx4 = llIdx4 + 1
                                            SQLQuery = "UPDATE ast SET "
                                            SQLQuery = SQLQuery + "astCPStatus = " & ilResetAst
                                            SQLQuery = SQLQuery + " WHERE astCode = " & llAst_0(llIdx)
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                                Exit Sub
                                            End If
                                        Next llIdx
                                    End If
                                End If
                                
                                If ilPartialPost And ilResetAst = 1 Then
                                    If (slMode = "F") Then
                                        llPartialSet_1 = llPartialSet_1 + (llReset_1_Cnt - 1)
                                        For llIdx = 0 To llReset_1_Cnt - 1 Step 1
                                            If imTerminate Then
                                                imDone = True
                                                Exit Sub
                                            End If
                                            
                                            
                                            'If
                                            
                                            llIdx5 = llIdx5 + 1
                                            SQLQuery = "UPDATE ast SET "
                                            SQLQuery = SQLQuery + "astCPStatus = " & ilResetAst
                                            SQLQuery = SQLQuery + " WHERE astCode = " & llAst_1(llIdx)
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                                Exit Sub
                                            End If
                                        Next llIdx
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Do
                        llIdx6 = llIdx6 + 1
                        If (Not ilBypassTimeCheck) Then
                            Do While Not rst_Ast1.EOF
                                llIdx7 = llIdx7 + 1
                                If imTerminate Then
                                    imDone = True
                                    Exit Sub
                                End If
                                ilResetAst = -1

                                '12/13/13: Obtain Pledge information from Dat
                                tlDatPledgeInfo1.lAttCode = rst_Ast1!astAtfCode
                                tlDatPledgeInfo1.lDatCode = rst_Ast1!astDatCode
                                tlDatPledgeInfo1.iVefCode = rst_Ast1!astVefCode
                                tlDatPledgeInfo1.sFeedDate = Format(rst_Ast1!astFeedDate, "m/d/yy")
                                tlDatPledgeInfo1.sFeedTime = Format(rst_Ast1!astFeedTime, "hh:mm:ssam/pm")
                                ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo1)

                                'If rst_Ast1!astCPStatus = 0 Then
                                    'astStatus is posting status info, astPledgeStatus is agreement status
                                    If (tgStatusTypes(gGetAirStatus(rst_Ast1!astStatus)).iPledged <> 2) Then
                                        '12/13/13
                                        'If (tgStatusTypes(gGetAirStatus(rst_Ast1!astPledgeStatus)).iPledged = 2) Then
                                        If (tgStatusTypes(gGetAirStatus(tlDatPledgeInfo1.iPledgeStatus)).iPledged = 2) Then
                                            If rst_Ast1!astCPStatus = 0 Then
                                                ilResetAst = 1
                                            End If
                                        ElseIf (tgStatusTypes(gGetAirStatus(tlDatPledgeInfo1.iPledgeStatus)).iPledged <> 1) Then '0(Live) or 3(no pledge times)
                                            'Pledge time
                                            If gTimeToLong(tlDatPledgeInfo1.sPledgeStartTime, False) <> gTimeToLong(rst_Ast1!astAirTime, False) Then
                                                If rst_Ast1!astCPStatus = 0 Then
                                                    ilResetAst = 1
                                                End If
                                            Else
                                                If rst_Ast1!astCPStatus = 1 Then
                                                    ilResetAst = 0
                                                End If
                                            End If
                                        Else
                                            'SQLQuery = "Select * FROM lst WHERE lstCode = " & rst_Ast1!astlsfCode
                                            'SQLQuery = "Select lstLogDate, lstLogTime, lstLogVefCode FROM lst WHERE lstCode = " & rst_Ast1!astlsfCode
                                            'Set rst_Lst = gSQLSelectCall(SQLQuery)
                                            llLstRet = gBinarySearchLst(rst_Ast1!astLsfCode)
                                            'If Not rst_Lst.EOF Then
                                            If llLstRet <> -1 Then
                                                If ilGetDAT Then
                                                    ilGetDAT = False
                                                    'Bypass hour test to eliminate time zone correction of lstLogTime
                                                    'SQLQuery = "SELECT shttTimeZone FROM shtt WHERE"
                                                    'SQLQuery = SQLQuery & " shttCode = " & rst_Ast1!astShfCode
                                                    'Set rst_Shtt = gSQLSelectCall(SQLQuery)
                                                    llShttRet = gBinarySearchShtt(rst_Ast1!astShfCode)
                                                    'If Not rst_Shtt.EOF Then
                                                    If llShttRet <> -1 Then
                                                        slZone = Trim$(tgShttInfo1(llShttRet).shttTimeZone)
                                                    Else
                                                        slZone = ""
                                                    End If
                                                    llDat = 0
                                                    SQLQuery = "SELECT * FROM dat WHERE"
                                                    SQLQuery = SQLQuery & " datAtfCode = " & tgAttInfo1(llAttRet).attCode
                                                    Set rst_DAT = gSQLSelectCall(SQLQuery)
                                                    Do While Not rst_DAT.EOF
                                                        If imTerminate Then
                                                            imDone = True
                                                            Exit Sub
                                                        End If
                                                        'gCreateUDTForDat rst_Dat, tlDat(UBound(tlDat))
                                                        gCreateUDTForDat rst_DAT, tlDat(llDat)
                                                        llDat = llDat + 1
                                                        If llDat = UBound(tlDat) Then
                                                            ReDim Preserve tlDat(0 To UBound(tlDat) + 30) As DATRST
                                                        End If
                                                        llIdx8 = llIdx8 + 1
                                                        rst_DAT.MoveNext
                                                    Loop
                                                    'ReDim Preserve tlDat(0 To UBound(tlDat) + 1) As DATRST
                                                    ReDim Preserve tlDat(0 To llDat) As DATRST
                                                End If
                                                llLogDate = DateValue(gAdjYear(tgLstInfo(llLstRet).lstLogDate))
                                                llLogTime = gTimeToLong(tgLstInfo(llLstRet).lstLogTime, False)

                                                If (slZone <> "") Then
                                                    If UBound(tlDat) <= LBound(tlDat) Then
                                                        ilZoneAdj = True
                                                    'ElseIf tlDat(LBound(tlDat)).iDACode <> 2 Then
                                                    ElseIf tgAttInfo1(llAttRet).attPledgeType <> "C" Then
                                                        ilZoneAdj = True
                                                    End If
                                                    If ilZoneAdj Then
                                                        ilVef = gBinarySearchVef(CLng(tgLstInfo(llLstRet).lstLogVefCode))
                                                        If ilVef <> -1 Then
                                                            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                                                                llIdx9 = llIdx9 + 1
                                                                If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = slZone Then
                                                                    If tgVehicleInfo(ilVef).sFed(ilZone) <> "*" Then
                                                                        ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
                                                                        llLogTime = llLogTime + 3600 * ilLocalAdj
                                                                        If llLogTime < 0 Then
                                                                            llLogTime = llLogTime + 86400
                                                                            llLogDate = llLogDate - 1
                                                                        ElseIf llLogTime > 86400 Then
                                                                            llLogTime = llLogTime - 86400
                                                                            llLogDate = llLogDate + 1
                                                                        End If
                                                                    End If
                                                                    Exit For
                                                                End If
                                                            Next ilZone
                                                        End If
                                                    End If
                                                End If
                                                'Find correct feed time
                                                For ilDat = LBound(tlDat) To UBound(tlDat) - 1 Step 1
                                                    llIdx10 = llIdx10 + 1
                                                    llFdStTime = gTimeToLong(tlDat(ilDat).sFdStTime, False)
                                                    llFdEdTime = gTimeToLong(tlDat(ilDat).sFdEdTime, True)
                                                    If llFdEdTime = llFdStTime Then
                                                        llFdEdTime = llFdEdTime + 1
                                                    End If
                                                    ilDayOk = False
                                                    Select Case Weekday(gAdjYear(Format$(llLogDate, "m/d/yy")))
                                                        Case vbMonday
                                                            If tlDat(ilDat).iFdMon = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                        Case vbTuesday
                                                            If tlDat(ilDat).iFdTue = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                        Case vbWednesday
                                                            If tlDat(ilDat).iFdWed = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                        Case vbThursday
                                                            If tlDat(ilDat).iFdThu = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                        Case vbFriday
                                                            If tlDat(ilDat).iFdFri = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                        Case vbSaturday
                                                            If tlDat(ilDat).iFdSat = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                        Case vbSunday
                                                            If tlDat(ilDat).iFdSun = 1 Then
                                                                ilDayOk = True
                                                            End If
                                                    End Select

                                                    If (ilDayOk) And (llLogTime >= llFdStTime) And (llLogTime < llFdEdTime) Then
                                                        llLogTime = llLogTime + gTimeToLong(tlDatPledgeInfo1.sPledgeStartTime, False) - llFdStTime
                                                        If llLogTime <> gTimeToLong(rst_Ast1!astAirTime, False) Then
                                                            If rst_Ast1!astCPStatus = 0 Then
                                                                ilResetAst = 1
                                                            End If
                                                        Else
                                                            If rst_Ast1!astCPStatus = 1 Then
                                                                ilResetAst = 0
                                                            End If
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilDat
                                            End If
                                        End If
                                    Else
                                        'If tgStatusTypes(gGetAirStatus(rst_Ast1!astPledgeStatus)).iPledged <> 2 Then
                                        If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo1.iPledgeStatus)).iPledged <> 2 Then
                                            ilResetAst = 1
                                        End If
                                    End If
                                'End If
                                'Don't unpost as posted as pledged could result in the two times matching
                                If ilResetAst = 1 Then
                                    If Not ilChangedAstCPStatusFlag Then
                                        ilChangedAstCPStatusFlag = True
                                        llChangeAstCPStatusForCPTT = llChangeAstCPStatusForCPTT + 1
                                    End If
                                    llChangedAstCPStatus = llChangedAstCPStatus + 1
                                    If slMode = "F" Then
                                        SQLQuery = "UPDATE ast SET "
                                        SQLQuery = SQLQuery + "astCPStatus = " & ilResetAst
                                        SQLQuery = SQLQuery + " WHERE astCode = " & rst_Ast1!astCode
                                        'cnn.Execute SQLQuery, rdExecDirect
                                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                            '6/10/16: Replaced GoSub
                                            'GoSub ErrHand:
                                            Screen.MousePointer = vbDefault
                                            gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                            Exit Sub
                                        End If
                                        llIdx11 = llIdx11 + 1
                                    End If
                                End If
                                rst_Ast1.MoveNext
                            Loop
                        End If
                        If imTerminate Then
                            imDone = True
                            Exit Sub
                        End If
                        
                        'Test to see if any spots aired or were they all not aired
                        ilSpotsAired = gDidAnySpotsAir(tgCpttInfo(llCpttIdx).cpttatfCode, slMoDate, slSuDate)
                        If ilSpotsAired Then
                            'We know at least one spot aired
                            ilSpotsAired = True
                        Else
                            'no aired spots were found
                            ilSpotsAired = False
                        End If
                        
                        SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
                        SQLQuery = SQLQuery + " AND astAtfCode = " & tgCpttInfo(llCpttIdx).cpttatfCode
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                        Set rst_Ast2 = gSQLSelectCall(SQLQuery)
                        If rst_Ast2.EOF Then
                            'Set CPTT as complete (tmCptt(llCpttIdx).cpttStatus = 2 = Not Aired)
                            If (tgCpttInfo(llCpttIdx).cpttStatus <> 1) Or (tgCpttInfo(llCpttIdx).cpttPostingStatus <> 2) Then
                                llPostComplete = llPostComplete + 1
                                If slMode = "F" Then
                                    SQLQuery = "UPDATE cptt SET "
                                    If ilSpotsAired Then
                                        SQLQuery = SQLQuery + "cpttStatus = 1" & ", " 'Complete spots aired
                                    Else
                                        SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete N0 spots aired
                                    End If
                                    SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
                                    SQLQuery = SQLQuery + " WHERE cpttCode = " & tgCpttInfo(llCpttIdx).cpttCode
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/10/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                        Exit Sub
                                    End If
                                    llIdx12 = llIdx12 + 1
                                End If
                            End If
                            Exit Do
                        Else
                            '12/13/13
                            'SQLQuery = "Select astCode, astPledgeStatus, astStatus FROM ast WHERE astCPStatus = 1"
                            SQLQuery = "Select * FROM ast WHERE astCPStatus = 1"
                            SQLQuery = SQLQuery + " AND astAtfCode = " & tgCpttInfo(llCpttIdx).cpttatfCode
                            SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                            Set rst_Ast2 = gSQLSelectCall(SQLQuery)
                            If Not rst_Ast2.EOF Then
                                llIdx13 = llIdx13 + 1
                                If (tgCpttInfo(llCpttIdx).cpttStatus <> 0) Or (tgCpttInfo(llCpttIdx).cpttPostingStatus <> 1) Then
                                    'D.S. 7/28/08  Avoid Not Carried spots from creating a partail week
                                    'This is where the spot was not found on the web because it was not
                                    'supposed to carried in the first place.  I added the outside if statement
                                    
                                    '12/13/13: Obtain Pledge information from Dat
                                    tlDatPledgeInfo2.lAttCode = rst_Ast2!astAtfCode
                                    tlDatPledgeInfo2.lDatCode = rst_Ast2!astDatCode
                                    tlDatPledgeInfo2.iVefCode = rst_Ast2!astVefCode
                                    tlDatPledgeInfo2.sFeedDate = Format(rst_Ast2!astFeedDate, "m/d/yy")
                                    tlDatPledgeInfo2.sFeedTime = Format(rst_Ast2!astFeedTime, "hh:mm:ssam/pm")
                                    ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo2)
                                    If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo2.iPledgeStatus)).iStatus <> 8 Then
                                        llPostPartial = llPostPartial + 1
                                        If slMode = "F" Then
                                            SQLQuery = "UPDATE cptt SET "
                                            SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
                                            SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
                                            SQLQuery = SQLQuery + " WHERE cpttCode = " & tgCpttInfo(llCpttIdx).cpttCode
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/10/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If (tgCpttInfo(llCpttIdx).cpttStatus <> 0) Or (tgCpttInfo(llCpttIdx).cpttPostingStatus <> 0) Then
                                    llPostCleared = llPostCleared + 1
                                    If slMode = "F" Then
                                        SQLQuery = "UPDATE cptt SET "
                                        SQLQuery = SQLQuery + "cpttStatus = 0" & ", "
                                        SQLQuery = SQLQuery + "cpttPostingStatus = 0"
                                        SQLQuery = SQLQuery + " WHERE cpttCode = " & tgCpttInfo(llCpttIdx).cpttCode
                                        'cnn.Execute SQLQuery, rdExecDirect
                                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                            '6/10/16: Replaced GoSub
                                            'GoSub ErrHand:
                                            Screen.MousePointer = vbDefault
                                            gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        Exit Do
                    Loop
                Else
                    If imTerminate Then
                        imDone = True
                        Exit Sub
                    End If
                    If (tgCpttInfo(llCpttIdx).cpttStatus <> 0) Or (tgCpttInfo(llCpttIdx).cpttPostingStatus <> 0) Then
                        llPostCleared = llPostCleared + 1
                        If slMode = "F" Then
                            SQLQuery = "UPDATE cptt SET "
                            SQLQuery = SQLQuery + "cpttStatus = 0" & ", "
                            SQLQuery = SQLQuery + "cpttPostingStatus = 0"
                            SQLQuery = SQLQuery + " WHERE cpttCode = " & tgCpttInfo(llCpttIdx).cpttCode
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                                Exit Sub
                            End If
                            llIdx14 = llIdx14 + 1
                        End If
                    End If
                End If
            End If
        End If
        slStartDate = Format$(Trim$(tgCpttInfo(llCpttIdx).CpttStartDate), sgShowDateForm)
        llCpttStartDate = DateValue(gAdjYear(gObtainPrevMonday(slStartDate)))
        If (llPrevCpttAtt = tgCpttInfo(llCpttIdx).cpttatfCode) And (llPrevCpttStartDate = llCpttStartDate) Then
            llDuplicates = llDuplicates + 1
            If slMode = "F" Then
                SQLQuery = "DELETE FROM cptt WHERE cpttCode = " & tgCpttInfo(llCpttIdx).cpttCode
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError smLogFileName, "CPTTCheck-mCheckAndFixCPTTs"
                    Exit Sub
                End If
            End If
        End If
        llPrevCpttAtt = tgCpttInfo(llCpttIdx).cpttatfCode
        llPrevCpttStartDate = llCpttStartDate
        
    Next llCpttIdx
    gFileChgdUpdate "cptt.mkd", True
    lacProgress.Caption = "Agreements Checked: " & Trim$(Str$(llAttCount)) & " of " & Trim$(Str$(lgCpttCount)) & ", Post Complete being Checked: " '& Trim$(Str$(llCpttCount)) & " of " & Trim$(Str$(llCpttTotalCount))
    txtWeekNum.Text = Trim$(Str$(llCpttCount))
    txtTtlWeeks.Text = Trim$(Str$(llCpttTotalCount))
    llTemp = llChangedAstCPStatus
    llTemp = llTemp - llPartialSet_0
    llTemp = llTemp - llPartialSet_1
    gLogMsg "", smLogFileName, False
    
    mPopListView smLogFileName, llTotalNoPrior, llTotalNoAfter, llTotalNoAdd, llTotalNoDelete, llChangedAstCPStatus, llChangeAstCPStatusForCPTT, llPostComplete, llPostPartial, llPostCleared, llDuplicates
    gLogMsg "", smLogFileName, False
    
    'not useful information
    'gLogMsg "**** Total Ast Set = " & llChangedAstCPStatus, smLogFileName, False
    'gLogMsg "**** Total Partial Ast Set to 0 = " & llPartialSet_0, smLogFileName, False
    'gLogMsg "**** Total Partial Ast Set to 1 = " & llPartialSet_1, smLogFileName, False
    'gLogMsg "**** Actual Total Ast Set = Total Ast Set - ((Total Partial set to 1) + (Total Partial set to 0)) = " & llTemp, smLogFileName, False
    gLogMsg "", smLogFileName, False
    
    gLogMsg "", smLogFileName, False
    gLogMsg "**** Discrepancies ", smLogFileName, False
    gLogMsg "        Exported Web Spots, Not Posted Local " & lmAstWebExported_NotPostedLocal, smLogFileName, False
    
    gLogMsg "", smLogFileName, False
    
    gLogMsg "**** Improbable Discrepancies. ", smLogFileName, False
    gLogMsg "        Spots Not Found on the Web.  " & lmAstNotFoundOnWeb, smLogFileName, False
    gLogMsg "        Most likely never exported to the web or the spot was modified and never re-exported. ", smLogFileName, False
    
    gLogMsg "", smLogFileName, False
    
    gLogMsg "**** Improbable Discrepancies. Web Spots Not Exported, but have been Posted Locally " & lmAstWebNotExported_PostedLocal, smLogFileName, False
    gLogMsg "        This could be due to spots being re-exported", smLogFileName, False
    
    
    lacProgress.Caption = "Agreements Checked: " & Trim$(Str$(llAttCount)) & " of " & Trim$(Str$(llAttTotalCount)) & ", Post Complete Checked: " '& Trim$(Str$(llCpttCount)) & " of " & Trim$(Str$(llCpttTotalCount))
    txtWeekNum.Text = Trim$(Str$(llCpttCount))
    txtTtlWeeks.Text = Trim$(Str$(llCpttTotalCount))
    
    frmCPTTCheck.Refresh
    gMsgBox "The Utility Program Completed Successfully."
    gLogMsg "", smLogFileName, False
    gLogMsg "The CPTT Utility Program Completed Successfully.", smLogFileName, False
    imDone = True
    cmdCancel.SetFocus
    DoEvents
    Exit Sub
    
BadAstRec:

     ilLine = Erl
     ilErrNo = Err.Number
     slDesc = Err.Description
    
     gLogMsg " ", smLogFileName, False
     gLogMsg "**** ERROR ***** " & SQLQuery & " ilOK = " & ilOk, smLogFileName, False
     gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, smLogFileName, False
     gLogMsg " ", smLogFileName, False
     ilOk = False
     Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CPTTCheck-mCheckCPTT"
End Sub


Private Sub mSetColumnWidth()
    Dim ilNoColumns As Integer
    Dim ilCol As Integer
    Dim llWidth As Long
 
    
    ilNoColumns = 9
    lbcView.ColumnHeaders.Item(1).Width = 0 'Removed so that user can't select row
    lbcView.ColumnHeaders.Item(9).Width = 0
    lbcView.ColumnHeaders.Item(3).Width = (lbcView.Width) / 9.2  'Station
    lbcView.ColumnHeaders.Item(4).Width = (lbcView.Width) / 12.7  '# Prior
    lbcView.ColumnHeaders.Item(5).Width = (lbcView.Width) / 5  'Agreement
    lbcView.ColumnHeaders.Item(6).Width = (lbcView.Width) / 12.7  '# After
    lbcView.ColumnHeaders.Item(7).Width = (lbcView.Width) / 12.7  '# to Add
    lbcView.ColumnHeaders.Item(8).Width = (lbcView.Width) / 11.5  '# to Delete
    For ilCol = 1 To ilNoColumns Step 1
        If ilCol <> 2 Then
            llWidth = llWidth + lbcView.ColumnHeaders.Item(ilCol).Width
        End If
    Next ilCol
    '150 was used to get scroll with correct
    lbcView.ColumnHeaders.Item(2).Width = lbcView.Width - llWidth - 2 * GRIDSCROLLWIDTH - ilNoColumns * 210
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    mSetColumnWidth
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If imDone Then
        Erase tmCPTTCheck
        Erase tmWebSpotsInfo
    End If
End Sub

Private Sub mPopListView(slFileName As String, llTotalNoPrior As Long, llTotalNoAfter As Long, llTotalNoAdd As Long, llTotalNoDelete As Long, llChangedAstCPStatus As Long, llChangeAstCPStatusForCPTT As Long, llPostedComplete As Long, llPostPartial As Long, llPostCleared As Long, llDuplicates As Long)
    Dim llLoop As Long
    Dim ilVef As Integer
    Dim slCallLetters As String
    Dim slSortDate As String
    Dim slVehName As String
    Dim slVehicle As String
    Dim slStation As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim mItem As ListItem
    
    If UBound(tmCPTTCheck) - 1 > 0 Then
        For llLoop = 0 To UBound(tmCPTTCheck) - 1 Step 1
            ilVef = gBinarySearchVef(CLng(tmCPTTCheck(llLoop).iVefCode))
            If ilVef <> -1 Then
                tmCPTTCheck(llLoop).sKey = tgVehicleInfo(ilVef).sVehicle
            Else
                tmCPTTCheck(llLoop).sKey = "Vehicle " & tmCPTTCheck(llLoop).iVefCode & " missing"
            End If
            slCallLetters = gGetCallLettersByShttCode(tmCPTTCheck(llLoop).iShfCode)
            If slCallLetters <> "" Then
                tmCPTTCheck(llLoop).sKey = tmCPTTCheck(llLoop).sKey & "|" & slCallLetters
            Else
                tmCPTTCheck(llLoop).sKey = tmCPTTCheck(llLoop).sKey & "|" & "Station " & tmCPTTCheck(llLoop).iShfCode & " missing"
            End If
            slSortDate = Trim$(Str$(tmCPTTCheck(llLoop).lAgreementStart))
            Do While Len(slSortDate) < 6
                slSortDate = "0" & slSortDate
            Loop
            tmCPTTCheck(llLoop).sKey = tmCPTTCheck(llLoop).sKey & "|" & slSortDate
        Next llLoop
        ArraySortTyp fnAV(tmCPTTCheck(), 0), UBound(tmCPTTCheck), 0, LenB(tmCPTTCheck(0)), 0, LenB(tmCPTTCheck(0).sKey), 0
    End If
    If UBound(tmCPTTCheck) > 0 Then
        slVehName = ""
        For llLoop = 0 To UBound(tmCPTTCheck) - 1 Step 1
            Set mItem = lbcView.ListItems.Add()
            ilVef = gBinarySearchVef(CLng(tmCPTTCheck(llLoop).iVefCode))
            If ilVef <> -1 Then
                slVehicle = tgVehicleInfo(ilVef).sVehicle
            Else
                slVehicle = "Vehicle " & tmCPTTCheck(llLoop).iVefCode & " missing"
            End If
            If StrComp(Trim$(slVehicle), slVehName, vbTextCompare) <> 0 Then
                mItem.Text = ""
                mItem.SubItems(1) = Trim$(slVehicle)
                slVehName = Trim$(slVehicle)
            End If
            slCallLetters = gGetCallLettersByShttCode(tmCPTTCheck(llLoop).iShfCode)
            If slCallLetters <> "" Then
                slStation = slCallLetters
            Else
                slStation = "Station " & tmCPTTCheck(llLoop).iShfCode & " missing"
            End If
            mItem.SubItems(2) = slStation
            If tmCPTTCheck(llLoop).iNoPrior > 0 Then
                mItem.SubItems(3) = Trim$(Str$(tmCPTTCheck(llLoop).iNoPrior))
            End If
            slStartDate = gAdjYear(Format$(tmCPTTCheck(llLoop).lAgreementStart, "m/d/yy"))
            If tmCPTTCheck(llLoop).lAgreementEnd <> DateValue("12/31/2069") Then
                slEndDate = gAdjYear(Format$(tmCPTTCheck(llLoop).lAgreementEnd, "m/d/yy"))
            Else
                slEndDate = "TFN"
            End If
            mItem.SubItems(4) = slStartDate & "-" & slEndDate
            If tmCPTTCheck(llLoop).iNoAfter > 0 Then
                mItem.SubItems(5) = Trim$(Str$(tmCPTTCheck(llLoop).iNoAfter))
            End If
            If tmCPTTCheck(llLoop).iNoAdd > 0 Then
                mItem.SubItems(6) = Trim$(Str$(tmCPTTCheck(llLoop).iNoAdd))
            End If
            If tmCPTTCheck(llLoop).iNoDelete > 0 Then
                mItem.SubItems(7) = Trim$(Str$(tmCPTTCheck(llLoop).iNoDelete))
            End If
            mItem.SubItems(8) = llLoop
            If StrComp(slFileName, "CpttCheckLog.Txt", vbTextCompare) = 0 Then
                gLogMsg "Vehicle: " & slVehName & ", Station: " & slStation & ", # Prior: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoPrior)) & ", Dates: " & slStartDate & "-" & slEndDate & ", # After: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoAfter)) & ", # to Add: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoAdd)) & ", # to Delete: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoDelete)), slFileName, False
            Else
                gLogMsg "Vehicle: " & slVehName & ", Station: " & slStation & ", # Prior: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoPrior)) & ", Dates: " & slStartDate & "-" & slEndDate & ", # After: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoAfter)) & ", # Added: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoAdd)) & ", # Deleted: " & Trim$(Str$(tmCPTTCheck(llLoop).iNoDelete)), slFileName, False
            End If
        Next llLoop
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = "Totals"
        mItem.SubItems(3) = Trim$(Str$(llTotalNoPrior))
        mItem.SubItems(5) = Trim$(Str$(llTotalNoAfter))
        mItem.SubItems(6) = Trim$(Str$(llTotalNoAdd))
        mItem.SubItems(7) = Trim$(Str$(llTotalNoDelete))
        If StrComp(slFileName, "CpttCheckLog.Txt", vbTextCompare) = 0 Then
            gLogMsg "**** " & CStr(llTotalNoAdd) & " Total Number of CPTT records within Agreement dates to be added ****", slFileName, False
            gLogMsg "**** " & CStr(llTotalNoDelete) & " Total Number of CPTT records within Agreement dates to be removed ****", slFileName, False
            gLogMsg "**** " & CStr(llTotalNoPrior) & " Total Number of Prior CPTT records to be removed ****", slFileName, False
            gLogMsg "**** " & CStr(llTotalNoAfter) & " Total Number of After CPTT records to be removed ****", slFileName, False
        Else
            gLogMsg "**** " & CStr(llTotalNoAdd) & " Total Number of CPTT records within Agreement dates added ****", slFileName, False
            gLogMsg "**** " & CStr(llTotalNoDelete) & " Total Number of CPTT records within Agreement dates removed ****", slFileName, False
            gLogMsg "**** " & CStr(llTotalNoPrior) & " Total Number of Prior CPTT records removed ****", slFileName, False
            gLogMsg "**** " & CStr(llTotalNoAfter) & " Total Number of After CPTT records removed ****", slFileName, False
        End If
    Else
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = "No Discrepancies within Agreements"
        gLogMsg "**** No Discrepancies within Agreements ****", slFileName, False
    End If
    If StrComp(slFileName, "CpttCheckLog.Txt", vbTextCompare) = 0 Then
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llPostedComplete)) & " CPTT missing Complete as Set"
        gLogMsg "**** " & CStr(llPostedComplete) & " CPTT missing Complete as set ****", slFileName, False
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llPostPartial)) & " CPTT missing Partial as Set"
        gLogMsg "**** " & CStr(llPostPartial) & " CPTT missing Partial as set ****", slFileName, False
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llPostCleared)) & " CPTT require to be Cleared"
        gLogMsg "**** " & CStr(llPostCleared) & " CPTT require to be Cleared ****", slFileName, False
    Else
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llPostedComplete)) & " CPTT Complete set"
        gLogMsg "**** " & CStr(llPostedComplete) & " CPTT Complete Posting flags set ****", slFileName, False
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llPostPartial)) & " CPTT Partial Set"
        gLogMsg "**** " & CStr(llPostPartial) & " CPTT Partially Posted flags set ****", slFileName, False
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llPostCleared)) & " CPTT Cleared"
        gLogMsg "**** " & CStr(llPostCleared) & " CPTT Cleared ****", slFileName, False
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llChangedAstCPStatus)) & " AST Posted Status Set"
        gLogMsg "**** " & CStr(llChangedAstCPStatus) & " AST Posted Status set ****", slFileName, False
        Set mItem = lbcView.ListItems.Add()
        mItem.Text = ""
        mItem.SubItems(1) = Trim$(Str$(llChangeAstCPStatusForCPTT)) & " CPTT set because AST Set"
        gLogMsg "**** " & CStr(llChangeAstCPStatusForCPTT) & " CPTT set because AST Set ****", slFileName, False
    End If
    Set mItem = lbcView.ListItems.Add()
    mItem.Text = ""
    If StrComp(slFileName, "CpttCheckLog.Txt", vbTextCompare) = 0 Then
        mItem.SubItems(1) = Trim$(Str$(llDuplicates)) & " Duplicates to be Removed"
        gLogMsg "**** " & CStr(llDuplicates) & " Total Number of Duplicates to be Removed ****", slFileName, False
    Else
        mItem.SubItems(1) = Trim$(Str$(llDuplicates)) & " Duplicates Removed"
        gLogMsg "**** " & CStr(llDuplicates) & " Total Number of Duplicates Removed ****", slFileName, False
    End If
    
    Set mItem = lbcView.ListItems.Add()
    mItem.Text = ""
    mItem.SubItems(1) = Trim$(Str$(lmAstNotFoundOnWeb)) & " Spots Not Found on the Web. Most likely never exported to the web or the spot was modified and never re-exported. "
    
    Set mItem = lbcView.ListItems.Add()
    mItem.Text = ""
    mItem.SubItems(1) = Trim$(Str$(lmAstWebExported_NotPostedLocal)) & " Exported by the Web, but Not Posted Locally"
    
    Set mItem = lbcView.ListItems.Add()
    mItem.Text = ""
    mItem.SubItems(1) = Trim$(Str$(lmAstWebNotExported_PostedLocal)) & " Not Exported by Web, but has been Posted Locally. This could be due to spots being re-exported."
    
    
    
    
End Sub

Private Sub mGetAirWeekAdj(llIndex As Long, rst_att As ADODB.Recordset)
    Dim ilPledgedStatus As Integer
    Dim ilDay As Integer
    Dim ilPdDay As Integer
    Dim ilFdDay As Integer
    Dim rst_DAT As ADODB.Recordset
    
    tmCPTTCheck(llIndex).iAirWeekAdj = 0
    SQLQuery = "SELECT * FROM dat WHERE"
    SQLQuery = SQLQuery & " datAtfCode = " & rst_att!attCode
    Set rst_DAT = gSQLSelectCall(SQLQuery)
    
    Do While Not rst_DAT.EOF
        ilPledgedStatus = tgStatusTypes(rst_DAT!datFdStatus).iPledged
        If ilPledgedStatus = 1 Then 'Delayed
            'If rst_dat!datDACode <> 2 Then
            If rst_att!attPledgeType <> "C" Then
                tmCPTTCheck(llIndex).iAirWeekAdj = 0
                Exit Do
            End If
            For ilDay = 0 To 6 Step 1
                Select Case ilDay
                    Case 0
                        If rst_DAT!datFdMon = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                    Case 1
                        If rst_DAT!datFdTue = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                    Case 2
                        If rst_DAT!datFdWed = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                    Case 3
                        If rst_DAT!datFdThu = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                    Case 4
                        If rst_DAT!datFdFri = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                    Case 5
                        If rst_DAT!datFdSat = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                    Case 6
                        If rst_DAT!datFdSun = 1 Then
                            ilFdDay = ilDay
                            Exit For
                        End If
                End Select
            Next ilDay
            For ilDay = 0 To 6 Step 1
                Select Case ilDay
                    Case 0
                        If rst_DAT!datPdMon = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                    Case 1
                        If rst_DAT!datPdTue = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                    Case 2
                        If rst_DAT!datPdWed = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                    Case 3
                        If rst_DAT!datPdThu = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                    Case 4
                        If rst_DAT!datPdFri = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                    Case 5
                        If rst_DAT!datPdSat = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                    Case 6
                        If rst_DAT!datPdSun = 1 Then
                            ilPdDay = ilDay
                            Exit For
                        End If
                End Select
            Next ilDay
            If ilPdDay >= ilFdDay Then
                tmCPTTCheck(llIndex).iAirWeekAdj = 0
                Exit Do
            Else
                tmCPTTCheck(llIndex).iAirWeekAdj = 7
            End If
            rst_DAT.MoveNext
        Else
            tmCPTTCheck(llIndex).iAirWeekAdj = 0
            Exit Do
        End If
    Loop

End Sub

Private Function mBinarySearchWebAstInfo(llCode As Long) As Long
    
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tmWebSpotsInfo)
    llMax = UBound(tmWebSpotsInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tmWebSpotsInfo(llMiddle).lAstCode Then
            'found the match
            mBinarySearchWebAstInfo = llMiddle
            Exit Function
        ElseIf llCode < tmWebSpotsInfo(llMiddle).lAstCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    
    mBinarySearchWebAstInfo = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchWebAstInfo: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
    mBinarySearchWebAstInfo = -1
    Exit Function
End Function

Private Function mGetWebAstRecs() As Integer

    Dim slWebAstRecs As String
    Dim slWebImports As String
    Dim slPathFileName As String
    Dim hlFrom As Integer
    Dim llCurMaxRecs As Long
    Dim slExportedFlag As String
    Dim slPostedFlag As String
    Dim slastCode As String
    Dim ilExportedFlag As Integer
    Dim ilPostedFlag As Integer
    Dim llAstCode As Long
    Dim llCount As Long
    Dim ilRet As Integer
    Dim llTtlNotPosted As Long
    Dim llTtlPostedNotExported As Long
    Dim llTtlPostedAndExported As Long
    Dim llFileSize As Long
    Dim ilAccessToWeb As Integer
    Dim llTotalWebAstRecords As Long
    Dim slTemp As String
    Dim slFileName As String
    Dim smWebImports As String
    
    On Error GoTo ErrHand
        
    lacProgress.Caption = "Standby...Gathering Web records."
    ReDim tmWebSpotsInfo(0 To 100000) As WEBSPOTSINFO
    llCurMaxRecs = 100000
    mGetWebAstRecs = True
    

    ' Get the total number of ast records currently on the web now.
    llTotalWebAstRecords = gExecWebSQLWithRowsEffected("Select Count(*) from Spots")
    If llTotalWebAstRecords = -1 Then
        gLogMsg "ERROR: IMPORT - Unable to obtain Web Ast record count.", "CpttFixLog.Txt", False
        Screen.MousePointer = vbDefault
        cmdCancel.Enabled = True
        mGetWebAstRecs = False
        Exit Function
    End If

    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    slTemp = slTemp & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    slFileName = "WebAstRecs_" & slTemp

    Call gLoadOption(sgWebServerSection, "WebImports", smWebImports)
    smWebImports = gSetPathEndSlash(smWebImports, True)
    slPathFileName = smWebImports & slFileName

    gLogMsg "Instructing Web Site to get AST records.", "CpttFixLog.Txt", False
    SQLQuery = "Insert into WorkStatus (FileName, Status, Msg1, Msg2, DTStamp) "
    SQLQuery = SQLQuery & "Values('" & slFileName & "', '1', 'Calling Get AST Records', '', '"
    SQLQuery = SQLQuery & Format(gNow(), sgShowDateForm) & " ')"
    ilRet = gExecWebSQLWithRowsEffected(SQLQuery)
    If ilRet = -1 Then
        gLogMsg "ERROR: Unable to insert new WorkStatus record.", "CpttFixLog.Txt", False
        Screen.MousePointer = vbDefault
        cmdCancel.Enabled = True
        mGetWebAstRecs = False
        Exit Function
    End If

    If Not gExecExtStoredProc(slFileName, "GetASTRecords.exe", False, False) Then
        gLogMsg "Error: " & "Unable to instruct Web site to Import Copy Rotation Comments", "CpttFixLog.Txt", False
        gLogMsg "", "CpttFixLog.Txt", False
        cmdCancel.Caption = "&Done"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    ilRet = mExCheckWebWorkStatus(slFileName)
    If ilRet = True Then
        gLogMsg "Get AST Records was Successful.", "CpttFixLog.Txt", False
    Else
        mGetWebAstRecs = False
        gLogMsg "Get AST Records Failed.", "CpttFixLog.Txt", False
        Screen.MousePointer = vbDefault
        Exit Function
    End If


    ' FTP the file to this PC from the web server.
    If Not gFTPFileFromWebServer(slPathFileName, slFileName) Then
        gLogMsg "No data was found to import at the present time.", "CpttFixLog.Txt", False
        Screen.MousePointer = vbDefault
        cmdCancel.Enabled = True
        mGetWebAstRecs = False
        Exit Function
    End If

    'debug
    'Call gLoadOption("WebServer", "WebImports", smWebImports)
    'smWebImports = gSetPathEndSlash(smWebImports)
    'slPathFileName = smWebImports & "WebAstRecs_DF7FN1C1_070919_115729.txt"

    
    'Open the and fill the array
    'ilRet = 0
    'hlFrom = FreeFile
    'Open slPathFileName For Input Access Read Lock Write As hlFrom
    ilRet = gFileOpen(slPathFileName, "Input Access Read Lock Write", hlFrom)
    If ilRet <> 0 Then
        mGetWebAstRecs = False
        Exit Function
    End If

    llFileSize = FileLen(slPathFileName)
    If (gUsingWeb) Then
        If llFileSize = 0 Then
            mGetWebAstRecs = False
            Exit Function
        End If
    End If
    
    llCount = 0
    llTtlNotPosted = 0
    llTtlPostedNotExported = 0
    llTtlPostedAndExported = 0
    
    ' Skip past the header definition record.
    Input #hlFrom, slastCode, ilPostedFlag, ilExportedFlag
    Do While Not EOF(hlFrom)
       Input #hlFrom, slastCode, slPostedFlag, slExportedFlag
        ilPostedFlag = CInt(slPostedFlag)
        ilExportedFlag = CInt(slExportedFlag)
        llAstCode = CLng(slastCode)
        'not posted or exported
        If ilPostedFlag = 0 And ilExportedFlag = 0 Then
            tmWebSpotsInfo(llCount).iFlag = 0
            tmWebSpotsInfo(llCount).lAstCode = llAstCode
            llTtlNotPosted = llTtlNotPosted + 1
        End If
        'posted and not exported
        If ilPostedFlag = 1 And ilExportedFlag = 0 Then
            tmWebSpotsInfo(llCount).iFlag = 1
            tmWebSpotsInfo(llCount).lAstCode = llAstCode
            llTtlPostedNotExported = llTtlPostedNotExported + 1
        End If
        'posted and exported
        If ilPostedFlag = 1 And ilExportedFlag = 1 Then
            tmWebSpotsInfo(llCount).iFlag = 2
            tmWebSpotsInfo(llCount).lAstCode = llAstCode
            llTtlPostedAndExported = llTtlPostedAndExported + 1
        End If
        llCount = llCount + 1
        
        If llCount = llCurMaxRecs Then
            llCurMaxRecs = llCurMaxRecs + 100000
            ReDim Preserve tmWebSpotsInfo(0 To llCurMaxRecs) As WEBSPOTSINFO
        End If
    Loop
    
    ReDim Preserve tmWebSpotsInfo(0 To llCount) As WEBSPOTSINFO
    ArraySortTyp fnAV(tmWebSpotsInfo(), 0), UBound(tmWebSpotsInfo), 0, LenB(tmWebSpotsInfo(0)), 0, -2, 0
    
    If llTotalWebAstRecords <> llCount Then
        gLogMsg "Error: The ast count in the file returned from the web server did not match the count reported by SQL Server", "CpttFixLog.Txt", False
        gLogMsg "     : SQL Server Count = " & llTotalWebAstRecords & ", File record count = " & llCount, "CpttFixLog.Txt", False
        Screen.MousePointer = vbDefault
        cmdCancel.Enabled = True
        mGetWebAstRecs = False
        Exit Function
    End If
    
    gLogMsg "**** Start Web Totals ****", "CpttFixLog.txt", False
    gLogMsg "Total Web Not Posted Not Exported: " & CStr(llTtlNotPosted), "CpttFixLog.txt", False
    gLogMsg "Total Web Posted and Not Exported: " & CStr(llTtlPostedNotExported), "CpttFixLog.txt", False
    gLogMsg "Total Web Posted and Exported: " & CStr(llTtlPostedAndExported), "CpttFixLog.txt", False
    gLogMsg "**** End Web Totals ****", "CpttFixLog.txt", False
    gLogMsg "", "WebAstLog.Txt", False
    gLogMsg "** Finished Getting Web AST Process **", "WebAstLog.Txt", False
    
    Close hlFrom
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCPTTCheck - mGetWebAstRecs"
    mGetWebAstRecs = False
    Exit Function
End Function

'***************************************************************************************
' JD 08-22-2007
' This function was added to handle a special case occurring in the function
' mCheckWebWorkStatus. We believe a network error is causing the error handler
' to fire. Adding retry code to the function mCheckWebWorkStatus itself did not
' seem feasable because we did not know where the error was actually occuring and
' simplying calling a resume next could cause even more trouble.
'
'***************************************************************************************
Private Function mExCheckWebWorkStatus(sFileName As String) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String

    On Error GoTo Err_Handler
    mExCheckWebWorkStatus = -1
    For ilLoop = 1 To 10
        ilRet = mCheckWebWorkStatus(sFileName)
        mExCheckWebWorkStatus = ilRet
        If ilRet <> -2 Then ' Retry only when this status is returned.
            Exit Function
        End If
        gLogMsg "mExCheckWebWorkStatus is retrying due to an error in mCheckWebWorkStatus", "CpttFixLog.txt", False
        DoEvents
        Sleep 2000  ' Delay for two seconds when retrying.
    Next
    If ilRet = -2 Then
        ilRet = -1  ' Keep the original error of -1 so all callers can process the error normally.
        gMsg = "A timeout has occured in frmCPTTCheck - mExCheckWebWorkStatus"
        gLogMsg gMsg, "CpttFixLog.txt", False
        gLogMsg " ", "CpttFixLog.txt", False
    End If
    Exit Function

Err_Handler:
    Screen.MousePointer = vbDefault
    mExCheckWebWorkStatus = -1
    gMsg = ""
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gMsg = "A general error has occured in frmWebExportSchdSpot - mExCheckWebWorkStatus: " & "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc
    gLogMsg gMsg, "CpttFixLog.txt", False
    gLogMsg " ", "CpttFixLog.txt", False
    Exit Function
End Function

Private Function mCheckWebWorkStatus(sFileName As String) As Integer
 
    'D.S. 6/22/05
    
    'input - sFilemane is the unique file name that is the key into the web
    'server database to check it's status
    
    'Web Server Status - 0 = Done, 1 = Working and 2 = Error
    
    'Loop while the web server is busy processing spots and emails
    'Check the server every 10 seconds Report status
    
    Dim sFTPAddress As String
    'Dim ilRet As Integer
    Dim llWaitTime As Long
    Dim ilModResult As Integer
    Dim imStatus As Integer
    Dim slResult As String
    Dim llNumRows As Long
    Dim ilTimedOut As Integer
    Dim llMaxTimeToWait As Long
    Dim slTemp As String
    
    'Debug information
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    
    'Number of Seconds to Sleep
    Const clNumSecsToSleep As Long = 2
    Const clSleepValue As Long = clNumSecsToSleep * cmOneSecond
    llMaxTimeToWait = 600   ' 20 minutes
    'llMaxTimeToWait = 30   ' 1 minute as a test.
    
    'Assuming clNumSecsToSleep is 10 then a mod value of 6 would
    'be 6 loops at 10 seconds each or 1 minute
    Const clModValue As Integer = 6
    
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    
    smWebWorkStatus = "WebWorkStatus_" & slTemp & "_" & sgUserName & ".txt"
    'smWebWorkStatus = "WebWorkStatus_" & slTemp & ".txt"
    
    
    On Error GoTo ErrHand
    
    mCheckWebWorkStatus = False
    If Not gHasWebAccess() Then
        Exit Function
    End If
    
    Call gLoadOption(sgWebServerSection, "FTPAddress", sFTPAddress)
    llWaitTime = 0
    imStatus = 1
    Do While imStatus = 1 And llWaitTime < llMaxTimeToWait

        Sleep clSleepValue
        SQLQuery = "Select Count(*) from WorkStatus Where FileName = " & "'" & sFileName & "'"
        llNumRows = gExecWebSQLWithRowsEffected(SQLQuery)
        If llNumRows = -1 Then
            'An error was returned
            imStatus = 2
        End If
        If llNumRows > 0 Then
            SQLQuery = "Select FileName, Status, Msg1, Msg2, DTStamp from WorkStatus Where FileName = " & "'" & sFileName & "'"
            'Get the status information from the web server database and write it to a file
            Call gRemoteExecSql(SQLQuery, smWebWorkStatus, "WebExports", True, True, 30)
            DoEvents
            Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports")
            llWaitTime = llWaitTime + 1
            ilModResult = llWaitTime Mod clModValue
            imStatus = CInt(smStatus)
            'Handle Web Error Condition
            If imStatus = 2 Then
                gLogMsg "   " & "The Web Server Returned an ERROR. See Below. ", "CpttFixLog.Txt", False
                gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "CpttFixLog.Txt", False
                Call gEndWebSession("CpttFixLog.Txt")
                mCheckWebWorkStatus = False
                Exit Function
            End If
            If ilModResult = 0 And imStatus = 1 Then
                DoEvents
                'SetResults "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
                gLogMsg "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "CpttFixLog.Txt", False
                gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "CpttFixLog.Txt", False
                DoEvents
            End If
        End If
    Loop
    
    If llWaitTime >= llMaxTimeToWait Then
        'We timed out
        gLogMsg "   " & "A timeout occured while waiting on the web server for a response.", "CpttFixLog.Txt", False
        Call gEndWebSession("CpttFixLog.Txt")
        mCheckWebWorkStatus = False
        Exit Function
        
    End If
    
    Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports")
    imStatus = CInt(smStatus)
    'Handle Web Error Condition
    If imStatus = 2 Then
        gLogMsg "   " & "The Web Server Returned an ERROR. See Below. ", "CpttFixLog.Txt", False
        gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "CpttFixLog.Txt", False
        Call gEndWebSession("CpttFixLog.Txt")
        mCheckWebWorkStatus = False
        Exit Function
    End If
    'SetResults "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
    gLogMsg "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "CpttFixLog.Txt", False
    gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "CpttFixLog.Txt", False
    mCheckWebWorkStatus = True
    
    Exit Function
 
ErrHand:
    Screen.MousePointer = vbDefault
    mCheckWebWorkStatus = -2
    gMsg = ""
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gMsg = "A general error has occured in frmCPTTCheck - mCheckWebWorkStatus: " & "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc
    gLogMsg gMsg, "CpttFixLog.txt", False
    Exit Function
End Function

Private Function mProcessWebWorkStatusResults(sFileName As String, sIniValue As String) As Boolean

    'D.S. 6/22/05
    'Open the file with the web server status information and set the variables

    Dim slLocation As String
    Dim hlFrom As Integer
    Dim ilRet  As Integer
    
    On Error GoTo ErrHand
    
    mProcessWebWorkStatusResults = False
    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
    slLocation = gSetPathEndSlash(slLocation, True)
    slLocation = slLocation & sFileName
    
    'On Error GoTo FileErrHand:
    'hlFrom = FreeFile
    'ilRet = 0
    'Open slLocation For Input Access Read As hlFrom
    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        gMsgBox "Error: frmWebExportSchdSpot-mProcessWebWorkStatusResults was unable to open the file."
        GoTo ErrHand
    End If
    
    'Skip past the header record
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    
    'Cover the case that the Web Server times out and does not create the second line in the file
    If smStatus = "Status" Then
        smStatus = "1"
        gLogMsg "Warning: " & "Had to Set smStatus to 1 because the Work Status File Only had the Header in it.", "CpttFixLog.Txt", False
    End If
    
    Close hlFrom
    mProcessWebWorkStatusResults = True
    Exit Function

'FileErrHand:
'    Close hlFrom
'    mProcessWebWorkStatusResults = True
'    'Cover the case that the Web Server times out and does not create the second line in the file
'    If smStatus = "Status" Then
'        smStatus = "1"
'        gLogMsg "Warning: FileErrHand " & "Had to Set smStatus to 1 because the Work Status File Only had the Header in it.", "CpttFixLog.Txt", False
'    End If
'
'    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebExportSchdSpot-mProcessWebWorkStatusResults: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "CpttFixLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mProcessWebWorkStatusResults = False
    Exit Function
End Function

'Private Sub mSetResults(Msg As String, FGC As Long)
'    lbcView.
'    lbcMsg.ListIndex = lbcMsg.ListCount - 1
'    lbcMsg.ForeColor = FGC
'    DoEvents
'    gLogMsg Msg, "MarketronImportLog.Txt", False
'End Sub
