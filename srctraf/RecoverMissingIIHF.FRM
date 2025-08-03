VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RecoverMissingIIHF 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recover Missing IIHF Records"
   ClientHeight    =   2670
   ClientLeft      =   465
   ClientTop       =   2340
   ClientWidth     =   7095
   ClipControls    =   0   'False
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
   Icon            =   "RecoverMissingIIHF.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2670
   ScaleWidth      =   7095
   Begin vb6RecoverMissingIIHF.CSI_Calendar edcDate 
      Height          =   330
      Left            =   3255
      TabIndex        =   5
      Top             =   135
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      Text            =   "4/19/2020"
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
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   4
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Left            =   2085
      TabIndex        =   4
      Top             =   1455
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6345
      Top             =   165
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Recover"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   1785
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   1785
      Width           =   1050
   End
   Begin VB.Label lacDate 
      Caption         =   " Month Start Date to Recover IIHF"
      Height          =   315
      Left            =   135
      TabIndex        =   6
      Top             =   165
      Width           =   3030
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   2190
      Width           =   6585
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1875
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "RecoverMissingIIHF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: RecoverMissingIIHF.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links screen code
Option Explicit
Option Compare Text
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
Dim imUserInput As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imMovingRecs As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmStartDate As Long
Dim lmEndDate As Long
Dim iihf_rst As ADODB.Recordset
Dim sdf_rst As ADODB.Recordset
Dim vef_rst As ADODB.Recordset
Private Sub cmcCancel_Click()
    If imMovingRecs Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'
'           1-29-01 Remove everything in file before genning the weeks summary
'                   previously only removed the vehicle it was processing.  Dormant ones remained on file.
Private Sub cmcGen_Click()
    Dim llPercent As Long
    Dim slDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slSQLQuery As String
    Dim ilRet As Integer
    Dim ilNoMonths As Integer
    Dim ilMonth As Integer
    Dim llRet As Long
    
    lacInfo.Visible = False
    If imMovingRecs Then
        Exit Sub
    End If
    slStartDate = Trim$(edcDate.Text)
    If slStartDate = "" Then
        MsgBox "Please enter the transaction date"
        Exit Sub
    End If
    If Not gIsDate(slStartDate) Then
        MsgBox "Transaction date in error"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    slStartDate = gObtainStartStd(slStartDate)
    slEndDate = gDecOneDay(gObtainStartStd(Format(Now, "ddddd")))
    ilNoMonths = 0
    slDate = slStartDate
    Do
        ilNoMonths = ilNoMonths + 1
        slDate = gIncOneDay(gObtainEndStd(slDate))
    Loop While gDateValue(slDate) < gDateValue(slEndDate)
    Screen.MousePointer = vbDefault
    ilRet = MsgBox("Recoving " & ilNoMonths & " Months of Missing records in IIHF, Ok to Continue", vbOKCancel + vbQuestion, "Question")
    If ilRet = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmcCancel.Caption = "&Cancel"
    imMovingRecs = True
    lacInfo.Visible = True
    lacInfo.Caption = ""
    slSQLQuery = "Select Count(1) as vefCount from VEF_Vehicles Join shtt On vefName = shttCallLetters"
    Set vef_rst = gSQLSelectCall(slSQLQuery)
    lmTotalNoRecs = ilNoMonths * vef_rst!vefCount
    lmProcessedNoRecs = 0
    plcGauge.Visible = True
    plcGauge.Value = 0
    slDate = slStartDate
    For ilMonth = 1 To ilNoMonths Step 1
        
        
        'slSQLQuery = "Select Distinct chfCntrNo, vefName, chfCode, vefCode, Count(1) "
        'slSQLQuery = slSQLQuery & "From sdf_Spot_Detail "
        'slSQLQuery = slSQLQuery & "Left Outer Join iihf_ImptInvHeader On sdfChfCode = iihfChfCode and sdfVefCode = iihfVefCode "
        'slSQLQuery = slSQLQuery & "Left Outer Join chf_Contract_Header On sdfChfCode = chfCode "
        'slSQLQuery = slSQLQuery & "Left Outer Join vef_Vehicles On sdfVefCode = vefCode "
        'slSQLQuery = slSQLQuery & "Where sdfVefCode In (Select vefCode from VEF_Vehicles Join shtt On vefName = shttCallLetters) "
        'slSQLQuery = slSQLQuery & "and sdfSchStatus = 'S' And sdfDate >= '" & Format(slDate, sgSQLDateForm) & "'"
        'slSQLQuery = slSQLQuery & "and sdfDate <= '" & Format(gObtainEndStd(slDate), sgSQLDateForm) & "'"
        'slSQLQuery = slSQLQuery & "and iihfInvStartDate = '" & Format(slDate, sgSQLDateForm) & "'"
        'slSQLQuery = slSQLQuery & "and iihfCode is Null Group By chfCntrNo, vefName, chfCode, vefCode"
                
        'Select Distinct chfCntrNo, vefName, chfCode, vefCode, shttCode, shttCallLetters, iihfCode, sdfLineNo, iihfsourceform, iihfInvStartDate, sdfDate, sdfSchStatus from SDF_Spot_Detail left outer join chf_contract_Header On sdfChfCode = chfCode left outer join vef_vehicles on sdfvefcode = vefcode  left outer join shtt on vefName = shttCallLetters Left Outer Join iihf_ImptInvHeader On sdfChfCode = iihfChfCode and sdfVefCode = iihfVefCode where sdfDate >= '2019-12-30' And sdfDate < '2020-01-27' and sdfSchStatus <> 'M' and iihfCode is Null and shttCode is not Null order by vefname
        
        slSQLQuery = "Select vefCode from VEF_Vehicles Join shtt On vefName = shttCallLetters"
        Set vef_rst = gSQLSelectCall(slSQLQuery)
        Do While Not vef_rst.EOF
            DoEvents
            slSQLQuery = "Select distinct sdfChfCode from sdf_Spot_Detail"
            slSQLQuery = slSQLQuery & " Where sdfVefCode = " & vef_rst!vefCode
            slSQLQuery = slSQLQuery & " and sdfSchStatus = 'S' And sdfDate >= '" & Format(slDate, sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " and sdfDate <= '" & Format(gObtainEndStd(slDate), sgSQLDateForm) & "'"
            
            Set sdf_rst = gSQLSelectCall(slSQLQuery)
            Do While Not sdf_rst.EOF
                DoEvents
                slSQLQuery = "Select iihfCode from iihf_ImptInvHeader"
                slSQLQuery = slSQLQuery & " Where iihfVefCode = " & vef_rst!vefCode
                slSQLQuery = slSQLQuery & " and iihfChfCode = " & sdf_rst!sdfchfCode
                slSQLQuery = slSQLQuery & " and iihfInvStartDate = '" & Format(slDate, sgSQLDateForm) & "'"
                Set iihf_rst = gSQLSelectCall(slSQLQuery)
                If iihf_rst.EOF Then
                    'Create IIHF
                    DoEvents
                    slSQLQuery = "Insert Into iihf_ImptInvHeader ( "
                    slSQLQuery = slSQLQuery & "iihfCode, "
                    slSQLQuery = slSQLQuery & "iihfVefCode, "
                    slSQLQuery = slSQLQuery & "iihfChfCode, "
                    slSQLQuery = slSQLQuery & "iihfInvStartDate, "
                    slSQLQuery = slSQLQuery & "iihfFileName, "
                    slSQLQuery = slSQLQuery & "iihfStnEstimateNo, "
                    slSQLQuery = slSQLQuery & "iihfStnInvoiceNo, "
                    slSQLQuery = slSQLQuery & "iihfStnContractNo, "
                    slSQLQuery = slSQLQuery & "iihfStnAdvtName, "
                    slSQLQuery = slSQLQuery & "iihfAmfCode, "
                    slSQLQuery = slSQLQuery & "iihfSourceForm, "
                    slSQLQuery = slSQLQuery & "iihfUnused "
                    slSQLQuery = slSQLQuery & ") "
                    slSQLQuery = slSQLQuery & "Values ( "
                    slSQLQuery = slSQLQuery & 0 & ", "
                    slSQLQuery = slSQLQuery & vef_rst!vefCode & ", "   'tlIIHF.ifVefCode & ", "
                    slSQLQuery = slSQLQuery & sdf_rst!sdfchfCode & ", "   'tlIIHF.lfChfCode & ", "
                    slSQLQuery = slSQLQuery & "'" & Format$(slDate, sgSQLDateForm) & "', "   'Format$(tlIIHF.sfInvStartDate, sgSQLDateForm) & "', "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("Manual Post") & "', "    'gFixQuote(tlIIHF.sfFileName) & "', "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "   'gFixQuote(tlIIHF.sfStnEstimateNo) & "', "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "    'gFixQuote(tlIIHF.sfStnInvoiceNo) & "', "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "   'gFixQuote(tlIIHF.sfStnContractNo) & "', "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', " 'gFixQuote(tlIIHF.sfStnAdvtName) & "', "
                    slSQLQuery = slSQLQuery & 0 & ", "   'tlIIHF.lfAmfCode & ", "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("C") & "', "  'gFixQuote(tlIIHF.sfSourceForm) & "', "
                    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "' "   'gFixQuote(tlIIHF.sfUnused) & "' "
                    slSQLQuery = slSQLQuery & ") "
                    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                End If
                sdf_rst.MoveNext
            Loop
            lmProcessedNoRecs = lmProcessedNoRecs + 1
            llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            plcGauge.Value = llPercent
            vef_rst.MoveNext
        Loop
        slDate = gIncOneDay(gObtainEndStd(slDate))
    Next ilMonth
    If lmProcessedNoRecs <> lmTotalNoRecs Then
        Screen.MousePointer = vbDefault
        MsgBox "Not all IIHF records recovered."
        Exit Sub
    End If
    'slSQLQuery = "Delete From rvf_Receivables "
    'slSQLQuery = slSQLQuery & " Where rvfTranDate < '" & Format(slDate, sgSQLDateForm) & "'"
    'llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    'If llRet <> 0 Then
    '    Screen.MousePointer = vbDefault
    '    MsgBox "Unable to complete removal task, error # " & ilRet & ". Restore database"
    '    ilRet = btrClose(hmRvf)
    '    btrDestroy hmRvf
    '    ilRet = btrClose(hmPhf)
    '    btrDestroy hmPhf
    '    Exit Sub
    'End If
    lacInfo.Visible = True
    imMovingRecs = False
    cmcCancel.Caption = "&Done"
    MsgBox "IIHF records recover completed "
    '9/27/18: Getting Invalid procedure call
    'cmcCancel.SetFocus
    cmcGen.Enabled = False
    Screen.MousePointer = vbDefault
    If imTerminate Then
        'tmcTimer.Enabled = False
        cmcCancel_Click
    End If
    Exit Sub
cmcExportErr:
    imTerminate = True
    ilRet = err
    Resume Next
End Sub

Private Sub Form_Load()
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String
    
    mInit
    If imTerminate Then
        tmcTimer.Enabled = True
    Else
        slDate = Format$(Now(), "m/d/yy")
        slMonth = Month(slDate)
        slYear = Year(slDate)
        llValue = Val(slMonth) * Val(slYear)
        ilValue = Int(10000 * Rnd(-llValue) + 1)
        llValue = ilValue
        ilValue = Int(10000 * Rnd(-llValue) + 1)
        slStr = Trim$(Str$(ilValue))
        Do While Len(slStr) < 4
            slStr = "0" & slStr
        Loop
        sgSpecialPassword = slStr
        'smSpecial = "Login" & slStr
        'smGGBGSpecial = "8373191"

        CSPWord.Show vbModal
        If Not igPasswordOk Then
            imTerminate = True
            tmcTimer.Enabled = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    iihf_rst.Close
    btrStopAppl
    
    Set RecoverMissingIIHF = Nothing   'Remove data segment

    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'
'      2-02-01 Move Opens to gcmcGen; right before the
'           generation of summary records.  The timer
'           wasnt activated until after opening; therefore
'           files could remain open and the backup
'           aborted
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    imTerminate = False
    mParseCmmdLine
    
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
'Rm**    btrDestroy hgHlf
    Screen.MousePointer = vbHourglass
    imMovingRecs = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    lmStartDate = lmNowDate
    Do While gWeekDayLong(lmStartDate) <> 0
        lmStartDate = lmStartDate - 1
    Loop
    lmEndDate = lmStartDate + 7 * 26 - 1
    lacTo.Caption = sgDBPath
    gCenterStdAlone RecoverMissingIIHF
    Screen.MousePointer = vbDefault
    'imMinElapsed = imTimeoutInterval - 2
    'If imTimeoutInterval > 0 Then
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
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
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim slChar As String
    slCommand = Command$
    igDirectCall = 0
    sgIniPath = ""
    sgCommandStr = "debug"
    igStdAloneMode = True 'Switch from/to stand alone mode
    sgCallAppName = ""
    ilTestSystem = False
    ilPos = InStr(1, slCommand, "/S:Test", 1)
    If ilPos > 0 Then
        ilTestSystem = True
    End If
    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    If ilPos > 0 Then
        ilTestSystem = False
    End If
    imUserInput = True
    igBkgdProg = 0
    'If imTimeoutInterval > 0 Then
    '    igBkgdProg = 5
    'End If
    slStr = "Guide"
    gInitStdAlone RecoverMissingIIHF, slStr, ilTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "RecoverMissingIIHF.Frm"
    gLogMsg "Current Mode: UserInput.", "RecoverMissingIIHF.txt", False
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'

    Screen.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload RecoverMissingIIHF
    igManUnload = NO
End Sub



Private Sub tmcTimer_Timer()
    tmcTimer.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
End Sub


Private Function mCheckRecLength() As Integer
'    If Not gRecLengthOk("Rvf.Btr", Len(tmRvf)) Then
'        mCheckRecLength = False
'        Exit Function
'    End If
'    If Not gRecLengthOk("Phf.Btr", Len(tmRvf)) Then
'        mCheckRecLength = False
'        Exit Function
'    End If

    mCheckRecLength = True
End Function


