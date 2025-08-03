VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAlertRpt 
   Caption         =   "Station Clearance Report"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "AffAlertRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7575
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5865
      FormDesignWidth =   7575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   240
      TabIndex        =   6
      Top             =   1725
      Width           =   6960
      Begin V81Affiliate.CSI_Calendar CalSelCFrom 
         Height          =   260
         Left            =   4770
         TabIndex        =   22
         Top             =   2595
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Text            =   "9/3/2020"
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
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.Frame frcClear 
         Caption         =   "Include Cleared Alerts"
         Height          =   915
         Left            =   120
         TabIndex        =   14
         Top             =   1725
         Width           =   6015
         Begin VB.CheckBox ckcClear 
            Caption         =   "Pool Alert"
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   24
            Top             =   300
            Width           =   1575
         End
         Begin VB.CheckBox ckcClear 
            Caption         =   "Affiliate Export"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   17
            Top             =   300
            Width           =   1575
         End
         Begin VB.CheckBox ckcClear 
            Caption         =   "Traffic Logs"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   16
            Top             =   300
            Width           =   1335
         End
         Begin VB.CheckBox ckcClear 
            Caption         =   "Contract"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lacClearFrom 
            Caption         =   "Effective Clear Date"
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   615
            Width           =   1500
         End
      End
      Begin VB.Frame frcAlert 
         Caption         =   "Include Alerts"
         Height          =   660
         Left            =   120
         TabIndex        =   10
         Top             =   1005
         Width           =   6015
         Begin VB.CheckBox ckcAlert 
            Caption         =   "Pool Alert"
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   23
            Top             =   300
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox ckcAlert 
            Caption         =   "Affiliate Export"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   13
            Top             =   300
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox ckcAlert 
            Caption         =   "Traffic Logs"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   300
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox ckcAlert 
            Caption         =   "Contract"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   765
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   1755
         Begin VB.OptionButton optSortby 
            Caption         =   "Alert Type"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   465
            Width           =   1365
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Date"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   210
            Value           =   -1  'True
            Width           =   1200
         End
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   20
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   19
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   18
      Top             =   225
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffAlertRpt.frx":08CA
         Left            =   1050
         List            =   "AffAlertRpt.frx":08CC
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   1170
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmAlertRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'*  frmAlertRpt - Print list of alerts created by Traffic by one of
'*      the following methods:  Contract- change proposal to complete;
'*      Spots - changing copy after log gen or moving/changing a spot;
'*      Exports required - reprinting or print final require an
'*      Affiliate export to be generated.
'*      Print list of alerts cleared by Traffic or Affiliate system.
'*
'*  Created 04/15/04 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit


Private Sub ckcClear_Click(Index As Integer)
       If ckcClear(Index).Value = vbChecked Then
        CalSelCFrom.SetEnabled (True)
    Else
        If ckcClear(0).Value = vbUnchecked And ckcClear(1).Value = vbUnchecked And ckcClear(2).Value = vbUnchecked Then
            CalSelCFrom.SetEnabled (False)
        End If
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmAlertRpt
End Sub

'       2-23-04 make start/end dates mandatory.  Now that AST records are
'       created, looping thru the earliest default of  1/1/70 and latest default of 12/31/2069
'       is not feasible
Private Sub cmdReport_Click()
    Dim sStartDate As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim sStartTime As String
    Dim slNow As String
    Dim slSelection As String, slAlert As String, slClear As String    'Dan changed line so all strings
    Dim slTypeC, slTypeL, slTypeForR As String  'Dan only slTypeForR is string; rest of line are variants
    Dim slEffClearDate As String
    Dim slSelectAlert, slSelectClear As String
    Dim AlertRst As ADODB.Recordset
    Dim sGenDate As String
    Dim sGenTime As String
    Dim sUserName As String
    Dim llContract As Long
    Dim slTypeForP As String        'Date: 10/18/2019   added Pool Alerts
    Dim slAllTypes As String        '9-3-20
    Dim slAllClearTypes As String   '9-3-20

    On Error GoTo ErrHand
  
    If ckcAlert(0).Value Or ckcAlert(1).Value Or ckcAlert(2).Value Or ckcAlert(3).Value Or ckcClear(0).Value Or ckcClear(1).Value Or ckcClear(2).Value Or ckcClear(3).Value Then
    Else
        MsgBox "At least one Alert or Clear item must be selected"
        ckcAlert(0).SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    sStartDate = ""
    If ckcClear(0).Value = vbChecked Or ckcClear(1).Value = vbChecked Or ckcClear(2).Value = vbChecked Then
        'user wants clear records
        If Trim$(CalSelCFrom.Text) <> "" Then
            sStartDate = Trim$(CalSelCFrom.Text)
            sStartDate = Format(sStartDate, "m/d/yyyy")
            dFWeek = CDate(sStartDate)
            'sgCrystlFormula2 = "'" & Format$(dFWeek, "mm") & "/" & Format$(dFWeek, "dd") & "/" & Format$(dFWeek, "yyyy") & "'"
            sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        End If
    Else
        sgCrystlFormula2 = ""
    End If
    
    If optSortby(0).Value = True Then          'Date
        sgCrystlFormula1 = "'D'"
    Else                                        'Type
        sgCrystlFormula1 = "'T'"
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    slSelection = ""
    slAlert = ""
    slClear = ""
    slTypeC = ""
    slTypeL = ""
    slTypeForR = ""
    slTypeForP = ""
    slSelection = ""
    slAllTypes = ""
    slAllClearTypes = ""
    
    'Include Ready alerts
'    If ckcAlert(0).Value = vbChecked Or ckcAlert(1).Value = vbChecked Or ckcAlert(2).Value = vbChecked Or ckcAlert(3).Value = vbChecked Then    'ttp 9941
'        slAlert = "((aufStatus = 'R') and "
'        If ckcAlert(0).Value = vbChecked Then
'            slTypeC = " aufType = 'C'"
'        End If
'        If ckcAlert(1).Value = vbChecked Then
'            If slTypeC <> "" Then
'                slTypeL = " or aufType = 'L' "
'            Else
'                slTypeL = " aufType  = 'L' "
'            End If
'        End If
'        If ckcAlert(2).Value = vbChecked Then
'            If slTypeC <> "" Or slTypeL <> "" Then
'                slTypeForR = " or aufType = 'R' or aufType = 'F' "
'            Else
'                slTypeForR = " (aufType = 'R' or aufType = 'F') "
'            End If
'        End If
'        If ckcAlert(3).Value = vbChecked Then
'            If slTypeC <> "" Or slTypeL <> "" Then
'                slTypeForP = " or aufType = 'U' and aufSubType = 'P' "
'            Else
'                slTypeForP = " aufType = 'U'  and aufSubType = 'P' "
'            End If
'        End If
'        slSelectAlert = slAlert & "(" & slTypeC & slTypeL & slTypeForR & slTypeForP & "))"
'    End If

    'ttp 9941 9-3-20 alter way to gather the data for sql query
    If ckcAlert(0).Value = vbChecked Or ckcAlert(1).Value = vbChecked Or ckcAlert(2).Value = vbChecked Or ckcAlert(3).Value = vbChecked Then    'ttp 9941
        slAlert = "(aufStatus = 'R') and "
        If ckcAlert(0).Value = vbChecked Then
            slAllTypes = " (aufType = 'C')"
        End If
        If ckcAlert(1).Value = vbChecked Then
            If slAllTypes = "" Then
                slAllTypes = " (aufType = 'L') "
            Else
                slAllTypes = slAllTypes & " or (aufType  = 'L') "
            End If
        End If
        If ckcAlert(2).Value = vbChecked Then
            If slAllTypes = "" Then
                slAllTypes = " (aufType = 'R' or aufType = 'F') "
            Else
                slAllTypes = slAllTypes & " or (aufType = 'R' or aufType = 'F') "
            End If
        End If
        If ckcAlert(3).Value = vbChecked Then
            If slAllTypes = "" Then
                slAllTypes = " (aufType = 'U' and aufSubType = 'P' )"
            Else
                slAllTypes = slAllTypes & " or (aufType = 'U'  and aufSubType = 'P' )"
            End If
        End If
        slSelectAlert = slAlert & "(" & slAllTypes & ")"
    End If


    'insert cleared alerts
    slEffClearDate = ""
    If sStartDate <> "" Then
        slEffClearDate = " and aufClearDate >= '" & Format$(dFWeek, sgSQLDateForm) & "'"
    End If
    
    If ckcClear(0).Value = vbChecked Or ckcClear(1).Value = vbChecked Or ckcClear(2).Value = vbChecked Or ckcClear(3).Value = vbChecked Then  'ttp 9941
'        If slAlert = "" Then
'            slClear = "((aufStatus = 'C'" & slEffClearDate & ") and "
'        Else
'            slAlert = "(" & slAlert & ")"
'            slClear = " or ((aufStatus = 'C' " & slEffClearDate & ") and "
'        End If
'        If ckcClear(0).Value = vbChecked Then
'            slTypeC = "aufType = 'C' "
'        End If
'        If ckcClear(1).Value = vbChecked Then
'            If slTypeC <> "" Then
'                slTypeL = " or aufType = 'L' "
'            Else
'                slTypeL = " aufType = 'L' "
'            End If
'        End If
'        If ckcClear(2).Value = vbChecked Then
'            If slTypeC <> "" Or slTypeL <> "" Then
'                slTypeForR = " or aufType = 'R' or aufType = 'F' "
'            Else
'                slTypeForR = " aufType = 'R' or aufType = 'F' "
'            End If
'        End If
'        If ckcClear(3).Value = vbChecked Then
'            If slTypeC <> "" Or slTypeL <> "" Then
'                slTypeForP = " or aufType = 'U' and aufSubType = 'P' "
'            Else
'                slTypeForP = " aufType = 'U' and aufSubType = 'P' "
'            End If
'        End If
'        slSelectClear = slClear & "(" & slTypeC & slTypeL & slTypeForR & slTypeForP & "))"

        'ttp 9941 9-3-20 alter way to gather the data for sql query
        If slAlert = "" Then
            slClear = "(aufStatus = 'C'" & slEffClearDate & ") and "
        Else
            slAlert = "(" & slAlert & ")"
            slClear = " or (aufStatus = 'C' " & slEffClearDate & ") and "
        End If
        If ckcClear(0).Value = vbChecked Then
            slAllClearTypes = " (aufType = 'C') "
        End If
        If ckcClear(1).Value = vbChecked Then
            If slAllClearTypes = "" Then
                slAllClearTypes = " (aufType = 'L') "
            Else
                slAllClearTypes = slAllClearTypes & " or (aufType = 'L') "
            End If
        End If
        If ckcClear(2).Value = vbChecked Then
            If slAllClearTypes = "" Then
                slAllClearTypes = " ( aufType = 'R' or aufType = 'F') "
            Else
                slAllClearTypes = slAllClearTypes & " or (aufType = 'R' or aufType = 'F') "
            End If
        End If
        If ckcClear(3).Value = vbChecked Then
            If slAllClearTypes = "" Then
                slAllClearTypes = " (aufType = 'U' and aufSubType = 'P') "
            Else
                slAllClearTypes = slAllClearTypes & " or (aufType = 'U' and aufSubType = 'P') "
            End If
        End If
        slSelectClear = slClear & "(" & slAllClearTypes & ")"
    End If
    
    slSelection = slSelectAlert & slSelectClear
        
    'SQLQuery = "select aufcode, aufchfcode, chfcntrno, vefname,  xxx.urfcode, yyy.urfcode from "
    'SQLQuery = "select aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufmoweekdate , "
    'SQLQuery = SQLQuery & "aufclearmethod, aufcleardate, aufcleartime, chfcntrno, vefname "
    SQLQuery = "select aufcode, aufcreateurfcode, aufcreateustcode, aufclearurfcode, aufclearustcode, aufChfCode, aufstatus,  "
    SQLQuery = SQLQuery & "urf_user_options.urfname as UrfCreateName_urfName, urf_clrUser_options.urfname as UrfClearName_urfName, ust_clr_options.ustname as UstClearName_ustName, ust_create_options.ustname as UstCreateName_ustName, chfCntrno  "
'    SQLQuery = SQLQuery & " from ((((((auf_alert_user left Outer Join chf_contract_header on aufchfcode = chfcode) "
'    SQLQuery = SQLQuery & " left outer join vef_vehicles on aufvefcode = vefcode) "
'    SQLQuery = SQLQuery & " left Outer Join urf_User_options on aufcreateurfcode = urf_user_options.urfcode) "
'    SQLQuery = SQLQuery & " left outer join urf_user_options urf_clruser_options on aufclearurfcode = urf_clruser_options.urfcode) "
'    SQLQuery = SQLQuery & " left outer join ust ust_clr_options on aufclearustcode = ust_clr_options.ustcode) "
'    SQLQuery = SQLQuery & " left outer join ust ust on aufcreateustcode = ust.ustcode) "
    SQLQuery = SQLQuery & " from (((((auf_alert_user left Outer Join urf_User_options on aufcreateurfcode = urf_user_options.urfcode) "
    SQLQuery = SQLQuery & " left outer join urf_user_options urf_clrUser_options on aufclearurfCode = urf_clruser_options.urfcode) "
    SQLQuery = SQLQuery & " left outer join ust ust_create_options on aufcreateustcode = ust_create_options.ustcode) "
    SQLQuery = SQLQuery & " left outer join ust ust_clr_options on aufclearustcode = ust_clr_options.ustcode) "
    SQLQuery = SQLQuery & " left Outer Join chf_contract_header on aufchfcode = chfcode) "
    SQLQuery = SQLQuery & " where " & slSelection
    
    Set AlertRst = gSQLSelectCall(SQLQuery)
    
    sGenDate = Format$(gNow(), "m/d/yyyy")
    sGenTime = Format$(gNow(), sgShowTimeWSecForm)

    While Not AlertRst.EOF
        sUserName = ""
        llContract = 0
        If AlertRst!aufChfCode > 0 And Not IsNull(AlertRst!chfCntrNo) Then
            llContract = AlertRst!chfCntrNo
        End If
        
        If AlertRst!aufStatus = "R" Then        'alerts
            If AlertRst!aufCreateUrfCode > 0 Then
                'sUserName = gDecryptField(Trim$(AlertRst(7).Value))  'create urf name
                sUserName = gDecryptField(Trim$(AlertRst!UrfCreateName_urfName))    'urf create name
            Else
                If AlertRst!aufCreateUstCode > 0 Then
                    'sUserName = Trim$(AlertRst(10).Value)       'create ust name
                    sUserName = Trim$(AlertRst!UstCreateName_ustName)   'UST create name
                End If
            End If
        Else                                    'clear
            If AlertRst!aufclearurfcode > 0 Then
                'sUserName = gDecryptField(Trim$(AlertRst(8).Value))    'clear urf
                sUserName = gDecryptField(Trim$(AlertRst!UrfClearName_urfName)) 'urf clear name
            Else
                If AlertRst!aufClearUstCode > 0 Then
                    'sUserName = Trim$(AlertRst(9).Value)            'clear ust
                    sUserName = Trim$(AlertRst!UstClearName_ustName)    'ust clear name
                End If
            End If
        End If
        SQLQuery = "INSERT INTO " & "GRF_Generic_Report"
        SQLQuery = SQLQuery & " (grfCode4,grfChfCode, grfGenDesc, grfGendate, grfGenTime) "
        SQLQuery = SQLQuery & " VALUES (" & AlertRst!aufCode & ", " & llContract & ", " & "'" & sUserName & "', " & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"    '", "
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "AlertRpt-cmdReport_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        AlertRst.MoveNext
    Wend
    
    AlertRst.Close
    SQLQuery = "Select * from grf_generic_report "
    SQLQuery = SQLQuery + "INNER JOIN  auf_alert_user on grfcode4 = aufcode "
    SQLQuery = SQLQuery + "LEFT OUTER JOIN vef_vehicles on aufvefcode = vefCode"
    SQLQuery = SQLQuery + " where ( grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    SQLQuery = SQLQuery
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AlertStatusSQL.rpt", "Alert"

    DoEvents
    'remove all the records just printed
    SQLQuery = "DELETE FROM grf_generic_report "
    SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "AlertRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    Screen.MousePointer = vbDefault

        
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AlertRpt-Click"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmAlertRpt
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmAlertRpt
    gCenterForm frmAlertRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sNowDate As String
    Dim dDate As Date
    
    frmAlertRpt.Caption = "Alert Status Report - " & sgClientName
    'slDate = Format$(Now, "m/d/yyyy")
    'Do While Weekday(slDate, vbMonday) <> vbMonday
    '    slDate = DateAdd("d", -1, slDate)
    'Loop
   ' txtOnAirDate.Text = Format$(slDate, sgShowDateForm)
    'txtOffAirDate.Text = Format$(DateAdd("d", 6, slDate), sgShowDateForm)
    CalSelCFrom.Move lacClearFrom.Left + lacClearFrom.Width, frcClear.Top + ckcClear(0).Top + ckcClear(0).Height
    sNowDate = Format$(gNow(), "m/d/yyyy")
    CalSelCFrom.SetEnabled (False)
   
    dDate = CDate(sNowDate)
    dDate = DateAdd("D", -7, dDate)
    'backup to Monday
    slDate = Format$(dDate, sgShowDateForm)
    slDate = gObtainPrevMonday(slDate)
    CalSelCFrom.Text = slDate
    
    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default

End Sub

Private Sub Form_Unload(Cancel As Integer)
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmAlertRpt = Nothing
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       '3-15-04 default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Sub optSortby_Click(Index As Integer)
Dim iLoop As Integer
Dim iIndex As Integer
    
   
End Sub
'


