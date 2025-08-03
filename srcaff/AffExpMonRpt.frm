VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExpMonRpt 
   Caption         =   "Export Monitoring Report"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffExpMonRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   7125
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3510
      Top             =   1080
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5385
      FormDesignWidth =   7125
   End
   Begin VB.CommandButton cmdCrystalTemp 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4590
      TabIndex        =   5
      Top             =   255
      Width           =   1935
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
      Height          =   3030
      Left            =   240
      TabIndex        =   1
      Top             =   1815
      Width           =   6705
      Begin V81Affiliate.CSI_Calendar cccEffStartDate 
         Height          =   270
         Left            =   1500
         TabIndex        =   6
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   476
         Text            =   "10/1/2007"
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
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.Label Label3 
         Caption         =   "Log Week Start"
         Height          =   225
         Left            =   225
         TabIndex        =   2
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   720
      Width           =   1935
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
         ItemData        =   "AffExpMonRpt.frx":08CA
         Left            =   1050
         List            =   "AffExpMonRpt.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   825
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmExpMonRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private imDone As Integer
Private smLogFileName As String
Dim tmEsfCode() As ESFCODE
Dim tmEdfCode() As EDFCODE



Private Sub cmdCrystalTemp_Click()

    RunGenExportMonitor
    
End Sub

Sub RunGenExportMonitor()

    Dim ilRet As Integer
    Dim rstESF As ADODB.Recordset
    Dim rstEDF As ADODB.Recordset
    Dim rstExpMonitor As ADODB.Recordset
    Dim rstESFCnt As ADODB.Recordset
    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim slRptName As String         'full report name of crystal .rpt
    Dim llUpper As Long
    Dim llMax As Long
    Dim llMaxEdf As Long
    Dim llLoop As Long
    Dim llIdx As Long
    Dim llAttCode As Long
    Dim ilFound As Boolean
    Dim slDate6 As String
    Dim slCalDateShort As String
    Dim slSQLDate As String
    Dim sGenDate As String
    Dim sGenTime As String

    On Error GoTo ErrHand
    'Set up date parameters
    slCalDateShort = Format$(cccEffStartDate.Text, sgShowDateForm)
    If gIsDate(slCalDateShort) = False Or Weekday(slCalDateShort, vbSunday) <> vbMonday Then
        Beep
        gMsgBox "Please enter a valid Monday date (m/d/yy)", vbCritical
        cccEffStartDate.SetFocus
        Exit Sub
    End If

    slSQLDate = Format$(cccEffStartDate.Text, sgSQLDateForm)

    slDate6 = CDate(slSQLDate) + 6
    slDate6 = Format$(slDate6, "yyyy-mm-dd")

    ilRet = gPopAttByDate(slSQLDate) ' Create array that holds all possble attcodes for this date

    If Not ilRet Then
        gLogMsg "**** ERROR: gPopAttByDate Failed.  Exiting the Program ****", smLogFileName, False
        imDone = True
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    cmdCrystalTemp.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    sGenDate = Format(gNow(), sgShowDateForm)    'current date and time used as key for prepass file to
                                              'access and clear
    sGenTime = Format(gNow(), sgShowTimeWSecForm)

    gUserActivityLog "S", sgReportListName & ": Prepass"

    llUpper = 0
    llMax = 500
    ReDim tmEsfCode(0 To llMax) As ESFCODE

    '**Getting the results from the Summary table using a date criteria**
    SQLQuery = " SELECT esfCode From ESF_Export_Summary "
    SQLQuery = SQLQuery & "Where esfExpDate <=  '" & slDate6 & "' And CONVERT(esfExpDate, SQL_DATE) +  esfNumDays >='" & slSQLDate & "'"
    Set rstESF = gSQLSelectCall(SQLQuery)

    ''**Loopin thru both result sets and writing the records that have no association with the Export tables**
'    Set rstExpMonitor = New Recordset
'        rstExpMonitor.Fields.Append "Vehicle", adChar, 40
'        rstExpMonitor.Fields.Append "Station", adChar, 40
'        rstExpMonitor.Fields.Append "Market", adChar, 60
'        rstExpMonitor.Fields.Append "PDate", adChar, 20
'        rstExpMonitor.Open
    If Not rstESF.EOF Then 'If the day entered has records then write to EDF array otherwise skip this step and write to the recordset for the ttx file
        While Not rstESF.EOF  'Writing to the ESF Array
            tmEsfCode(llUpper).lCode = Trim$(rstESF!ESFCODE)
            llUpper = llUpper + 1
            If llUpper = llMax Then
                llMax = llMax + 500
                ReDim Preserve tmEsfCode(0 To llMax) As ESFCODE
            End If
            rstESF.MoveNext
        Wend
        ReDim Preserve tmEsfCode(0 To llUpper) As ESFCODE 'ReDim the array
            
        llUpper = 0        ''**Getting the results from the Detail table using the values rendered from the Summary table**
        llMaxEdf = 500
        ReDim tmEdfCode(0 To llMaxEdf) As EDFCODE
        For llLoop = 0 To UBound(tmEsfCode) - 1 Step 1
            If tmEsfCode(llLoop).lCode > 0 Then
                SQLQuery = "SELECT Distinct edfAttCode, edfEsfCode "
                SQLQuery = SQLQuery + " FROM EDF_Export_Detail"
                SQLQuery = SQLQuery + " WHERE (edfEsfCode = " & tmEsfCode(llLoop).lCode & ")"
                Set rstEDF = gSQLSelectCall(SQLQuery)
            End If
            
            If Not rstEDF.EOF Then 'Writing to the EDF Array
                While Not rstEDF.EOF
                    tmEdfCode(llUpper).lEsfCode = CLng(rstEDF!edfEsfCode)
                    tmEdfCode(llUpper).lAttCode = CLng(rstEDF!edfAttCode)
                    llUpper = llUpper + 1
                    If llUpper = llMaxEdf Then
                        llMaxEdf = llMaxEdf + 500
                        ReDim Preserve tmEdfCode(0 To llMaxEdf) As EDFCODE
                    End If
                    rstEDF.MoveNext
                Wend
            End If
        Next llLoop
        
        ReDim Preserve tmEdfCode(0 To llUpper) As EDFCODE
        If UBound(tmEdfCode) - 1 > 1 Then
            ArraySortTyp fnAV(tmEdfCode(), 0), UBound(tmEdfCode), 0, LenB(tmEdfCode(0)), 0, -2, 0
        End If
        For llIdx = 0 To UBound(tgAttExpMon) - 1 Step 1 '
            'Do a binarysearch after sort first  gBinarySearchEDF has the previous array stored to compare with tgAttExpMon  ' use -1
            ilFound = False
            llAttCode = mBinarySearchEDF(Trim$(tgAttExpMon(llIdx).lCode)) 'cccc
            If llAttCode >= 0 Then
                ilFound = True
            End If
            'Write to the rst for Crystal
            If ilFound Then
                '3-1-18 change to use prepass vs ADO report
'                rstExpMonitor.AddNew
'                rstExpMonitor.Fields("Vehicle") = Trim$(tgAttExpMon(llIdx).sVehName)
'                rstExpMonitor.Fields("Station") = Trim$(tgAttExpMon(llIdx).sCallLetters)
'                rstExpMonitor.Fields("Market") = Trim$(tgAttExpMon(llIdx).sMarket)
'                rstExpMonitor.Fields("PDate") = Trim$(slSQLDate)
'                rstExpMonitor.Update

                On Error GoTo ErrHand
                SQLQuery = "INSERT INTO " & "SMR"
                SQLQuery = SQLQuery & " (smrDMAMarket, smrZipCode, smrString1, smrString2, "
                SQLQuery = SQLQuery & " smrGendate, smrGenTime) "
                SQLQuery = SQLQuery & "VALUES ( '" & Trim$(tgAttExpMon(llIdx).sMarket) & "', " & "'" & Trim$(slSQLDate) & "', " & "'" & Trim$(tgAttExpMon(llIdx).sVehName) & "', " & "'" & Trim$(tgAttExpMon(llIdx).sCallLetters) & "', "
                SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"   '", "
                ilRet = 0
                cnn.BeginTrans
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "VehVisualRpt-mInsertGRFDiscrep"
                    cnn.RollbackTrans
                    Exit Sub
                End If
                If ilRet = 0 Then
                   cnn.CommitTrans
                End If
                
            End If
        Next llIdx
    Else             'No reason to do a binarysearch - the previous array has no values
        '11-8-11 no data, no need to create anything
'        For llIdx = 0 To UBound(tgAttExpMon) - 1 Step 1
'            'Write to the rst for Crystal
'            rstExpMonitor.AddNew
'            rstExpMonitor.Fields("Vehicle") = Trim$(tgAttExpMon(llIdx).sVehName)
'            rstExpMonitor.Fields("Station") = Trim$(tgAttExpMon(llIdx).sCallLetters)
'            rstExpMonitor.Fields("Market") = Trim$(tgAttExpMon(llIdx).sMarket)
'            rstExpMonitor.Fields("PDate") = Trim$(slSQLDate)
'            rstExpMonitor.Update
'        Next llIdx
    End If

    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    sgCrystlFormula1 = "For the week of " & Format$(Trim$(slSQLDate), "MM/DD/YY")
    
'    slRptName = "afMonExportsTTX"
'    slExportName = "afMonExportsTTX"
    slRptName = "afMonExports"
    slExportName = "afMonExports"
    
    SQLQuery = "Select * from smr where ( smrGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND smrGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & ")"

    gUserActivityLog "E", sgReportListName & ": Prepass"
    'frmCrystal.gActiveCrystalReports ilExportType, ilRptDest, slRptName & ".rpt", slExportName, rstExpMonitor
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName & ".rpt", slExportName
 
    cmdCrystalTemp.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    SQLQuery = "DELETE FROM SMR "
    SQLQuery = SQLQuery & " WHERE (smrGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and smrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"

    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "RunGenExportMonitor-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans

    Screen.MousePointer = vbDefault

    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExpMonRpt-RunReport"
End Sub

Private Sub cmdDone_Click()
    
    Unload frmExpMonRpt
    Erase tmEsfCode
    Erase tmEdfCode

End Sub


Private Sub cmdReturn_Click()

    frmReports.Show
    Unload frmExpMonRpt
    
End Sub
Private Sub Form_Load()

    Dim slDate As String
        
    frmExpMonRpt.Caption = "Export Monitoring Report - " & sgClientName
    slDate = Format$(gNow(), "m/d/yyyy")
    slDate = gObtainNextMonday(slDate)
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop

    cccEffStartDate.Text = Format$(slDate, sgShowDateForm)
    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = True

End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmExpMonRpt
    gCenterForm frmExpMonRpt

End Sub

Private Sub Form_Initialize()

    mInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmEsfCode
    Erase tmEdfCode
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmExpMonRpt = Nothing

End Sub

Public Function mBinarySearchEDF(llCode As Long) As Long
    
    'Returns the index number of tgEdfCod that matches the Code that was passed in
    'Note: for this to work tgEdfCode was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    On Error GoTo ErrHand
    
    llMin = LBound(tmEdfCode)
    llMax = UBound(tmEdfCode) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tmEdfCode(llMiddle).lAttCode Then
            'found the match
            mBinarySearchEDF = llMiddle
            Exit Function
        ElseIf llCode < tmEdfCode(llMiddle).lAttCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchEDF = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mBinarySearchEDF: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    mBinarySearchEDF = -1
    Exit Function
    
End Function

