VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStationPersonRpt 
   Caption         =   "Station Personnel"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffStationPerson.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4590
      TabIndex        =   4
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
      Height          =   4750
      Left            =   240
      TabIndex        =   1
      Top             =   1785
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   7440
         Picture         =   "AffStationPerson.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Select Stations from File.."
         Top             =   480
         Width           =   360
      End
      Begin VB.CheckBox ckcMissingPersonnelOnly 
         Caption         =   "Missing Personnel Only"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox lbcStation 
         Height          =   3180
         ItemData        =   "AffStationPerson.frx":0E34
         Left            =   4320
         List            =   "AffStationPerson.frx":0E3B
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   3645
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   4350
         TabIndex        =   9
         Top             =   480
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
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
         ItemData        =   "AffStationPerson.frx":0E42
         Left            =   1050
         List            =   "AffStationPerson.frx":0E44
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   825
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   2175
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3480
      Top             =   720
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6630
      FormDesignWidth =   9360
   End
End
Attribute VB_Name = "frmStationPersonRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hmFrom As Integer

'constant index within export/import line
Private Const ATTCODEINDEX = 1
Private Const VENDORINDEX = 2  '1
Private Const LOGDATE = 3  '2
Private Const SPOTCOUNT = 4
Private Const PROCESSDATETIME = 5

Private imSort1 As Integer           '0 = station, 1 = vehicle,2 = Vendor
Private imSort2 As Integer           '0 = none, 1 = station, 2 =vehicle, 3 = vendor
Private imSort3 As Integer           '0 = none, 1 = station, 2 =vehicle, 3 = vendor
Private imChkAllStationsIgnore As Integer
Private smUsingUnivision As String * 1

Private rstATT As ADODB.Recordset
Private imIncludeCodes As Integer   'include or exclude the code list
Private imUseCodes() As Integer     'array of stations to include

Private Function mEnableGenerateReportButton()
    'Enable GENERATE REPORT when ALL filters are set
    If (lbcStation.SelCount > 0 Or chkAllStations.Value = vbChecked) Then
        cmdReport.Enabled = True
    Else
        cmdReport.Enabled = False
    End If
End Function






Private Sub chkAllStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkAllStationsIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStation.ListCount > 0 Then
        imChkAllStationsIgnore = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkAllStationsIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    
    'Enable GENERATE REPORT when ALL filters are set
    mEnableGenerateReportButton
    Screen.MousePointer = vbDefault

End Sub


Private Sub cmdReport_Click()
 
    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim slRptName As String         'full report name of crystal .rpt
    Dim slFromFile As String
    Dim llFromDate As Long
    Dim sFromDate As String
    Dim llToDate As Long
    Dim sToDate As String
    Dim slGenDate As String
    Dim slGenTime As String
    Dim ilOk As Integer
    Dim dFWeek As Date
    Dim blIsExport As Boolean
    Dim slFilePath As String
    Dim llCount As Long
    Dim ilLoop As Integer
    Dim ilTemp As Integer
    Dim llVefCode As Long
    Dim ilVefCode As Integer
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim blVendorFound As Boolean
    Dim blStationFound As Boolean
    Dim ilVendorId As Integer
    Dim ilShttCode As Integer
    Dim llAttCode As Long
    Dim slService As String
    Dim slType As String * 1
    Dim sGenDate As String      'generation date for filtering prepass records
    Dim sGenTime As String      'generation time for filtering prepass records
    
    On Error GoTo ErrHand
    
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
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    '8/28/2018 Set value for "Missing Personnel Only"       FYM
    sgCrystlFormula1 = "N"
    If (ckcMissingPersonnelOnly.Value = vbChecked) Then
        sgCrystlFormula1 = "Y"
    End If
    
    slRptName = "AfStationPersonRpt.rpt"
    slExportName = "AfStationPerson"
       
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    slGenDate = Format$(gNow(), "m/d/yyyy")
    slGenTime = Format$(gNow(), sgShowTimeWSecForm)

    Screen.MousePointer = vbHourglass
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    'create list array for selected STATIONS    Date: 8/2/2018    FYM
    ReDim imUseCodes(0 To 0) As Integer
    gObtainCodes lbcStation, imIncludeCodes, imUseCodes()        'build array of which codes to incl/excl

    mWriteStationReport slGenDate, slGenTime
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
        
    SQLQuery = "Select * from afr "
    SQLQuery = SQLQuery & "LEFT outer Join shtt On afr.afrAstCode = shtt.shttCode "
    SQLQuery = SQLQuery & "LEFT outer Join artt On artt.arttShttCode = shtt.shttCode "
    SQLQuery = SQLQuery + " where ( afrGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"
    
    'display the report
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName

    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "", "frmStationPersonRpt" & "-cmdReport"
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
    gHandleError "AffErrorLog.txt", "frmStationPersonRpt-Click"
End Sub

Private Sub cmdDone_Click()
    Unload frmStationPersonRpt

End Sub


Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmStationPersonRpt
    
End Sub

'TTP 9943 - Add ability to import stations for report selectivity
Private Sub cmdStationListFile_Click()
    Dim slCurDir As String
    slCurDir = CurDir
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    ' Import from the Selected File
    gSelectiveStationsFromImport lbcStation, chkAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim sNowDate As String
    ReDim tmVendorList(0 To 0) As VendorInfo

    frmStationPersonRpt.Caption = "Station Personnel Report - " & sgClientName

    imChkAllStationsIgnore = False
    chkAllStations.Value = vbUnchecked
    lbcStation.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcStation.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStation.ItemData(lbcStation.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop
    chkAllStations.Value = vbUnchecked
End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmStationPersonRpt
    gCenterForm frmStationPersonRpt
    gPopExportTypes cboFileType
    cboFileType.Enabled = True
    cmdReport.Enabled = False   'enable only after ALL filters are set  8/2/2018    FYM
    
End Sub
Private Sub Form_Initialize()
    mInit

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmStationPersonRpt = Nothing

End Sub


Private Sub mWriteStationReport(slGenDate As String, slGenTime As String)
    Dim SQLQuery As String
    Dim rst_cct As Recordset
    Dim blFound As Boolean
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT shttcode from shtt where shttType = 0"
    Set rst_cct = gSQLSelectCall(SQLQuery)
    Do While Not rst_cct.EOF

        blFound = gTestIncludeExclude(rst_cct!shttCode, imIncludeCodes, imUseCodes())
        If blFound Then
            SQLQuery = "INSERT INTO afr (afrGenDate, afrGenTime, "      'gen date & time
            SQLQuery = SQLQuery & "afrAstCode) "                        'call letters code
            SQLQuery = SQLQuery & " Values ( '" & Format$(slGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "',"
            SQLQuery = SQLQuery & rst_cct!shttCode & ") "
            
            cnn.BeginTrans
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "", "frmStationPersonRpt" & "-mWriteVendorReport"
                Exit Sub
            End If
            cnn.CommitTrans
        End If
        
        rst_cct.MoveNext
    Loop
    Exit Sub
        
ErrHand:
    Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "frmStationPersonRpt-mWriteVendorReport"
    Exit Sub
End Sub

Public Function mStripDoubleQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    Dim slQuote As String * 1
    slQuote = """"
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = slQuote Then
                sOutStr = sOutStr & " "
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    mStripDoubleQuote = sOutStr
    Exit Function
End Function







Private Sub lbcStation_Click()
  If imChkAllStationsIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = vbChecked Then
        imChkAllStationsIgnore = True
        'chkListBox.Value = False
        chkAllStations.Value = vbUnchecked
        imChkAllStationsIgnore = False
    End If
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
End Sub


