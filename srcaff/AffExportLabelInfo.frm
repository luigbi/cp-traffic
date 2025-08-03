VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportLabelInfo 
   Caption         =   "Mailing Labels"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   ControlBox      =   0   'False
   Icon            =   "AffExportLabelInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10275
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   1110
      TabIndex        =   20
      Top             =   4575
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "&Browse"
      Height          =   300
      Left            =   4800
      TabIndex        =   19
      Top             =   4575
      Width           =   1170
   End
   Begin VB.ListBox lbcContacts 
      Height          =   1230
      ItemData        =   "AffExportLabelInfo.frx":08CA
      Left            =   4440
      List            =   "AffExportLabelInfo.frx":08CC
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contact"
      Height          =   1215
      Left            =   4440
      TabIndex        =   13
      Top             =   960
      Width           =   2355
      Begin VB.OptionButton optContact 
         Caption         =   "Titles"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   870
         Width           =   2010
      End
      Begin VB.OptionButton optContact 
         Caption         =   "Affidavit Contact"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   540
         Width           =   1935
      End
      Begin VB.OptionButton optContact 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   255
         Value           =   -1  'True
         Width           =   2010
      End
   End
   Begin VB.OptionButton optSP 
      Caption         =   "Both"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   12
      Top             =   4200
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optSP 
      Caption         =   "Stations"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton optSP 
      Caption         =   "People"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   4200
      Width           =   1170
   End
   Begin VB.CheckBox chkListBox 
      Caption         =   "All"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.ListBox lbcVehAff 
      Height          =   2595
      ItemData        =   "AffExportLabelInfo.frx":08CE
      Left            =   240
      List            =   "AffExportLabelInfo.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   4125
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2595
      ItemData        =   "AffExportLabelInfo.frx":08D2
      Left            =   6960
      List            =   "AffExportLabelInfo.frx":08D4
      TabIndex        =   6
      Top             =   1050
      Width           =   3060
   End
   Begin VB.TextBox txtOffAirDate 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   945
   End
   Begin VB.TextBox txtOnAirDate 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   4560
      Width           =   1770
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5040
      Top             =   120
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5055
      FormDesignWidth =   10275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Vehicles"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lbcFile 
      Caption         =   "Export File"
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   4560
      Width           =   780
   End
   Begin VB.Label lacContact 
      Alignment       =   2  'Center
      Caption         =   "Contact"
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   6930
      TabIndex        =   7
      Top             =   720
      Width           =   3045
   End
   Begin VB.Label Label4 
      Caption         =   "End Date:"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Start Date:"
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   270
      Width           =   1170
   End
End
Attribute VB_Name = "frmExportLabelInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmExportLabelInfo - Mailing Labels
'*
'*  Created July,1998 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imChkListBoxIgnore As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private hmMsg As Integer
Private hmTo As Integer
Private smContact As String
Private smPhone As String
Private smFax As String
Private smEmail As String

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    'Dim NewForm As New frmViewReport
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    'grdVehAff.MoveFirst
    'For i = 0 To grdVehAff.Rows
    '    grdVehAff.SelBookmarks.Add grdVehAff.Bookmark
    '    grdVehAff.MoveNext
    'Next i
    If lbcVehAff.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllContacts_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
   
End Sub

Private Sub cmcBrowse_Click()
    'Label Info Export
    Dim slCurDir As String
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    'CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    '"(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    CommonDialog1.fileName = "Aflabels.csv"
    CommonDialog1.ShowSave
    ' Display name of selected file
    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmExportLabelInfo
End Sub

Private Sub LoadContactInfo(sCurContact As String, sCurPhone As String, iShttCode As Integer)
    On Error GoTo ErrHand
    Dim rst_tnt As ADODB.Recordset
    Dim rst_artt As ADODB.Recordset
    Dim slContactTitle As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim iltntCode As Integer
    Dim ilLen As Integer
    Dim ilPos As Integer

    smContact = ""
    smPhone = ""
    smFax = ""
    smEmail = ""

    If optContact(0).Value Then       ' no contact information
        Exit Sub
    End If

    smContact = Trim$(sCurContact)
    smPhone = Trim$(sCurPhone)
    If optContact(1).Value Then
        ' The user wants the Affidavit Contact. If the current contact is not blank, then try to use this.
        ' Note however that if we cannot find an exact match then the fax and email will be blank because it is not defined in the att table.
        If Len(smContact) > 0 Then
            ilLen = Len(smContact)
            ilPos = InStr(smContact, " ")
            If ilPos > 0 Then
                slFirstName = Left(smContact, ilPos - 1)
                slLastName = Trim(right(smContact, ilLen - ilPos))
            Else
                slFirstName = Trim(smContact)
            End If
            slFirstName = gFixQuote(slFirstName)
            slLastName = gFixQuote(slLastName)
            
            SQLQuery = "Select arttFirstName, arttLastName, arttPhone, arttFax, arttEmail From artt Where arttShttCode = " & iShttCode & " And arttFirstName = '" & slFirstName & "' And arttLastName = '" & slLastName & "'"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            If Not rst_artt.EOF Then
                smPhone = Trim$(rst_artt!arttPhone)
                smFax = Trim$(rst_artt!arttFax)
                smEmail = Trim$(rst_artt!arttEmail)
            End If
            rst_artt.Close
            Exit Sub
        End If
        ' We don't have a default Affidavit Contact. We'll need to look this up.
        SQLQuery = "Select arttFirstName, arttLastName, arttPhone, arttFax, arttEmail From artt Where arttShttCode = " & iShttCode & " And arttAffContact = '1'"
        Set rst_artt = gSQLSelectCall(SQLQuery)
        If Not rst_artt.EOF Then
            smContact = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
            smPhone = Trim$(rst_artt!arttPhone)
            smFax = Trim$(rst_artt!arttFax)
            smEmail = Trim$(rst_artt!arttEmail)
        End If
        rst_artt.Close
        Exit Sub
    End If

    ' Obtain the title selected in the lbcContacts list box.
    slContactTitle = lbcContacts.Text
    slContactTitle = gFixQuote(slContactTitle)
    ' Get the tntCode from the tnt table for this title.
    SQLQuery = "Select tntCode From tnt Where tntTitle = '" & slContactTitle & "'"
    Set rst_tnt = gSQLSelectCall(SQLQuery)
    If rst_tnt.EOF Then
        ' This could only occur if someone deleted the title after the user loaded this screen.
        gLogMsg "ERROR: Title information was not found for " & slContactTitle, "LableExportLog.Txt", False
        Exit Sub
    End If
    iltntCode = rst_tnt!tntCode

    ' Otherwise we need to lookup the name from the arttTable.
    ' The affidavit contact is not defined in the agreement. Obtain it from the
    SQLQuery = "Select arttFirstName, arttLastName, arttPhone, arttFax, arttEmail From artt Where arttShttCode = " & iShttCode & " And arttTntCode = " & iltntCode
    Set rst_artt = gSQLSelectCall(SQLQuery)
    If Not rst_artt.EOF Then
        smContact = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
        smPhone = Trim$(rst_artt!arttPhone)
        smFax = Trim$(rst_artt!arttFax)
        smEmail = Trim$(rst_artt!arttEmail)
    End If
    rst_artt.Close
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportLabelInfo-LoadContactInfo"
End Sub

Private Sub cmdExport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName, sVehicles, sStations As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim sStationType As String
    Dim iNoCDs As Integer
    Dim iLoop As Integer
    Dim iType As Integer
    Dim sContact As String
    Dim sOutput As String
    Dim ilIdx As Integer
    Dim slTest As String
    Dim sToFile As String
    Dim sDateTime As String
    Dim iRet As Integer
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim sMsgFileName As String
    Dim slStr As String
    Dim slSiteClient As String
    Dim slSiteAddr1 As String
    Dim slSiteAddr2 As String
    Dim slSiteAddr3 As String
    Dim sLastStation As String
    Dim ilTotalRecords As Long
    Dim cprst As ADODB.Recordset
    Dim rst_MaxCD As ADODB.Recordset
    Dim rst_Main As ADODB.Recordset
    Dim rst_site As ADODB.Recordset

    On Error GoTo ErrHand

    sToFile = Trim$(txtFile.Text)
    If Len(sToFile) < 1 Then
        gMsgBox "You must enter a file name to export to"
        Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    lbcMsg.Clear

    ' Validate Input Info
    If lbcVehAff.SelCount <= 0 Then
        gMsgBox "Vehicle must be selected.", vbOKOnly
        Exit Sub
    End If
    
    sStartDate = Trim$(txtOnAirDate.Text)
    If sStartDate = "" Then
        sStartDate = "1/1/1970"
    End If
    sEndDate = Trim$(txtOffAirDate.Text)
    If sEndDate = "" Then
        sEndDate = "12/31/2069"
    End If
    If gIsDate(sStartDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        txtOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        txtOffAirDate.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    iRet = 0
    'On Error GoTo cmdExportErr:
    'sDateTime = FileDateTime(sToFile)
    iRet = gFileExist(sToFile)
    If iRet = 0 Then
        Screen.MousePointer = vbDefault
        sDateTime = gFileDateTime(sToFile)
        iRet = gMsgBox("Export Previously Created " & sDateTime & " Continue with Export by Replacing File?", vbOKCancel, "File Exist")
        If iRet = vbCancel Then
            gLogMsg "** Terminated Because Export File Existed **", "LableExportLog.Txt", False
            Close #hmTo
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Kill sToFile
    End If
    On Error GoTo 0
    'iRet = 0
    'On Error GoTo cmdExportErr:
    'hmTo = FreeFile
    'Open sToFile For Output As hmTo
    iRet = gFileOpen(sToFile, "Output", hmTo)
    If iRet <> 0 Then
        gLogMsg "** Terminated because " & sToFile & " failed to open. **", "LableExportLog.Txt", False
        Close #hmTo
        Screen.MousePointer = vbDefault
        gMsgBox "Open Error #" & Str$(Err.Numner) & sToFile, vbOKOnly, "Open Error"
        Exit Sub
    End If
    gLogMsg "** Storing Output into " & sToFile & " **", "LableExportLog.Txt", False
    cmdExport.Enabled = False
    On Error GoTo 0

    If optSP(0).Value Then                          'stations
        sStationType = "shttType = 0"
    ElseIf optSP(1).Value Then                      'people
        sStationType = "shttType = 1"
    Else
        sStationType = ""                           'both
    End If

    sStartDate = Format(sStartDate, "m/d/yyyy")
    sgStdDate = sStartDate
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sDateRange = "(attOffAir >= '" & Format$(sStartDate, sgSQLDateForm) & "') And (attDropDate >= '" & Format$(sStartDate, sgSQLDateForm) & "') And (attOnAir <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
    sVehicles = ""

    If chkListBox.Value = 0 Then    ' = 0 Then                        'User did NOT select all vehicles
        For i = 0 To lbcVehAff.ListCount - 1 Step 1
            If lbcVehAff.Selected(i) Then
                If Len(sVehicles) = 0 Then
                    sVehicles = "(attVefCode = " & lbcVehAff.ItemData(i) & ")"
                Else
                    sVehicles = sVehicles & " OR (attVefCode = " & lbcVehAff.ItemData(i) & ")"
                End If
            End If
        Next i
    End If

    ' Obtain the Client name from the Site options table.
    slSiteClient = ""
    slSiteAddr1 = ""
    slSiteAddr2 = ""
    slSiteAddr3 = ""
    SQLQuery = "Select spfGClient, spfGAddr1, spfGAddr2, spfGAddr3 From SPF_Site_Options"
    Set rst_site = gSQLSelectCall(SQLQuery)
    If Not rst_site.EOF Then
        slSiteClient = rst_site!spfgClient
        slSiteAddr1 = rst_site!spfGAddr1
        slSiteAddr2 = rst_site!spfGAddr2
        slSiteAddr3 = rst_site!spfGAddr3
    End If
    rst_site.Close
    
    'Determine Max number of CD
    If sStationType <> "" Then
        SQLQuery = "Select MAX(attNoCDs) from att, shtt"
    Else
        SQLQuery = "Select MAX(attNoCDs) from att"
    End If
    SQLQuery = SQLQuery + " WHERE ((" & sDateRange & ")"
    If sStationType <> "" Then
        SQLQuery = SQLQuery + " AND ((attShfCode = shttCode)" & " And " & sStationType & ")"
    End If
    If sVehicles <> "" Then
        SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
    End If
    '10/29/14: Bypass Service agreements
    SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
    SQLQuery = SQLQuery + ")"
    imTerminate = False
    imExporting = True
    ilTotalRecords = 0
    Set rst_MaxCD = gSQLSelectCall(SQLQuery)
    'D.S. Avoid invalid use of Null error
    If rst_MaxCD(0).Value > 0 Then
        iNoCDs = rst_MaxCD(0).Value
        For iLoop = 1 To iNoCDs Step 1
            SQLQuery = "SELECT *"
            SQLQuery = SQLQuery + " FROM VEF_Vehicles, shtt, att"
            SQLQuery = SQLQuery + " WHERE (vefCode = attVefCode"
            SQLQuery = SQLQuery + " AND attshfCode = shttCode  "
            SQLQuery = SQLQuery + " AND attNoCDs >= " + Trim$(Str$(iLoop))
            SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
            If sStationType <> "" Then
                SQLQuery = SQLQuery + " AND (" & sStationType & ")"
            End If
            If sVehicles <> "" Then
                SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
            End If
            '10/29/14: Bypass Service agreements
            SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"

            ' Default sort order
            SQLQuery = SQLQuery + ")" + "ORDER BY shttCallLetters, vefName"

            Set rst_Main = gSQLSelectCall(SQLQuery)
            Do While Not rst_Main.EOF
                If sLastStation <> rst_Main!shttCallLetters Then
                    sLastStation = rst_Main!shttCallLetters
                    SetResults "Exporting labels for " & Trim$(rst_Main!shttCallLetters), RGB(0, 0, 0)
                End If
                ' Obtain the correct contact person.
                Call LoadContactInfo(rst_Main!attACName, rst_Main!attACPhone, rst_Main!shttCode)
                slStr = """" & Trim$(rst_Main!vefName) & """" & ","
                slStr = slStr + """" & Trim$(smContact) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttCallLetters) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttAddress1) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttAddress2) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttCity) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttState) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttCountry) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!shttZip) & """" & ","
                'slStr = slStr + """" & Format$(rst_Main!attAgreeStart, "mm/dd/yy") & "-" & Format$(rst_Main!attAgreeEnd, "mm/dd/yy") & """" & ","
                slStr = slStr + """" & Format$(sStartDate, "mm/dd/yy") & "-" & Format$(sEndDate, "mm/dd/yy") & """" & ","
                slStr = slStr + """" & Trim$(smPhone) & """" & ","
                slStr = slStr + """" & Trim$(smFax) & """" & ","
                slStr = slStr + """" & Trim$(smEmail) & """" & ","
                slStr = slStr + """" & Trim$(slSiteClient) & """" & ","
                slStr = slStr + """" & Trim$(slSiteAddr1) & """" & ","
                slStr = slStr + """" & Trim$(slSiteAddr2) & """" & ","
                slStr = slStr + """" & Trim$(slSiteAddr3) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!attLabelID) & """" & ","
                slStr = slStr + """" & Trim$(rst_Main!attLabelShipInfo) & """"
                Print #hmTo, slStr
                rst_Main.MoveNext
                DoEvents
                ilTotalRecords = ilTotalRecords + 1
                If imTerminate Then
                    Print #hmTo, "** User Terminated **"
                    Close #hmTo
                    imExporting = False
                    cmdExport.Enabled = True
                    Screen.MousePointer = vbDefault
                    cmdCancel.SetFocus
                    SetResults "Export Canceled", RGB(255, 0, 0)
                    Exit Sub
                End If
            Loop
        Next iLoop
        rst_Main.Close
        Screen.MousePointer = vbDefault
    Else
        SetResults "No labels were found for", RGB(255, 0, 0)
        SetResults lbcVehAff.Text & " " & sStartDate & " - " & sEndDate, RGB(255, 0, 0)
        Screen.MousePointer = vbDefault
        Close #hmTo
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Done"
        imExporting = False
        Exit Sub
    End If
    cmdExport.Enabled = True
    Screen.MousePointer = vbDefault
    Close #hmTo
    SetResults Trim$(Str(ilTotalRecords)) & " labels exported.", RGB(0, 155, 0)
    SetResults "Export Complete.", RGB(0, 155, 0)
    cmdCancel.Caption = "&Done"
    imExporting = False
    Exit Sub

'cmdExportErr:
'    iRet = Err
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportLabelInfo-cmdExport"
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmExportLabelInfo
    gCenterForm frmExportLabelInfo
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmExportLabelInfo.Caption = "Export Label Info - " & sgClientName
    imChkListBoxIgnore = False
    'SQLQuery = "SELECT vef.vefName from vef WHERE ((vef.vefvefCode = 0 AND vef.vefType = 'C') OR vef.vefType = 'L' OR vef.vefType = 'A')"
    'SQLQuery = SQLQuery + " ORDER BY vef.vefName"
    'Set rst = gSQLSelectCall(SQLQuery)
    'While Not rst.EOF
    '    grdVehAff.AddItem "" & rst(0).Value & ""
    '    rst.MoveNext
    'Wend
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    txtOnAirDate.Text = Format$(slDate, sgShowDateForm)
    txtOffAirDate.Text = Format$(DateAdd("d", 6, slDate), sgShowDateForm)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    'gPopExportTypes cboFileType     '3-15-04 populate export types
    'cboFileType.Enabled = False     'current defaulted to display, disallow export type selectivity
    ilRet = gPopTitleNames()        'get all the title names
    'place titles in list box for selection
    For iLoop = 0 To UBound(tgTitleInfo) - 1 Step 1
'        If iLoop = 0 Then
'            lbcContacts.AddItem "None"
'        End If
        lbcContacts.AddItem Trim$(tgTitleInfo(iLoop).sTitle)
        lbcContacts.ItemData(lbcContacts.NewIndex) = tgTitleInfo(iLoop).iCode
    Next iLoop
    lbcContacts.ListIndex = 0           'default to none
    txtFile.Text = sgExportDirectory & "Aflabels.csv"
    'lbcVehAff.Height = lbcContacts.Height + (lbcContacts.Top - lbcVehAff.Top)    'calc fullheight of vehicle box
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExportLabelInfo = Nothing
End Sub

Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub optContact_Click(Index As Integer)
    If Index < 2 Then           'no contact or using affidavit contact, dont show titles list box
        ' lbcContacts.Move lacContact.Left, lacContact.Top + lacContact.Height + 12, lacContact.Width, lbcVehAff.Height
        lbcContacts.Visible = False
        lacContact.Visible = False
        'lbcVehAff.Height = lbcContacts.Height + (lbcContacts.Top - lbcVehAff.Top)    'calc fullheight of vehicle box
    Else
        lbcContacts.Visible = True
        lacContact.Visible = True
        'lbcVehAff.Height = lacContact.Top - lbcVehAff.Top - 120
    End If
End Sub


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
    ilRet = 0
    slNowDate = Format$(gNow(), sgShowDateForm)
    slToFile = sgMsgDirectory & "ExptLabelInfo.Txt"
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
    Print #hmMsg, "** Export Label Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub SetResults(Msg As String, FGC As Long)
    lbcMsg.AddItem Msg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = FGC
    DoEvents
End Sub

