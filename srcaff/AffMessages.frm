VERSION 5.00
Begin VB.Form frmMessages 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Viewer"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmcPrt 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7485
      Top             =   4905
   End
   Begin VB.PictureBox pbcPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   2790
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   6
      Top             =   2145
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer tmcUsers 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1260
      Top             =   5010
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   615
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lbcUsers 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      ItemData        =   "AffMessages.frx":0000
      Left            =   4905
      List            =   "AffMessages.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ListBox lbcShowFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      ItemData        =   "AffMessages.frx":0004
      Left            =   240
      List            =   "AffMessages.frx":0006
      TabIndex        =   3
      Top             =   1395
      Width           =   9015
   End
   Begin VB.ListBox lbcFileSelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      ItemData        =   "AffMessages.frx":0008
      Left            =   240
      List            =   "AffMessages.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4300
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4965
      TabIndex        =   1
      Top             =   4890
      Width           =   2010
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2520
      TabIndex        =   0
      Top             =   4890
      Width           =   2010
   End
   Begin VB.Image imcPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8415
      Picture         =   "AffMessages.frx":000C
      Top             =   4740
      Width           =   480
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private smChoice As String
Private imNewDisplay As Integer
Dim tmAffMessages(0 To 99) As AFFMESSAGES   ' JD 02-09-24 Added 1 for the wegener check utility
'Dan m 9/28/11 using EmailGeneric.frm for email
Private smFileAttachment As String
Private smFileAttachmentName As String
Private Const FILESET As String = "XFileSetX"
'8886
Private myLogger As CLogger

Private Sub cmdEMail_Click()
    Dim ilRet As Integer
    Dim fs As New FileSystemObject
    
    If fs.FILEEXISTS(smFileAttachment) Then
        ilRet = MsgBox("Would you like to attach ** " & smFileAttachmentName & " ** to your email?", vbYesNo)
        If ilRet = vbNo Then
            smFileAttachment = ""
        End If
    Else
        smFileAttachment = ""
    End If
    If lbcFileSelect.ListIndex >= 0 Then
        sgEMailGenericTitle = tmAffMessages(lbcFileSelect.ItemData(lbcFileSelect.ListIndex)).Name
    Else
        sgEMailGenericTitle = ""
    End If
    Set ogEmailer = New CEmail
    ogEmailer.Attachment = smFileAttachment
    EmailGeneric.isCounterpointService = True
    EmailGeneric.isZipAttachment = True
    EmailGeneric.Show vbModal
    Set ogEmailer = Nothing
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload frmMessages
End Sub

Private Sub Form_Load()

    gCenterForm frmMessages
    LoadlbcFileSelect
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '8886
    Set myLogger = Nothing
    Erase tmAffMessages
    Set frmMessages = Nothing
    
End Sub

Private Sub imcPrt_Click()
    Dim ilCurrentLineNo As Integer
    Dim ilLinesPerPage As Integer
    Dim slRecord As String
    Dim slHeading As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    If lbcShowFile.ListCount <= 0 Then
        Exit Sub
    End If
    pbcPrinting.Visible = True
    DoEvents
    ilCurrentLineNo = 0
    ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    ilRet = 0
    On Error GoTo imcPrtErr:
   ' slHeading = "Printing " & sgFileAttachmentName & " for " & Trim$(sgUserName) & " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    slHeading = "Printing " & smFileAttachmentName & " for " & Trim$(sgUserName) & " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    '6/12/16: Replaced GoSub
    'GoSub mHeading1
    mHeader1 slHeading, ilCurrentLineNo, ilRet
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    'Output Information
    For ilLoop = 0 To lbcShowFile.ListCount - 1 Step 1
        slRecord = "    " & lbcShowFile.List(ilLoop)
        '6/12/16: Replaced GoSub
        'GoSub mLineOutput
        mLineOutput slHeading, slRecord, ilCurrentLineNo, ilLinesPerPage, ilRet
        If ilRet <> 0 Then
            Printer.EndDoc
            On Error GoTo 0
            pbcPrinting.Visible = False
            Exit Sub
        End If
    Next ilLoop
    Printer.EndDoc
    On Error GoTo 0
    'pbcPrinting.Visible = False
    tmcPrt.Enabled = True
    Exit Sub
'mHeading1:  'Output file name and date
'    Printer.Print slHeading
'    If ilRet <> 0 Then
'        Return
'    End If
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Printer.Print " "
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
'mLineOutput:
'    If ilCurrentLineNo >= ilLinesPerPage Then
'        Printer.NewPage
'        If ilRet <> 0 Then
'            Return
'        End If
'        ilCurrentLineNo = 0
'        GoSub mHeading1
'        If ilRet <> 0 Then
'            Return
'        End If
'    End If
'    Printer.Print slRecord
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
imcPrtErr:
    ilRet = Err.Number
        gMsgBox "Printing Error #  " & Str$(ilRet), vbCritical
    Resume Next
End Sub

Private Sub lbcFileSelect_Click()
    Dim slLocation As String
    'Init
    lbcUsers.Visible = False
    lbcUsers.Clear
    lbcShowFile.Clear
    slLocation = ""

    'clear the horz. scroll bar if its there
    SendMessageByNum lbcShowFile.hwnd, LB_SETHORIZONTALEXTENT, 0, 0
    'Dan M 4/25/13 added lbcUsers
    If Not mIsFileSet(lbcFileSelect.ListIndex) Then
        slLocation = sgDBPath & "Messages\" & tmAffMessages(lbcFileSelect.ItemData(lbcFileSelect.ListIndex)).fileName
        smFileAttachmentName = tmAffMessages(lbcFileSelect.ItemData(lbcFileSelect.ListIndex)).fileName
        mDisplayFile slLocation
        smFileAttachment = slLocation
    End If
End Sub
Sub LoadlbcFileSelect()

    Dim llCount As Long
    Dim ilRet As Integer
    Dim blCSVAffidate As Boolean
    Dim rst_Saf As ADODB.Recordset
    
    'D.S. Added a lot of new entries below.  Also added a test to make sure the file exists before adding to the list.
    'Note: There are currently 4 empty position.  Add to them first.  Position does not matter as they are sorted when added to the list.
          'Of course the indexes for the .Name and .filename must match.
          
    blCSVAffidate = False
    SQLQuery = "Select safFeatures1, safFeatures3, safFeatures5, safFeatures6 From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rst_Saf = gSQLSelectCall(SQLQuery, "frmMessage: LoadlbcFileSelect")
    If Not rst_Saf.EOF Then
        If (Asc(rst_Saf!safFeatures5) And CSVAFFIDAVITIMPORT) = CSVAFFIDAVITIMPORT Then
            blCSVAffidate = True
        End If
    End If
    rst_Saf.Close
    
    tmAffMessages(0).Name = "Affiliate Error Log"
    tmAffMessages(1).Name = "BIA Import Log"
    tmAffMessages(2).Name = "CPTT Fix Log"
    tmAffMessages(3).Name = "DDFReorg Log"
    tmAffMessages(4).Name = "Export CNC Spots"
    tmAffMessages(5).Name = "Export ISCI"
    tmAffMessages(6).Name = "Export Label Info"
    tmAffMessages(7).Name = "Export RCS"
    tmAffMessages(8).Name = "Stations Not Exported"
    tmAffMessages(9).Name = "Web Email Log"
    tmAffMessages(10).Name = "Web Export Log"
    tmAffMessages(11).Name = "Web Import Log"
    tmAffMessages(12).Name = "Agreements A/E Not Assigned"
    tmAffMessages(13).Name = "Update Station Info"
    tmAffMessages(14).Name = "Marketron Export Logs"
    tmAffMessages(15).Name = "Marketron Import Logs"
    tmAffMessages(16).Name = "Report Log"
    tmAffMessages(17).Name = "IDC Logs"
    tmAffMessages(18).Name = "ISCI XRef Export Log"
    tmAffMessages(19).Name = "Univision Export Log"
    tmAffMessages(20).Name = "StarGuide Export Log"
    tmAffMessages(21).Name = "Wegener Compel Export Log"
    tmAffMessages(22).Name = "X-Digital Export Log"
    tmAffMessages(22).Name = "X-Digital Logs"
    tmAffMessages(23).Name = "C and C Export Result"
    tmAffMessages(24).Name = "IDC Export Result"
    tmAffMessages(25).Name = "ISCI Export Result"
    tmAffMessages(26).Name = "ISCI XRef Export Result"
    tmAffMessages(27).Name = "Marketron Export Result"
    tmAffMessages(28).Name = "RCS Export Result"
    tmAffMessages(29).Name = "Univision Export Result"
    tmAffMessages(30).Name = "StarGuide Export Result"
    tmAffMessages(31).Name = "Wegener Export Result"
    tmAffMessages(32).Name = "X-Digital Export Result"
    tmAffMessages(33).Name = "Counterpoint Affidavit System Export Result"
    tmAffMessages(34).Name = "Email-Weekly sent mass emails"
    tmAffMessages(35).Name = "Email-Previous week sent mass emails"
    tmAffMessages(36).Name = "Email-Improperly formatted emails"
    tmAffMessages(37).Name = "Update issues"
    tmAffMessages(38).Name = "Wegener IPump Logs"
    tmAffMessages(39).Name = "Re-Import Affiliate Spots"
    tmAffMessages(40).Name = "Illegal Characters Found"
    tmAffMessages(41).Name = "Task Blocked"
    tmAffMessages(42).Name = "Pool Unassigned Log"
    tmAffMessages(43).Name = "Wegener Compel Import Log"
    tmAffMessages(44).Name = "Agreement Log"
    tmAffMessages(45).Name = "Spot Utilities Log"
    tmAffMessages(46).Name = "Affiliate Web Error Log"
    tmAffMessages(47).Name = "Ast Check Util"
    tmAffMessages(48).Name = "Clear Events"
    tmAffMessages(49).Name = "Archive"
    tmAffMessages(50).Name = "Archive Removal Detail"
    tmAffMessages(51).Name = "CSI Unzip"
    tmAffMessages(52).Name = "CSI Backup"
    tmAffMessages(53).Name = "CSI Server Log.txt"
    tmAffMessages(54).Name = "FastAddVerbose"
    tmAffMessages(55).Name = "Email PreviousWeekly Log"
    tmAffMessages(56).Name = "Export Dallas"
    tmAffMessages(57).Name = "Export Engineering"
    tmAffMessages(58).Name = "Export Matrix"
    tmAffMessages(59).Name = "Export ReRate"
    tmAffMessages(60).Name = "Export Air Wave"
    tmAffMessages(61).Name = "Export Cart"
    tmAffMessages(62).Name = "Export Gen"
    tmAffMessages(63).Name = "Fast Add Summary"
    tmAffMessages(64).Name = "Fast End Summary"
    tmAffMessages(65).Name = "FTP Log"
    tmAffMessages(66).Name = "Import Research"
    tmAffMessages(67).Name = "Import Act1"
    tmAffMessages(68).Name = "Imprt Copy"
    tmAffMessages(69).Name = "Import Radar"
    tmAffMessages(70).Name = "Label Export Log"
    tmAffMessages(71).Name = "Log Activity For CPTT"
    tmAffMessages(72).Name = "Multicast Missing"
    tmAffMessages(73).Name = "No Time Zone Stations"
    tmAffMessages(74).Name = "Not Marked Complete"
    tmAffMessages(75).Name = "Radar Export Log"
    tmAffMessages(76).Name = "Reconcile"
    tmAffMessages(77).Name = "Remote SQL Calls"
    tmAffMessages(78).Name = "Set Copy Inventory Dates"
    tmAffMessages(79).Name = "SMF Check"
    tmAffMessages(80).Name = "SSF Check"
    tmAffMessages(81).Name = "SSF Fix"
    tmAffMessages(82).Name = "Station Information"
                                                            tmAffMessages(83).Name = ""
    tmAffMessages(84).Name = "Stations Not Found"
    tmAffMessages(85).Name = "Station Spot Builder"
    tmAffMessages(86).Name = "Traffic Errors"
    tmAffMessages(87).Name = "Update Errors"
    tmAffMessages(88).Name = "Vendor Conversion"
    tmAffMessages(89).Name = "Web Activity Log"
    tmAffMessages(90).Name = "WebEmail Log"
                                                            tmAffMessages(91).Name = ""
    tmAffMessages(92).Name = "Web Export Retry Log"
    tmAffMessages(93).Name = "Web Export Summary"
                                                            tmAffMessages(94).Name = ""
                                                            tmAffMessages(95).Name = ""
    tmAffMessages(96).Name = "Web No Email Address"
    
    tmAffMessages(97).Name = ""
    'If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        If blCSVAffidate Then
            tmAffMessages(97).Name = "CSV Affidavit Import Log"
        End If
    'Else
    '    mnuImportCSVAffidavit.Visible = False
    'End If
    
    tmAffMessages(0).fileName = "AffErrorLog.Txt"
    tmAffMessages(1).fileName = "BIAImportLog.Txt"
    tmAffMessages(2).fileName = "CpttFixLog.Txt"
    tmAffMessages(3).fileName = "DDFReorg.Txt"
    tmAffMessages(4).fileName = "ExptCnCSpots.Txt"
    tmAffMessages(5).fileName = "ExptISCI.Txt"
    tmAffMessages(6).fileName = "ExptLabelInfo.Txt"
    tmAffMessages(7).fileName = "ExptRCS.Txt"
    tmAffMessages(8).fileName = "StationsNotExported.Txt"
    tmAffMessages(9).fileName = "WebEmailLog.Txt"
    tmAffMessages(10).fileName = "WebExportLog.Txt"
    tmAffMessages(11).fileName = "WebImportLog.Txt"
    tmAffMessages(12).fileName = "AgreementsAENotAssigned.Txt"
    tmAffMessages(13).fileName = "UpdateStationStatus.txt"
    tmAffMessages(14).fileName = "MarketronExportLog.txt"
    tmAffMessages(15).fileName = "MarketronImportLog.txt"
    tmAffMessages(14).fileName = FILESET & "MarketronExport"
    tmAffMessages(15).fileName = FILESET & "MarketronImport"
    tmAffMessages(16).fileName = "ReportLog.txt"
    tmAffMessages(17).fileName = FILESET & "Idc"
    tmAffMessages(17).fileName = "IdcExportLog.txt"
    tmAffMessages(18).fileName = "ISCIXRefExportLog.txt"
    tmAffMessages(19).fileName = "UnivisionExportLog.txt"
    tmAffMessages(20).fileName = "StarGuideExportLog.txt"
    tmAffMessages(21).fileName = "WegenerExportLog.txt"
    tmAffMessages(22).fileName = "XDigitalExportLog.txt"
    tmAffMessages(22).fileName = FILESET & "XDigital"
    tmAffMessages(23).fileName = "CnCResultList.txt"
    tmAffMessages(24).fileName = "IDCResultList.txt"
    tmAffMessages(25).fileName = "ISCIResultList.txt"
    tmAffMessages(26).fileName = "ISCIXRefResultList.txt"
    tmAffMessages(27).fileName = "MarketronResultList.txt"
    tmAffMessages(28).fileName = "RCSResultList.txt"
    tmAffMessages(29).fileName = "UnivisionResultList.txt"
    tmAffMessages(30).fileName = "StarGuideResultList.txt"
    tmAffMessages(31).fileName = "WegenerResultList.txt"
    tmAffMessages(32).fileName = "XDSResultList.txt"
    tmAffMessages(33).fileName = "CSIWebResultList.txt"
    tmAffMessages(34).fileName = "EmailWeeklyLog.txt"
    tmAffMessages(35).fileName = "EmailPreviousWeeklyLog.txt"
    tmAffMessages(36).fileName = "EmailFormatImproper.txt"
    tmAffMessages(37).fileName = "UpdateErrors.txt"
    tmAffMessages(38).fileName = FILESET & "iPump"
    tmAffMessages(39).fileName = "ReImportAffiliateSpots.Txt"
    tmAffMessages(40).fileName = "AffBadCharLog.Txt"
    tmAffMessages(41).fileName = "TaskBlocked_*.Txt"
    tmAffMessages(42).fileName = "PoolUnassignedLog_*.Txt"
    tmAffMessages(43).fileName = "WegenerImportResult_*.Txt"
    tmAffMessages(43).fileName = "WegenerImportResult_*.Txt"
    tmAffMessages(44).fileName = "AffAgreementLog.Txt"
    tmAffMessages(45).fileName = "AffUtilsLog.Txt"
    tmAffMessages(46).fileName = "AffWebErrorLog.txt"
    tmAffMessages(47).fileName = "ASTCheckUtility.txt"
    tmAffMessages(48).fileName = "ClearEvents.Txt"
    tmAffMessages(49).fileName = "csiArchive.txt"
    tmAffMessages(50).fileName = "csiArchiveRemovalDetail.txt"
    tmAffMessages(51).fileName = "CSIUnzip.Txt"
    tmAffMessages(52).fileName = "CSI_Backup.txt"
    tmAffMessages(53).fileName = "CSI_Server_Log.txt"
    tmAffMessages(54).fileName = "FastAddVerbose.Txt"
    tmAffMessages(55).fileName = "EmailPreviousWeekly_Log.txt"
    tmAffMessages(56).fileName = "ExpDall.Txt"
    tmAffMessages(57).fileName = "ExpEngr.Txt"
    tmAffMessages(58).fileName = "ExpMatrix.Txt"
    tmAffMessages(59).fileName = "ExportReRate.txt"
    tmAffMessages(60).fileName = "ExptAirWave.Txt"
    tmAffMessages(61).fileName = "ExptCart.Txt"
    tmAffMessages(62).fileName = "ExptGen.Txt"
    tmAffMessages(63).fileName = "FastAddSummary.txt"
    tmAffMessages(64).fileName = "FastEndSummary.Txt"
    tmAffMessages(65).fileName = "FTPLog.txt"
    tmAffMessages(66).fileName = "ImportResearch.Txt"
    tmAffMessages(67).fileName = "ImptAct1.Txt"
    tmAffMessages(68).fileName = "ImptCopy.Txt"
    tmAffMessages(69).fileName = "ImptRad.Txt"
    tmAffMessages(70).fileName = "LableExportLog.Txt"
    tmAffMessages(71).fileName = "LogActivityFor_CPTT.txt"
    tmAffMessages(72).fileName = "MulticastMissing.txt"
    tmAffMessages(73).fileName = "NoTimeZoneStations.Txt"
    tmAffMessages(74).fileName = "NotMarkedComplete.Txt"
    tmAffMessages(75).fileName = "RadarExportLog.Txt"
    tmAffMessages(76).fileName = "Reconcile.Txt"
    tmAffMessages(77).fileName = "RemoteSQLCalls.txt"
    tmAffMessages(78).fileName = "SetCopyInventoryDates.txt"
    tmAffMessages(79).fileName = "SMFCheck.Txt"
    tmAffMessages(80).fileName = "SSFCheck.Txt"
    tmAffMessages(81).fileName = "SSFFix.Txt"
    tmAffMessages(82).fileName = "StationInformation.Txt"
                                                            tmAffMessages(83).fileName = ""
    tmAffMessages(84).fileName = "StationsNotFound.Txt"
    tmAffMessages(85).fileName = "StationSpotBuilder.txt"
    tmAffMessages(86).fileName = "TrafficErrors.Txt"
    tmAffMessages(87).fileName = "UpdateErrors.txt"
    tmAffMessages(88).fileName = "VendorConversion.txt"
    tmAffMessages(89).fileName = "WebActivityLog.Txt"
    tmAffMessages(90).fileName = "WebEmailLog.Txt"
                                                            tmAffMessages(91).fileName = ""
    tmAffMessages(92).fileName = "WebExpRetryLog.txt"
    tmAffMessages(93).fileName = "WebExpSummary.Txt"
                                                            tmAffMessages(94).fileName = ""
                                                            tmAffMessages(95).fileName = ""
    tmAffMessages(96).fileName = "WebNoEmailAddress.Txt"
    tmAffMessages(97).fileName = "WebNoEmailAddress.Txt"
    
    'If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        If blCSVAffidate Then
            tmAffMessages(99).fileName = "CSVAffidavitImportLog.Txt"
        End If
    'Else
    '    mnuImportCSVAffidavit.Visible = False
    'End If

    For llCount = 0 To UBound(tmAffMessages)
        If Trim$(tmAffMessages(llCount).fileName) <> "" Then
            'Only add to the list if the file exists in their Messages folder
            ilRet = gFileExist(sgMsgDirectory & Trim$(tmAffMessages(llCount).fileName))
            If ilRet = 0 Then
                lbcFileSelect.AddItem tmAffMessages(llCount).Name
                lbcFileSelect.ItemData(lbcFileSelect.NewIndex) = llCount
            End If
        End If
    Next
    
    ilRet = LoadWegenerCheckLogFile() ' JD 02-09-24 added
    
End Sub
Function LoadWegenerCheckLogFile() As Integer
' JD 02-09-24 Added support for the wegener check utility
' These log files have a date/time stamp and we want the latest one in the folder.
Dim CurrFile As String
Dim CurrDate As Date
Dim NewestDate As Date
Dim NewestFile As String

CurrFile = Dir$(sgMsgDirectory & "WegenerUtility_*.txt", vbNormal)
Do While Len(CurrFile) > 0
    CurrDate = FileDateTime(sgMsgDirectory & CurrFile)
    If CurrDate > NewestDate Then
        NewestDate = CurrDate
        NewestFile = CurrFile
    End If
    CurrFile = Dir$()
Loop
If Len(NewestFile) > 0 Then
    tmAffMessages(98).fileName = NewestFile
    tmAffMessages(98).Name = "Wegner Check Utilty"

    lbcFileSelect.AddItem "Wegner Check Utilty"
    lbcFileSelect.ItemData(lbcFileSelect.NewIndex) = 98
End If
End Function
Private Sub mInit()
    '8886
   ' Dim myLogger As CLogger

    imcPrt.Picture = frmDirectory!imcPrinter.Picture
'    sgFileAttachment = ""
    sgFileAttachment = ""
    pbcPrinting.Move lbcShowFile.Left + lbcShowFile.Width / 2 - pbcPrinting.Width / 2, lbcShowFile.Top + lbcShowFile.Height / 2 - pbcPrinting.Height / 2
    gAdjustScreenMessage Me, pbcPrinting
    Set myLogger = New CLogger
    With myLogger
        .CleanThisFolder = messages
       ' 8462 save for 90 days
'        .CleanFolder , , 60
       .CleanFolder
    End With
   ' Set myLogger = Nothing
End Sub

Private Sub lbcUsers_Click()
    Dim slLocation As String

    lbcShowFile.Clear
    smFileAttachmentName = lbcUsers.Text
    slLocation = sgDBPath & "Messages\" & smFileAttachmentName
    smFileAttachment = slLocation
    mDisplayFile slLocation
    
End Sub

Private Sub mDisplayFile(sLocation As String)

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim llRet As Long
    Dim slTemp As String
    Dim slRetString As String
    Dim llMaxWidth As Long
    Dim llValue As Long
    Dim llRg As Long
    Dim slCurDir As String
    
    slCurDir = CurDir
    frmMessages.pbcArial.Width = 8925
    'Make Sure we start out each time without a horizontal scroll bar
    llValue = 0
    If imNewDisplay Then
        llRet = SendMessageByNum(lbcShowFile.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llMaxWidth = 0
    If fs.FILEEXISTS(sLocation) Then
        Set tlTxtStream = fs.OpenTextFile(sLocation, ForReading, False)
    Else
        lbcShowFile.Clear
        lbcShowFile.AddItem "** No Data Available **"
        sgFileAttachment = ""
        Exit Sub
    End If
    slTemp = ""

    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        lbcShowFile.AddItem slRetString
        If (frmMessages.pbcArial.TextWidth(slRetString)) > llMaxWidth Then
            llMaxWidth = (frmMessages.pbcArial.TextWidth(slRetString))
        End If
    Loop

    'Show a horzontal scroll bar if needed
    If llMaxWidth > lbcShowFile.Width Then
        llValue = llMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcShowFile.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    imNewDisplay = False
    tlTxtStream.Close
    ChDir slCurDir
    
End Sub

Private Sub lbcUsers_GotFocus()

    tmcUsers.Enabled = False
    
End Sub

Private Sub lbcUsers_Scroll()

    tmcUsers.Enabled = False
    tmcUsers.Enabled = True
    
End Sub

Private Sub pbcPrinting_Paint()

    pbcPrinting.CurrentX = (pbcPrinting.Width - pbcPrinting.TextWidth("Printing Message Information....")) / 2
    pbcPrinting.CurrentY = (pbcPrinting.Height - pbcPrinting.TextHeight("Printing Message Information....")) / 2 - 30
    pbcPrinting.Print "Printing Message Information...."
    
End Sub

Private Sub tmcPrt_Timer()

    tmcPrt.Enabled = False
    pbcPrinting.Visible = False
    
End Sub

Private Sub tmcUsers_Timer()

    tmcUsers.Enabled = False
    imNewDisplay = True
End Sub
Private Function mIsFileSet(ilIndex As Integer) As Boolean
    Dim blRet As Boolean
    Dim ilPos As Integer
    Dim slFiles As String
   ' Dim slFileName As String
    Dim slFolder As String
    Dim slMainName As String
    '8886
    Dim myFolder As Folder
    Dim myFile As file
    
    blRet = False
    slMainName = tmAffMessages(lbcFileSelect.ItemData(ilIndex)).fileName
    ilPos = InStr(1, slMainName, FILESET)
    If ilPos > 0 Then
        blRet = True
        slFolder = sgDBPath & "Messages\"
        slFiles = Mid(slMainName, ilPos + Len(FILESET))
        '8886
'        slFileName = ""
'        slFileName = Dir(slFolder & slFiles & "*LOG_??-??-??.txt")
'        Do While slFileName > ""
'            lbcUsers.AddItem slFileName
'            slFileName = Dir()
'        Loop
        Set myFolder = myLogger.myFile.GetFolder(slFolder)
        For Each myFile In myFolder.Files()
            If myLogger.IsLogFile(myFile.Name) Then
                If InStr(1, myFile.Name, slFiles, vbTextCompare) Then
                    lbcUsers.AddItem myFile.Name
                End If
            End If
        Next
        lbcUsers.Visible = True
    End If
    ilPos = InStr(1, slMainName, "TaskBlocked_")
    If ilPos > 0 Then
        blRet = True
        slFolder = sgDBPath & "Messages\"
        '8886
'        slFileName = Dir(slFolder & slMainName)
'        Do While slFileName > ""
'            lbcUsers.AddItem slFileName
'            slFileName = Dir()
'        Loop
        Set myFolder = myLogger.myFile.GetFolder(slFolder)
        For Each myFile In myFolder.Files()
            If InStr(1, myFile.Name, "TaskBlocked_", vbTextCompare) Then
                lbcUsers.AddItem myFile.Name
            End If
        Next
        lbcUsers.Visible = True
    End If
    ilPos = InStr(1, slMainName, "PoolUnassignedLog_")
    If ilPos > 0 Then
        blRet = True
        slFolder = sgDBPath & "Messages\"
        '8886
'        slFileName = Dir(slFolder & slMainName)
'        Do While slFileName > ""
'            lbcUsers.AddItem slFileName
'            slFileName = Dir()
'        Loop
        Set myFolder = myLogger.myFile.GetFolder(slFolder)
        For Each myFile In myFolder.Files()
            If InStr(1, myFile.Name, "PoolUnassignedLog_", vbTextCompare) Then
                lbcUsers.AddItem myFile.Name
            End If
        Next
        lbcUsers.Visible = True
    End If
    ilPos = InStr(1, slMainName, "WegenerImportResult_")
    If ilPos > 0 Then
        blRet = True
        slFolder = sgDBPath & "Messages\"
        '8886
'        slFileName = Dir(slFolder & slMainName)
'        Do While slFileName > ""
'            lbcUsers.AddItem slFileName
'            slFileName = Dir()
'        Loop
        Set myFolder = myLogger.myFile.GetFolder(slFolder)
        For Each myFile In myFolder.Files()
            If InStr(1, myFile.Name, "WegenerImportResult_", vbTextCompare) Then
                lbcUsers.AddItem myFile.Name
            End If
        Next
        lbcUsers.Visible = True
    End If
    mIsFileSet = blRet

End Function


Private Sub mLineOutput(slHeading As String, slRecord As String, ilCurrentLineNo As Integer, ilLinesPerPage As Integer, ilRet As Integer)
    On Error GoTo imcPrtErr:
    If ilCurrentLineNo >= ilLinesPerPage Then
        Printer.NewPage
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
        ilCurrentLineNo = 0
        mHeader1 slHeading, ilCurrentLineNo, ilRet
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
    End If
    Printer.Print slRecord
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
imcPrtErr:
    ilRet = Err.Number
    gMsgBox "Printing Error #  " & Str$(ilRet), vbCritical
    Resume Next
End Sub

Private Sub mHeader1(slHeading As String, ilCurrentLineNo As Integer, ilRet As Integer)
    On Error GoTo imcPrtErr:
    Printer.Print slHeading
    If ilRet <> 0 Then
        'Return
        Exit Sub
    End If
    ilCurrentLineNo = ilCurrentLineNo + 1
    Printer.Print " "
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
imcPrtErr:
    ilRet = Err.Number
    gMsgBox "Printing Error #  " & Str$(ilRet), vbCritical
    Resume Next
End Sub
