VERSION 5.00
Begin VB.Form Messages 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
   ClipControls    =   0   'False
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmcPrt 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7485
      Top             =   5190
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
      TabIndex        =   8
      Top             =   2430
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer tmcUsers 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1260
      Top             =   5295
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
      TabIndex        =   7
      Top             =   5325
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      TabIndex        =   6
      Top             =   1350
      Visible         =   0   'False
      Width           =   1800
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
      ItemData        =   "Messages.frx":0000
      Left            =   4920
      List            =   "Messages.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   405
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
      ItemData        =   "Messages.frx":0004
      Left            =   240
      List            =   "Messages.frx":0006
      TabIndex        =   3
      Top             =   1680
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
      ItemData        =   "Messages.frx":0008
      Left            =   240
      List            =   "Messages.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   405
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
      Top             =   5175
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
      Top             =   5175
      Width           =   2010
   End
   Begin VB.Image imcPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8415
      Picture         =   "Messages.frx":000C
      Top             =   5025
      Width           =   480
   End
   Begin VB.Label lblHeader 
      Caption         =   "Message Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   315
      TabIndex        =   4
      Top             =   60
      Width           =   4200
   End
End
Attribute VB_Name = "Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Messages.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Private imAllClicked As Integer
Private imSetAll As Integer
Private smChoice As String
Private imNewDisplay As Integer
Private smFileAttachment As String
Private smFileAttachmentName As String
'8723
Private bmIsFileSet As Boolean
Private Const FILESET As String = "XFileSetX"
'8886
Private myLogger As CLogger
Private bmIsWebAPI As Boolean

Private Sub ckcAll_Click()

    Dim Value As Integer
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer

    lbcShowFile.Clear
    tmcUsers.Enabled = False
    tmcUsers.Enabled = True

    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    Else
        lbcShowFile.Clear
    End If

    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If lbcUsers.ListCount > 0 Then
            llRg = CLng(lbcUsers.ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcUsers.hWnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
    End If

End Sub

Private Sub cmdEmail_Click()

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slFileName As String
    Dim fs As New FileSystemObject
    Dim tlTxtStream As TextStream
    'Dan M 9/16/09 smFileAttachment blank for invoice and logs and files that don't exist.  Find invoicing/logs if lbcusers visible and at least one selected
    If smFileAttachment <> "" Or lbcUsers.SelCount > 0 Then
  '  If smFileAttachment <> "" Then
        ilRet = MsgBox("Would you like to attach ** " & smFileAttachmentName & " ** to your email?", vbYesNo)
        If ilRet = vbNo Then
            smFileAttachment = ""
        Else
            'Logs and Invoicing are a special case.  Each users has their own file
            'so we need to be able to build a single temp file from one or many files.
            If smChoice = "Logs" Or smChoice = "Invoicing" Then
                Select Case smChoice
                '5676
                    Case "Logs"
                        'slFileName = "C:\CSI\Logs.txt"
                        slFileName = sgRootDrive & "CSI\Logs.txt"
                    Case "Invoicing"
                        'slFileName = "C:\CSI\InvError.txt"
                        slFileName = sgRootDrive & "CSI\InvError.txt"
                End Select

                'If the temp file is already there it gets overwritten
                fs.CreateTextFile slFileName, True, False
                Set tlTxtStream = fs.OpenTextFile(slFileName, ForWriting, False, TristateFalse)

                For ilLoop = 0 To lbcShowFile.ListCount - 1 Step 1
                    tlTxtStream.WriteLine (lbcShowFile.List(ilLoop))
                Next ilLoop
                tlTxtStream.Close
                smFileAttachment = slFileName
            End If
            If smChoice = "Web API" Then
                smFileAttachment = sgDBPath & "Messages\WebAPI\" & smFileAttachment
            End If
        End If
    End If
    'Dan 9/6/11 replace with generic email
    'PBEmail.Show vbModal
    Set ogEmailer = New CEmail
    ogEmailer.Attachment = smFileAttachment
    EmailGeneric.isCounterpointService = True
    EmailGeneric.isZipAttachment = True
    EmailGeneric.Show vbModal
    Set ogEmailer = Nothing
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload Messages
End Sub

Private Sub Form_Load()
    gCenterStdAlone Messages
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '8886
    Set myLogger = Nothing

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
    ilLinesPerPage = (Printer.height - 1440) / Printer.TextHeight("TEST") - 1
    ilRet = 0
    On Error GoTo imcPrtErr:
    slHeading = "Printing " & smFileAttachmentName & " for " & Trim$(tgUrf(0).sRept) & " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    '6/6/16: Replaced GoSub
    'GoSub mHeading1
    mHeading1 ilRet, slHeading, ilCurrentLineNo
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    'Output Information
    For ilLoop = 0 To lbcShowFile.ListCount - 1 Step 1
        slRecord = "    " & lbcShowFile.List(ilLoop)
        '6/6/16: Replaced GoSub
        'GoSub mLineOutput
        mLineOutput ilRet, slHeading, ilCurrentLineNo, slRecord, ilLinesPerPage
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
'        '6/6/16: Replaced GoSub
'        'GoSub mHeading1
'        mHeading1 ilRet, slHeading, ilCurrentLineNo
'        If ilRet <> 0 Then
'            Return
'        End If
'    End If
'    Printer.Print slRecord
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
imcPrtErr:
    ilRet = err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub

Private Sub lbcFileSelect_Click()

    Dim ilLoop As Integer
    Dim slLocation As String
    Dim llRet As Long
    Dim ilPos As Integer

    'Init
    lbcUsers.Visible = False
    ckcAll.Visible = False
    ckcAll.Value = vbUnchecked
    lbcShowFile.Clear
    lbcUsers.Clear
    slLocation = ""
    bmIsWebAPI = False
    
    'clear the horz. scroll bar if its there
    llRet = SendMessageByNum(lbcShowFile.hWnd, LB_SETHORIZONTALEXTENT, 0, 0)

    'Find out which group was selected; adv, logs, inv. etc.
    For ilLoop = 0 To lbcFileSelect.ListCount - 1 Step 1
        If lbcFileSelect.Selected(ilLoop) Then
            smChoice = Trim$(lbcFileSelect.Text)
            Exit For
        End If
    Next ilLoop

    Select Case smChoice
        Case "Advertiser Merge"
            slLocation = sgDBPath & "Messages\MergeAdf.txt"
            smFileAttachmentName = "Advertiser Merge"
        Case "Advertiser Product Merge"
            slLocation = sgDBPath & "Messages\MergeProd.txt"
            smFileAttachmentName = "Advertiser Product Merge"
        Case "Affiliate - Aired Spot Import"
            slLocation = sgDBPath & "Messages\ImptAiredSpots.txt"
            smFileAttachmentName = "Affiliate - Aired Spot Import"
        Case "Affiliate - Scheduled Spot Export"
            slLocation = sgDBPath & "Messages\ExptSchdSpots.txt"
            smFileAttachmentName = "Affiliate - Scheduled Spot Export"
        Case "Agency Merge"
            slLocation = sgDBPath & "Messages\MergeAgf.txt"
            smFileAttachmentName = "Agency Merge"
        Case "Background Schedule Messages"
            slLocation = sgDBPath & "Messages\BkgdSchdErrors.txt"
            smFileAttachmentName = "Background Schedule Messages"
        Case "Backup"
            slLocation = sgDBPath & "Messages\Backup.txt"
            smFileAttachmentName = "Backup"
        Case "Contract Swap Check"
            slLocation = sgDBPath & "Messages\ContractSwapCheck.Csv"
            smFileAttachmentName = "Contract Swap Check"
        Case "CSIStart"
            slLocation = sgRootDrive & "CSI\CSIStart.txt"
            smFileAttachmentName = "CSIStart"
        Case "CSISetup"
            ilPos = InStr(1, sgDBPath, "\Prod", vbTextCompare)
            If ilPos > 0 Then
                slLocation = Left$(sgDBPath, ilPos) & "CSISetup.txt"
            Else
                ilPos = InStr(1, sgDBPath, "\Test", vbTextCompare)
                If ilPos > 0 Then
                    slLocation = Left$(sgDBPath, ilPos) & "CSISetup.txt"
                Else
                    ilPos = InStr(1, sgDBPath, "\Dev", vbTextCompare)
                    If ilPos > 0 Then
                        slLocation = Left$(sgDBPath, ilPos) & "CSISetup.txt"
                    End If
                End If
            End If
            smFileAttachmentName = "CSISetup"
        Case "CSIUnzip"
            slLocation = sgDBPath & "Messages\CSIUnzip.txt"
            smFileAttachmentName = "CSIUnzip"
        Case "DBUnzip"
            slLocation = sgDBPath & "Messages\DBUnzip.txt"
            smFileAttachmentName = "DBUnzip"
        Case "DDFReorg"
            slLocation = sgDBPath & "Messages\DDFREorg.txt"
            smFileAttachmentName = "DDFReorg"
        Case "Export - Automation"
            slLocation = sgDBPath & "Messages\ExptGen.Txt"
            smFileAttachmentName = "Export - Automation"
        Case "Export - Dallas"
            slLocation = sgDBPath & "Messages\ExpDall.Txt"
            smFileAttachmentName = "Export - Dallas"
        Case "Export - ENCO"
            'slLocation = sgDBPath & "Messages\ExptEnco.Txt"
            slLocation = FILESET & "ExptEnco*"
            smFileAttachmentName = "Export - Enco"
        Case "Export - Network Inventory"
            slLocation = sgDBPath & "Messages\ExpCncNI.Txt"
            smFileAttachmentName = "Export - Network Inventory"
        Case "Export - Engineering Feed"
            slLocation = sgDBPath & "Messages\ExpEngr.Txt"
            smFileAttachmentName = "Export - Engineering"
        Case "Export - Phoenix"
            slLocation = sgDBPath & "Messages\ExpPhnx.Txt"
            smFileAttachmentName = "Export - Phoenix"
        Case "Export - Scheduled Spots"
            slLocation = sgDBPath & "Messages\ExpCncSS.Txt"
            smFileAttachmentName = "Export - Scheduled Spots"
        Case "Export - Audio MP2"
            slLocation = sgDBPath & "Messages\ExptMP2.Txt"
            smFileAttachmentName = "Export - Audio MP2"
        Case "Export - Great Plains"
            slLocation = sgDBPath & "Messages\ExportGreatPlains.Txt"
            smFileAttachmentName = "Export - Great Plains"
        Case "Export - Get Paid"
            slLocation = sgDBPath & "Messages\ExportGetPaid.Txt"
            smFileAttachmentName = "Export - Get Paid"
        Case "Export - Invoice"
            slLocation = sgDBPath & "Messages\ExportInvoice.Txt"
            smFileAttachmentName = "Export - Invoice"
        Case "File Check"
            slLocation = sgDBPath & "Messages\BtrCheck.txt"
            smFileAttachmentName = "File Check"
        Case "File Fix"
            slLocation = sgDBPath & "Messages\BtrFix.txt"
            smFileAttachmentName = "File Fix"
        Case "Import - Automation"
            slLocation = sgDBPath & "Messages\ImptGen.txt"
            smFileAttachmentName = "Import - Automation"
        Case "Import - Act1"
            slLocation = sgDBPath & "Messages\ImptAct1.txt"
            smFileAttachmentName = "Import - Act1"
        Case "Import - Radar"
            slLocation = sgDBPath & "Messages\ImptRad.txt"
            smFileAttachmentName = "Import - Radar"
        Case "Import - Satellite"
            slLocation = sgDBPath & "Messages\ImptSat.txt"
            smFileAttachmentName = "Import - Satellite"
        Case "Import - CSV"
            slLocation = sgDBPath & "Messages\ImptCSV.txt"
            smFileAttachmentName = "Import - CSV"
        '6/5/14: Add Import vehicle
        Case "Import - Vehicle"
            slLocation = sgDBPath & "Messages\ImptVeh.txt"
            smFileAttachmentName = "Import - Vehicle"
        Case "Invoice Check"
            slLocation = sgDBPath & "Messages\InvCheck.csv"
            smFileAttachmentName = "Invoice Check"
        Case "Invoicing"
            lbcUsers.Visible = True
            ckcAll.Visible = True
            mPopInvoiceListBox Messages, lbcUsers, "InvError"
            smFileAttachmentName = "Invoicing"
        Case "Logs"
            lbcUsers.Visible = True
            ckcAll.Visible = True
            mPopInvoiceListBox Messages, lbcUsers, "Logs"
            smFileAttachmentName = "Logs"
        Case "Export - Matrix"
            slLocation = sgDBPath & "Messages\ExpMatrix.txt"
            smFileAttachmentName = "Matrix Export"
        Case "Export - Corporate"
            slLocation = sgDBPath & "Messages\Corporate Export.Txt"
            smFileAttachmentName = "Export - Corporate"
        Case "Reallocation"
            slLocation = sgDBPath & "Messages\Realloc.Txt"
            smFileAttachmentName = "Reallocation"
        Case "Rep-Net Link"
            slLocation = sgDBPath & "Messages\RepNetLink.Txt"
            smFileAttachmentName = "Rep-Net Link"
        Case "Salespeople Merge"
            slLocation = sgDBPath & "Messages\MergeSlf.txt"
            smFileAttachmentName = "Salespeople Merge"
        Case "Set Credit Messages"
            slLocation = sgDBPath & "Messages\SetCreditErrors.txt"
            smFileAttachmentName = "Set Credit Messages"
        Case "Set Copy Inventory Dates"
            slLocation = sgDBPath & "Messages\SetCopyInventoryDates.txt"
            smFileAttachmentName = "Set Copy Inventory Dates"
        Case "SMFCheck"
            slLocation = sgDBPath & "Messages\SMFCheck.txt"
            smFileAttachmentName = "SMFCheck"
        Case "SSFCheck"
            slLocation = sgDBPath & "Messages\SSFCheck.txt"
            smFileAttachmentName = "SSFCheck"
        Case "SSFFix"
            slLocation = sgDBPath & "Messages\SSFFix.txt"
            smFileAttachmentName = "SSFFix"
        Case "Station Feed - Copy Export"
            slLocation = sgDBPath & "Messages\StnFdCpy.txt"
            smFileAttachmentName = "Station Feed - Copy Export"
        Case "Station Feed - Export All Instructions"
            slLocation = sgDBPath & "Messages\ExpInst.txt"
            smFileAttachmentName = "Station Feed - Export All Instructions"
        Case "Station Feed - Export All Spots"
            slLocation = sgDBPath & "Messages\ExpAlSpt.txt"
            smFileAttachmentName = "Station Feed - Export All Spots"
        Case "Station Feed - Export Regional Spots"
            slLocation = sgDBPath & "Messages\ExpRgSpt.txt"
            smFileAttachmentName = "Station Feed - Export Regional Spots"
        Case "Station Feed - Export Station Feed"
            slLocation = sgDBPath & "Messages\ExpStnFd.txt"
            smFileAttachmentName = "Station Feed - Export Station Feed"
        Case "Traffic Messages"
            slLocation = sgDBPath & "Messages\TrafficErrors.txt"
            smFileAttachmentName = "Traffic Messages"
        Case "vCreative"
            slLocation = sgDBPath & "Messages\vCreativeLog.txt"
            smFileAttachmentName = "vCreative"
        Case "Vehicle Merge"
            slLocation = sgDBPath & "Messages\MergeVef.txt"
            smFileAttachmentName = "Vehicle Merge"
        Case "Reconcile"
            slLocation = sgDBPath & "Messages\Reconcile.txt"
            smFileAttachmentName = "Reconcile"
        Case "Sport Contract Check"
            slLocation = sgDBPath & "Messages\SportChk.csv"
        Case "Stations Not Found - Import Region"
            slLocation = sgDBPath & "Messages\StationsNotFound.Txt"
            smFileAttachmentName = "Stations Not Found - Import Regions"
        Case "Update issues"
            slLocation = sgDBPath & "Messages\UpdateErrors.txt"
            smFileAttachmentName = "Update Issues"
        Case "Export - Sales Force"
            slLocation = sgDBPath & "Messages\ExportSalesForce.txt"
            smFileAttachmentName = "Export - Sales Force"
        Case "Export - Efficio"
            slLocation = sgDBPath & "Messages\EfficioExport.txt"
            smFileAttachmentName = "Export - Efficio"
        Case "Export - Tableau"
            slLocation = sgDBPath & "Messages\ExpTableau.txt"
            smFileAttachmentName = "Tableau Export"
        '8723
        Case "Filemaker To Orders Issues"
            slLocation = FILESET & "FilemakerToOrdersIssues"
            smFileAttachmentName = "Filemaker To Order Issues"
        Case "Set Credit"
            slLocation = sgDBPath & "Messages\SetCredit.txt"
            smFileAttachmentName = "Set Credit"
        Case "Days to Pay"
            slLocation = sgDBPath & "Messages\AvgToPay.csv"
            smFileAttachmentName = "Days to Pay"
        Case "Pool Unassigned Log"
            slLocation = FILESET & "PoolUnassignedLog_"
            smFileAttachmentName = "Pool Unassigned Log"
        Case "Export - RAB"                                     '1-30-20
            slLocation = sgDBPath & "Messages\ExpRAB.txt"
            smFileAttachmentName = "RAB Export"
        Case "Export - CRE"                                     '1-11-21
            slLocation = sgDBPath & "Messages\ExpCustomRevenueExport.txt"
            smFileAttachmentName = "Custom Revenue Export"
        Case "Log Affiliate Posted"
            slLocation = sgDBPath & "Messages\LogAffiliatePosted.txt"
            smFileAttachmentName = "Log Affiliate Posted"
        Case "Over-Delivered CPM Impressions"
            slLocation = FILESET & "ContractOverDelivered_*"
            smFileAttachmentName = "Over-Delivered CPM Impressions"
        Case "Web API"  ' 7-19-23
            bmIsWebAPI = True
            slLocation = sgDBPath & "Messages\WebAPI"
            lbcUsers.Visible = True
            ckcAll.Visible = False
            mLoadWebAPIFiles (slLocation)
            bmIsFileSet = True
            smFileAttachmentName = "Web API Log"
            
    End Select

    'Logs and invoicing are special cases; all the rest are handled from here
    If smChoice <> "Logs" And smChoice <> "Invoicing" And smChoice <> "Web API" Then
        lbcShowFile.Clear
        '8723--allow logs with different dates
        If Not mIsFileSet(slLocation) Then
            'Dan M 9/16/09 reversed order-- mdisplay file will blank out smFileAttachment if it doesn't exist
            smFileAttachment = slLocation
            bmIsFileSet = False
            mDisplayFile slLocation
        End If
    End If
End Sub
Private Sub mLoadWebAPIFiles(sPath As String)
    Dim myFolder As Folder
    Dim MyFile As file

    Set myFolder = myLogger.MyFile.GetFolder(sPath)
    
    For Each MyFile In myFolder.Files()
        lbcUsers.AddItem MyFile.Name
    Next
End Sub
Private Function mPopInvoiceListBox(frm As Form, lbcCtrl As control, slFileName As String) As Integer
'
'   ilRet = mPopInvoiceListBox (MainForm, cbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       cbcCtrl (I)- List box control that will be populated with names
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    Dim ilRecLen As Integer     'URF record length
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilRet As Integer
    Dim slLocations As String
    Dim fs As New FileSystemObject

    hlUrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mPopInvoiceListBoxErr
    gBtrvErrorMsg ilRet, "mPopInvoiceListBox (btrOpen):" & "Urf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlUrf)  'Get and save record length
    lbcCtrl.Clear

    ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        gUrfDecrypt tlUrf
        If (tlUrf.sDelete <> "Y") Then
            gFindMatch tlUrf.sRept, 0, lbcCtrl
            If gLastFound(lbcCtrl) < 0 Then
                If tlUrf.iCode > 1 Then
                    slLocations = sgDBPath & "Messages\" & slFileName & CStr(tlUrf.iCode) & ".txt"
                    If fs.FILEEXISTS(slLocations) Then
                        If Trim$(tlUrf.sRept) <> "" Then
                            lbcCtrl.AddItem " " & Trim$(tlUrf.sRept)
                        Else
                            lbcCtrl.AddItem " " & Trim$(tlUrf.sName)
                        End If
                        lbcCtrl.ItemData(lbcCtrl.NewIndex) = tlUrf.iCode
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        On Error GoTo mPopInvoiceListBoxErr
        gBtrvErrorMsg ilRet, "mPopInvoiceListBox (btrGetFirst):" & "Urf.Btr", frm
        On Error GoTo 0
    End If
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Function
mPopInvoiceListBoxErr:
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    gDbg_HandleError "Messages: mPopInvoiceListBox"
End Function


Private Sub mInit()
    Dim fs As New FileSystemObject
    Dim slFileName As String
    
    slFileName = sgDBPath & "Messages\WebAPI"   ' 7-19-23
    If Not fs.FolderExists(slFileName) Then
        fs.CreateFolder (slFileName)
    End If

    imcPrt.Picture = IconTraf!imcPrinter.Picture    'IconTraf!imcCamera.Picture
    smFileAttachment = ""
    imAllClicked = False
    imSetAll = True
    
    mPopulateMessageNames
    
    pbcPrinting.Move lbcShowFile.Left + lbcShowFile.Width / 2 - pbcPrinting.Width / 2, lbcShowFile.Top + lbcShowFile.height / 2 - pbcPrinting.height / 2
    gAdjustScreenMessage Me, pbcPrinting
    '8886 match affiliate, clean folder
    Set myLogger = New CLogger
    myLogger.CleanThisFolder = FoldersToClean.Messages
    myLogger.CleanFolder
    
End Sub

Private Sub mHandleInvOrLogErrFiles()

    Dim ilLoop As Integer
    Dim slLocation As String
    Dim slBaseFileName As String
    Dim fs As New FileSystemObject

    Select Case smChoice
        Case "Invoicing"
            slBaseFileName = "InvError"
            smFileAttachment = "Invoicing"
        Case "Logs"
            slBaseFileName = "Logs"
            smFileAttachment = "Logs"
    End Select

    lbcShowFile.Clear
    'Show the slelected user's name then display their file
    For ilLoop = 0 To lbcUsers.ListCount - 1 Step 1
        If lbcUsers.Selected(ilLoop) Then
            slLocation = sgDBPath & "Messages\" & slBaseFileName & lbcUsers.ItemData(ilLoop) & ".txt"
            lbcShowFile.AddItem "*** " & Trim$(lbcUsers.List(ilLoop)) & " ***"
            If fs.FILEEXISTS(slLocation) Then
                mDisplayFile slLocation
            End If
        End If
    Next ilLoop

End Sub

Private Sub lbcUsers_Click()

    Dim ilLoop As Integer
    Dim slLocation As String
    
    lbcShowFile.Clear
    If bmIsWebAPI Then  ' 7-19-23
        slLocation = sgDBPath & "Messages\WebAPI\" & lbcUsers.Text
        smFileAttachment = lbcUsers.Text
        mDisplayFile slLocation
        Exit Sub
    End If
    
    '8723
    If bmIsFileSet Then
        smFileAttachmentName = lbcUsers.Text
        slLocation = sgDBPath & "Messages\" & smFileAttachmentName
        smFileAttachment = slLocation
        mDisplayFile slLocation
    Else
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked
            imSetAll = True
        End If
    
        For ilLoop = 0 To lbcUsers.ListCount - 1 Step 1
            If lbcUsers.Selected(ilLoop) Then
                tmcUsers.Enabled = False
                tmcUsers.Enabled = True
                Exit For
            End If
        Next ilLoop
    End If

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

    Messages.pbcArial.Width = 8925
    'Make Sure we start out each time without a horizontal scroll bar
    llValue = 0
    If imNewDisplay Then
        llRet = SendMessageByNum(lbcShowFile.hWnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llMaxWidth = 0
    If fs.FILEEXISTS(sLocation) Then
        Set tlTxtStream = fs.OpenTextFile(sLocation, ForReading, False)
    Else
        lbcShowFile.Clear
        lbcShowFile.AddItem "** No Data Available **"
        smFileAttachment = ""
        Exit Sub
    End If
    slTemp = ""

    Screen.MousePointer = vbHourglass
    DoEvents
    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        lbcShowFile.AddItem slRetString
        If (Messages.pbcArial.TextWidth(slRetString)) > llMaxWidth Then
            llMaxWidth = (Messages.pbcArial.TextWidth(slRetString))
        End If
    Loop
    
    'Show a horzontal scroll bar if needed
    If llMaxWidth > lbcShowFile.Width Then
        llValue = llMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcShowFile.hWnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    Screen.MousePointer = vbDefault
    imNewDisplay = False
    tlTxtStream.Close
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
    pbcPrinting.CurrentY = (pbcPrinting.height - pbcPrinting.TextHeight("Printing Message Information....")) / 2 - 30
    pbcPrinting.Print "Printing Message Information...."
End Sub

Private Sub tmcPrt_Timer()
    tmcPrt.Enabled = False
    pbcPrinting.Visible = False
End Sub

Private Sub tmcUsers_Timer()
    tmcUsers.Enabled = False
    imNewDisplay = True
    mHandleInvOrLogErrFiles

End Sub
Public Sub mPopulateMessageNames()
ReDim slMessageNames(0 To 68) As String         'all message filenames.  Increase this array as new text files are added and add the filename to end of list.  list box is sorted
Dim ilLoopOnNames As Integer

        slMessageNames(0) = "Advertiser Merge"
        slMessageNames(1) = "Advertiser Product Merge"
        slMessageNames(2) = "Affiliate - Aired Spot Import"
        slMessageNames(3) = "Affiliate - Scheduled Spot Export"
        slMessageNames(4) = "Agency Merge"
        slMessageNames(5) = "Background Schedule Messages"
        slMessageNames(6) = "Backup"
        slMessageNames(7) = "Contract Swap Check"
        slMessageNames(8) = "CSISetup"
        slMessageNames(9) = "CSIStart"
        slMessageNames(10) = "CSIUnzip"
        slMessageNames(11) = "DBUnzip"
        slMessageNames(12) = "Export - Audio MP2"
        slMessageNames(13) = "Export - Automation"
        slMessageNames(14) = "Export - Dallas"
        slMessageNames(15) = "Export - Efficio"
        slMessageNames(16) = "Export - ENCO"
        slMessageNames(17) = "Export - Engineering Feed"
        slMessageNames(18) = "Export - Get Paid"
        slMessageNames(19) = "Export - Great Plains"
        slMessageNames(20) = "Export - Invoice"
        slMessageNames(21) = "Export - Matrix"
        slMessageNames(22) = "Export - Network Inventory"
        slMessageNames(23) = "Export - Phoenix"
        slMessageNames(24) = "Export - Sales Force"
        slMessageNames(25) = "Export - Scheduled Spots"
        slMessageNames(26) = "File Check"
        slMessageNames(27) = "File Fix"
        slMessageNames(28) = "Import - Act1"
        slMessageNames(29) = "Import - Automation"
        slMessageNames(30) = "Import - CSV"
        slMessageNames(31) = "Import - RADAR"
        slMessageNames(32) = "Import - Satellite"
        slMessageNames(33) = "Import - Vehicle"
        slMessageNames(34) = "Invoice Check"
        slMessageNames(35) = "Invoicing"
        slMessageNames(36) = "Logs"
        slMessageNames(37) = "Export - Corporate"
        slMessageNames(38) = "Reallocation"
        slMessageNames(39) = "Reconcile"
        slMessageNames(40) = "Rep-Net Link"
        slMessageNames(41) = "Salespeople Merge"
        slMessageNames(42) = "Set Copy Inventory Dates"
        slMessageNames(43) = "Set Credit Messages"
        slMessageNames(44) = "Show Fix"
        slMessageNames(45) = "SMFCheck"
        slMessageNames(46) = "Sport Contract Check"
        slMessageNames(47) = "SsfCheck"
        slMessageNames(48) = "SSFFix"
        slMessageNames(49) = "Station Feed - Copy Export"
        slMessageNames(50) = "Station Feed - Export All Instructions"
        slMessageNames(51) = "Station Feed - Export All Spots"
        slMessageNames(52) = "Station Feed - Export Regional Spots"
        slMessageNames(53) = "Station Feed - Export Station Feed"
        slMessageNames(54) = "Stations Not Found - Import Region"
        slMessageNames(55) = "Traffic Messages"
        slMessageNames(56) = "Update issues"
        slMessageNames(57) = "VCreative"
        slMessageNames(58) = "Vehicle Merge"
        slMessageNames(59) = "Export - Tableau"         '7-16-15
        '8723
        slMessageNames(60) = "Filemaker To Orders Issues"
        slMessageNames(61) = "Set Credit"
        slMessageNames(62) = "Days to Pay"
        slMessageNames(63) = "Pool Unassigned Log"
        slMessageNames(64) = "Export - RAB"                 '1-30-20
        slMessageNames(65) = "Log Affiliate Posted"                 '1-30-20
        slMessageNames(66) = "Export - CRE"                 '1-11-21
        slMessageNames(67) = "Over-Delivered CPM Impressions"
        slMessageNames(68) = "Web API"      ' 7-19-23
        For ilLoopOnNames = LBound(slMessageNames) To UBound(slMessageNames)
            lbcFileSelect.AddItem slMessageNames(ilLoopOnNames)
            lbcFileSelect.ItemData(lbcFileSelect.NewIndex) = ilLoopOnNames
        Next ilLoopOnNames
        
        Exit Sub
End Sub


Private Sub mHeading1(ilRet As Integer, slHeading As String, ilCurrentLineNo As Integer)
    On Error GoTo mHeading1Err:
    Printer.Print slHeading
    If ilRet <> 0 Then
        'Return
        Exit Sub
    End If
    ilCurrentLineNo = ilCurrentLineNo + 1
    Printer.Print " "
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
mHeading1Err:
    ilRet = err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub

Private Sub mLineOutput(ilRet As Integer, slHeading As String, ilCurrentLineNo As Integer, slRecord As String, ilLinesPerPage As Integer)
    On Error GoTo mLineOutputErr:
    If ilCurrentLineNo >= ilLinesPerPage Then
        Printer.NewPage
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
        ilCurrentLineNo = 0
        '6/6/16: Replaced GoSub
        'GoSub mHeading1
        mHeading1 ilRet, slHeading, ilCurrentLineNo
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
    End If
    Printer.Print slRecord
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
mLineOutputErr:
    ilRet = err.Number
    MsgBox "Printing Error # " & str$(ilRet)
    Resume Next
End Sub
Private Function mIsFileSet(slFileToTest As String) As Boolean
    '8723
    Dim blRet As Boolean
    Dim ilPos As Integer
    Dim slFiles As String
   ' Dim slFileName As String
    Dim slFolder As String
    Dim slMainName As String
    '8886
    Dim myFolder As Folder
    Dim MyFile As file
    Dim blAsteria As Boolean
    
    blRet = False
    slMainName = slFileToTest
    ilPos = InStr(1, slMainName, FILESET)
    If ilPos > 0 Then
        lbcUsers.Clear
        blRet = True
        slFolder = sgDBPath & "Messages\"
        slFiles = Mid(slMainName, ilPos + Len(FILESET))
        blAsteria = False
        If InStr(1, slFiles, "*", vbTextCompare) = Len(slFiles) Then
            blAsteria = True
            slFiles = Left(slFiles, Len(slFiles) - 1)
        End If
        '8886
'        slFileName = ""
'        slFileName = Dir(slFolder & slFiles & "*LOG_??-??-??.txt")
'        Do While slFileName > ""
'            lbcUsers.AddItem slFileName
'            slFileName = Dir()
'        Loop
        Set myFolder = myLogger.MyFile.GetFolder(slFolder)
        If blAsteria Then
            For Each MyFile In myFolder.Files()
                If InStr(1, MyFile.Name, slFiles, vbTextCompare) > 0 Then
                    lbcUsers.AddItem MyFile.Name
                End If
            Next
        Else
            For Each MyFile In myFolder.Files()
                If myLogger.IsLogFile(MyFile.Name) Then
                    If InStr(1, MyFile.Name, slFiles, vbTextCompare) Then
                        lbcUsers.AddItem MyFile.Name
                    End If
                End If
            Next
        End If
        lbcUsers.Visible = True
        bmIsFileSet = True
    End If
'    ilPos = InStr(1, slMainName, "TaskBlocked_")
'    If ilPos > 0 Then
'        blRet = True
'        slFolder = sgDBPath & "Messages\"
'        slFileName = Dir(slFolder & slMainName)
'        Do While slFileName > ""
'            lbcUsers.AddItem slFileName
'            slFileName = Dir()
'        Loop
'        lbcUsers.Visible = True
'
'    End If
    mIsFileSet = blRet

End Function
