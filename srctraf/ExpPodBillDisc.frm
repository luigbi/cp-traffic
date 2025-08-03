VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpPodBillDisc 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   10245
   Begin VB.Frame frcAmazon 
      Height          =   1455
      Left            =   8160
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox edcRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "Region/Endpoint - Example: USEast1, USEast2, USWest1 or USWest2"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox edcPrivateKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   9
         ToolTipText     =   "The Private Key Assigned by AWS"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox edcAccessKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "The Access Key Assigned by AWS"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox edcBucketName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox ckcKeepLocalFile 
         Caption         =   "Keep Local File"
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox edcAmazonSubfolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         ToolTipText     =   "(Optional) Amazon Web Bucket Subfolder Name.   Example: Counterpoint"
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PrivateKey"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "AccessKey"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "BucketName"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lacExportFilename 
         Caption         =   "lacExportFilename"
         Height          =   255
         Left            =   7800
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Folder (optional)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.PictureBox plcTo 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   5685
      TabIndex        =   31
      Top             =   1680
      Width           =   5745
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   30
         Width           =   5625
      End
   End
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7800
      Top             =   2160
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7320
      Top             =   2160
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
      Left            =   5505
      TabIndex        =   13
      Top             =   3120
      Width           =   1050
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   3795
      TabIndex        =   12
      Top             =   3120
      Width           =   1050
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6105
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5835
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6450
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcTo 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
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
      Left            =   7200
      TabIndex        =   4
      Top             =   1680
      Width           =   1485
   End
   Begin VB.CheckBox ckcAmazon 
      Caption         =   "Upload to Amazon Web bucket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   3075
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   3075
   End
   Begin VB.TextBox edcMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   600
   End
   Begin VB.TextBox edcYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   1
      Top             =   600
      Width           =   720
   End
   Begin VB.TextBox edcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6840
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label lacSaveIn 
      Appearance      =   0  'Flat
      Caption         =   "Save In"
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
      Height          =   210
      Left            =   360
      TabIndex        =   29
      Top             =   1710
      Width           =   810
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   3360
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label lacStartYear 
      Appearance      =   0  'Flat
      Caption         =   "Start Year"
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
      Height          =   225
      Left            =   2040
      TabIndex        =   16
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label lacMonth 
      Appearance      =   0  'Flat
      Caption         =   "Start Month"
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
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   630
      Width           =   1185
   End
   Begin VB.Label lacContract 
      Appearance      =   0  'Flat
      Caption         =   "Contract #"
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
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   1140
      Width           =   1065
   End
End
Attribute VB_Name = "ExpPodBillDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpPodBillDisc.Frm
'
' Release: 1.0
' Date   : 3/12/21
'
' Description:
'   This file contains the Export "Ad Server Billing Discrepancy" input screen code
Option Explicit
Option Compare Text
Dim imExporting As Integer
Dim imTerminate As Integer
Dim imFirstActivate As Integer
Dim imExportOption As Integer       'lbcExport.ItemData(lbcExport.ListIndex)
Dim smClientName As String
Dim myBucket As CsiToAmazonS3.ApiCaller
Dim hmMsg As Integer   'From file hanle
Dim hmPodBillDisc As Integer
Dim lmNowDate As Long   'Todays date
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract
Dim smExportName As String
Dim smExportOptionName As String
Dim smExportFilename As String
' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
Dim tmMnfSS() As MNF                    'array of Sales Sources MNF
Dim tmMnfGroups() As MNF
Dim lmStdStartDates(0 To 2) As Long   'start dates by std Broadcast Calendar
Dim lmCalStartDates(0 To 2) As Long   'start dates by std Broadcast Calendar
Dim lmStartDate As Long 'overall StartDate for Query (Min btwn Std and Cal)
Dim lmEndDate As Long   'overall EndDate for Query (Max btwn Std and Cal)

Private Sub ckcAmazon_Click()
    If ckcAmazon.Value = vbChecked Then
        frcAmazon.Left = 120
        frcAmazon.Top = 1440
        frcAmazon.Visible = True
    Else
        frcAmazon.Visible = False
    End If
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDateTime As String
    Dim slMonthHdr As String * 36
    Dim ilSaveMonth As Integer
    Dim ilYear As Integer
    Dim slStart As String
    Dim slTimeStamp As String
    Dim ilHowManyDefined As Integer
    Dim ilHowMany As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim slFNMonth As String
    Dim olRs As ADODB.Recordset
    Dim slErrorMessage As String
    Dim blLogSuccess As Boolean
    
    lacInfo(0).Visible = True
    lacInfo(1).Visible = False

    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    'Amazon Enabled?
    If ckcAmazon.Value = vbChecked Then
        If edcBucketName.Text = "" Or edcRegion.Text = "" Or edcAccessKey.Text = "" Or edcPrivateKey.Text = "" Then ckcAmazon.Value = vbUnchecked
    End If
    
    'Verify data input
    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    slStr = edcMonth.Text             'month in text form (jan..dec, or 1-12
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
    End If
    If ilSaveMonth = 0 Then
        edcMonth.SetFocus                 'invalid month
        gAutomationAlertAndLogHandler "Month is Not Valid", vbOkOnly + vbApplicationModal, "Start Month"
        Exit Sub
    End If
    
    slFNMonth = Mid$(slMonthHdr, (ilSaveMonth - 1) * 3 + 1, 3)          'get the text month (jan...dec)
    slStr = edcYear.Text
    ilYear = gVerifyYear(slStr)
    If ilYear = 0 Then
        edcYear.SetFocus                 'invalid year
        'MsgBox "Year is Not Valid", vbOkOnly + vbApplicationModal, "Start Year"
        gAutomationAlertAndLogHandler "Year is Not Valid", vbOkOnly + vbApplicationModal, "Start Year"
        Exit Sub
    End If

    lmCntrNo = 0                'ths is for debugging on a single contract
    slStr = edcContract
    If slStr <> "" Then
        lmCntrNo = Val(slStr)
    End If

    'smExportFile contains the name to use which has been moved to edcTo.Text
    smExportName = Trim$(edcTo.Text)
    If Len(smExportName) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Sub
    End If

    If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then
        smExportName = Trim$(sgExportPath) & smExportName
    End If

    ilRet = 0
    'On Error GoTo cmcExportErr:
    'slDateTime = FileDateTime(smExportName)
    ilRet = gFileExist(smExportName)
    If ilRet = 0 Then
        'file already exists, do not overwrite
        'MsgBox "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        gAutomationAlertAndLogHandler "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        
        Exit Sub
        'Kill smExportName
    End If

    If Not mOpenMsgFile() Then          'open message file
         cmcCancel.SetFocus
         Exit Sub
    End If
    On Error GoTo 0
    ilRet = 0
    'Print #hmMsg, "** Storing Output into " & smExportName & " **"
    gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
    
    gAutomationAlertAndLogHandler "* StartMonth=" & edcMonth.Text
    gAutomationAlertAndLogHandler "* StartYear=" & edcYear.Text
    gAutomationAlertAndLogHandler "* Contract#=" & edcContract.Text
    If ckcAmazon.Value = vbChecked Then
        gAutomationAlertAndLogHandler "* AmazonBucket=True"
    Else
        gAutomationAlertAndLogHandler "* AmazonBucket=False"
    End If
    lacInfo(0).Visible = True
    lacInfo(0).Caption = "Exporting...": lacInfo(0).Refresh
    Screen.MousePointer = vbHourglass
    imExporting = True

    'Get Std Bcast Cal Billing periods
    slStart = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(str$(ilYear))
    gBuildStartDates slStart, 1, 2, lmStdStartDates() 'build array of std start & end dates
    'Get Calendar Billing periods
    slStart = Trim$(str$(ilSaveMonth)) & "/01/" & Trim$(str(ilYear))
    gBuildStartDates slStart, 4, 2, lmCalStartDates() 'build array of std start & end dates

    lmStartDate = IIF(lmCalStartDates(1) < lmStdStartDates(1), lmCalStartDates(1), lmStdStartDates(1))
    lmEndDate = IIF(lmCalStartDates(2) > lmStdStartDates(2), lmCalStartDates(2) - 1, lmStdStartDates(2) - 1)
    
    slErrorMessage = mQueryDatabaseCSV(olRs)
    If slErrorMessage = "No errors" Then
        slErrorMessage = mExport(olRs, smExportName)
    End If
    
    If Mid(slErrorMessage, 1, 9) = "No errors" Then
        lacInfo(0).Caption = "Export created... " & slErrorMessage: lacInfo(0).Refresh
        'Print #hmMsg, "** Export Created: " & slErrorMessage
        gAutomationAlertAndLogHandler "** Export Created: " & slErrorMessage
        'Amazon
        If ckcAmazon.Value = vbChecked And edcBucketName.Text <> "" And edcRegion.Text <> "" And edcAccessKey.Text <> "" And edcPrivateKey.Text <> "" Then
            'Print #hmMsg, "** Uploading " & smExportFilename & " to " & edcBucketName.Text & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            'TTP 10504 - Amazon web bucket upload cleanup
            gAutomationAlertAndLogHandler "** Uploading " & smExportFilename & " to " & AmazonBucketFolder(edcBucketName.Text, edcAmazonSubfolder.Text) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            lacInfo(0).Caption = "Uploading " & smExportFilename
            lacInfo(0).Refresh
            DoEvents
            Set myBucket = New CsiToAmazonS3.ApiCaller
            On Error Resume Next
            err = 0
            'TTP 10504 - Amazon web bucket upload cleanup
'            If edcAmazonSubfolder.Text <> "" Then
'                edcAmazonSubfolder.Text = Trim(Replace(edcAmazonSubfolder.Text, "\", "/"))
'                If right(edcAmazonSubfolder.Text, 1) <> "/" Then edcAmazonSubfolder.Text = edcAmazonSubfolder.Text & "/"
'                If Left(edcAmazonSubfolder.Text, 1) = "/" Then edcAmazonSubfolder.Text = Mid(edcAmazonSubfolder.Text, 2)
'                '3/1/21 - added Folder support: "|" to split Bucket Name and Subfolder
'                myBucket.UploadAmazonBucketFile edcBucketName.Text + "|" + edcAmazonSubfolder.Text, edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, edcTo.Text, False
'            Else
'                myBucket.UploadAmazonBucketFile edcBucketName.Text, edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, edcTo.Text, False
'            End If
            myBucket.UploadAmazonBucketFile AmazonBucketFolder(edcBucketName.Text, edcAmazonSubfolder.Text), edcRegion.Text, edcAccessKey.Text, edcPrivateKey.Text, edcTo.Text, False
            If err <> 0 Then
                lacInfo(0).Caption = "Error Uploading " & smExportFilename & " - " & err & " - " & Error(err)
                'Print #hmMsg, "** Error Uploading " & smExportFilename & " - " & err & " - " & Error(err) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                gAutomationAlertAndLogHandler "** Error Uploading " & smExportFilename & " - " & err & " - " & Error(err) & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Else
                If myBucket.ErrorMessage <> "" Then
                    lacInfo(0).Caption = "Error Uploading " & smExportFilename
                    'Print #hmMsg, "** Error Uploading " & smExportFilename & " - " & Replace(myBucket.ErrorMessage, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    gAutomationAlertAndLogHandler "** Error Uploading " & smExportFilename & " - " & Replace(myBucket.ErrorMessage, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                Else
                    lacInfo(0).Caption = "Sucess Uploading " & smExportFilename
                    'Print #hmMsg, "** Finished Uploading " & smExportFilename & " - " & Replace(myBucket.Message, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    gAutomationAlertAndLogHandler "** Finished Uploading " & smExportFilename & " - " & Replace(myBucket.Message, vbCrLf, ";") & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    If ckcKeepLocalFile.Value = vbUnchecked Then
                        'We want to remove the Local File
                        Kill edcTo.Text
                        'Print #hmMsg, "** Deleted Local Export File : " & smExportName & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        gAutomationAlertAndLogHandler "** Deleted Local Export File : " & smExportName & " - " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                    End If
                End If
            End If
            Set myBucket = Nothing
        End If
    Else
        lacInfo(0).Caption = "Errors writing ..." & slErrorMessage: lacInfo(0).Refresh
        'Print #hmMsg, "Export Terminated, " & "Errors writing ..." & slErrorMessage
        gAutomationAlertAndLogHandler "Export Terminated, " & "Errors writing ..." & slErrorMessage
    End If
    
    Set olRs = Nothing
    Close #hmMsg
    cmcExport.Enabled = False
    cmcCancel.Caption = "&Done"
    If igExportType <= 1 Then       'ok to set focus if manual mode
        cmcCancel.SetFocus
    End If
    Screen.MousePointer = vbDefault
    imExporting = False
    Exit Sub
    
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub

Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Csv"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo.Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
    If edcTo.Text = "" Then
        edcTo.Text = smExportName
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcMonth_Change()
Dim slStr As String
    If Len(edcMonth) = 3 Then
        gCtrlGotFocus edcMonth
        If igExportType <= 1 And Not imFirstActivate Then
            tmcClick_Timer
        End If
    End If
End Sub

Private Sub edcMonth_Click()
    gCtrlGotFocus edcMonth
End Sub

Private Sub edcMonth_GotFocus()
    gCtrlGotFocus edcMonth
End Sub

Private Sub edcMonth_LostFocus()
    If igExportType <= 1 And Not imFirstActivate Then
        tmcClick_Timer
    End If
End Sub

Private Sub edcTo_Change()
    'get Filename from the full path and filename (used as the object name to upload to Amazon)
    cmcExport.Enabled = False
    Dim lsFilename As String
    Dim liSeparator As Integer
    lsFilename = edcTo.Text
    If lsFilename = "" Then Exit Sub
    liSeparator = InStrRev(lsFilename, "\")
    smExportFilename = Mid(lsFilename, liSeparator + 1)
    lacExportFilename.Caption = smExportFilename
    cmcExport.Enabled = True
End Sub

Private Sub edcYear_Change()
    If Len(edcYear.Text) = 4 Then
        gCtrlGotFocus edcYear
        If igExportType <= 1 And Not imFirstActivate Then
            tmcClick_Timer
        End If
    End If
End Sub

Private Sub edcYear_Click()
    gCtrlGotFocus edcYear
End Sub

Private Sub edcYear_GotFocus()
    gCtrlGotFocus edcYear
End Sub

Private Sub edcYear_LostFocus()
    If igExportType <= 1 And Not imFirstActivate Then
        tmcClick_Timer
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
    
    tmcClick_Timer
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    
    'check perms
    If imExportOption = EXP_ADSERVERBILLDISC Then
        If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER Then
            lacInfo(0).AddItem "Ad Server Discrepancy Export Disabled"
            imTerminate = True
            Exit Sub
        End If
    End If
    
    If igExportType <= 1 Then
        'manual from exports or manual from traffic
        Me.WindowState = vbNormal
        cmcExport.Enabled = True
    Else
        'Running in Auto mode
        Me.WindowState = vbMinimized
        If imExportOption = EXP_ADSERVERBILLDISC Then
            cmcExport.Enabled = True
            'gOpenTmf
            'tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
            'tmcSetTime.Enabled = True
            'gUpdateTaskMonitor 1, "ME"
            cmcExport_Click
            'gUpdateTaskMonitor 2, "ME"
            imTerminate = True
        End If
    End If
    
    tmcClick.Interval = 2000    '2 seconds
    tmcClick.Enabled = True
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slMonthStr As String * 36
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilVff As Integer
    Dim ilLoop As Integer
    Dim slLocation As String
    Dim slReturn As String * 130
    Dim slFileName As String
    Dim ilVefInx As Integer
    Dim StartPeriod As Integer
    slMonthStr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imExporting = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    'lmNowDate = 44242 '2/15/21, lmNowDate = 44211 '1/15/21

    gCenterStdAlone Me
    
    'get Export Option #
    imExportOption = ExportList!lbcExport.ItemData(ExportList!lbcExport.ListIndex)
    If imExportOption = EXP_ADSERVERBILLDISC Then
        smExportOptionName = "AdServerBillDisc"
    Else
        smExportOptionName = ""
    End If
    
    'a timing issue prevents the filename from showing in the text box.
    If igExportType >= 4 Then   'igExportType As Integer  '0=Manual; 1=From Traffic,8=PodBillDisc
        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
    End If
    
    'determine default month year
    slDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'    'Default to last month, based on today's date
'    If Val(slMonth) = 1 Then
'        ilMonth = 12
'        ilYear = Val(slYear) - 1
'    Else
'        ilMonth = Val(slMonth) - 1
'        ilYear = Val(slYear)
'    End If
    
    'Default to current month, based on today's system date
    ilMonth = Val(slMonth)
    ilYear = Val(slYear)
    
    edcMonth.Text = Mid$(slMonthStr, (ilMonth - 1) * 3 + 1, 3)
    edcYear.Text = Trim$(str$(ilYear))
    
    If igExportType <= 1 Then                    'igExportType As Integer  '0=Manual; 1=From Traffic, ..8=PodBillDisc
        'Disable Options based on Site if needed (for Interactive mode)
    Else
        On Error GoTo mObtainIniValuesErr
        'find exports.ini
        sgIniPath = gSetPathEndSlash(sgIniPath, True)
        If igDirectCall = -1 Then
            slFileName = sgIniPath & "Exports.Ini"
        Else
            slFileName = CurDir$ & "\Exports.Ini"
        End If
        
        'allow month start period adjustment (automation), add # of months to filename. Exports.ini "StartPeriod=#" - if StartPeriod=2 then the Start Month / Year is set to 2 Months prior to current date / passed in date
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "StartPeriod", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'Use the Start Month/Year that's already defaulted (1 month prior to today)
        Else
            StartPeriod = Trim$(gStripChr0(slReturn))
            If StartPeriod > 0 Then
                edcMonth.Text = MonthName(Month(DateAdd("m", -StartPeriod, gObtainEndStd(Format$(lmNowDate, "m/d/yy")))), True)
                edcYear.Text = Year(DateAdd("m", -StartPeriod, gObtainEndStd(Format$(lmNowDate, "m/d/yy"))))
            End If
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Export", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'default to the export path
            sgExportPath = sgExportPath
        Else
            sgExportPath = Trim$(gStripChr0(slReturn))
        End If
        sgExportPath = gSetPathEndSlash(sgExportPath, True)

        'Amazon Support
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "BucketName", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcBucketName.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcBucketName.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "BucketFolder", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcAmazonSubfolder.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcAmazonSubfolder.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Region", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcRegion.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcRegion.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "AccessKey", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcAccessKey.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcAccessKey.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "PrivateKey", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            edcPrivateKey.Text = ""
        Else
            slCode = Trim$(gStripChr0(slReturn))
            edcPrivateKey.Text = slCode
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "KeepLocalFile", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ckcKeepLocalFile.Value = vbUnchecked           'delete local file after upload to Amazon
        Else
            If InStr(1, slReturn, "Yes", vbTextCompare) > 0 Then
                ckcKeepLocalFile.Value = vbChecked         'Keep local file
            Else
                ckcKeepLocalFile.Value = vbUnchecked       'delete local file after upload to Amazon
            End If
        End If
        'TTP 9992
        If edcBucketName.Text <> "" And edcRegion.Text <> "" And edcAccessKey.Text <> "" And edcPrivateKey.Text <> "" Then
            'If INI provides all 4 AWS values then Check Amazon
            ckcAmazon.Value = vbChecked
        Else
            ckcAmazon.Value = vbUnchecked
        End If
    End If

    smClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
        tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smClientName = Trim$(tmMnf.sName)
        End If
    End If
    
    If igExportType >= 4 Then      'need to setup the filename if background mode
        tmcClick_Timer
    End If
    
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
    Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub

mObtainIniValuesErr:
    Resume Next

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
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgDBPath & "Messages\" & "ExpAdServerBillDisc.Txt"
    sgMessageFile = slToFile
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                'MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                'MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            'MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, "Export Ad Server Billing Discrepancy: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    'Print #hmMsg, ""
    gAutomationAlertAndLogHandler "Export Ad Server Billing Discrepancy: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    
    mOpenMsgFile = True
    Exit Function
End Function

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
    igParentRestarted = False
    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Me
    igManUnload = NO
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If smExportOptionName = "AdServerBillDisc" Then
        plcScreen.Print "Ad Server Billing Discrepancy Export"
    Else
        plcScreen.Print smExportOptionName & " Export"
    End If
End Sub

Private Sub tmcClick_Timer()
Dim slRepeat As String
Dim ilRet As Integer
Dim slDateTime As String
Dim slMonthBy As String
Dim slMonthHdr As String
Dim slStr As String
Dim ilSaveMonth As Integer
Dim slFNMonth As String
Dim ilYear As Integer
Dim slExtension As String * 4
    smExportFilename = ""
    tmcClick.Enabled = False
    'Determine name of export (.csv file)
    slExtension = ".csv"
    slRepeat = "A"
    If imExportOption = EXP_ADSERVERBILLDISC Then
        slExtension = ".csv"
        slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
        slStr = edcMonth.Text             'month in text form (jan..dec, or 1-12
        gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
        If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
            ilSaveMonth = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
'                MsgBox "Month is Not Valid", vbOkOnly + vbApplicationModal, "Start Month"
                Exit Sub
            End If
        End If
    
        'DoEvents
        slFNMonth = Mid$(slMonthHdr, (ilSaveMonth - 1) * 3 + 1, 3)          'get the text month (jan...dec)
        slStr = edcYear.Text
        ilYear = gVerifyYear(slStr)
        If ilYear = 0 Then
'            MsgBox "Year is Not Valid", vbOkOnly + vbApplicationModal, "Start Year"
            Exit Sub
        End If
    End If
    slMonthBy = slMonthBy & Trim$(slFNMonth) & Trim$(str$(ilYear)) & "-"
    'build Filename
    Do
        ilRet = 0
        smExportFilename = Trim$(smExportOptionName) & " " & slMonthBy & Format(gNow, "mmddyy") & gFileNameFilter(Trim$(slRepeat & " " & Trim$(smClientName))) & slExtension
        smExportName = Trim$(sgExportPath) & Trim$(smExportOptionName) & " " & slMonthBy & Format(gNow, "mmddyy")
        smExportName = Trim$(smExportName) & gFileNameFilter(slRepeat & " " & Trim$(smClientName)) & slExtension             '2-27-14
        ilRet = gFileExist(smExportName)
        If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
            slRepeat = Chr(Asc(slRepeat) + 1)
        End If
    Loop While ilRet = 0
    edcTo.Text = smExportName
    edcTo.Visible = True
    Exit Sub
End Sub

Private Sub tmcSetTime_Timer()
'    If imExportOption = EXP_ADSERVERBILLDISC Then
'        gUpdateTaskMonitor 0, "PBD"
'    End If
End Sub

Function mExport(ByRef olRs As Recordset, smExportName) As String
    'Create a CSV file, give me a recordset and a (fully qualified path\filename.ext) filename
    'Makes Headers from recordset Column Names
    'Makes rows from Data
    Dim slErrorMessage As String
    Dim olFileSys As FileSystemObject
    Dim olCsv As TextStream
    Dim slPath As String
    Dim slRowToWrite As String
    Dim slComma As String
    Dim slAppendLine As String
    Dim olField As Field
    Dim slFormattedString As String
    Dim slHeader As String
    Dim llRecords As Long
    Dim slBillCycle As String
    Dim slLineStartDate As String
    Dim slLineEndDate As String
    Dim ilMatchCntr As Integer
    Dim ilRdfCode As Integer
    Dim ilAdfCode As Integer
    Dim ilVefCode As Integer
    Dim slRdfName As String
    Dim slAdfName As String
    Dim slVefName As String
    Dim slProduct As String
    Dim ilRet As Integer
    
    On Error GoTo ERRORBOX
    slComma = ","
    Set olFileSys = New FileSystemObject
    Set olCsv = olFileSys.OpenTextFile(smExportName, ForWriting, True)
    
    'Get Header
    slHeader = mGetHeaderString(olRs)
    
    'Write Header
    olCsv.WriteLine slHeader
    
    'Check Records
    If olRs.EOF And olRs.BOF Then
        mExport = "There are no records to export"
        GoTo finish
    End If
    
    'Wite Data rows
    olRs.MoveFirst
    Do While Not olRs.EOF
        slRowToWrite = ""
        ilMatchCntr = False
        'check if the record falls in the Selected Month based on Contract BillCycle
        slBillCycle = olRs.Fields("xchfBillCycle").Value
        slLineStartDate = olRs.Fields("xStartDate").Value
        slLineEndDate = olRs.Fields("xEndDate").Value
        If slBillCycle = "C" Then
            'Check if Contract is in Date span using Monthly Cal
            If DateValue(slLineEndDate) > DateValue(Format(lmCalStartDates(1), "m/d/yy")) And DateValue(slLineStartDate) < DateValue(Format(lmCalStartDates(2), "m/d/yy")) Then
                ilMatchCntr = True
            End If
        Else
            'Check if Contract is in Date span using Std BCast Cal
            If DateValue(slLineEndDate) >= DateValue(Format(lmStdStartDates(1), "m/d/yy")) And DateValue(slLineStartDate) < DateValue(Format(lmStdStartDates(2), "m/d/yy")) Then
                ilMatchCntr = True
            End If
        End If
        If ilMatchCntr = True Then
            'RDF Daypart name lookup
            ilRdfCode = olRs.Fields("xRdfCode").Value
            slRdfName = ""
            ilRet = gBinarySearchRdf(ilRdfCode)
            If ilRet <> -1 Then
                slRdfName = Trim(tgMRdf(ilRet).sName)
            End If
            
            'ADF Name Lookup
            ilAdfCode = olRs.Fields("xAdfCode").Value
            slAdfName = ""
            ilRet = gBinarySearchAdf(ilAdfCode)
            If ilRet <> -1 Then
                slAdfName = Trim$(tgCommAdf(ilRet).sName)
            End If
            
            'Vef Name Lookup
            ilVefCode = olRs.Fields("xVefCode").Value
            slVefName = ""
            ilRet = gBinarySearchVef(ilVefCode)
            If ilRet <> -1 Then
                slVefName = Trim(tgMVef(ilRet).sName)
            End If
            
            'Get the Product name from the Results
            slProduct = Trim(olRs.Fields("sAdvertiser/Product").Value) 'We have the product, but were going to lookup the Advertiser and insert the value into this column (For Performance)
        
            For Each olField In olRs.Fields
                slFormattedString = mWriteField(olField)
                If slFormattedString <> "~[IGNORE]~" Then
                    'not sure if need to test for first error string.
                    If slFormattedString = "Error reading records in mExport" Or slFormattedString = "Error in function mWriteField" Then
                        mExport = slFormattedString
                        GoTo finish
                    End If
                    
                    Select Case olField.Name
                        Case "sAdvertiser/Product"
                            'TTP 10670 - Ad Server Billing Discrepancy: comma in advertiser name causes rows to get misaligned when opened in Excel
                            slFormattedString = """" & Trim(slAdfName) & "/" & Trim(slProduct) & """"
                        Case "sVehicle"
                            slFormattedString = """" & Trim(slVefName) & """"
                        Case "sAd Location"
                            'JW 7/26/23 - added Quotes around Ad Location
                            slFormattedString = """" & Trim(slRdfName) & """"
                    End Select
                    
                    If slRowToWrite <> "" Then slRowToWrite = slRowToWrite & slComma
                    slRowToWrite = slRowToWrite & slFormattedString
                End If
            Next olField
            olCsv.WriteLine slRowToWrite
            llRecords = llRecords + 1
        End If
        olRs.MoveNext
    Loop
    olRs.Close
    olCsv.Close
    slErrorMessage = "No errors, " & llRecords & " rows exported.."
    mExport = slErrorMessage

finish:
    Set olField = Nothing
    Set olFileSys = Nothing
    Set olCsv = Nothing
    Exit Function

ERRORBOX:
    mExport = "Error reading records in mExport:" & err & "-" & Error(err)
    Set olField = Nothing
    Set olFileSys = Nothing
    Set olCsv = Nothing
    GoTo finish
End Function

Private Function mGetHeaderString(ByRef olRs As Recordset) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'Get the headers from the recordset columns (skipping the 1st Letter, which is a format indicator)
    mGetHeaderString = ""
    Dim slComma As String
    Dim olField As Field
    slComma = ","
    For Each olField In olRs.Fields
            If Left(olField.Name, 1) <> "x" Then
            If mGetHeaderString <> "" Then mGetHeaderString = mGetHeaderString & slComma
            mGetHeaderString = mGetHeaderString & Mid(olField.Name, 2)
        End If
    Next olField
End Function

Private Function mWriteField(olField As Field) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'return a formatted string for the CSV export, based on the 1st letter (format indicator) of ColumnName
    Dim slReturnName As String
    Dim ilWriteToLineCode As Integer
    On Error GoTo ERRORBOX
    Select Case Left(olField.Name, 1)
        Case "s" 'String ("x")
            If IsNull(olField.Value) Then
                mWriteField = """"""
            Else
                mWriteField = Chr(34) & Trim(olField.Value) & Chr(34) 'Quotted
            End If
        Case "i" 'integer (0)
            If IsNull(olField.Value) Then
                mWriteField = ""
            Else
                mWriteField = Trim(Int(Val(olField.Value)))
            End If
        Case "c" 'currency (#.00)
            If IsNull(olField.Value) Then
                mWriteField = ""
            Else
                mWriteField = Format(Val(olField.Value), "#0.00")
            End If
        'could write other handlers here if needed
        Case "x" 'ignore me
            mWriteField = "~[IGNORE]~"
        Case Else
            mWriteField = Chr(34) & Trim(olField.Value) & Chr(34) 'Quotted
    End Select
    Exit Function
    
ERRORBOX:
    mWriteField = "Error in function mWriteField"
End Function

Private Function mQueryDatabaseCSV(ByRef olRs As Recordset) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'This returns the Column Names Exactly as they should appear in the export, with a prepended format character (like s=string, n=numeric); example: "sCall Letters"
    Dim slErrorMessage As String
    Dim SQLQuery As String
    Dim blNeedAnd As Boolean 'For Query building
    Dim lLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    
    slStartDate = Format(lmStartDate, "yyyy-mm-dd")
    slEndDate = Format(lmEndDate, "yyyy-mm-dd")
    ilYear = Year(slEndDate)
    ilMonth = Month(slEndDate)
    SQLQuery = SQLQuery & "SELECT "
    SQLQuery = SQLQuery & " chf.chfBillCycle                                                AS 'xchfBillCycle', "
    SQLQuery = SQLQuery & " pcf.pcfStartDate                                                AS 'xStartDate', "
    SQLQuery = SQLQuery & " pcf.pcfEndDate                                                  AS 'xEndDate', "
    SQLQuery = SQLQuery & " pcf.pcfrdfcode                                                  as 'xRdfCode',"
    SQLQuery = SQLQuery & " chf.chfadfcode                                                  as 'xAdfCode',"
    SQLQuery = SQLQuery & " pcf.pcfVefCode                                                  as 'xVefCode',"
    SQLQuery = SQLQuery & " chf.chfCntrNo                                                   AS 'iContract #', "
    'TTP 10681 - Ad Server Billing Discrepancy Export: add External Contract Number and External Line Number
    SQLQuery = SQLQuery & " chf.chfExtCntrNo                                                AS 'iExtContract #', "
    SQLQuery = SQLQuery & " LTRIM(RTRIM(chf.chfProduct))                                    AS 'sAdvertiser/Product', "
    SQLQuery = SQLQuery & " pcf.pcfPodCPMID                                                 AS 'iID', "
    'TTP 10681 - Ad Server Billing Discrepancy Export: add External Contract Number and External Line Number
    SQLQuery = SQLQuery & " pcf.pcfExtLineNo                                                AS 'iExtLineID', "
    SQLQuery = SQLQuery & " NULL                                                            AS 'sVehicle', " 'need to lookup vef.vefName using [xVefCode]
    SQLQuery = SQLQuery & " avf.avfName                                                     AS 'sAd Server Vendor', "
    SQLQuery = SQLQuery & " NULL                                                            AS 'sAd Location', " 'need to lookup rdf.rdfName using [xRdfCode]
    SQLQuery = SQLQuery & " RIGHT('00'+RTRIM(CAST(Month(pcf.pcfStartDate) AS CHAR(2))),2) + '/' + "
    SQLQuery = SQLQuery & "   RIGHT('00'+RTRIM(CAST(Day(pcf.pcfStartDate) AS CHAR(2))),2) +'/' + "
    SQLQuery = SQLQuery & "   RIGHT('00'+RTRIM(CAST(Year(pcf.pcfStartDate) AS CHAR(4))),2) + "
    SQLQuery = SQLQuery & "   ' - ' + "
    SQLQuery = SQLQuery & "   RIGHT('00'+RTRIM(CAST(Month(pcf.pcfEndDate) AS CHAR(2))),2) + '/' + "
    SQLQuery = SQLQuery & "   RIGHT('00'+RTRIM(CAST(Day(pcf.pcfEndDate) AS CHAR(2))),2) +'/' + "
    SQLQuery = SQLQuery & "   RIGHT('00'+RTRIM(CAST(Year(pcf.pcfEndDate) AS CHAR(4))),2)    AS 'sDate Range', "
    
    '09/28/2022 - JW - Ad server billing discrepancy export: show "baked-in" price type
    'SQLQuery = SQLQuery & " IF(pcf.pcfPriceType = 'C','CPM','Flat Rate')                    AS 'sPrice Type', "
    SQLQuery = SQLQuery & " IF(pcf.pcfPriceType = 'C','CPM',IF(pcf.pcfDeliveryType = 1,'Baked-in','Flat Rate')) AS 'sPrice Type', "

    SQLQuery = SQLQuery & " pcf.pcfImpressionGoal                                           AS 'iImpressions Ordered', "
    SQLQuery = SQLQuery & " SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfImpressions,0))           AS 'iTotal Billed', "
    SQLQuery = SQLQuery & " IF(CAST(SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0)) AS Integer) > pcf.pcfImpressionGoal, "
    SQLQuery = SQLQuery & "     if(CAST(CAST(saffeatures7 AS BINARY(1)) AS INTEGER) & 128 = 128, "
    SQLQuery = SQLQuery & "         SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0)), "
    SQLQuery = SQLQuery & "         pcf.pcfImpressionGoal"
    SQLQuery = SQLQuery & "     ), "
    SQLQuery = SQLQuery & "     SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0))"
    SQLQuery = SQLQuery & " )                                                               AS 'iTotal Invoiced',  "
    SQLQuery = SQLQuery & " ibfCurrent.ibfImpressions                                       AS 'iPosted', "
    SQLQuery = SQLQuery & " IF(SUM(IF(ibfCurrent.ibfBilled='Y',ibfCurrent.ibfBilledImpression,0)) <> 0,"
    'SQLQuery = SQLQuery & "     --When running the export for a month that has been invoiced:"
    'SQLQuery = SQLQuery & "     ---> difference should be the [total invoiced] minus the [impressions ordered]. "
    SQLQuery = SQLQuery & "     (IF(CAST(SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0)) AS Integer) > pcf.pcfImpressionGoal, "
    SQLQuery = SQLQuery & "         if(CAST(CAST(saffeatures7 AS BINARY(1)) AS INTEGER) & 128 = 128, "
    SQLQuery = SQLQuery & "             SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0)), "
    SQLQuery = SQLQuery & "             pcf.pcfImpressionGoal"
    SQLQuery = SQLQuery & "         ), "
    SQLQuery = SQLQuery & "         SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0))"
    SQLQuery = SQLQuery & "      )) - pcf.pcfImpressionGoal                                             "
    SQLQuery = SQLQuery & " ,"
    'SQLQuery = SQLQuery & "     --When running the export for a month that has not been invoiced:"
    'SQLQuery = SQLQuery & "     ---> difference should be the [posted amount] for the month plus [total invoiced] minus the [impressions ordered]. "
    SQLQuery = SQLQuery & "     (IF(CAST(SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0)) AS Integer) > pcf.pcfImpressionGoal, "
    SQLQuery = SQLQuery & "         if(CAST(CAST(saffeatures7 AS BINARY(1)) AS INTEGER) & 128 = 128, "
    SQLQuery = SQLQuery & "             SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0)), "
    SQLQuery = SQLQuery & "             pcf.pcfImpressionGoal"
    SQLQuery = SQLQuery & "         ), "
    SQLQuery = SQLQuery & "         SUM(IF(ibfAll.ibfBilled='Y',ibfAll.ibfBilledImpression,0))"
    SQLQuery = SQLQuery & "      ) + IF(ibfCurrent.ibfImpressions IS NULL,0,ibfCurrent.ibfImpressions))"
    SQLQuery = SQLQuery & "      - pcf.pcfImpressionGoal                                        "
    SQLQuery = SQLQuery & " )                                                               AS 'iDifference', "
    'TTP 10751 - Ad Server Billing Discrepancy export: posted impressions of 0 for the month will not appear as "missing"
    'SQLQuery = SQLQuery & " IF (ibfCurrent.ibfImpressions IS NULL,'Missing','')             AS 'sStatus' "
    SQLQuery = SQLQuery & " IF (ibfCurrent.ibfImpressions IS NULL OR ibfCurrent.ibfImpressions=0,'Missing','') AS 'sStatus' "
    
    SQLQuery = SQLQuery & "FROM "
    SQLQuery = SQLQuery & "    ""CHF_Contract_Header"" chf"
    SQLQuery = SQLQuery & "    JOIN ""pcf_Pod_CPM_Cntr"" pcf ON pcf.pcfChfCode = chf.chfCode AND pcf.pcfDelete = 'N' AND pcf.pcfType <> 'P' AND pcf.pcfStartDate <=pcf.pcfEndDate"
    SQLQuery = SQLQuery & "    JOIN ""RDF_Standard_Daypart"" rdf ON rdf.rdfCode = pcf.pcfrdfcode"
    SQLQuery = SQLQuery & "    JOIN ""ADF_Advertisers"" adf ON adf.adfCode = chf.chfadfcode"
    SQLQuery = SQLQuery & "    JOIN ""VEF_Vehicles"" vef ON vef.vefCode = pcf.pcfVefCode"
    SQLQuery = SQLQuery & "    LEFT JOIN ""VFF_Vehicle_Features"" vff ON vff.vffVefCode = vef.vefCode"
    SQLQuery = SQLQuery & "    LEFT JOIN ""avf_AdVendor"" avf ON avf.avfCode = vff.vffAvfCode"
    SQLQuery = SQLQuery & "    LEFT JOIN ""ibf_Impression_Bill"" ibfAll     ON ibfAll.ibfCntrNo = chf.chfCntrNo     AND ibfAll.ibfPodCPMID = pcf.pcfPodCPMID "
    SQLQuery = SQLQuery & "    LEFT JOIN ""ibf_Impression_Bill"" ibfCurrent ON ibfCurrent.ibfCntrNo = chf.chfCntrNo AND ibfCurrent.ibfPodCPMID = pcf.pcfPodCPMID AND ibfCurrent.ibfBillYear=" & ilYear & " AND ibfCurrent.ibfBillMonth=" & ilMonth & " "
    SQLQuery = SQLQuery & "    JOIN ""SAF_Schd_Attributes"" saf ON saf.safVefCode=0 "
    SQLQuery = SQLQuery & "WHERE "
    If lmCntrNo = 0 Then
        'all Contracts
        SQLQuery = SQLQuery & "    chf.chfAdServerDefined = 'Y' AND "
        SQLQuery = SQLQuery & "    chf.chfStartdate <= chf.chfEndDate AND "
        SQLQuery = SQLQuery & "    chf.chfStartdate <= '" & slEndDate & "' AND "
        SQLQuery = SQLQuery & "    chf.chfEndDate >= '" & slStartDate & "' AND "
        SQLQuery = SQLQuery & "    chf.chfDelete <> 'Y' AND "
        'TTP 10685 - Ad Server Billing Discrepancy export: filters out lines with 0 impressions ordered
        'SQLQuery = SQLQuery & "    chf.chfStatus in ('H','O','G','N') AND "
        'SQLQuery = SQLQuery & "    pcf.pcfImpressionGoal <> 0 "
        SQLQuery = SQLQuery & "    chf.chfStatus in ('H','O','G','N') "
    Else
        'Specific Contract
        SQLQuery = SQLQuery & "    chf.chfAdServerDefined = 'Y' AND "
        SQLQuery = SQLQuery & "    chf.chfStartdate <= chf.chfEndDate AND "
        SQLQuery = SQLQuery & "    chf.chfStartdate <= '" & slEndDate & "' AND "
        SQLQuery = SQLQuery & "    chf.chfEndDate >= '" & slStartDate & "' AND "
        SQLQuery = SQLQuery & "    chf.chfDelete <> 'Y' AND"
        SQLQuery = SQLQuery & "    chf.chfStatus in ('H','O','G','N') AND "
        'TTP 10685 - Ad Server Billing Discrepancy export: filters out lines with 0 impressions ordered
        'SQLQuery = SQLQuery & "    pcf.pcfImpressionGoal <> 0 AND "
        SQLQuery = SQLQuery & "    chf.chfCntrNo = " & lmCntrNo
    End If
    SQLQuery = SQLQuery & "GROUP BY "
    SQLQuery = SQLQuery & "    saf.saffeatures7, "
    SQLQuery = SQLQuery & "    chf.chfCntrNo, "
    SQLQuery = SQLQuery & "    chf.chfExtCntrNo, " 'TTP 10681 - Ad Server Billing Discrepancy Export: add External Contract Number and External Line Number
    SQLQuery = SQLQuery & "    chf.chfBillCycle, "
    SQLQuery = SQLQuery & "    chf.chfProduct, "
    SQLQuery = SQLQuery & "    chf.chfadfcode, "
    SQLQuery = SQLQuery & "    pcf.pcfStartDate, "
    SQLQuery = SQLQuery & "    pcf.pcfEndDate, "
    SQLQuery = SQLQuery & "    pcf.pcfPriceType, "
    SQLQuery = SQLQuery & "    pcf.pcfDeliveryType, " '09/28/2022 - JW - Ad server billing discrepancy export: show "baked-in" price type
    SQLQuery = SQLQuery & "    pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & "    pcf.pcfExtLineNo, " 'TTP 10681 - Ad Server Billing Discrepancy Export: add External Contract Number and External Line Number
    SQLQuery = SQLQuery & "    pcf.pcfImpressionGoal, "
    SQLQuery = SQLQuery & "    pcf.pcfVefCode, "
    SQLQuery = SQLQuery & "    avf.avfName, "
    SQLQuery = SQLQuery & "    rdf.rdfName, "
    SQLQuery = SQLQuery & "    pcf.pcfrdfcode, "
    SQLQuery = SQLQuery & "    ibfCurrent.ibfImpressions "
    
    On Error GoTo ERRORBOX
    Set olRs = gSQLSelectCall(SQLQuery)
    mQueryDatabaseCSV = "No errors"
    Exit Function
    
ERRORBOX:
    mQueryDatabaseCSV = "Problem with query in mQueryDatabase. "
End Function

