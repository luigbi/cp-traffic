VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpCntrLine 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   9885
   Begin V81TrafficExports.CSI_Calendar CSI_CalEnd 
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Text            =   "01/08/2024"
      BackColor       =   16776960
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin V81TrafficExports.CSI_Calendar CSI_CalStart 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Text            =   "01/08/2024"
      BackColor       =   16776960
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
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
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin VB.Frame frcAmazon 
      Height          =   1455
      Left            =   7680
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   9615
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
      Begin VB.CheckBox ckcKeepLocalFile 
         Caption         =   "Keep Local File"
         Height          =   195
         Left            =   6000
         TabIndex        =   13
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox edcBucketName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox edcAccessKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   11
         ToolTipText     =   "The Access Key Assigned by AWS"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox edcPrivateKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6000
         PasswordChar    =   "*"
         TabIndex        =   12
         ToolTipText     =   "The Private Key Assigned by AWS"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox edcRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         ToolTipText     =   "Region/Endpoint - Example: USEast1, USEast2, USWest1 or USWest2"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Folder (optional)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lacExportFilename 
         Caption         =   "lacExportFilename"
         Height          =   255
         Left            =   7800
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "BucketName"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "AccessKey"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PrivateKey"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
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
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   4
      Top             =   960
      Width           =   1200
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   3075
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
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
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
      Left            =   7560
      TabIndex        =   6
      Top             =   1560
      Width           =   1485
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6450
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5835
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6105
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1935
      Visible         =   0   'False
      Width           =   525
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
      TabIndex        =   14
      Top             =   3000
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
      Left            =   5505
      TabIndex        =   1
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7320
      Top             =   2040
   End
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7800
      Top             =   2040
   End
   Begin VB.PictureBox plcTo 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   5685
      TabIndex        =   21
      Top             =   1560
      Width           =   5745
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   5625
      End
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6840
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active End Date"
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
      Height          =   195
      Left            =   3120
      TabIndex        =   31
      Top             =   480
      Width           =   1410
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
      TabIndex        =   30
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label lacSelCFrom 
      Appearance      =   0  'Flat
      Caption         =   "Active Start Date"
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
      TabIndex        =   29
      Top             =   510
      Width           =   1545
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   3360
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
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
      TabIndex        =   27
      Top             =   1590
      Width           =   810
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   6300
   End
End
Attribute VB_Name = "ExpCntrLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCntrLine.Frm
'
' Description:
'   This file contains TTP 10233 -Audacy: "Contract line export" input screen code
'   8/17/21 - JW
Option Explicit
Option Compare Text
Dim imExporting As Integer
Dim imTerminate As Integer
Dim imFirstActivate As Integer
Dim imExportOption As Integer       'lbcExport.ItemData(lbcExport.ListIndex)
Dim smClientName As String
Dim myBucket As CsiToAmazonS3.ApiCaller
Dim hmMsg As Integer                'From file hanle
Dim lmNowDate As Long               'Todays date
Dim lmCntrNo As Long                'for debugging purposes to filter a single contract
Dim smExportName As String
Dim smExportOptionName As String
Dim smExportFilename As String
Dim imNumberDecPlaces As Integer
Dim imAdjDecPlaces As Integer
Dim sm1or2PlaceRating As String
    
'MsgBox parameters
Const vbOkOnly = 0                  'OK button only
Const vbCritical = 16               'Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0

'CHF
Dim hmCHF As Integer                'Contract header file handle
Dim tmChf As CHF
Dim imCHFRecLen As Integer

'CLF
Dim hmClf As Integer                'Contract line file handle
Dim tmClfSrchKey As CLFKEY0         'CLF record image
Dim imClfRecLen As Integer          'CLF record length
Dim tmClf As CLF

'CFF
Dim hmCff As Integer
Dim tmCff As CFF
Dim ilCff As Integer
Dim imCffRecLen As Integer

'MNF
Dim hmMnf As Integer        'Multiname file handle
Dim imMnfRecLen As Integer  'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF

'MNF Salesperson stuff
Dim tmMnfSS() As MNF        'array of Sales Sources MNF
Dim tmMnfGroups() As MNF

'CEF
Dim hmCef As Integer        'comment for other type
Dim tmCef As CEF
Dim imCefRecLen As Integer
Dim tmCefSrchKey0 As LONGKEY0

'DRF
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim imDrfRecLen As Integer  'DRF record length

'DPF
Dim hmDpf As Integer        'Demo Plus file handle
Dim tmDpf As DPF            'DPF record image
Dim imDpfRecLen As Integer  'DPF record length

'SOF
Dim hmSof As Integer        'Office file handle
Dim imSofRecLen As Integer  'SOF record length
Dim tmSof() As SOF

'hmDef
Dim hmDef As Integer

'hmRaf
Dim hmRaf As Integer        'RAF file handle

'SBF
Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim imSbfRecLen As Integer  'SBF record length
Dim tlSBFTypes As SBFTypes
Dim tlSbf() As SBF

'Research Arrays
Dim lmRtg() As Integer
Dim lmGrimp() As Long
Dim lmGRP() As Long
Dim lmCost() As Long

'Export Data Array
Dim tmWOLINEEXPDATA() As EXPWOINVLN

Private Sub ckcAmazon_Click()
    If ckcAmazon.Value = vbChecked Then
        frcAmazon.Left = 120
        frcAmazon.Top = 1320
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
    ReDim tmSof(0) As SOF
    
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
    'slStr = edcMonth.Text             'month in text form (jan..dec, or 1-12
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
    End If
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    
    lmCntrNo = 0                'ths is for debugging on a single contract
    slStr = edcContract
    If slStr <> "" Then
        lmCntrNo = Val(slStr)
    End If

    If lmCntrNo = 0 Then
        If CSI_CalStart.Text = "" Then
            Beep
            CSI_CalStart.SetFocus
            Exit Sub
        End If
        
        If CSI_CalEnd.Text = "" Then
            Beep
            CSI_CalEnd.SetFocus
            Exit Sub
        End If
        
        If CSI_CalStart.Text <> "" And CSI_CalEnd.Text <> "" Then
            If DateValue(CSI_CalEnd.Text) < DateValue(CSI_CalStart.Text) Then
                Beep
                CSI_CalEnd.SetFocus
                Exit Sub
            End If
        End If
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
    
    ilRet = gFileExist(smExportName)
    If ilRet = 0 Then
        'file already exists, do not overwrite
        ''MsgBox "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        gAutomationAlertAndLogHandler "Filename already exists, enter new name", vbOkOnly + vbApplicationModal, "Save In"
        Exit Sub
    End If

    If Not mOpenMsgFile() Then          'open message file
         cmcCancel.SetFocus
         Exit Sub
    End If
    ilRet = 0
    
    gAutomationAlertAndLogHandler "Contract Line Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    'Print #hmMsg, "** Storing Output into " & smExportName & " **"
    gAutomationAlertAndLogHandler "* Storing Output into " & smExportName
    gAutomationAlertAndLogHandler "* OptionName=" & smExportOptionName
    gAutomationAlertAndLogHandler "* ActiveStartDate=" & CSI_CalStart.Text
    gAutomationAlertAndLogHandler "* ActiveEndDate=" & CSI_CalEnd.Text
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

    '---------------------------------------------
    'Open Some Pervasive File handles
    hmCef = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        Exit Sub
    End If
    
    hmDrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    imDrfRecLen = Len(tmDrf)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        Exit Sub
    End If
    
    hmDpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    imDpfRecLen = Len(tmDpf)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        Exit Sub
    End If
    
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    'imSofRecLen = Len(hmSof)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        Exit Sub
    End If
    
    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
        Exit Sub
    End If
    
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        Exit Sub
    End If
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmSbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpCntrLine
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        ilRet = btrClose(hmDef)
        btrDestroy hmDef
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        Exit Sub
    End If
    imSbfRecLen = Len(tmSbf)
    
    '-----------------------------------------------------
    'Build array of the vehicle group codes and names
    ilRet = gObtainMnfForType("H", slTimeStamp, tmMnfGroups())

    '-----------------------------------------------------
    'Query Database
    slErrorMessage = mQueryDatabase(olRs)
    
    If Mid(slErrorMessage, 1, 9) = "No errors" Then
        'use Recordset for each Contract/proposal, build array of Contract Lines with Research
        gAutomationAlertAndLogHandler "Processing Data..."
        If Mid(slErrorMessage, 1, 9) = "No errors" Then
            slErrorMessage = mProcessData(olRs)
        Else
            gAutomationAlertAndLogHandler "Process Data Error: " & slErrorMessage
        End If
    
        '-------------------------------------------------
        'Export
        If Mid(slErrorMessage, 1, 9) = "No errors" Then
            gAutomationAlertAndLogHandler "Exporting Data..."
            slErrorMessage = mExportData(olRs, smExportName, ",")
        Else
            gAutomationAlertAndLogHandler "ProcessData Error: " & slErrorMessage
        End If
    Else
        gAutomationAlertAndLogHandler "Query Database Error: " & slErrorMessage
    End If
    
    '-----------------------------------------------------
    If Mid(slErrorMessage, 1, 9) = "No errors" Then
        lacInfo(0).Caption = "Export created... " & slErrorMessage: lacInfo(0).Refresh
        'Print #hmMsg, "** Export Created: " & slErrorMessage
        gAutomationAlertAndLogHandler "** Export Completed: " & slErrorMessage
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

Private Sub CSI_CalEnd_Change()
    gCtrlGotFocus CSI_CalEnd
    If igExportType <= 1 And Not imFirstActivate Then
        mGetExportFilename
    End If
End Sub

Private Sub CSI_CalEnd_GotFocus()
    gCtrlGotFocus CSI_CalEnd
End Sub

Private Sub CSI_CalStart_Change()
    gCtrlGotFocus CSI_CalStart
    If igExportType <= 1 And Not imFirstActivate Then
        mGetExportFilename
    End If
End Sub

Private Sub CSI_CalStart_GotFocus()
    gCtrlGotFocus CSI_CalStart
End Sub

Private Sub edcContract_Change()
    mGetExportFilename
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    
    mGetExportFilename
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
    If imExportOption = EXP_AUDACYINVLINE Then
    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
        If (tgSpfx.iInvExpFeature And INVEXP_AUDACYLINE) <> INVEXP_AUDACYLINE Then
            lacInfo(0).Caption = "Contract Line Export Disabled"
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
        If imExportOption = EXP_AUDACYINVLINE Then
            cmcExport.Enabled = True
            'gOpenTmf
            'tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
            'tmcSetTime.Enabled = True
            'gUpdateTaskMonitor 1, "WL"
            cmcExport_Click
            'gUpdateTaskMonitor 2, "WL"
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

    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    
    ilRet = btrClose(hmDef)
    btrDestroy hmRaf
    
    ilRet = btrClose(hmRaf)
    btrDestroy hmDef
    
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    
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
    Dim ilStartPeriod As Integer
    Dim ilStartMonth As Integer
    Dim ilMonths As Integer
    Dim ilCalendar As Integer '0=Std, 1=Cal
    
    slMonthStr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imExporting = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    
    gCenterStdAlone Me
    
    'get Export Option #
    imExportOption = ExportList!lbcExport.ItemData(ExportList!lbcExport.ListIndex)
    If imExportOption = EXP_AUDACYINVLINE Then
        smExportOptionName = "CntrLine"
    Else
        smExportOptionName = ""
    End If
    
    'get "Research in" unit of measure
    If tgSpf.sSAudData = "H" Then
        imNumberDecPlaces = 1
        imAdjDecPlaces = 10
    ElseIf tgSpf.sSAudData = "N" Then
        imNumberDecPlaces = 2
        imAdjDecPlaces = 100
    ElseIf tgSpf.sSAudData = "U" Then
        imNumberDecPlaces = 3
        imAdjDecPlaces = 1000
    Else
        imNumberDecPlaces = 0
        imAdjDecPlaces = 1
    End If
        
    
    'determine default month year
    slDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    
    'Default to current month, based on today's system date
    ilMonth = Val(slMonth)
    ilYear = Val(slYear)

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
        
        'Calendar: can be set to Cal (for calendar month) or Std (for standard broadcast month), and is used when calculating the starting month to determine which calendar type to use.
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Calendar", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'Use the Start Month/Year that's already defaulted (1 month prior to today)
            ilCalendar = 0
        Else
            If UCase(Trim$(gStripChr0(slReturn))) = "CAL" Then
                ilCalendar = 1
            End If
        End If
                
        'StartPeriod: the number of months specified as the "StartPeriod" value will be subtracted from the current month number using the "Calendar" setting to determine the starting month number for the active start date. This is the "rolling month" option. For example, if the current month using today's date is June 2021, and the StartPeriod is set to 1, the export would set the active start date to the first day of May 2021 (one month prior to the current month).
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "StartPeriod", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'Use the Start Month/Year that's already defaulted (1 month prior to today)
            CSI_CalStart.Text = gObtainStartStd(Format$(lmNowDate, "m/d/yy"))
        Else
            ilStartPeriod = Val(Trim$(gStripChr0(slReturn)))
            If ilStartPeriod > 0 Then
                If ilCalendar = 1 Then
                    CSI_CalStart.Text = gObtainStartCal(DateAdd("m", -ilStartPeriod, gObtainEndCal(Format$(lmNowDate, "m/d/yy"))))
                Else
                    CSI_CalStart.Text = gObtainStartStd(DateAdd("m", -ilStartPeriod, gObtainEndStd(Format$(lmNowDate, "m/d/yy"))))
                End If
            End If
        End If
        
        'StartMonth: this is an optional setting that can be set to a number between 1 and 12, with 1 for January, 2 for February, etc. When used, this setting makes the export always start from the specified month of the current year, overriding the StartPeriod parameter (if defined). This is the "fixed start month" option, and should be used if the export always needs to start from the same month.
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "StartMonth", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'Do nothing
        Else
            ilStartMonth = Val(Trim$(gStripChr0(slReturn)))
            If ilStartMonth > 0 And ilStartMonth < 13 Then
                If ilCalendar = 1 Then
                    CSI_CalStart.Text = gObtainStartCal(ilStartMonth & "/1/" & ilYear)
                Else
                    CSI_CalStart.Text = gObtainStartStd(ilStartMonth & "/1/" & ilYear)
                End If
            End If
        End If
        
        'Months: the number of months for the export to use when determining the active start and end dates, starting from the starting month as determined by the StartPeriod or StartMonth value.
        ilRet = GetPrivateProfileString(sgExportIniSectionName, "Months", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'Do nothing
            CSI_CalEnd.Text = gObtainEndStd(Format$(lmNowDate, "m/d/yy"))
        Else
            ilMonths = Val(Trim$(gStripChr0(slReturn))) - 1
            If ilMonths > -1 Then
                If CSI_CalStart.Text <> "" Then
                    If ilCalendar = 1 Then
                        CSI_CalEnd.Text = gObtainEndCal(DateAdd("m", ilMonths, DateValue(gObtainEndCal(CSI_CalStart.Text))))
                    Else
                        ilMonth = Month(CSI_CalStart.Text)
                        ilYear = Year(CSI_CalStart.Text)
                        CSI_CalEnd.Text = gObtainEndStd(DateAdd("m", ilMonths + 1, DateValue(ilMonth & "/15/" & ilYear)))
                    End If
                End If
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
    
    '------------------------
    'Get Client name or Abv
    hmMnf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    imMnfRecLen = Len(tmMnf)  'Get and save ADF record length
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        imTerminate = True
        Exit Sub
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
        mGetExportFilename
    End If
    
    Screen.MousePointer = vbDefault
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
    slToFile = sgDBPath & "Messages\" & Trim(smExportOptionName) & ".Txt"
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
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, "Contract Line Export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    'Print #hmMsg, ""
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
    If smExportOptionName = "CntrLine" Then
        plcScreen.Print "Contract Line Export"
    Else
        plcScreen.Print smExportOptionName & " Export"
    End If
End Sub

Private Sub mGetExportFilename()
    Dim slRepeat As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slStr As String
    Dim ilYear As Integer
    Dim slExtension As String * 4
    Dim lmStart As Long         'Starting date
    Dim lmEnd As Long           'Ending date
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilTemp As Integer
    Dim slStart As String
    Dim slEnd As String

    'Start Date
    lmStart = gDateValue(CSI_CalStart.Text)
    gObtainYearMonthDayStr CSI_CalStart.Text, True, slYear, slMonth, slDay
    slStart = Trim$(slMonth) & Trim$(slDay) & Mid(slYear, 3, 2)
    
    'End Date
    lmEnd = gDateValue(CSI_CalEnd.Text)
    gObtainYearMonthDayStr CSI_CalEnd.Text, True, slYear, slMonth, slDay
    slEnd = Trim$(slMonth) & Trim$(slDay) & Mid(slYear, 3, 2)
    
    smExportFilename = ""
    tmcClick.Enabled = False
    
    'Determine name of export
    slExtension = ".csv"
    slRepeat = "A"
    Select Case imExportOption
        Case EXP_AUDACYINVLINE
            slExtension = ".csv"
    End Select
    
    'build Filename
    Do
        ilRet = 0
        If Val(edcContract.Text) > 0 Then
            smExportFilename = Trim$(smExportOptionName) & " Cntr" & Val(edcContract.Text) & " " & gFileNameFilter(Trim$(slRepeat & " " & Trim$(smClientName))) & slExtension
            smExportName = Trim$(sgExportPath) & Trim$(smExportOptionName) & " Cntr" & Val(edcContract.Text) & " "
            'TTP 10559 - Contract Line export: append generation date and time to filename
            smExportName = smExportName & Format(Date, "MMDDYY") & "_" & Format(Time, "HHMM")
            smExportName = Trim$(smExportName) & gFileNameFilter(slRepeat & " " & Trim$(smClientName)) & slExtension
        Else
            smExportFilename = Trim$(smExportOptionName) & " " & slStart & "-" & slEnd & " " & gFileNameFilter(Trim$(slRepeat & " " & Trim$(smClientName))) & slExtension
            smExportName = Trim$(sgExportPath) & Trim$(smExportOptionName) & " " & slStart & "-" & slEnd & " "
            'TTP 10559 - Contract Line export: append generation date and time to filename
            smExportName = smExportName & Format(Date, "MMDDYY") & "_" & Format(Time, "HHMM")
            smExportName = Trim$(smExportName) & gFileNameFilter(slRepeat & " " & Trim$(smClientName)) & slExtension
        End If
        
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
'    If imExportOption = EXP_AUDACYINVLINE Then
'        gUpdateTaskMonitor 0, "PBD"
'    End If
End Sub

Private Function mGetHeaderString() As String
    mGetHeaderString = "Contract Number,External Version Number,Proposal Version Number,Contract Type,Salesperson,Salesperson email,Salesperson ID,Sales Office,Sales Office ID,Agency Name,Agency ID,External Agency ID,Advertiser Name,Advertiser ID,External Advertiser ID,Product Name,Cash/Trade,Trade Percentage,Air Time/NTR,Demo,Status,Revenue Set 1,Revenue Set 2,Revenue Set 3,Vehicle Name,Vehicle ID,Market,Research,SubCompany,Format,SubTotals,Spot Length,Daypart,Line Type,Line Start Date,Line End Date,Total Units,Price Type,Total Gross,Rating,Line CPP,Line GRPs,CPM,Average Audience,Line Gross Impressions,NTR_Billing Date,NTR Description,NTR Type,Amount Per NTR Item,Number of NTR Items"
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

Private Function mQueryDatabase(ByRef olRs As Recordset) As String
    'This returns the Column Names Exactly as they should appear in the export, with a prepended format character (like s=string, n=numeric); example: "sCall Letters"
    Dim slErrorMessage As String
    Dim SQLQuery As String
    Dim blNeedAnd As Boolean 'For Query building
    Dim lLoop As Integer
    Dim lmStart As Long
    Dim lmEnd As Long
    Dim slStartDate As String
    Dim slEndDate As String
    
    lmStart = gDateValue(CSI_CalStart.Text) 'Start Date
    lmEnd = gDateValue(CSI_CalEnd.Text) 'End Date
    slStartDate = Format(lmStart, "yyyy-mm-dd") 'SQL Start Date
    slEndDate = Format(lmEnd, "yyyy-mm-dd") 'SQL End Date
    SQLQuery = " SELECT "
    SQLQuery = SQLQuery & " chf.chfCode AS 'lchfCode', "
    SQLQuery = SQLQuery & " chf.chfCntrNo AS 'lContract_Number', "
    SQLQuery = SQLQuery & " chf.chfextRevNo AS 'lExternal_Version_Number', "
    SQLQuery = SQLQuery & " chf.chfPropVer AS 'iProposal_Version_Number', "
    SQLQuery = SQLQuery & " CASE chf.chfType "
    SQLQuery = SQLQuery & " WHEN 'C' THEN 'Standard' "
    SQLQuery = SQLQuery & " WHEN 'V' THEN 'Reservation' "
    SQLQuery = SQLQuery & " WHEN 'R' THEN 'Direct Response' "
    SQLQuery = SQLQuery & " WHEN 'Q' THEN 'Per Inquiry' "
    SQLQuery = SQLQuery & " WHEN 'S' THEN 'PSA' "
    SQLQuery = SQLQuery & " WHEN 'M' THEN 'Promo' "
    SQLQuery = SQLQuery & " ELSE '' END AS 'sContract_Type', "
    SQLQuery = SQLQuery & " chf.chfslfCode1 AS 'iSalesperson_ID', "
    SQLQuery = SQLQuery & " slf.slfsofcode AS 'iSales_Office_ID', "
    SQLQuery = SQLQuery & " isnull(agf.agfCode,0) AS 'iAgency_ID', "
    SQLQuery = SQLQuery & " rtrim(agfx.agfxRefId) AS 'sExternal_Agency_ID', "
    SQLQuery = SQLQuery & " adf.adfCode AS 'iAdvertiser_ID', "
    SQLQuery = SQLQuery & " rtrim(adfx.adfxRefId) AS 'iExternal_Advertiser_ID', "
    SQLQuery = SQLQuery & " rtrim(chf.chfProduct) AS 'sProduct_Name', "
    SQLQuery = SQLQuery & " if(chf.chfPctTrade>0,'T','C') AS 'sCash/Trade', "
    SQLQuery = SQLQuery & " chf.chfPctTrade AS 'iTrade_Percentage', "
    SQLQuery = SQLQuery & " rtrim(mnfDemo.mnfName) AS 'sDemo', "
    SQLQuery = SQLQuery & " CASE chf.chfStatus "
    SQLQuery = SQLQuery & " WHEN 'W' THEN 'Working Proposal' "
    SQLQuery = SQLQuery & " WHEN 'D' THEN 'Rejected' "
    SQLQuery = SQLQuery & " WHEN 'C' THEN 'Completed Proposal' "
    SQLQuery = SQLQuery & " WHEN 'I' THEN 'Unapproved Proposal' "
    SQLQuery = SQLQuery & " WHEN 'H' THEN 'Approved Hold' "
    SQLQuery = SQLQuery & " WHEN 'G' THEN 'Approved Hold' "
    SQLQuery = SQLQuery & " WHEN 'N' THEN 'Approved Order' "
    SQLQuery = SQLQuery & " WHEN 'O' THEN 'Approved Order' "
    SQLQuery = SQLQuery & " ELSE '' END AS 'sStatus', "
    SQLQuery = SQLQuery & " rtrim(mnfRevSet1.mnfName)       AS 'sRevenue_Set_1', "
    SQLQuery = SQLQuery & " rtrim(mnfRevSet2.mnfName)       AS 'sRevenue_Set_2', "
    SQLQuery = SQLQuery & " rtrim(mnfRevSet3.mnfName)       AS 'sRevenue_Set_3' "
    
    SQLQuery = SQLQuery & " FROM "
    SQLQuery = SQLQuery & " ""CHF_Contract_Header"" chf "
    SQLQuery = SQLQuery & " JOIN (SELECT "
    SQLQuery = SQLQuery & " chfCntrNo, MAX(chfCntRevNo) as chfCntRevNo "
    SQLQuery = SQLQuery & " FROM "
    SQLQuery = SQLQuery & " ""CHF_Contract_Header"" chf"
    If lmCntrNo = 0 Then
        'all Contracts in the Date Range
        SQLQuery = SQLQuery & " WHERE "
        SQLQuery = SQLQuery & " chf.chfStartDate <= '" & slEndDate & "' AND "
        SQLQuery = SQLQuery & " chf.chfEndDate >= '" & slStartDate & "'"
        'SQLQuery = SQLQuery & " AND chf.chfStatus in ('W','H','O','G','N','C','I')"
    Else
        'Specific Contract
        SQLQuery = SQLQuery & " WHERE "
        SQLQuery = SQLQuery & " chf.chfCntrNo = " & lmCntrNo
        'SQLQuery = SQLQuery & " AND chf.chfStatus in ('W','H','O','G','N','C','I')"
    End If
    SQLQuery = SQLQuery & " GROUP BY chfCntrNo "
    SQLQuery = SQLQuery & " ) LastRev on LastRev.ChfCntrNo = chf.ChfCntrNo and LastRev.chfCntRevNo = chf.chfCntRevNo "
    SQLQuery = SQLQuery & " LEFT JOIN ""SLF_Salespeople"" slf ON slf.slfcode = chf.chfslfCode1 "
    SQLQuery = SQLQuery & " LEFT JOIN ""AGF_Agencies"" agf ON agf.agfCode = chf.chfagfCode "
    SQLQuery = SQLQuery & " LEFT JOIN ""AGFX_Agencies"" agfx ON agfx.agfxCode = agf.agfCode "
    SQLQuery = SQLQuery & " LEFT JOIN ""ADF_Advertisers"" adf ON adf.adfCode = chf.chfadfCode "
    SQLQuery = SQLQuery & " LEFT JOIN ""ADFX_Advertisers"" adfx ON adfx.adfxCode = adf.adfCode "
    SQLQuery = SQLQuery & " LEFT JOIN ""MNF_Multi_Names"" mnfDemo ON mnfDemo.mnfCode = chf.chfmnfDemo1 "
    SQLQuery = SQLQuery & " LEFT JOIN ""MNF_Multi_Names"" mnfRevSet1 ON mnfRevSet1.mnfCode = chf.chfmnfRevBk1 "
    SQLQuery = SQLQuery & " LEFT JOIN ""MNF_Multi_Names"" mnfRevSet2 ON mnfRevSet2.mnfCode = chf.chfmnfRevBk2 "
    SQLQuery = SQLQuery & " LEFT JOIN ""MNF_Multi_Names"" mnfRevSet3 ON mnfRevSet3.mnfCode = chf.chfmnfRevBk3 "
    
    'Debug.Print SQLQuery
    
    On Error GoTo ERRORBOX
    Set olRs = gSQLSelectCall(SQLQuery)
    mQueryDatabase = "No errors"
    Exit Function
    
ERRORBOX:
    mQueryDatabase = "Problem with query in mQueryDatabase. "
End Function

'
'               Search the array of vehicle groups (tmMnfGroups)
'               <input> ilMnfCode = Multiname code
'               Return : -1 if not found, else index to the vehicle group item
Private Function mBinarySearchMnfVehicleGroup(ilMnfCode As Integer)
    Dim ilMiddle As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    ilMin = LBound(tmMnfGroups)
    ilMax = UBound(tmMnfGroups) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilMnfCode = tmMnfGroups(ilMiddle).iCode Then
            'found the match
            mBinarySearchMnfVehicleGroup = ilMiddle
            Exit Function
        ElseIf ilMnfCode < tmMnfGroups(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchMnfVehicleGroup = -1
    'Sort by code so that binary search can be used
    If UBound(tmMnfGroups) > 1 Then
        ArraySortTyp fnAV(tmMnfGroups(), 0), UBound(tmMnfGroups), 0, LenB(tmMnfGroups(0)), 0, -1, 0
    End If
End Function

Function mProcessData(ByRef olRs As Recordset) As String
    'Create a CSV file, give me a recordset and a (fully qualified path\filename.ext) filename
    'Makes Headers from recordset Column Names
    'Makes rows from Data
    Dim slErrorMessage As String
    'Dim olFileSys As FileSystemObject
    Dim slPath As String
    
    'Loops / timers
    Dim llRecords As Long
    Dim llMatch As Long
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim llLoopPkg As Long
    Dim ilLoop As Integer
    Dim ilClf As Integer
    'Lookup Values
    Dim slLineType As String
    Dim slPriceType As String
    Dim ilAdfCode As Integer
    Dim ilAgfCode As Integer
    Dim ilRdfCode As Integer
    Dim ilLastSlspID As Integer
    Dim ilSlspID As Integer
    Dim ilSlspEmailID As Integer
    Dim ilVefCode As Integer
    Dim iTVefCode As Integer
    Dim slAdfName As String
    Dim slAgfName As String
    Dim slRdfName As String
    Dim slSlspName As String
    Dim slSlspEmail As String
    Dim ilLastSlsOfficeID As Integer
    Dim slSofName As String
    Dim ilLastVefID As Integer
    Dim slVefName As String
    Dim slMnfVehGp3Mkt As String
    Dim slMnfVehGp5Rsch As String
    Dim slMnfVehGp6Sub As String
    Dim slMnfVehGp4Fmt As String
    Dim slvefMnfVehGp2 As String
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Integer
    Dim slNTRType As String
    Dim slCBS As String
    
    'Pkg Hangling
    Dim ilHiddenLines As Integer
    Dim ilPackageDnfCode As Integer
    'Ranges
    Dim llFltStart As Long      'CFF Start Long
    Dim llFltEnd As Long        'CFF End Long
    Dim llClfStartDate As Long  'CLF Start Long
    Dim llClfEndDate As Long    'CLF End Long
    Dim slCLFStartDate As String 'CLF Start String
    Dim slCLFEndDate As String  'CLF Date String
    Dim slStartDate As String   'Generic Start Date String
    Dim slEndDate As String     'Generic End Date String
    Dim slDate As String        'Generic Date String
    'Research
    Dim slGrossNet As String    'G=Gross, N=Net; when using Research subs, return in Gross or Net
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llTotalGrImp As Long
    Dim llTotalGRP As Long
    Dim llTotalCPP As Long
    Dim llPop As Long
    Dim llTotalAvgAud As Long
    Dim ilTotalAvgRtg As Integer
    Dim llTotalCPM As Long
    Dim llPopEst As Long
    Dim llAvgAud As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim llRate As Long
    Dim slStr As String
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llDate As Long
    Dim llDate2 As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilSpots As Integer
    Dim ilDay As Integer
    Dim ilUpperWk As Integer
    Dim ilTotLnSpts As Integer
    Dim ilWinx As Integer
    Dim llSpots As Long
    ReDim tmWOLINEEXPDATA(0) As EXPWOINVLN
    ReDim ilInputDays(0 To 6) As Integer    'valid days of the week for audience retrieval
    Dim ilFoundSpot As Integer
    On Error GoTo ERRORBOX
    'On Error GoTo 0
        
    slGrossNet = "G"
    '-------------------------------------------
    'Check Records
    If olRs.EOF And olRs.BOF Then
        mProcessData = "There are no records to export"
        GoTo finish
    End If
    
    '-------------------------------------------
    'Loop through Contracts in Recordset
    olRs.MoveFirst
    Do While Not olRs.EOF
        '------------------------------------------------------------------------------------------------
        'Contract Header
        '------------------------------------------------------------------------------------------------
        '-------------------------------------------
        '"sSalesperson"
        ilSlspID = olRs.Fields("iSalesperson_ID").Value
        
        If ilLastSlspID <> ilSlspID Then 'lookup SlsPsn Name and Email
            'get Salesperson Name
            gObtainSalespersonName ilSlspID, slSlspName, True
            '"sSalesperson_email"
            slSlspEmail = ""
            ilSlspEmailID = -1
            'Salemsan Comment ID: SalemanID -> UserID -> Email (SLF ->  URFSLFCode -> URF)
            For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) Step 1
                If tgPopUrf(ilLoop).iSlfCode = ilSlspID Then
                    ilSlspEmailID = tgPopUrf(ilLoop).lEMailCefCode
                    Exit For
                End If
            Next ilLoop
            If ilSlspEmailID <> -1 Then
                tmCefSrchKey0.lCode = ilSlspEmailID ' Look for the comment for this Saleman (user)
                imCefRecLen = Len(tmCef)
                ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slSlspEmail = gStripChr0(tmCef.sComment)
                End If
            End If
            ilLastSlspID = ilSlspID
        End If
        
        '-------------------------------------------
        '"sSales_Office"
        If ilLastSlsOfficeID <> olRs.Fields("iSales_Office_ID").Value Then
            ilLastSlsOfficeID = olRs.Fields("iSales_Office_ID").Value
            slSofName = mGetSOFName(olRs.Fields("iSales_Office_ID").Value)
        End If
        
        '-------------------------------------------
        '"sAgency_Name"
        ilAgfCode = IIF(IsNull(olRs.Fields("iAgency_ID").Value), 0, olRs.Fields("iAgency_ID").Value)
        'TTP 10500
        If ilAgfCode = 0 Or ilAgfCode = -1 Then
            slAgfName = "Direct"
        Else
            slAgfName = ""
        End If
        ilRet = gBinarySearchAgf(ilAgfCode)
        If ilRet <> -1 Then
            slAgfName = Trim$(tgCommAgf(ilRet).sName)
        End If
        
        '-------------------------------------------
        '"sAdvertiser_Name" (ADF Name Lookup)
        ilAdfCode = olRs.Fields("iAdvertiser_ID").Value
        slAdfName = ""
        ilRet = gBinarySearchAdf(ilAdfCode)
        If ilRet <> -1 Then
            slAdfName = Trim$(tgCommAdf(ilRet).sName)
        End If

        '-------------------------------------------
        'Going Perversive here...
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, olRs.Fields("lChfCode").Value, False, tgChfCT, tgClfCT(), tgCffCT())
        
        gUnpackDate tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), slStartDate
        llStartDate = gDateValue(slStartDate)
        
        gUnpackDate tgChfCT.iEndDate(0), tgChfCT.iEndDate(1), slEndDate
        llEndDate = gDateValue(slEndDate)
        
        sm1or2PlaceRating = gSet1or2PlaceRating(tgChfCT.iAgfCode)
        lacInfo(0).Caption = "Processing " & llRecords & " Records..": lacInfo(0).Refresh
        
        'Debug.Print "CHFCode:" & olRs.Fields("lChfCode").Value & ", Cntr: " & olRs.Fields("lContract_Number").Value
        '------------------------------------------------------------------------------------------------
        'Contract Lines
        '------------------------------------------------------------------------------------------------
        For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
            tmClf = tgClfCT(ilClf).ClfRec
            ReDim lmRtg(0 To 105)                        'setup arrays for return values from audtolnresearch
            ReDim lmGrimp(0 To 105)
            ReDim lmGRP(0 To 105)
            ReDim lmCost(0 To 105)                       'setup arrays for return values from audtolnresearch
            ReDim llWklyspots(0 To 105) As Long          'sched lines weekly # spots
            ReDim llWklyAvgAud(0 To 105) As Long         'sched lines weekly avg aud
            ReDim llWklyRates(0 To 105) As Long          'sched lines weekly rates
            ReDim llWklyPopEst(0 To 105) As Long
            
            'initialize for the next line
            llSpots = 0
            llPop = -1
            llTotalAvgAud = 0
            llTotalCPM = 0
            llTotalGRP = 0
            ilTotalAvgRtg = 0
            dlTotalCost = 0 'TTP 10439 - Rerate 21,000,000
            llTotalCPP = 0
            ilTotLnSpts = 0
            llTotalGrImp = 0
            llPopEst = 0
                                            
            'Flight StartDate
            gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slCLFStartDate
            llClfStartDate = gDateValue(slCLFStartDate)
            'Flight EndDate
            gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slCLFEndDate
            llClfEndDate = gDateValue(slCLFEndDate)
            'CBS?
            slCBS = "N"
            If llClfEndDate < llClfStartDate Then slCBS = "Y"
            
            'If llClfEndDate >= llClfStartDate Then
                'obtain population and demo codes by schedule line
                ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, tmClf.iDnfCode, 0, tgChfCT.iMnfDemo(0), llPop)
                If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                    llOvStartTime = 0
                    llOvEndTime = 0
                Else
                    'override times exist
                    gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                    gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                End If
                '------------------------------------------------------------------------------------------------
                'Contract Flights
                '------------------------------------------------------------------------------------------------
                ilCff = tgClfCT(ilClf).iFirstCff
                ilFoundSpot = False
                slPriceType = ""
                Do While ilCff <> -1
                    tmCff = tgCffCT(ilCff).CffRec
                    If tmClf.sType = "H" Or tmClf.sType = "S" Then 'process on hidden & std lines (no packages)
                        'lacInfo(0).Caption = "Processing Research (Contract " & tgChfCT.lCntrNo & ").."
                        If igDOE >= 500 Then
                            lacInfo(0).Refresh
                            igDOE = 0
                        End If
                        igDOE = igDOE + 1

                        ilFoundSpot = True
                        For ilLoop = 0 To 6                     'init all days to not airing, setup for research results later
                            ilInputDays(ilLoop) = False
                        Next ilLoop
                        
                        
                        'Debug.Print "CLFCode:" & tmClf.lCode
                        
                        gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slDate
                        llFltStart = gDateValue(slDate)
                        'backup start date to Monday
                        ilLoop = gWeekDayLong(llFltStart)
                        Do While ilLoop <> 0
                            llFltStart = llFltStart - 1
                            ilLoop = gWeekDayLong(llFltStart)
                        Loop
                        gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slDate
                        llFltEnd = gDateValue(slDate)
                        
                        '-------------------------------------------
                        'Price Type
                        Select Case tmCff.sPriceType
                            Case "N": slPriceType = "N/C"
                            Case "M": slPriceType = "MG"
                            Case "B": slPriceType = "Bonus"
                            Case "S": slPriceType = "Spinoff"
                            Case "P": slPriceType = "Package"
                            Case "R": slPriceType = "Recapturable"
                            Case "A": slPriceType = "ADU"
                            Case "T": slPriceType = "Paid"
                            Case Else: slPriceType = ""
                        End Select
                        
                        '-------------------------------------------
                        'Loop thru the flight by week and build the number of spots for each week
                        For llDate2 = llFltStart To llFltEnd Step 7
                            If llDate2 >= llClfStartDate And llDate2 <= llClfEndDate Then
                                If tmCff.sDyWk = "W" Then               'weekly
                                    ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                     For ilDay = 0 To 6 Step 1
                                        If (llDate2 + ilDay >= llFltStart) And (llDate2 + ilDay <= llFltEnd) Then
                                            If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        End If
                                     Next ilDay
                                Else                                    'daily
                                     If ilLoop + 6 < llFltEnd Then      'we have a whole week
                                        ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                        For ilDay = 0 To 6 Step 1
                                            If tmCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        Next ilDay
                                     Else                               'do partial week
                                        For llDate = llDate2 To llFltEnd Step 1
                                            ilDay = gWeekDayLong(llDate)
                                            ilSpots = ilSpots + tmCff.iDay(ilDay)
                                            If tmCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        Next llDate
                                    End If
                                End If
        
                                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmClf.iDnfCode, tmClf.iVefCode, 0, tgChfCT.iMnfDemo(0), llDate2, llDate2, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                'Loop and build avg aud, spots, & spots per week
                                'calc week index
                                ilUpperWk = (llDate2 - llClfStartDate) / 7 + 1
                                llWklyspots(ilUpperWk - 1) = ilSpots
                                ilTotLnSpts = ilTotLnSpts + ilSpots
                                'implement net option
                                llRate = gGetGrossOrNetFromRate(tmCff.lActPrice, slGrossNet, tgChfCT.iAgfCode)
                                llWklyRates(ilUpperWk - 1) = llRate
                                llWklyAvgAud(ilUpperWk - 1) = llAvgAud
                                llWklyPopEst(ilUpperWk - 1) = llPopEst
                                ilUpperWk = ilUpperWk + 1
                                
                            End If
                        Next llDate2
                    End If 'process on hidden & std lines (no packages)
                    ilCff = tgCffCT(ilCff).iNextCff                 'get next flight record from mem
                Loop                                                'while ilcff <> -1
                
                If ilFoundSpot = True Then
                    ReDim lmRtg(0 To ilUpperWk)                        'setup arrays for return values from audtolnresearch
                    ReDim lmGrimp(0 To ilUpperWk)
                    ReDim lmGRP(0 To ilUpperWk)
                    ReDim lmCost(0 To ilUpperWk)                       'setup arrays for return values from audtolnresearch
                    ReDim Preserve llWklyspots(0 To ilUpperWk) As Long          'sched lines weekly # spots
                    ReDim Preserve llWklyAvgAud(0 To ilUpperWk) As Long         'sched lines weekly avg aud
                    ReDim Preserve llWklyRates(0 To ilUpperWk) As Long          'sched lines weekly rates
                    ReDim Preserve llWklyPopEst(0 To ilUpperWk) As Long
                    
                    '-------------------------------------------
                    'RDF Daypart name lookup
                    ilRdfCode = tmClf.iRdfCode
                    slRdfName = ""
                    ilRet = gBinarySearchRdf(ilRdfCode)
                    If ilRet <> -1 Then
                        slRdfName = Trim(tgMRdf(ilRet).sName)
                    End If
                    
                    '-------------------------------------------
                    'Line Type
                    slLineType = ""
                    Select Case tmClf.sType
                        Case "S": slLineType = "C"
                        Case "O": slLineType = "P"
                        Case "H": slLineType = "H"
                    End Select
                    
                    '-------------------------------------------
                    '"sVehicle_Name" (Vef Name Lookup)
                    ilVefCode = tmClf.iVefCode
                    If ilLastVefID <> ilVefCode Then
                        slVefName = ""
                        iTVefCode = -1
                        ilRet = gBinarySearchVef(ilVefCode)
                        If ilRet <> -1 Then
                            iTVefCode = ilRet
                            slVefName = Trim(tgMVef(ilRet).sName)
                            ilLastVefID = ilVefCode
                        End If
                        
                        '-------------------------------------------
                        '"sMarket"
                        ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp3Mkt)
                        slMnfVehGp3Mkt = ""
                        If ilRet >= 0 Then
                            slMnfVehGp3Mkt = Trim$(tmMnfGroups(ilRet).sName)
                        End If
                                    
                        '-------------------------------------------
                        '"sResearch"
                        ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp5Rsch)
                        slMnfVehGp5Rsch = ""
                        If ilRet >= 0 Then
                            slMnfVehGp5Rsch = Trim$(tmMnfGroups(ilRet).sName)
                        End If
                        
                        '-------------------------------------------
                        '"sSub-Company"
                        ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp6Sub)
                        slMnfVehGp6Sub = ""
                        If ilRet >= 0 Then
                            slMnfVehGp6Sub = Trim$(tmMnfGroups(ilRet).sName)
                        End If
                        
                        '-------------------------------------------
                        '"sFormat"
                        ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp4Fmt)
                        slMnfVehGp4Fmt = ""
                        If ilRet >= 0 Then
                            slMnfVehGp4Fmt = Trim$(tmMnfGroups(ilRet).sName)
                        End If
                        
                        '-------------------------------------------
                        '"sSubTotals"
                        ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp2)
                        slvefMnfVehGp2 = ""
                        If ilRet >= 0 Then
                            slvefMnfVehGp2 = Trim$(tmMnfGroups(ilRet).sName)
                        End If
                    End If
        
                    'Debug.Print "StdHidden: " & slVefName
                    'Debug.Print "-pop:" & llPop & ", Spots:" & llSpots
                    '-------------------------------------------
                    'Schedule line complete, get its avg aud data for the line
                    'gAvgAudToLnResearch sm1or2PlaceRating, False, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), llTotalCost, llTotalAvgAud, lmRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                    gAvgAudToLnResearch sm1or2PlaceRating, False, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), dlTotalCost, llTotalAvgAud, lmRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                    'Debug.Print "-Spots:" & llSpots & " ,TotalCost:" & llTotalCost & " ,TotalAvgRtg:" & ilTotalAvgRtg & " ,TotalGrImp:" & llTotalGrImp & ",TotalGRP:" & llTotalGRP & " ,TotalCPP:" & llTotalCPP & " ,TotalCPM:" & llTotalCPM & " ,TotalAvgAud:" & llTotalAvgAud
                    
                    '-------------------------------------------
                    'Std/Hidden Contract Related data
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lContract_Number = olRs.Fields("lContract_Number").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lExternal_Version_Number = olRs.Fields("lExternal_Version_Number").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iProposal_Version_Number = olRs.Fields("iProposal_Version_Number").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAdvertiser_ID = olRs.Fields("iAdvertiser_ID").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sExternal_Advertiser_ID = IIF(IsNull(olRs.Fields("iExternal_Advertiser_ID")), "", olRs.Fields("iExternal_Advertiser_ID").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAdvertiser_Name = Trim(slAdfName)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAgency_ID = IIF(IsNull(olRs.Fields("iAgency_ID").Value), -1, olRs.Fields("iAgency_ID").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sExternal_Agency_ID = IIF(IsNull(olRs.Fields("sExternal_Agency_ID")), "", olRs.Fields("sExternal_Agency_ID").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAgency_Name = Trim(slAgfName)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSales_Office_ID = olRs.Fields("iSales_Office_ID").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSales_Office = Trim(gStripChr0(slSofName))
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSalesperson_ID = olRs.Fields("iSalesperson_ID").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSalesperson = slSlspName
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSalesperson_email = slSlspEmail
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iTrade_Percentage = olRs.Fields("iTrade_Percentage").Value
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sCashTrade = Trim(gStripChr0(olRs.Fields("sCash/Trade").Value))
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sContract_Type = Trim(gStripChr0(olRs.Fields("sContract_Type").Value))
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sDemo = IIF(IsNull(olRs.Fields("sDemo")), "", olRs.Fields("sDemo").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sProduct_Name = Trim(gStripChr0(olRs.Fields("sProduct_Name").Value))
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_1 = IIF(IsNull(olRs.Fields("sRevenue_Set_1")), "", olRs.Fields("sRevenue_Set_1").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_2 = IIF(IsNull(olRs.Fields("sRevenue_Set_2")), "", olRs.Fields("sRevenue_Set_2").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_3 = IIF(IsNull(olRs.Fields("sRevenue_Set_3")), "", olRs.Fields("sRevenue_Set_3").Value)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sStatus = Trim(gStripChr0(olRs.Fields("sStatus").Value))
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lChfCode = olRs.Fields("lchfCode").Value
                    '-------------------------------------------
                    'Line related data
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).dAverage_Audience = llTotalAvgAud
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lCPM = llTotalCPM
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lLine_GRPs = llTotalGRP
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iRating = ilTotalAvgRtg
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).dTotal_Gross = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iLine_CPP = llTotalCPP
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iLine_Gross_Impressions = llTotalGrImp
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSpot_Length = tmClf.iLen
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iTotal_Units = ilTotLnSpts
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sDaypart = slRdfName
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sPrice_Type = slPriceType
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sLine_End_Date = slCLFEndDate
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sLine_Start_Date = slCLFStartDate
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sLine_Type = slLineType
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iDnfCode = tmClf.iDnfCode
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lPop = llPop
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sCBS = slCBS
                    '-------------------------------------------
                    'Line vehicle related data
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iVehicle_ID = ilVefCode
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sMarket = Trim(slMnfVehGp3Mkt)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sResearch = Trim(slMnfVehGp5Rsch)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sFormat = Trim(slMnfVehGp4Fmt)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSubCompany = Trim(slMnfVehGp6Sub)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSubtotals = Trim(slvefMnfVehGp2)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sVehicle_Name = Trim(slVefName)
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAir_TimeNTR = "A"
                    tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iPkLineNo = tmClf.iPkLineNo
                    '-------------------------------------------
                    'Next Contract Line (Looking for NonPackage Lines)
                    ReDim Preserve tmWOLINEEXPDATA(0 To UBound(tmWOLINEEXPDATA) + 1) As EXPWOINVLN
                    'Record Counter
                    If igDOE >= 500 Then
                        lacInfo(0).Caption = "Processing " & llRecords & " Records..": lacInfo(0).Refresh
                        igDOE = 0
                    End If
                    igDOE = igDOE + 1
                    llRecords = llRecords + 1
                End If  'ilFoundSpot = True
            'End If 'CBS
        Next ilClf  'get next Contract Line
        
        '------------------------------------------------------------------------------------------------
        'Get Packages (If this is a Package line, gather all lines Hidden, create Package Research totals)
        '------------------------------------------------------------------------------------------------
        llSpots = 0
        dlTotalCost = 0 'TTP 10439 - Rerate 21,000,000
        llSpots = 0
        llTotalAvgAud = 0
        llTotalCPM = 0
        llTotalGRP = 0
        ilTotalAvgRtg = 0
        llTotalCPP = 0
        ilTotLnSpts = 0
        llTotalGrImp = 0
        If UBound(tmWOLINEEXPDATA) > LBound(tmWOLINEEXPDATA) Then   'found at least 1 line with research totals
            'Compute package #'s
            For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                tmClf = tgClfCT(ilClf).ClfRec
                'Debug.Print "PKG-CLFCode:" & tmClf.lCode
                'TTP 10453 - Contract Line export: "combination of price, spots and audience exceeded 21,000,000" warning message appearing
                If tmClf.sType <> "H" And tmClf.sType <> "S" And tmClf.sType <> "" And tmClf.iLine <> 0 Then
                    'Flight StartDate
                    gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slCLFStartDate
                    llClfStartDate = gDateValue(slCLFStartDate)
                    'Flight EndDate
                    gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slCLFEndDate
                    llClfEndDate = gDateValue(slCLFEndDate)
                    'CBS?
                    slCBS = "N"
                    If llClfEndDate < llClfStartDate Then slCBS = "Y"
                    ilPackageDnfCode = -1
                    
                    'If slCBS <> "Y" Then                       'We Want CBS vehicles to Export
                        llPop = -1
                        ilHiddenLines = 0
                        ReDim lmRtg(0 To 0)                     'setup arrays for return values from audtolnresearch
                        ReDim lmGrimp(0 To 0)
                        ReDim lmGRP(0 To 0)
                        ReDim lmCost(0 To 0)                    'setup arrays for return values from audtolnresearch
                        
                        For llLoopPkg = LBound(tmWOLINEEXPDATA) To UBound(tmWOLINEEXPDATA) - 1 Step 1
                            If tmClf.iLine = tmWOLINEEXPDATA(llLoopPkg).iPkLineNo And tmClf.lChfCode = olRs.Fields("lChfCode").Value Then
                                If tmWOLINEEXPDATA(llLoopPkg).sCBS <> "Y" Then
                                    tmWOLINEEXPDATA(llLoopPkg).iPkLineNo = 0 'Flag so this wont get calc'd again
                                    ReDim Preserve lmCost(0 To ilHiddenLines) As Long
                                    ReDim Preserve lmGrimp(0 To ilHiddenLines) As Long
                                    ReDim Preserve lmGRP(0 To ilHiddenLines) As Long
                                    lmCost(ilHiddenLines) = tmWOLINEEXPDATA(llLoopPkg).dTotal_Gross 'TTP 10439 - Rerate 21,000,000
                                    lmGrimp(ilHiddenLines) = tmWOLINEEXPDATA(llLoopPkg).iLine_Gross_Impressions
                                    lmGRP(ilHiddenLines) = tmWOLINEEXPDATA(llLoopPkg).lLine_GRPs
                                    'determine if varying populations across the weeks (demo estimates) or across the lines
                                    If tmWOLINEEXPDATA(llLoopPkg).iDnfCode > 0 Then
                                        If tgSpf.sDemoEstAllowed = "Y" Then
                                            If llPop < 0 Then
                                                llPop = tmWOLINEEXPDATA(llLoopPkg).lPop
                                            ElseIf llPop <> tmWOLINEEXPDATA(llLoopPkg).lPop And tmWOLINEEXPDATA(llLoopPkg).lPop <> 0 Then
                                                llPop = 0
                                            End If
                                        Else
                                            If llPop < 0 Then
                                                llPop = tmWOLINEEXPDATA(llLoopPkg).lPop
                                            ElseIf llPop <> tmWOLINEEXPDATA(llLoopPkg).lPop And tmWOLINEEXPDATA(llLoopPkg).lPop <> 0 Then
                                                llPop = 0
                                            End If
                                        End If
                                    End If
                                    If tmWOLINEEXPDATA(llLoopPkg).iTotal_Units > 0 Then
                                        If tmWOLINEEXPDATA(llLoopPkg).iDnfCode > 0 Then
                                            ilPackageDnfCode = tmWOLINEEXPDATA(llLoopPkg).iDnfCode
                                        End If
                                    End If
                                    ilHiddenLines = ilHiddenLines + 1
                                End If 'CBS?
                            End If 'Pkg Line = Line#
                        Next llLoopPkg
                        
                        'Get total package spots, cost his ok as each week of hidden lines must match the package line cost
                        llSpots = 0
                        slPriceType = ""
                        ilCff = tgClfCT(ilClf).iFirstCff
                        Do While ilCff <> -1
                            tmCff = tgCffCT(ilCff).CffRec
                            For ilLoop = 0 To 6                 'init all days to not airing, setup for research results later
                                ilInputDays(ilLoop) = False
                            Next ilLoop
                            ilSpots = 0
                            gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                            llFltStart = gDateValue(slStr)
                            'backup start date to Monday
                            ilLoop = gWeekDayLong(llFltStart)
                            Do While ilLoop <> 0
                                llFltStart = llFltStart - 1
                                ilLoop = gWeekDayLong(llFltStart)
                            Loop
                            gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                            llFltEnd = gDateValue(slStr)
                            
                            '-------------------------------------------
                            'Price Type
                            Select Case tmCff.sPriceType
                                Case "N": slPriceType = "N/C"
                                Case "M": slPriceType = "MG"
                                Case "B": slPriceType = "Bonus"
                                Case "S": slPriceType = "Spinoff"
                                Case "P": slPriceType = "Package"
                                Case "R": slPriceType = "Recapturable"
                                Case "A": slPriceType = "ADU"
                                Case "T": slPriceType = "Paid"
                                Case Else: slPriceType = ""
                            End Select

                            '-------------------------------------------
                            'Loop thru the flight by week and build the number of spots for each week
                            For llDate2 = llFltStart To llFltEnd Step 7
                                If llDate2 >= llStartDate And llDate2 <= llEndDate Then
                                    If tmCff.sDyWk = "W" Then            'weekly
                                        ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                    Else                                        'daily
                                         If ilLoop + 6 < llFltEnd Then          'we have a whole week
                                            ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                         Else                                   'do partial week
                                            For llDate = llDate2 To llFltEnd Step 1
                                                ilDay = gWeekDayLong(llDate)
                                                ilSpots = ilSpots + tmCff.iDay(ilDay)
                                            Next llDate
                                        End If
                                    End If
                                    llSpots = llSpots + ilSpots
                                End If                  'if llDate2 >= llStartDate and llDate2 <= llEndDAte
                            Next llDate2
                            ilCff = tgCffCT(ilCff).iNextCff            'get next flight record from mem
                        Loop 'while ilcff <> -1
                                                   
                        '-------------------------------------------
                        'Pkg RDF Daypart name lookup
                        ilRdfCode = tmClf.iRdfCode
                        slRdfName = ""
                        ilRet = gBinarySearchRdf(ilRdfCode)
                        If ilRet <> -1 Then
                            slRdfName = Trim(tgMRdf(ilRet).sName)
                        End If
                        
                        '-------------------------------------------
                        'Pkg Line Type
                        slLineType = ""
                        Select Case tmClf.sType
                            Case "S": slLineType = "C"
                            Case "O": slLineType = "P"
                            Case "H": slLineType = "H"
                        End Select
                        
                        '-------------------------------------------
                        '"sVehicle_Name" (Vef Name Lookup)
                        ilVefCode = tmClf.iVefCode
                        If ilLastVefID <> ilVefCode Then
                            slVefName = ""
                            iTVefCode = -1
                            ilRet = gBinarySearchVef(ilVefCode)
                            If ilRet <> -1 Then
                                iTVefCode = ilRet
                                slVefName = Trim(tgMVef(ilRet).sName)
                                ilLastVefID = ilVefCode
                            End If
                            
                            '-------------------------------------------
                            '"sMarket"
                            ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp3Mkt)
                            slMnfVehGp3Mkt = ""
                            If ilRet >= 0 Then
                                slMnfVehGp3Mkt = Trim$(tmMnfGroups(ilRet).sName)
                            End If
                                        
                            '-------------------------------------------
                            '"sResearch"
                            ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp5Rsch)
                            slMnfVehGp5Rsch = ""
                            If ilRet >= 0 Then
                                slMnfVehGp5Rsch = Trim$(tmMnfGroups(ilRet).sName)
                            End If
                            
                            '-------------------------------------------
                            '"sSub-Company"
                            ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp6Sub)
                            slMnfVehGp6Sub = ""
                            If ilRet >= 0 Then
                                slMnfVehGp6Sub = Trim$(tmMnfGroups(ilRet).sName)
                            End If
                            
                            '-------------------------------------------
                            '"sFormat"
                            ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp4Fmt)
                            slMnfVehGp4Fmt = ""
                            If ilRet >= 0 Then
                                slMnfVehGp4Fmt = Trim$(tmMnfGroups(ilRet).sName)
                            End If
                            
                            '-------------------------------------------
                            '"sSubTotals"
                            ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp2)
                            slvefMnfVehGp2 = ""
                            If ilRet >= 0 Then
                                slvefMnfVehGp2 = Trim$(tmMnfGroups(ilRet).sName)
                            End If
                        End If
                        
                        'Obtain package research totals
                        'Debug.Print "Pkg: " & slVefName
                        'Debug.Print "-pop:" & llPop & ", Spots:" & llSpots
                        'gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                        gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                        'Debug.Print "-Pkg-Spots:" & llSpots & " ,TotalCost:" & llTotalCost & " ,TotalAvgRtg:" & ilTotalAvgRtg & " ,TotalGrImp:" & llTotalGrImp & ",TotalGRP:" & llTotalGRP & " ,TotalCPP:" & llTotalCPP & " ,TotalCPM:" & llTotalCPM & " ,TotalAvgAud:" & llTotalAvgAud
                                            
                        '-------------------------------------------
                        'Price Type
                        slPriceType = ""
                        Select Case tmCff.sPriceType
                            Case "N": slPriceType = "N/C"
                            Case "M": slPriceType = "MG"
                            Case "B": slPriceType = "Bonus"
                            Case "S": slPriceType = "Spinoff"
                            Case "P": slPriceType = "Package"
                            Case "R": slPriceType = "Recapturable"
                            Case "A": slPriceType = "ADU"
                            Case "T": slPriceType = "Paid"
                            Case Else: slPriceType = ""
                        End Select
                        '-------------------------------------------
                        'Package Contract Related data
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lContract_Number = olRs.Fields("lContract_Number").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lExternal_Version_Number = olRs.Fields("lExternal_Version_Number").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iProposal_Version_Number = olRs.Fields("iProposal_Version_Number").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAdvertiser_ID = olRs.Fields("iAdvertiser_ID").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sExternal_Advertiser_ID = IIF(IsNull(olRs.Fields("iExternal_Advertiser_ID")), "", olRs.Fields("iExternal_Advertiser_ID").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAdvertiser_Name = Trim(slAdfName)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAgency_ID = IIF(IsNull(olRs.Fields("iAgency_ID").Value), -1, olRs.Fields("iAgency_ID").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sExternal_Agency_ID = IIF(IsNull(olRs.Fields("sExternal_Agency_ID")), "", olRs.Fields("sExternal_Agency_ID").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAgency_Name = Trim(slAgfName)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSales_Office_ID = olRs.Fields("iSales_Office_ID").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSales_Office = Trim(gStripChr0(slSofName))
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSalesperson_ID = olRs.Fields("iSalesperson_ID").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSalesperson = slSlspName
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSalesperson_email = slSlspEmail
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iTrade_Percentage = olRs.Fields("iTrade_Percentage").Value
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sCashTrade = Trim(gStripChr0(olRs.Fields("sCash/Trade").Value))
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sContract_Type = Trim(gStripChr0(olRs.Fields("sContract_Type").Value))
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sDemo = IIF(IsNull(olRs.Fields("sDemo")), "", olRs.Fields("sDemo").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sProduct_Name = Trim(gStripChr0(olRs.Fields("sProduct_Name").Value))
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_1 = IIF(IsNull(olRs.Fields("sRevenue_Set_1")), "", olRs.Fields("sRevenue_Set_1").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_2 = IIF(IsNull(olRs.Fields("sRevenue_Set_2")), "", olRs.Fields("sRevenue_Set_2").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_3 = IIF(IsNull(olRs.Fields("sRevenue_Set_3")), "", olRs.Fields("sRevenue_Set_3").Value)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sStatus = Trim(gStripChr0(olRs.Fields("sStatus").Value))
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lChfCode = olRs.Fields("lchfCode").Value
                        '-------------------------------------------
                        'Line related data
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).dAverage_Audience = llTotalAvgAud
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lCPM = llTotalCPM
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lLine_GRPs = llTotalGRP
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iRating = ilTotalAvgRtg
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).dTotal_Gross = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iLine_CPP = llTotalCPP
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iLine_Gross_Impressions = llTotalGrImp
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSpot_Length = tmClf.iLen
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iTotal_Units = llSpots
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sDaypart = slRdfName
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sPrice_Type = slPriceType
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sLine_End_Date = slCLFEndDate
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sLine_Start_Date = slCLFStartDate
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sLine_Type = slLineType
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iPkLineNo = 0
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iDnfCode = ilPackageDnfCode
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lPop = llPop
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sCBS = slCBS
                        '-------------------------------------------
                        'Line vehicle related data
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iVehicle_ID = ilVefCode
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sMarket = Trim(slMnfVehGp3Mkt)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sResearch = Trim(slMnfVehGp5Rsch)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sFormat = Trim(slMnfVehGp4Fmt)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSubCompany = Trim(slMnfVehGp6Sub)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSubtotals = Trim(slvefMnfVehGp2)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sVehicle_Name = Trim(slVefName)
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAir_TimeNTR = "A"
                        tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iPkLineNo = 0
                        
                        '-------------------------------------------
                        'Next Contract Line (looking for Packages)
                        llSpots = 0
                        llTotalAvgAud = 0
                        llTotalCPM = 0
                        llTotalGRP = 0
                        ilTotalAvgRtg = 0
                        dlTotalCost = 0 'TTP 10439 - Rerate 21,000,000
                        llTotalCPP = 0
                        ilTotLnSpts = 0
                        llTotalGrImp = 0
                        ReDim Preserve lmCost(0 To 0) As Long
                        ReDim Preserve lmGrimp(0 To 0) As Long
                        ReDim Preserve lmGRP(0 To 0) As Long
                        ReDim Preserve tmWOLINEEXPDATA(0 To UBound(tmWOLINEEXPDATA) + 1) As EXPWOINVLN
                        'Record Counter
                        'lacInfo(0).Caption = "Processing Research (Contract " & tgChfCT.lCntrNo & ").."
                        If igDOE >= 500 Then
                            lacInfo(0).Refresh
                            igDOE = 0
                        End If
                        igDOE = igDOE + 1
                        llRecords = llRecords + 1
                    'End If 'CBS?
                End If
            Next ilClf
        End If                      'found at least 1 line with research totals
        
        '------------------------------------------------------------------------------------------------
        'Get NTR
        '------------------------------------------------------------------------------------------------
        ReDim tlSbf(0 To 0) As SBF
        tlSBFTypes.iNTR = True          'include NTR billing
        tlSBFTypes.iInstallment = False      'exclude Installment billing
        tlSBFTypes.iImport = False           'exclude rep import billing
        ilRet = gObtainSBF(ExpCntrLine, hmSbf, olRs.Fields("lChfCode").Value, slStartDate, slEndDate, tlSBFTypes, tlSbf(), 0)     '11-28-06 add last parm to indicate which key to use
        For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1
            tmSbf = tlSbf(llSbf)
            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
            llDate = gDateValue(slDate)
                        
            '-------------------------------------------
            '"sVehicle_Name" (Vef Name Lookup)
            ilVefCode = tmSbf.iBillVefCode
            If ilLastVefID <> ilVefCode Then
                slVefName = ""
                iTVefCode = -1
                ilRet = gBinarySearchVef(ilVefCode)
                If ilRet <> -1 Then
                    iTVefCode = ilRet
                    slVefName = Trim(tgMVef(ilRet).sName)
                    ilLastVefID = ilVefCode
                End If
                '-------------------------------------------
                '"sMarket"
                ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp3Mkt)
                slMnfVehGp3Mkt = ""
                If ilRet >= 0 Then
                    slMnfVehGp3Mkt = Trim$(tmMnfGroups(ilRet).sName)
                End If
                '-------------------------------------------
                '"sResearch"
                ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp5Rsch)
                slMnfVehGp5Rsch = ""
                If ilRet >= 0 Then
                    slMnfVehGp5Rsch = Trim$(tmMnfGroups(ilRet).sName)
                End If
                '-------------------------------------------
                '"sSub-Company"
                ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp6Sub)
                slMnfVehGp6Sub = ""
                If ilRet >= 0 Then
                    slMnfVehGp6Sub = Trim$(tmMnfGroups(ilRet).sName)
                End If
                '-------------------------------------------
                '"sFormat"
                ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp4Fmt)
                slMnfVehGp4Fmt = ""
                If ilRet >= 0 Then
                    slMnfVehGp4Fmt = Trim$(tmMnfGroups(ilRet).sName)
                End If
                '-------------------------------------------
                '"sSubTotals"
                ilRet = mBinarySearchMnfVehicleGroup(tgMVef(iTVefCode).iMnfVehGp2)
                slvefMnfVehGp2 = ""
                If ilRet >= 0 Then
                    slvefMnfVehGp2 = Trim$(tmMnfGroups(ilRet).sName)
                End If
            End If
            
            '-------------------------------------------
            'slNtrType
            tmMnfSrchKey.iCode = tmSbf.iMnfItem
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slNTRType = Trim$(gStripChr0(tmMnf.sName))
            End If
            '-------------------------------------------
            'Contract Related data
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lContract_Number = olRs.Fields("lContract_Number").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lExternal_Version_Number = olRs.Fields("lExternal_Version_Number").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iProposal_Version_Number = olRs.Fields("iProposal_Version_Number").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAdvertiser_ID = olRs.Fields("iAdvertiser_ID").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sExternal_Advertiser_ID = IIF(IsNull(olRs.Fields("iExternal_Advertiser_ID")), "", olRs.Fields("iExternal_Advertiser_ID").Value)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAdvertiser_Name = Trim(slAdfName)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAgency_ID = olRs.Fields("iAgency_ID").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sExternal_Agency_ID = IIF(IsNull(olRs.Fields("sExternal_Agency_ID")), "", olRs.Fields("sExternal_Agency_ID").Value)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAgency_Name = Trim(slAgfName)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSales_Office_ID = olRs.Fields("iSales_Office_ID").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSales_Office = Trim(gStripChr0(slSofName))
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iSalesperson_ID = olRs.Fields("iSalesperson_ID").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSalesperson = slSlspName
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSalesperson_email = slSlspEmail
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iTrade_Percentage = olRs.Fields("iTrade_Percentage").Value
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sCashTrade = Trim(gStripChr0(olRs.Fields("sCash/Trade").Value))
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sContract_Type = Trim(gStripChr0(olRs.Fields("sContract_Type").Value))
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sDemo = IIF(IsNull(olRs.Fields("sDemo")), "", olRs.Fields("sDemo").Value)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sProduct_Name = Trim(gStripChr0(olRs.Fields("sProduct_Name").Value))
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_1 = IIF(IsNull(olRs.Fields("sRevenue_Set_1")), "", olRs.Fields("sRevenue_Set_1").Value)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_2 = IIF(IsNull(olRs.Fields("sRevenue_Set_2")), "", olRs.Fields("sRevenue_Set_2").Value)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sRevenue_Set_3 = IIF(IsNull(olRs.Fields("sRevenue_Set_3")), "", olRs.Fields("sRevenue_Set_3").Value)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sStatus = Trim(gStripChr0(olRs.Fields("sStatus").Value))
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).lChfCode = olRs.Fields("lchfCode").Value
            '-------------------------------------------
            'ntr related data
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sAir_TimeNTR = "N"
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iVehicle_ID = ilVefCode
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sVehicle_Name = Trim(slVefName)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sMarket = Trim(slMnfVehGp3Mkt)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSubCompany = Trim(slMnfVehGp6Sub)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sSubtotals = Trim(slvefMnfVehGp2)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sFormat = Trim(slMnfVehGp4Fmt)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sResearch = Trim(slMnfVehGp5Rsch)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iAmount_Per_NTR_Item = gLongToStrDec(tmSbf.lGross, 2)
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).iNumber_of_NTR_Items = tmSbf.iNoItems
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sNTR_Billing_Date = slDate
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sNTR_Description = Trim(gStripChr0(tmSbf.sDescr))
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).dTotal_Gross = tmSbf.iNoItems * tmSbf.lGross 'TTP 10439 - Rerate 21,000,000
            tmWOLINEEXPDATA(UBound(tmWOLINEEXPDATA)).sNTR_Type = slNTRType

            '-------------------------------------------
            'Next NTR record
            If igDOE >= 500 Then
                lacInfo(0).Caption = "Processing " & llRecords & " Records..": lacInfo(0).Refresh
                igDOE = 0
            End If
            igDOE = igDOE + 1
            ReDim Preserve tmWOLINEEXPDATA(0 To UBound(tmWOLINEEXPDATA) + 1) As EXPWOINVLN
            llRecords = llRecords + 1
        Next llSbf
        
        '-------------------------------------------
        'Move to Next Contract Record
        olRs.MoveNext
    Loop 'Next Cotract
    
    olRs.Close
    
    slErrorMessage = "No errors, " & llRecords & " rows processed.."
    mProcessData = slErrorMessage

finish:
    'Set olFileSys = Nothing
    Exit Function

ERRORBOX:
    mProcessData = "Error reading records in mExport:" & err & "-" & Error(err)
    'Set olFileSys = Nothing
    GoTo finish
End Function

Function mExportData(olRs As ADODB.Recordset, smExportName, Optional slDelimeter = ",") As String
    'Create a CSV file, give me a recordset and a (fully qualified path\filename.ext) filename
    'Makes Headers from recordset Column Names
    'Makes rows from Data
    Dim slErrorMessage As String
    Dim olCsv As TextStream
    Dim slFormattedString As String
    Dim slHeader As String
    Dim llRecords As Long
    Dim llLoop As Long
    Dim olFileSys As FileSystemObject
    
    On Error GoTo ERRORBOX
    Set olFileSys = New FileSystemObject
    Set olCsv = olFileSys.OpenTextFile(smExportName, ForWriting, True)
    
    '------------------------------------------
    'Get Header
    slHeader = mGetHeaderString()
    
    '------------------------------------------
    'Write Header
    olCsv.WriteLine slHeader
    llRecords = 0
    '------------------------------------------
    'Read Array, write CSV
    For llLoop = 0 To UBound(tmWOLINEEXPDATA) - 1
        slFormattedString = ""
        'Contract Number
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).lContract_Number
        slFormattedString = slFormattedString & ","
        'External Version Number
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).lExternal_Version_Number
        slFormattedString = slFormattedString & ","
        'Proposal Version Number
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iProposal_Version_Number
        slFormattedString = slFormattedString & ","
        'Contract Type
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sContract_Type & """"
        slFormattedString = slFormattedString & ","
        'Salesperson
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sSalesperson & """"
        slFormattedString = slFormattedString & ","
        'Salesperson email
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sSalesperson_email & """"
        slFormattedString = slFormattedString & ","
        'Salesperson ID
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iSalesperson_ID
        slFormattedString = slFormattedString & ","
        'Sales Office
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sSales_Office & """"
        slFormattedString = slFormattedString & ","
        'Sales Office ID
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iSales_Office_ID
        slFormattedString = slFormattedString & ","
        'Agency Name
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sAgency_Name & """"
        slFormattedString = slFormattedString & ","
        'Agency ID
        If tmWOLINEEXPDATA(llLoop).iAgency_ID = -1 Then
            slFormattedString = slFormattedString & ""
        Else
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iAgency_ID
        End If
        slFormattedString = slFormattedString & ","
        'External Agency ID
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sExternal_Agency_ID & """"
        slFormattedString = slFormattedString & ","
        'Advertiser Name
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sAdvertiser_Name & """"
        slFormattedString = slFormattedString & ","
        'Advertiser ID
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).iAdvertiser_ID & """"
        slFormattedString = slFormattedString & ","
        'External Advertiser ID
        '6/13/23 - 'External Advertiser ID' value should be quoted but Jason says: let's not change it. That export is feeding into another system. I'm not sure if that other system can handle suddenly having quotes around that field. They're only putting numbers and dashes into that field, so as long as they keep doing that, they should be fine. I don't think they typically open that file in Excel anyhow, it just feeds that other system
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).sExternal_Advertiser_ID
        slFormattedString = slFormattedString & ","
        'Product Name
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sProduct_Name & """"
        slFormattedString = slFormattedString & ","
        'Cash/Trade
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sCashTrade & """"
        slFormattedString = slFormattedString & ","
        'Trade Percentage
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iTrade_Percentage
        slFormattedString = slFormattedString & ","
        'Air Time/NTR
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sAir_TimeNTR & """"
        slFormattedString = slFormattedString & ","
        'Demo
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sDemo & """"
        slFormattedString = slFormattedString & ","
        'Status
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sStatus & """"
        slFormattedString = slFormattedString & ","
        'Revenue Set 1
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sRevenue_Set_1 & """"
        slFormattedString = slFormattedString & ","
        'Revenue Set 2
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sRevenue_Set_2 & """"
        slFormattedString = slFormattedString & ","
        'Revenue Set 3
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sRevenue_Set_3 & """"
        slFormattedString = slFormattedString & ","
        'Vehicle Name
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sVehicle_Name & """"
        slFormattedString = slFormattedString & ","
        'Vehicle ID
        slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iVehicle_ID
        slFormattedString = slFormattedString & ","
        'Market
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sMarket & """"
        slFormattedString = slFormattedString & ","
        'Research
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sResearch & """"
        slFormattedString = slFormattedString & ","
        'Sub-Company
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sSubCompany & """"
        slFormattedString = slFormattedString & ","
        'Format
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sFormat & """"
        slFormattedString = slFormattedString & ","
        'SubTotals
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sSubtotals & """"
        slFormattedString = slFormattedString & ","
        
        'Spot Length
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iSpot_Length
        End If
        slFormattedString = slFormattedString & ","
        
        'Daypart
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sDaypart & """"
        slFormattedString = slFormattedString & ","
        
        'Line Type
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sLine_Type & """"
        slFormattedString = slFormattedString & ","
        If tmWOLINEEXPDATA(llLoop).sLine_Start_Date <> "" And tmWOLINEEXPDATA(llLoop).sLine_End_Date <> "" Then
            If DateValue(tmWOLINEEXPDATA(llLoop).sLine_Start_Date) > DateValue(tmWOLINEEXPDATA(llLoop).sLine_End_Date) Then
                'Line Start Date
                slFormattedString = slFormattedString & """CBS"""
                slFormattedString = slFormattedString & ","
                'Line End Date
                slFormattedString = slFormattedString & """CBS"""
                slFormattedString = slFormattedString & ","
            Else
                'Line Start Date
                slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sLine_Start_Date & """"
                slFormattedString = slFormattedString & ","
                'Line End Date
                slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sLine_End_Date & """"
                slFormattedString = slFormattedString & ","
            End If
        Else
            'Line Start Date
            slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sLine_Start_Date & """"
            slFormattedString = slFormattedString & ","
            'Line End Date
            slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sLine_End_Date & """"
            slFormattedString = slFormattedString & ","
        End If
        
        'Total Units
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iTotal_Units
        End If
        slFormattedString = slFormattedString & ","
        
        'sPrice Type
        slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sPrice_Type & """"
        slFormattedString = slFormattedString & ","
        
        'Total Gross
        slFormattedString = slFormattedString & gDblToStrDec(tmWOLINEEXPDATA(llLoop).dTotal_Gross, 2) 'TTP 10439 - Rerate 21,000,000
        slFormattedString = slFormattedString & ","
        
        'Rating
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & gIntToStrDec(tmWOLINEEXPDATA(llLoop).iRating, IIF(sm1or2PlaceRating = "2", 2, 1))
        End If
        slFormattedString = slFormattedString & ","
        
        'Line CPP
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iLine_CPP
        End If
        slFormattedString = slFormattedString & ","
        
        'Line GRPs
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).lLine_GRPs / 10
        End If
        slFormattedString = slFormattedString & ","
        
        'CPM
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).lCPM / 100
        End If
        slFormattedString = slFormattedString & ","
        
        'Average Audience
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & gLongToStrDec(tmWOLINEEXPDATA(llLoop).dAverage_Audience, imNumberDecPlaces)
        End If
        slFormattedString = slFormattedString & ","
        
        'Line Gross Impressions
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "A" Then
            slFormattedString = slFormattedString & gLongToStrDec(tmWOLINEEXPDATA(llLoop).iLine_Gross_Impressions, imNumberDecPlaces)
        End If
        slFormattedString = slFormattedString & ","
        
        'NTR Billing Date
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "N" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).sNTR_Billing_Date
        End If
        slFormattedString = slFormattedString & ","
        
        'NTR Description
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "N" Then
            slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sNTR_Description & """"
        End If
        slFormattedString = slFormattedString & ","
        
        'NTR Type
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "N" Then
            slFormattedString = slFormattedString & """" & tmWOLINEEXPDATA(llLoop).sNTR_Type & """"
        End If
        slFormattedString = slFormattedString & ","
        
        'Amount Per NTR Item
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "N" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iAmount_Per_NTR_Item
        End If
        slFormattedString = slFormattedString & ","
        
        'Number of NTR Items
        If tmWOLINEEXPDATA(llLoop).sAir_TimeNTR = "N" Then
            slFormattedString = slFormattedString & tmWOLINEEXPDATA(llLoop).iNumber_of_NTR_Items
        End If
        
        olCsv.WriteLine slFormattedString
        
        If igDOE >= 100 Then
            lacInfo(0).Caption = llRecords & " rows exported..": lacInfo(0).Refresh
            igDOE = 0
        End If
        igDOE = igDOE + 1
        llRecords = llRecords + 1
    Next llLoop
    
    olCsv.Close
    slErrorMessage = "No errors, Exported " & llRecords & " rows exported.."
    'Print #hmMsg, "** " & slErrorMessage & " **"
    gAutomationAlertAndLogHandler "** " & slErrorMessage & " **"
    mExportData = slErrorMessage

finish:
    Set olCsv = Nothing
    Set olFileSys = Nothing
    Exit Function

ERRORBOX:
    slErrorMessage = "Error reading records in mExport:" & err & "-" & Error(err)
    mExportData = slErrorMessage
    Set olCsv = Nothing
    Set olFileSys = Nothing
    GoTo finish
End Function

Private Function mGetSOFName(ByVal lSofCode As Integer) As String
    Dim rst As Recordset
    Dim slName As String
    Dim llCounter As Long
    Dim slSql As String
    If UBound(tmSof) = 0 Then
        slSql = "Select sofCode, ltrim(rtrim(sofName)) as sofName From SOF_Sales_Offices"
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            Do While Not rst.EOF
                tmSof(UBound(tmSof)).iCode = rst!sofcode
                tmSof(UBound(tmSof)).sName = Trim$(rst!sofName)
                If lSofCode = rst!sofcode Then slName = rst!sofName
                ReDim Preserve tmSof(0 To UBound(tmSof) + 1)
                rst.MoveNext
            Loop
        End If
    End If
    If slName = "" Then
        For llCounter = 0 To UBound(tmSof) - 1
            If tmSof(llCounter).iCode = lSofCode Then
                slName = tmSof(llCounter).sName
                Exit For
            End If
        Next llCounter
    End If
    mGetSOFName = slName
    Set rst = Nothing
End Function

