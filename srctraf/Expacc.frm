VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpACC 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3840
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   7095
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
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   7095
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   4905
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   21
      Top             =   1050
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Expacc.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   25
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   26
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2130
      MaxLength       =   10
      TabIndex        =   6
      Top             =   795
      Width           =   930
   End
   Begin VB.CommandButton cmcStartDate 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3060
      Picture         =   "Expacc.frx":2E1A
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   795
      Width           =   195
   End
   Begin VB.TextBox edcEndDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   9
      Top             =   795
      Width           =   930
   End
   Begin VB.CommandButton cmcEndDate 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6585
      Picture         =   "Expacc.frx":2F14
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   795
      Width           =   195
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6435
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
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
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1515
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6450
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5835
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6105
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3135
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
      Left            =   5415
      TabIndex        =   4
      Top             =   330
      Width           =   1485
   End
   Begin VB.PictureBox plcTo 
      Height          =   375
      Left            =   1035
      ScaleHeight     =   315
      ScaleWidth      =   4245
      TabIndex        =   2
      Top             =   300
      Width           =   4305
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   4185
      End
   End
   Begin VB.Frame frcTranType 
      Caption         =   "Transaction Types"
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   90
      TabIndex        =   11
      Top             =   1230
      Width           =   3045
      Begin VB.OptionButton rbcType 
         Caption         =   "Receivables (All except HI)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   795
         Width           =   2745
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Cash (PI, PO)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   2085
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Invoices (IN, AN, HI)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   2775
      End
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
      Left            =   2400
      TabIndex        =   15
      Top             =   3360
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
      TabIndex        =   16
      Top             =   3360
      Width           =   1050
   End
   Begin VB.ListBox lbcSOffice 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   6015
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   450
      TabIndex        =   28
      Top             =   3030
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Transaction Start Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   795
      Width           =   2025
   End
   Begin VB.Label lacExportDate 
      Appearance      =   0  'Flat
      Caption         =   "Transaction End Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3615
      TabIndex        =   8
      Top             =   795
      Width           =   1995
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      Caption         =   "To File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   375
      Width           =   810
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   3315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1905
      TabIndex        =   20
      Top             =   2745
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "ExpACC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ExpACC.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export spot (clearance) input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
'Receivables
Dim hmRvf As Integer    'Receivable file handle
Dim tmRvf As RVF
Dim imRvfRecLen As Integer        'Rvf record length
'Payment History
Dim hmPhf As Integer    'Rate Card item file handle
Dim tmPhf As RVF
Dim imPhfRecLen As Integer        'Rpf record length
'Product
Dim hmPrf As Integer        'Prf Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length
Dim tmPrfSrchKey As LONGKEY0  'Prf key record image
Dim tmSOfficeCode() As SORTCODE
Dim smSOfficeCodeTag As String
Dim tmSalesOffice() As SALESOFFICE
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer    '1=Start Date; 2=Export Date
Dim smTranStart As String
Dim smTranEnd As String
Dim tmSbfSrchKey0 As SBFKEY0
Dim imSbfRecLen As Integer  'SBF record length
Dim hmSbf As Integer        'Special Billing file handle
Dim tmSbf As SBF            'SBF record image

Dim hmIhf As Integer
Dim tmIhf As IHF        'IHF record image
Dim tmIhfSrchKey0 As INTKEY0    'IHF key record image
Dim imIhfRecLen As Integer        'IHF record length


Dim hmItf As Integer
Dim tmItf As ITF        'ITF record image
Dim tmItfSrchKey0 As INTKEY0    'ITF key record image
Dim imItfRecLen As Integer        'ITF record length

Dim hmCHF As Integer
Dim tmChf As CHF    'Used when updating only to eliminate conflict
Dim tmChfSrchKey1 As CHFKEY1  'CHF key record image (contract #)
Dim imCHFRecLen As Integer
Dim hmClf As Integer
Dim hmCff As Integer
Dim tmChfAcc As CHF
Dim tmClfAcc() As CLFLIST
Dim tmCffAcc() As CFFLIST
Private Type EXPACCINFO
    lCntrNo As Long
    sInstallDefined As String * 1
    iVefCode As Integer
    sGLCashNo As String * 20
    sGLTradeNo As String * 20
    dGross As Double              '7-8-08 lGross As Long:  change all occurances of this field to double to
                                'prevent overflow error
End Type
Dim tmExpAccInfo() As EXPACCINFO

Private Sub cmcCalDn_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
        edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        imCalMonth = imCalMonth - 1
        If imCalMonth <= 0 Then
            imCalMonth = 12
            imCalYear = imCalYear - 1
        End If
        pbcCalendar_Paint
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
        edcEndDate.SetFocus
    End If
End Sub

Private Sub cmcCalUp_Click()
    If imDateBox = 1 Then
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
        edcStartDate.SetFocus
    Else
        imCalMonth = imCalMonth + 1
        If imCalMonth > 12 Then
            imCalMonth = 1
            imCalYear = imCalYear + 1
        End If
        pbcCalendar_Paint
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
        edcEndDate.SetFocus
    End If
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    imDateBox = -1
End Sub

Private Sub cmcEndDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
End Sub

Private Sub cmcEndDate_GotFocus()
    Dim slStr As String
    If imDateBox <> 2 Then
        plcCalendar.Visible = False
        slStr = edcEndDate.Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = 2
    plcCalendar.Move edcEndDate.Left + edcEndDate.Width + cmcEndDate.Width - plcCalendar.Width, edcEndDate.Top + edcEndDate.Height
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slDateTime As String
    ReDim tmExpAccInfo(0 To 0) As EXPACCINFO
    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    sgMessageFile = sgDBPath & "Messages\" & "ExptAccounting.Txt"
    smTranStart = Trim$(edcStartDate.Text)
    If Len(smTranStart) = 0 Then
        smTranStart = "1/1/1970"
    End If
    If Not gValidDate(smTranStart) Then
        ''MsgBox "Start Date is Not Valid", vbOkOnly + vbApplicationModal, "Start Date"
        gAutomationAlertAndLogHandler "Start Date is Not Valid", vbOkOnly + vbApplicationModal, "Start Date"
        edcStartDate.SetFocus
        Exit Sub
    End If
    smTranEnd = Trim$(edcEndDate.Text)
    If Len(smTranEnd) = 0 Then
        smTranEnd = "12/31/2069"
    End If
    If Not gValidDate(smTranEnd) Then
        ''MsgBox "End Date is Not Valid", vbOkOnly + vbApplicationModal, "End Date"
        gAutomationAlertAndLogHandler "End Date is Not Valid", vbOkOnly + vbApplicationModal, "End Date"
        edcEndDate.SetFocus
        Exit Sub
    End If
    slToFile = Trim$(edcTo.Text)
    If Len(slToFile) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Sub
    End If
    'If InStr(slToFile, ":") = 0 Then
    If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
        slToFile = sgExportPath & slToFile
    End If
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        'hmTo = FreeFile
        'Open slToFile For Append As hmTo
        ilRet = gFileOpen(slToFile, "Append", hmTo)
        If ilRet <> 0 Then
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            edcTo.SetFocus
            Exit Sub
        End If
    Else
        ilRet = 0
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            edcTo.SetFocus
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    imExporting = True
    
    gAutomationAlertAndLogHandler "** Export Accounting **"
    gAutomationAlertAndLogHandler "* Storing Output into: " & slToFile
    gAutomationAlertAndLogHandler "* StartDate =" & edcStartDate.Text
    gAutomationAlertAndLogHandler "* EndDate =" & edcEndDate.Text
    If rbcType(0).Value = True Then gAutomationAlertAndLogHandler "* TransType = Invoices (IN, AN, HI)"
    If rbcType(1).Value = True Then gAutomationAlertAndLogHandler "* TransType = Cash (PI, PO)"
    If rbcType(2).Value = True Then gAutomationAlertAndLogHandler "* TransType = Receivables (All except HI)"
    
    ilRet = mReadRec()
    If ilRet Then
        lacInfo(0).Caption = "Export Successfully Completed"
        gAutomationAlertAndLogHandler "Export Successfully Completed"
    Else
        lacInfo(0).Caption = "Export Failed"
        gAutomationAlertAndLogHandler "Export Failed"
    End If
    lacInfo(1).Caption = "Export File: " & slToFile
    lacInfo(0).Visible = True
    lacInfo(1).Visible = True
    Close hmTo
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    cmcExport.Enabled = False
    Screen.MousePointer = vbDefault
    imExporting = False
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub

Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
    imDateBox = -1
End Sub

Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub

Private Sub cmcStartDate_GotFocus()
    Dim slStr As String
    If imDateBox <> 1 Then
        plcCalendar.Visible = False
        slStr = edcStartDate.Text
        If gValidDate(slStr) Then
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Else
            lacDate.Visible = False
        End If
    End If
    imDateBox = 1
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
    gCtrlGotFocus ActiveControl
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
End Sub

Private Sub cmcTo_GotFocus()
    plcCalendar.Visible = False
    imDateBox = -1
    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
End Sub

Private Sub edcEndDate_Change()
    Dim slStr As String
    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub

Private Sub edcEndDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    If imDateBox <> 2 Then
        plcCalendar.Visible = False
    End If
    plcCalendar.Move edcEndDate.Left + edcEndDate.Width + cmcEndDate.Width - plcCalendar.Width, edcEndDate.Top + edcEndDate.Height
    imDateBox = 2
    gCtrlGotFocus edcEndDate
End Sub

Private Sub edcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub

Private Sub edcStartDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    If imDateBox <> 1 Then
        plcCalendar.Visible = False
    End If
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
    imDateBox = 1
    gCtrlGotFocus edcStartDate
End Sub

Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
End Sub

Private Sub edcTo_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    plcCalendar.Visible = False
    imDateBox = -1
    lacInfo(0).Visible = False
    lacInfo(1).Visible = False
    gCtrlGotFocus ActiveControl
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
    frcTranType.Visible = False
    frcTranType.Visible = True
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
        plcTo.Visible = False
        plcTo.Visible = True
        frcTranType.Visible = False
        frcTranType.Visible = True
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmSOfficeCode
    Erase tmSalesOffice
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmPhf)
    btrDestroy hmPhf
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmIhf)
    btrDestroy hmIhf
    ilRet = btrClose(hmItf)
    btrDestroy hmItf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    
    Set ExpACC = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    plcCalendar.Visible = False
    imDateBox = -1
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If imDateBox = 1 Then
        slStr = edcStartDate.Text
    ElseIf imDateBox = 2 Then
        slStr = edcEndDate.Text
    End If
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
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
    Dim ilSof As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slOffice As String
    Dim slRegion As String
    Dim slCode As String
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    mInitBox
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpACC
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf)
    hmPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpACC
    On Error GoTo 0
    imPhfRecLen = Len(tmPhf)
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpACC
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpACC
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf) 'btrRecordLength(hmSbf)    'Get Sbf size

    hmIhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpACC
    On Error GoTo 0
    imIhfRecLen = Len(tmIhf)  'Get and save IHF record length

    hmItf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmItf, "", sgDBPath & "Itf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpACC
    On Error GoTo 0
    imItfRecLen = Len(tmItf)  'Get and save ARF record length

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CHF)", ExpACC
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ARF record length


    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CLF)", ExpACC
    On Error GoTo 0


    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CFF)", ExpACC
    On Error GoTo 0


    gCenterStdAlone ExpACC
    ilRet = gObtainAdvt()   'Build into tgCommAdf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainAgency() 'Build into tgCommAgf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainSalesperson() 'Build into tgMSlf
    If ilRet = False Then
        imTerminate = True
    End If
    smSOfficeCodeTag = ""
    ilRet = gPopOfficeSourceBox(ExpACC, lbcSOffice, tmSOfficeCode(), smSOfficeCodeTag)
    ReDim tmSalesOffice(0 To UBound(tmSOfficeCode)) As SALESOFFICE
    For ilSof = 0 To UBound(tmSOfficeCode) - 1 Step 1  'lbcSOfficeCode.ListCount - 1 Step 1
        slNameCode = tmSOfficeCode(ilSof).sKey    'lbcSOfficeCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "/", slOffice)
        ilRet = gParseItem(slName, 1, "/", slRegion)
        tmSalesOffice(ilSof).sOffice = slOffice
        tmSalesOffice(ilSof).sRegion = slRegion
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmSalesOffice(ilSof).iCode = Val(slCode)
    Next ilSof
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    imDateBox = -1
    imCalType = 0   'Standard
    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeExportRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize stand alone mode    *
'*                                                     *
'*******************************************************
Private Sub mMakeExportRec(tlRvf As RVF)


    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilSof As Integer
    Dim slAdfName As String
    Dim slProd As String
    Dim slAgfName As String
    Dim slSellVeh As String
    Dim slAirVeh As String
    Dim slCntrNo As String
    Dim slTranDate As String
    Dim slInvDate As String
    Dim slCheckNo As String
    Dim slInvNo As String
    Dim slTranType As String
    Dim slAction As String
    Dim slNet As String
    Dim slGross As String
    Dim slAgeMonth As String
    Dim slAgeYear As String
    Dim slSlsp As String
    Dim slSlspOffice As String
    Dim slCash As String
    Dim slGLRevNo As String
    Dim slEnteredDate As String
    Dim slRecord As String
    Dim ilExp As Integer
    Dim ilCount As Integer
    Dim dlTotalGross As Double          'change from long to double to prevent overflow
                                        'change all occurences of gLongToStrDec to gDblToStrDec
    Dim slTotalGross As String
    Dim slRunningNet As String
    Dim slRunningGross As String
    Dim ilRunningCount As Integer

    slAdfName = ""
    'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
    '    If tgCommAdf(ilLoop).iCode = tlRvf.iAdfCode Then
        ilLoop = gBinarySearchAdf(tlRvf.iAdfCode)
        If ilLoop <> -1 Then
            slAdfName = """" & Trim$(tgCommAdf(ilLoop).sName) & """"
    '        Exit For
        End If
    'Next ilLoop
    slProd = ""
    If tlRvf.lPrfCode > 0 Then
        tmPrfSrchKey.lCode = tlRvf.lPrfCode
        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) Then
            slProd = """" & Trim$(tmPrf.sName) & """"
        End If
    End If
    slAgfName = ""
    If tlRvf.iAgfCode > 0 Then
        'For ilLoop = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
        '    If tgCommAgf(ilLoop).iCode = tlRvf.iAgfCode Then
            ilLoop = gBinarySearchAgf(tlRvf.iAgfCode)
            If ilLoop <> -1 Then
                slAgfName = """" & Trim$(tgCommAgf(ilLoop).sName) & ", " & Trim$(tgCommAgf(ilLoop).sCityID) & """"
        '        Exit For
            End If
        'Next ilLoop
    End If
    slSellVeh = ""
    If tlRvf.iBillVefCode > 0 Then
        'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilLoop).iCode = tlRvf.iBillVefCode Then
            ilLoop = gBinarySearchVef(tlRvf.iBillVefCode)
            If ilLoop <> -1 Then
                slSellVeh = """" & Trim$(tgMVef(ilLoop).sName) & """"
        '        Exit For
            End If
        'Next ilLoop
    End If
    slAirVeh = ""
    If tlRvf.iAirVefCode > 0 Then
        'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilLoop).iCode = tlRvf.iAirVefCode Then
            ilLoop = gBinarySearchVef(tlRvf.iAirVefCode)
            If ilLoop <> -1 Then
                slAirVeh = """" & Trim$(tgMVef(ilLoop).sName) & """"
        '        Exit For
            End If
        'Next ilLoop
    End If
    If tlRvf.lCntrNo > 0 Then
        slCntrNo = Trim$(str$(tlRvf.lCntrNo))
    Else
        slCntrNo = ""
    End If
    gUnpackDate tlRvf.iTranDate(0), tlRvf.iTranDate(1), slTranDate
    slTranDate = gAdjYear(slTranDate)
    gUnpackDate tlRvf.iInvDate(0), tlRvf.iInvDate(1), slInvDate
    slInvDate = gAdjYear(slInvDate)
    '6/9/15: Check number changed to string
    'If tlRvf.lCheckNo > 0 Then
    '    slCheckNo = Trim$(str$(tlRvf.lCheckNo))
    'Else
    '    slCheckNo = ""
    'End If
    slCheckNo = Trim$(tlRvf.sCheckNo)     '6-16-20 remove the Str$  slCheckNo = Trim$(str$(tlRvf.sCheckNo))
    If tlRvf.lInvNo > 0 Then
        slInvNo = Trim$(str$(tlRvf.lInvNo))
    Else
        slInvNo = ""
    End If
    slTranType = Trim$(tlRvf.sTranType)
    slAction = Trim$(tlRvf.sAction)
    gPDNToStr tlRvf.sGross, 2, slGross
    gPDNToStr tlRvf.sNet, 2, slNet
    slAgeMonth = Trim$(str$(tlRvf.iAgePeriod))
    slAgeYear = Trim$(str$(tlRvf.iAgingYear))
    slSlsp = ""
    slSlspOffice = ""
    If tlRvf.iSlfCode > 0 Then
        For ilLoop = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            If tgMSlf(ilLoop).iCode = tlRvf.iSlfCode Then
                slSlsp = """" & Trim$(tgMSlf(ilLoop).sFirstName) & " " & Trim$(tgMSlf(ilLoop).sLastName) & """"
                For ilSof = 0 To UBound(tmSalesOffice) - 1 Step 1  'lbcSOfficeCode.ListCount - 1 Step 1
                    If tgMSlf(ilLoop).iSofCode = tmSalesOffice(ilSof).iCode Then
                        slSlspOffice = """" & Trim$(tmSalesOffice(ilSof).sOffice) & """"
                        Exit For
                    End If
                Next ilSof
                Exit For
            End If
        Next ilLoop
    End If
    slCash = tlRvf.sCashTrade

    ilCount = 0
    dlTotalGross = 0
    For ilExp = 0 To UBound(tmExpAccInfo) - 1 Step 1
        If (tmExpAccInfo(ilExp).iVefCode = tlRvf.iBillVefCode) Then
            ilCount = ilCount + 1
            dlTotalGross = dlTotalGross + tmExpAccInfo(ilExp).dGross
        End If
    Next ilExp
    slTotalGross = gDblToStrDec(dlTotalGross, 2)
    gUnpackDate tlRvf.iDateEntrd(0), tlRvf.iDateEntrd(1), slEnteredDate
    slEnteredDate = gAdjYear(slEnteredDate)
    slRunningNet = "0.00"
    slRunningGross = "0.00"
    ilRunningCount = 0
    For ilExp = 0 To UBound(tmExpAccInfo) - 1 Step 1
        If (tmExpAccInfo(ilExp).iVefCode = tlRvf.iBillVefCode) Then
            ilRunningCount = ilRunningCount + 1
            If slCash <> "T" Then
                slGLRevNo = Trim$(tmExpAccInfo(ilExp).sGLCashNo)
            Else
                slGLRevNo = Trim$(tmExpAccInfo(ilExp).sGLTradeNo)
            End If
            If ((slGLRevNo >= "A") And (slGLRevNo <= "Z")) Or ((slGLRevNo >= "0") And (slGLRevNo <= "9")) Then
            Else
                slGLRevNo = ""
            End If
            gPDNToStr tlRvf.sGross, 2, slGross
            gPDNToStr tlRvf.sNet, 2, slNet
            If ilCount > 1 Then
                If ilRunningCount < ilCount Then
                    slGross = gDivStr(gMulStr(gDblToStrDec(tmExpAccInfo(ilExp).dGross, 2), slGross), slTotalGross)
                    slRunningGross = gAddStr(slRunningGross, slGross)
                    slNet = gDivStr(gMulStr(gDblToStrDec(tmExpAccInfo(ilExp).dGross, 2), slNet), slTotalGross)
                    slRunningNet = gAddStr(slRunningNet, slNet)
                Else
                    slGross = gSubStr(slGross, slRunningGross)
                    slNet = gSubStr(slNet, slRunningNet)
                End If
            End If
            slRecord = slAdfName & "," & slProd & "," & slAgfName & "," & slSellVeh & "," & slAirVeh & "," & slCntrNo & "," & slTranDate & "," & slInvDate & "," & slCheckNo & "," & slInvNo & "," & slTranType & "," & slAction & "," & slNet & "," & slGross & "," & slAgeMonth & "," & slAgeYear & "," & slSlsp & "," & slSlspOffice & "," & slCash & "," & slGLRevNo & "," & slEnteredDate
            Print #hmTo, slRecord
        End If
    Next ilExp
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOk As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer

    llNoRec = gExtNoRec(imRvfRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmRvf   'Clear any previous extend operation
    'Note:  Key 4 is in contract number order
    ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        'Do While ilRet = BTRV_ERR_NONE
        Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") 'Set extract limits (all records)
        tlCharTypeBuff.sType = "A"
        ilOffSet = gFieldOffset("Rvf", "RvfType")
        ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        gPackDate smTranStart, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
        ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        gPackDate smTranEnd, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
        ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilRet = btrExtAddField(hmRvf, 0, imRvfRecLen)  'Extract record
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                    imRvfRecLen = Len(tmRvf)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        ilOk = False
                        If rbcType(0).Value Then
                            If (StrComp(Trim$(tmRvf.sTranType), "IN", 1) = 0) Or (StrComp(Trim$(tmRvf.sTranType), "AN", 1) = 0) Or (StrComp(Trim$(tmRvf.sTranType), "HI", 1) = 0) Then
                                ilOk = True
                            End If
                        ElseIf rbcType(1).Value Then
                            If (StrComp(Trim$(tmRvf.sTranType), "PI", 1) = 0) Or (StrComp(Trim$(tmRvf.sTranType), "PO", 1) = 0) Then
                                ilOk = True
                            End If
                        ElseIf rbcType(2).Value Then
                            If StrComp(Trim$(tmRvf.sTranType), "HI", 1) <> 0 Then
                                ilOk = True
                            End If
                        End If
                        If ilOk Then
                            If mGetDollarsForInstallmentDistribution(tmRvf) Then
                                mMakeExportRec tmRvf
                            End If
                        End If
                        ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
                        Loop
                    Loop
                End If
            End If
        End If
    End If

    If rbcType(2).Value = False Then
        llNoRec = gExtNoRec(imPhfRecLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hmPhf   'Clear any previous extend operation
        ilRet = btrGetFirst(hmPhf, tmPhf, imPhfRecLen, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            'Do While ilRet = BTRV_ERR_NONE
            Call btrExtSetBounds(hmPhf, llNoRec, -1, "UC", "RVF", "") 'Set extract limits (all records)
            gPackDate smTranStart, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            ilRet = btrExtAddLogicConst(hmPhf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            gPackDate smTranEnd, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            ilRet = btrExtAddLogicConst(hmPhf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            ilRet = btrExtAddField(hmPhf, 0, imPhfRecLen)  'Extract record
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrExtGetNext(hmPhf, tmPhf, imPhfRecLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                        imPhfRecLen = Len(tmPhf)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmPhf, tmPhf, imPhfRecLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE

                            ilOk = False
                            If rbcType(0).Value Then
                                If (StrComp(Trim$(tmPhf.sTranType), "IN", 1) = 0) Or (StrComp(Trim$(tmPhf.sTranType), "AN", 1) = 0) Or (StrComp(Trim$(tmPhf.sTranType), "HI", 1) = 0) Then
                                    ilOk = True
                                End If
                            ElseIf rbcType(1).Value Then
                                If (StrComp(Trim$(tmPhf.sTranType), "PI", 1) = 0) Or (StrComp(Trim$(tmPhf.sTranType), "PO", 1) = 0) Then
                                    ilOk = True
                                End If
                            ElseIf rbcType(2).Value Then
                            End If
                            If ilOk Then
                                If mGetDollarsForInstallmentDistribution(tmPhf) Then
                                    mMakeExportRec tmPhf
                                End If
                            End If
                            ilRet = btrExtGetNext(hmPhf, tmPhf, imPhfRecLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmPhf, tmPhf, imPhfRecLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            End If
        End If
    End If

    mReadRec = True
    Exit Function

    On Error GoTo 0
    mReadRec = False
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpACC
    igManUnload = NO
End Sub

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If imDateBox = 1 Then
                    edcStartDate.Text = Format$(llDate, "m/d/yy")
                    edcStartDate.SelStart = 0
                    edcStartDate.SelLength = Len(edcStartDate.Text)
                    imBypassFocus = True
                    edcStartDate.SetFocus
                    Exit Sub
                ElseIf imDateBox = 2 Then
                    edcEndDate.Text = Format$(llDate, "m/d/yy")
                    edcEndDate.SelStart = 0
                    edcEndDate.SelLength = Len(edcEndDate.Text)
                    imBypassFocus = True
                    edcEndDate.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDateBox = 1 Then
        edcStartDate.SetFocus
    ElseIf imDateBox = 2 Then
        edcEndDate.SetFocus
    End If
End Sub

Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub

Private Sub rbcType_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    imDateBox = -1
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Export Accounting"
End Sub

Private Function mGetDollarsForInstallmentDistribution(tlRvf As RVF) As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slEndDate As String
    Dim llMoStartDate As Long
    Dim ilSpots As Integer
    Dim llPrice As Long
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilLoop As Integer
    Dim slGLCashNo As String
    Dim slGLTradeNo As String
    Dim slGLCashMMNo As String
    Dim slGLTradeMMNo As String
    Dim ilIndex As Integer
    Dim ilVpf As Integer

    For ilLoop = 0 To UBound(tmExpAccInfo) - 1 Step 1
        If (tmExpAccInfo(ilLoop).lCntrNo = tlRvf.lCntrNo) And (tmExpAccInfo(ilLoop).iVefCode = tlRvf.iBillVefCode) Then
            mGetDollarsForInstallmentDistribution = True
            Exit Function
        End If
    Next ilLoop
    ReDim tmExpAccInfo(0 To 0) As EXPACCINFO
    imCHFRecLen = Len(tmChf)
    tmChfSrchKey1.lCntrNo = tlRvf.lCntrNo
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlRvf.lCntrNo) And (tmChf.sSchStatus <> "F")
        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlRvf.lCntrNo) And (tmChf.sSchStatus = "F") Then
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChf.lCode, False, tmChfAcc, tmClfAcc(), tmCffAcc())
        If Not ilRet Then
            mGetDollarsForInstallmentDistribution = False
            Exit Function
        End If
        For ilClf = LBound(tmClfAcc) To UBound(tmClfAcc) - 1 Step 1
            'Bypass Package lines
            If (tmClfAcc(ilClf).ClfRec.sType <> "O") And (tmClfAcc(ilClf).ClfRec.sType <> "A") And (tmClfAcc(ilClf).ClfRec.sType <> "E") Then
                ilVpf = gBinarySearchVpf(tmClfAcc(ilClf).ClfRec.iVefCode)
                If ilVpf <> -1 Then
                    slGLCashNo = Trim$(tgVpf(ilVpf).sBilledRevenue)
                    slGLTradeNo = Trim$(tgVpf(ilVpf).sBilledTrade)
                Else
                    mGetDollarsForInstallmentDistribution = False
                    Exit Function
                End If
                'If tmChf.sInstallDefined <> "Y" Then
                '    ilFound = False
                '    For ilLoop = 0 To UBound(tmExpAccInfo) - 1 Step 1
                '        If (tmExpAccInfo(ilLoop).iVefCode = tmClfAcc(ilClf).ClfRec.iVefCode) Then
                '            ilFound = True
                '        End If
                '    Next ilLoop
                '    If Not ilFound Then
                '        tmExpAccInfo(UBound(tmExpAccInfo)).lCntrNo = tlRvf.lCntrNo
                '        tmExpAccInfo(UBound(tmExpAccInfo)).iVefCode = tmClfAcc(ilClf).ClfRec.iVefCode
                '        tmExpAccInfo(UBound(tmExpAccInfo)).sInstallDefined = tmChf.sInstallDefined
                '        tmExpAccInfo(UBound(tmExpAccInfo)).lGross = -1
                '        tmExpAccInfo(UBound(tmExpAccInfo)).sGLCashNo = slGLCashNo
                '        tmExpAccInfo(UBound(tmExpAccInfo)).sGLTradeNo = slGLTradeNo
                '        ReDim Preserve tmExpAccInfo(0 To UBound(tmExpAccInfo) + 1) As EXPACCINFO
                '    End If
                'Else
                    llPrice = 0
                    ilCff = tmClfAcc(ilClf).iFirstCff
                    Do While ilCff <> -1
                        gUnpackDate tmCffAcc(ilCff).CffRec.iStartDate(0), tmCffAcc(ilCff).CffRec.iStartDate(1), slStartDate
                        gUnpackDate tmCffAcc(ilCff).CffRec.iEndDate(0), tmCffAcc(ilCff).CffRec.iEndDate(1), slEndDate
                        llStartDate = gDateValue(slStartDate)
                        slStartDate = gObtainPrevMonday(slStartDate)
                        llMoStartDate = gDateValue(slStartDate)
                        llEndDate = gDateValue(slEndDate)
                        If llStartDate <= llEndDate Then
                            ilIndex = -1
                            For ilLoop = 0 To UBound(tmExpAccInfo) - 1 Step 1
                                If (tmExpAccInfo(ilLoop).iVefCode = tmClfAcc(ilClf).ClfRec.iVefCode) Then
                                    ilIndex = ilLoop
                                End If
                            Next ilLoop
                            If ilIndex = -1 Then
                                ilIndex = UBound(tmExpAccInfo)
                                tmExpAccInfo(ilIndex).lCntrNo = tlRvf.lCntrNo
                                tmExpAccInfo(ilIndex).iVefCode = tmClfAcc(ilClf).ClfRec.iVefCode
                                tmExpAccInfo(ilIndex).sInstallDefined = tmChf.sInstallDefined
                                tmExpAccInfo(ilIndex).dGross = 0
                                tmExpAccInfo(ilIndex).sGLCashNo = slGLCashNo
                                tmExpAccInfo(ilIndex).sGLTradeNo = slGLTradeNo
                                ReDim Preserve tmExpAccInfo(0 To ilIndex + 1) As EXPACCINFO
                            End If
                            For llDate = llMoStartDate To llEndDate Step 7
                                If tmCffAcc(ilCff).CffRec.sDyWk = "D" Then
                                    ilSpots = 0
                                    For ilDay = 0 To 6 Step 1
                                        ilSpots = ilSpots + tmCffAcc(ilCff).CffRec.iDay(ilDay)
                                    Next ilDay
                                Else
                                    ilSpots = tmCffAcc(ilCff).CffRec.iSpotsWk + tmCffAcc(ilCff).CffRec.iXSpotsWk
                                End If
                                llPrice = 0
                                If ilSpots > 0 Then
                                    If tmCffAcc(ilCff).CffRec.sPriceType = "T" Then
                                        llPrice = ilSpots * tmCffAcc(ilCff).CffRec.lActPrice
                                    End If
                                End If
                                tmExpAccInfo(ilIndex).dGross = tmExpAccInfo(ilIndex).dGross + llPrice
                            Next llDate
                        End If
                        ilCff = tmCffAcc(ilCff).iNextCff
                    Loop
                'End If
            End If
        Next ilClf
        'SBF: NTR and Multi-Media
        imSbfRecLen = Len(tmSbf)
        tmSbfSrchKey0.lChfCode = tmChf.lCode
        tmSbfSrchKey0.iDate(0) = 0
        tmSbfSrchKey0.iDate(1) = 0
        tmSbfSrchKey0.sTranType = " "
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tmChf.lCode)
            If (tmSbf.sTranType = "I") Then
                ilVpf = gBinarySearchVpf(tmSbf.iBillVefCode)
                If ilVpf <> -1 Then
                    slGLCashNo = Trim$(tgVpf(ilVpf).sBilledRevenue)
                    slGLTradeNo = Trim$(tgVpf(ilVpf).sBilledTrade)
                Else
                    mGetDollarsForInstallmentDistribution = False
                    Exit Function
                End If
                If tmSbf.iIhfCode <= 0 Then
                    'NTR
                    ilIndex = -1
                    For ilLoop = 0 To UBound(tmExpAccInfo) - 1 Step 1
                        If (tmExpAccInfo(ilLoop).iVefCode = tmSbf.iBillVefCode) And (Trim$(tmExpAccInfo(ilLoop).sGLCashNo) = slGLCashNo) And (Trim$(tmExpAccInfo(ilLoop).sGLTradeNo) = slGLTradeNo) Then
                            ilIndex = ilLoop
                            Exit For
                        End If
                    Next ilLoop
                    If ilIndex = -1 Then
                        ilIndex = UBound(tmExpAccInfo)
                        tmExpAccInfo(ilIndex).lCntrNo = tlRvf.lCntrNo
                        tmExpAccInfo(ilIndex).iVefCode = tmSbf.iBillVefCode
                        tmExpAccInfo(ilIndex).sInstallDefined = tmChf.sInstallDefined
                        tmExpAccInfo(ilIndex).dGross = 0
                        tmExpAccInfo(ilIndex).sGLCashNo = slGLCashNo
                        tmExpAccInfo(ilIndex).sGLTradeNo = slGLTradeNo
                        ReDim Preserve tmExpAccInfo(0 To ilIndex + 1) As EXPACCINFO
                    End If
                    tmExpAccInfo(ilIndex).dGross = tmExpAccInfo(ilIndex).dGross + tmSbf.lGross * tmSbf.iNoItems
                Else
                    'Multi-Media
                    slGLCashMMNo = slGLCashNo
                    slGLTradeMMNo = slGLTradeNo
                    tmIhfSrchKey0.iCode = tmSbf.iIhfCode
                    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmIhf.iItfCode > 0 Then
                            tmItfSrchKey0.iCode = tmIhf.iItfCode
                            ilRet = btrGetEqual(hmItf, tmItf, imItfRecLen, tmItfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                If Trim$(tmItf.sGLCash) <> "" Then
                                    slGLCashMMNo = Trim$(tmItf.sGLCash)
                                End If
                                If Trim$(tmItf.sGLTrade) <> "" Then
                                    slGLTradeMMNo = Trim$(tmItf.sGLTrade)
                                End If
                            End If
                        End If
                    End If
                    ilIndex = -1
                    For ilLoop = 0 To UBound(tmExpAccInfo) - 1 Step 1
                        If (tmExpAccInfo(ilLoop).iVefCode = tmSbf.iBillVefCode) And (Trim$(tmExpAccInfo(ilLoop).sGLCashNo) = slGLCashMMNo) And (Trim$(tmExpAccInfo(ilLoop).sGLTradeNo) = slGLTradeMMNo) Then
                            ilIndex = ilLoop
                            Exit For
                        End If
                    Next ilLoop
                    If ilIndex = -1 Then
                        ilIndex = UBound(tmExpAccInfo)
                        tmExpAccInfo(ilIndex).lCntrNo = tlRvf.lCntrNo
                        tmExpAccInfo(ilIndex).iVefCode = tmSbf.iBillVefCode
                        tmExpAccInfo(ilIndex).sInstallDefined = tmChf.sInstallDefined
                        tmExpAccInfo(ilIndex).dGross = 0
                        tmExpAccInfo(ilIndex).sGLCashNo = slGLCashMMNo
                        tmExpAccInfo(ilIndex).sGLTradeNo = slGLTradeMMNo
                        ReDim Preserve tmExpAccInfo(0 To ilIndex + 1) As EXPACCINFO
                    End If
                    tmExpAccInfo(ilIndex).dGross = tmExpAccInfo(ilIndex).dGross + tmSbf.lGross * tmSbf.iNoItems
                End If
            End If
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        mGetDollarsForInstallmentDistribution = True
        Exit Function
    Else
        mGetDollarsForInstallmentDistribution = False
        Exit Function
    End If
End Function

