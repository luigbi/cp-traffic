VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrSchedRpt 
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7065
   Begin VB.Frame frcOption 
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
      Height          =   1980
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   6495
      Begin VB.CheckBox ckcSchOption 
         Caption         =   "Show all fields"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4275
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   2310
   End
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
      FormDesignHeight=   4290
      FormDesignWidth =   7065
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4455
      TabIndex        =   7
      Top             =   720
      Width           =   1920
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   6
      Top             =   240
      Width           =   2685
   End
   Begin VB.Frame frcOutput 
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
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   480
         Width           =   2190
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2310
      End
   End
End
Attribute VB_Name = "EngrSchedRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrSchedRpt - a report of the filtered Schedule
'*                 from Snapshot button.  All this form
'*                 does is to determine how to output the
'*                 results
'*
'*  Created 3-28-06
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit


Private Sub cmdDone_Click()
    igReturnCallStatus = CALLCANCELLED            'dont produce the report flag
    Unload EngrSchedRpt
End Sub

Private Sub cmdReport_Click()
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim SQLQuery As String
    Dim ilListIndex As Integer
    Dim slSQLFromDate As String
    Dim slSQLToDAte As String
    Dim slDate As String
    
    Dim ilRptDest As Integer            'disply, print, save as file
    Dim slRptName As String
    Dim slExportName As String
    Dim slRptType As String
    Dim llResult As Long
    Dim ilExportType As Integer
    Dim llGridRow As Long
    Dim slStr As String
    Dim llTime As Long
    Dim llAirDate As Long
    Dim slFilter As String              'filters selected by user
    Dim ilLoop As Integer
    Dim slOperator As String * 2        'operator for filter
  
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbDefault
    sgReturnCallName = ""   'assume no exporting, no export filename
    sgReturnOption = ""
    If igRptIndex = SCHED_RPT Then              'snapshot from schedule screen
        If ckcSchOption.Value = vbChecked Then      'include all fields
            sgReturnOption = "ALL"
            sgReturnCallName = "SchedAll"       'this is for the export name only if Save To Selected
        Else
            sgReturnCallName = "Sched"          'this is for the export name only if Save To selected
        End If
    ElseIf igRptIndex = LIBRARYEVENT_RPT Then        'snapshot from library screen
        If ckcSchOption.Value = vbChecked Then      'include all fields
            sgReturnOption = "ALL"
            sgReturnCallName = "LibEvtSSAll"       'this is for the export name only if Save To Selected
        Else
            sgReturnCallName = "LibEvtSS"          'this is for the export name only if Save To selected
        End If
    ElseIf igRptIndex = ASAIRCOMPARE_RPT Then
        If ckcSchOption.Value = vbChecked Then      'include all fields
            sgReturnOption = "ALL"
            sgReturnCallName = "AsAirCompareAll"       'this is for the export name only if Save To Selected
        Else
            sgReturnCallName = "AsAirCompare"          'this is for the export name only if Save To selected
        End If
    End If
    
    If optRptDest(0).Value = True Then
        igRptDest = 0
        'sgReturnOption = ""     'no exporting, clear the export name
    ElseIf optRptDest(1).Value = True Then
        igRptDest = 1
        'sgReturnOption = ""     'no exporting, clear the export  name
    ElseIf optRptDest(2).Value = True Then
        igRptDest = 2
        igExportType = cboFileType.ListIndex
    Else
         Exit Sub
    End If
    igReturnCallStatus = CALLDONE            'produce the report flag
    Unload EngrSchedRpt
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Site Rpt-cmdReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Site Rpt-cmdReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdReturn_Click()
    EngrReports.Show
    Unload EngrSchedRpt
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrSchedRpt
    gCenterForm EngrSchedRpt
End Sub

Private Sub Form_Load()
    'EngrUserRpt.Caption = "User - " & sgClientName
    gPopExportTypes cboFileType
    cboFileType.Enabled = False
    'gChangeCaption frcOption
    If igRptIndex = SCHED_RPT Then      'from snapshot
        ckcSchOption.Visible = True
    ElseIf igRptIndex = LIBRARY_RPT Then    'from snapshot
        ckcSchOption.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrSchedRpt = Nothing
End Sub
Private Sub optDisplay_Click(Index As Integer)

End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub
Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub
