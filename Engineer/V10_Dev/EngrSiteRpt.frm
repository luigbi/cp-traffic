VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrSiteRpt 
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
      TabIndex        =   6
      Top             =   1860
      Width           =   6495
      Begin VB.TextBox edcTo 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox edcFrom 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frcOldNew 
         Caption         =   "Show"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
         Begin VB.OptionButton optOldNew 
            Caption         =   "History"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optOldNew 
            Caption         =   "Current"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label lacChangeTo 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lacChangeDates 
         Caption         =   "Enter change dates- From"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
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
      Caption         =   "Done"
      Height          =   375
      Left            =   4455
      TabIndex        =   9
      Top             =   1200
      Width           =   1920
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4275
      TabIndex        =   8
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   7
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
Attribute VB_Name = "EngrSiteRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrSiteRpt - a report of Site options
'*
'*  Created September,  2004
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private Sub cmdDone_Click()
    Unload EngrSiteRpt
End Sub

Private Sub cmdReport_Click()
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim SQLQuery As String
    Dim ilListIndex As Integer
    Dim slSQLFromDate As String
    Dim slSQLToDAte As String
    Dim slDate As String
    Dim ilLoop As Integer
  
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass


    If optRptDest(0).Value = True Then
       ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilRptDest = 2
        ilExportType = cboFileType.ListIndex
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
  
    slRptName = slRptName & ".rpt"      'concatenate the crystal report name plus extension
    SQLQuery = "Select * from SOE_Site_Option, UIE_User_Info where soeCurrent = 'Y' and soeuiecode = uiecode"

    
    EngrCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
    Screen.MousePointer = vbDefault
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
    Unload EngrSiteRpt
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrSiteRpt
    gCenterForm EngrSiteRpt
End Sub

Private Sub Form_Load()
    'EngrUserRpt.Caption = "User - " & sgClientName
    gPopExportTypes cboFileType
    cboFileType.Enabled = False
    gChangeCaption frcOption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrSiteRpt = Nothing
End Sub

Private Sub optDisplay_Click(Index As Integer)

End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub optOldNew_Click(Index As Integer)
    If Index = 0 Then               'current
        lacChangeDates.Visible = False
        lacChangeTo.Visible = False
        edcFrom.Visible = False
        edcTo.Visible = False
    Else
        lacChangeDates.Visible = True
        lacChangeTo.Visible = True
        edcFrom.Visible = True
        edcTo.Visible = True
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub
