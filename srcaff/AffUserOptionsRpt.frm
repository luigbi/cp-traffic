VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmUserOptionsRpt 
   Caption         =   "User Options"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffUserOptionsRpt.frx":0000
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
      Height          =   3030
      Left            =   240
      TabIndex        =   1
      Top             =   1785
      Width           =   6705
      Begin VB.CheckBox ckcDormantUsers 
         Caption         =   "Include Dormant Users"
         Height          =   255
         Left            =   315
         TabIndex        =   9
         Top             =   315
         Width           =   2235
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
         ItemData        =   "AffUserOptionsRpt.frx":08CA
         Left            =   1050
         List            =   "AffUserOptionsRpt.frx":08CC
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
End
Attribute VB_Name = "frmUserOptionsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCrystalTemp_Click()

    mUserOptionsReport
    
End Sub
Sub mUserOptionsReport()

    Dim ilRet As Integer
    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim slRptName As String         'full report name of crystal .rpt
    
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
    
    cmdCrystalTemp.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    'guide gets to see key codes
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        sgCrystlFormula1 = 1
    Else
        sgCrystlFormula1 = 0
    End If
    'see dormant?
    sgCrystlFormula2 = ckcDormantUsers.Value
    slRptName = "afUserOptions.rpt"
    slExportName = "afUserOptions"
    gUserActivityLog "E", sgReportListName & ": Prepass"
    'dan todo change for rollback
    'frmCrystal.gCrystlReports "", ilExportType, ilRptDest, slRptName, slExportName, True
    frmCrystal.gCrystlReports "", ilExportType, ilRptDest, slRptName, slExportName
    
    cmdCrystalTemp.Enabled = True
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmUserOptionsRpt-mUserlOptionsReport"
End Sub

Private Sub cmdDone_Click()
    
    Unload frmUserOptionsRpt

End Sub


Private Sub cmdReturn_Click()

    frmReports.Show
    Unload frmUserOptionsRpt
    
End Sub
Private Sub Form_Load()

        
    frmUserOptionsRpt.Caption = "User Options Report - " & sgClientName

End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmUserOptionsRpt
    gCenterForm frmUserOptionsRpt
    gPopExportTypes cboFileType
    cboFileType.Enabled = True

End Sub

Private Sub Form_Initialize()

    mInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmUserOptionsRpt = Nothing

End Sub

