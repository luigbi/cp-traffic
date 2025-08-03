VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmRptNoSel 
   Caption         =   "Site Options"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffRptNoSel.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
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
      FormDesignHeight=   5835
      FormDesignWidth =   7125
   End
   Begin VB.CommandButton cmdCrystalTemp 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   255
      Width           =   1935
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4590
      TabIndex        =   1
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
         ItemData        =   "AffRptNoSel.frx":08CA
         Left            =   1050
         List            =   "AffRptNoSel.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   825
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmRptNoSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim smExportName As String      'name given to a SAVE-TO file
    Dim smRptName As String         'full report name of crystal .rpt

Private Sub cmdCrystalTemp_Click()

    cmdCrystalTemp.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    mNoOptionsReport
    
    cmdCrystalTemp.Enabled = True               're-enable Gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    
End Sub
Sub mNoOptionsReport()

    Dim ilRet As Integer
    Dim ilRptDest As Integer        'output to display, print, save to
    'Dim slExportName As String      'name given to a SAVE-TO file
    Dim ilExportType As Integer     'SAVE-TO output type
    'Dim slRptName As String         'full report name of crystal .rpt
    
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
'    slRptName = "afSiteOptions.rpt"
'    slExportName = "afSiteOptions"
    'dan to do changed with 58 rollback
'    frmCrystal.gCrystlReports "", ilExportType, ilRptDest, smRptName, smExportName, True
    frmCrystal.gCrystlReports "", ilExportType, ilRptDest, smRptName, smExportName

    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRptNoSel-mNoOptionsReport"
End Sub

Private Sub cmdDone_Click()
    
    Unload frmRptNoSel

End Sub


Private Sub cmdReturn_Click()

    frmReports.Show
    Unload frmRptNoSel
    
End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmRptNoSel
    gCenterForm frmRptNoSel
    gPopExportTypes cboFileType
    cboFileType.Enabled = True
    mAdjustForSpecificReport
End Sub
Private Sub mAdjustForSpecificReport()
    Me.Caption = "Site Options"
    smRptName = "afSiteOptions.rpt"
    smExportName = "afSiteOptions"

End Sub

Private Sub Form_Initialize()

    mInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmRptNoSel = Nothing

End Sub

