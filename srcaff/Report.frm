VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form Report 
   Appearance      =   0  'Flat
   Caption         =   "Reports"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   3255
   ClientWidth     =   9375
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
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   9375
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer crviewer1 
      Height          =   3855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _cx             =   7011
      _cy             =   6800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   0   'False
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin VB.PictureBox pbcStart 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   5850
      Width           =   45
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public bmShowExportForm As Boolean


Public Sub crviewer1_ExportButtonClicked(UseDefault As Boolean)

    fCrViewerExport.Show vbModal
    If bmShowExportForm Then
        UseDefault = True
    Else
        UseDefault = False
    End If
    bmShowExportForm = False
End Sub

Private Sub CRViewer1_GotFocus()
On Error GoTo ERRORBOX
    With crviewer1
        .Width = Report.Width - 250
        .Height = Report.Height - 500  'need the 500 for scrollbar
        .Left = 0
        .Top = 0
        .ViewReport
    End With
    Exit Sub
ERRORBOX:
    MsgBox "problem with report viewer", vbExclamation + vbOKOnly, "Report Viewer"
End Sub


Private Sub Form_Load()
    Dim myConnections As CCsiSystemConnection
    
    Set myConnections = New CCsiSystemConnection
    If myConnections.System = 1 Then
        gUserActivityLog "E", sgReportListName & ": Prepass"
    End If
    gUserActivityLog "S", sgReportListName & ": Display"
    With crviewer1  'these are what are different than default.
        .DisplayTabs = False
        .DisplayGroupTree = False
        .EnableAnimationCtrl = False
        .EnableCloseButton = False
        .EnableRefreshButton = False
        .EnableStopButton = False
        If Not ogReport Is Nothing Then
            .ReportSource = ogReport.ReportBin
            If ogReport.iLastPrintJob > 1 Then
                .DisplayGroupTree = True
            End If
        End If
    End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
'this if shouldn't be necessary;  we only display this after working with last print job, so should be able to close object and move on;  but program will crash!
'not closing object and going back into loop fixes issue, but not sure why.
        If ogReport.iLastPrintJob > 1 Then
           ogReport.UserWantsToQuit = True
       Else
            Set ogReport = Nothing
       End If
        Set Report = Nothing   'Remove data segment
        gUserActivityLog "E", sgReportListName & ": Display"
        'Dan 8/26/11 run this by Dick
        Unload Me
        
End Sub


