VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "CRViewer.dll"
Begin VB.Form EngrViewReport 
   Caption         =   "Crystal Reports"
   ClientHeight    =   4335
   ClientLeft      =   2055
   ClientTop       =   2355
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6945
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      lastProp        =   600
      _cx             =   11033
      _cy             =   6800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "EngrViewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************
' Copyright: Counterpoint Software, Inc.
' Date: September 2004
' Name: EngrViewReport
'**************************************************************************
Public Report As New CRAXDRT.Report

'**************************************************************************
' Copyright: Counterpoint Software, Inc. 2002
' Created by: Doug Smith
' Date: August 2002
'
' Display the report on the screen
'**************************************************************************
Private Sub Form_Load()
    EngrViewReport.Caption = "Crystal Reports - " & sgClientName
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
     
End Sub

'**************************************************************************
' Copyright: Counterpoint Software, Inc. 2002
' Created by: Doug Smith
' Date: August 2002
'**************************************************************************
Private Sub Form_Resize()
 
 With CRViewer1
    .Top = 0
    .Left = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
    Unload EngrViewReport
End Sub
