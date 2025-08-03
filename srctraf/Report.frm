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
   Begin VB.Label lbcNoReport 
      Caption         =   "There is no report to display"
      Height          =   1215
      Left            =   795
      TabIndex        =   2
      Top             =   795
      Width           =   2475
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
Private bmPrintButtonClicked As Boolean

Private Sub crviewer1_PrintButtonClicked(UseDefault As Boolean)
'    '11/30/11 added to stop issue of failure to print because of multi reports
'    '12/14/11 now lost ability to limit pages: add new function to reporthelper
'    If Not ogReport Is Nothing Then
'        If ogReport.iLastPrintJob > 1 Then
'            'ogReport.PrintOutWithExportForm
'            ogReport.PrintOut 1
'            UseDefault = False
'        Else
'            UseDefault = True
'        End If
'    Else
'        UseDefault = True
'    End If
    'TTP 10511 - Contract Snapshot: display report then printing to printer or print to PDF causes Traffic crash on the Cloud when contract has NTR and all pages are printed (detail, NTR, and summary page)
    'Check if we're viewing a Contract Report..
    If Mid(UCase(ogReport.Reports.Keys(0)), 1, 2) = "BR" Then
        crviewer1.ShowLastPage
    End If
    bmPrintButtonClicked = True             '3-1-17 User wants to print after displaying
End Sub

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
        .Height = Report.Height - 500 'need the 500 for scrollbar
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
        .EnableDrillDown = False            '9-17-15 disable drill down because user cannot get back to full report
                                            'if enabling it, set displaytabs = true & enableclosebutton = true; all reports are not drill down compatible (code wise).
                                            'also, in .rpt can set Show all headers in drill down if drill down will be enabled
    On Error Resume Next
        If Not ogReport Is Nothing Then
            .ReportSource = ogReport.ReportBin
            If Not .ReportSource Is Nothing Then
                If ogReport.iLastPrintJob > 1 Then
                    .DisplayGroupTree = True
                End If
                lbcNoReport.Visible = False
            Else
              .Visible = False
              lbcNoReport.Visible = True
            End If
        End If
    End With
End Sub

Private Sub Form_Resize()
    '4/11/22 - JW - TTP 10446 - Reports: Restore Min/Max buttons & Resize Report viewer
    If Me.WindowState = vbMinimized Then Exit Sub
    crviewer1.Width = Me.ScaleWidth
    crviewer1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
'this if shouldn't be necessary;  we only display this after working with last print job, so should be able to close object and move on;  but program will crash!
'not closing object and going back into loop fixes issue, but not sure why.
'               2-7-17 'do not set ogreport to nothing; it has to continue to use for archiving (if invoicing).  all others places do clean it up
'               .userWantsToQuit does nothing;not used
'        If ogReport.iLastPrintJob > 1 Then
'           ogReport.UserWantsToQuit = True
'       Else
'            Set ogReport = Nothing
'       End If
    If bmPrintButtonClicked Then            '3-1-17 if printed after displaying, need to reestablish the reports for multi-reports, and only 1 of them selected,
                                            'If not reestablishing the reports, error occurs
        mReOpenReports
    End If
    Set Report = Nothing   'Remove data segment
    gUserActivityLog "E", sgReportListName & ": Display"
    'Dan 8/26/11 run this by Dick
    Unload Me
End Sub

'3-1-17 restore the dictionary .  For some reason, when only 1 report of multi-reports are displayed, then Printed from viewer, the reports (dictionary)
'gets set to something; whereby an error is created in the following exports.  this happens only with invoices since all other reports do not execute any
'other exports after printing.
Private Sub mReOpenReports()
   ' Dim vlArray As Variant
    Dim vlArrayKeys As Variant
    Dim ilCounter As Integer
    Dim myTempReport As CRAXDRT.Report
    Dim ilRet As Integer
    Dim slName As String
    Dim ilTotal As Integer
    
    ilTotal = ogReport.Reports.Count
    If ilTotal = 1 Then
      '  vlArray = ogReport.Reports.Items
        vlArrayKeys = ogReport.Reports.Keys
        ogReport.Reports.RemoveAll
        For ilCounter = 0 To ilTotal - 1
            'Set myTempReport = vlArray(ilCounter)
            slName = vlArrayKeys(ilCounter)
            ilRet = ogReport.OpenReport(slName, False)
            ogReport.DiscardSavedData = True
        Next ilCounter
    End If
End Sub


