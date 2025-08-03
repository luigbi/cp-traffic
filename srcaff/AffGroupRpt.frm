VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmGroupRpt 
   Caption         =   "Affiliate Group Report"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffGroupRpt.frx":0000
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
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4590
      TabIndex        =   8
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
      Left            =   165
      TabIndex        =   5
      Top             =   1740
      Width           =   6705
      Begin VB.Frame FraSortBy 
         Caption         =   "Select by"
         Height          =   2160
         Left            =   150
         TabIndex        =   9
         Top             =   315
         Width           =   2835
         Begin VB.OptionButton optSortBy 
            Caption         =   "Owner"
            Height          =   255
            HelpContextID   =   1
            Index           =   6
            Left            =   135
            TabIndex        =   16
            Top             =   1020
            Width           =   1080
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Vehicle"
            Height          =   255
            HelpContextID   =   1
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   1785
            Width           =   1380
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Time Zone"
            Height          =   255
            HelpContextID   =   1
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   1530
            Width           =   1185
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "DMA Market"
            Height          =   255
            HelpContextID   =   1
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   255
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Format"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   510
            Width           =   1140
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "MSA Market"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   765
            Width           =   1230
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "State"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   1275
            Width           =   1155
         End
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4590
      TabIndex        =   6
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
         ItemData        =   "AffGroupRpt.frx":08CA
         Left            =   1050
         List            =   "AffGroupRpt.frx":08CC
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1290
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmGroupRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub mRunReportCode()

    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String

    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex    'get the export type selected
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    gUserActivityLog "S", sgReportListName & ": Prepass"
    If optSortby(1).Value = True Then
       slRptName = "AfDMAMkt"               'DMA Market
       SQLQuery = "Select mktName, mktGroupName from mkt"
    ElseIf optSortby(2).Value = True Then
        slRptName = "AfFormat"                              'Format
        SQLQuery = "Select fmtName, fmtGroupName from fmt_Station_Format"
    ElseIf optSortby(3).Value = True Then
        slRptName = "AfMSAMkt"                'MSA Market
        SQLQuery = "Select metName, metGroupName from met"
    ElseIf optSortby(4).Value = True Then
        slRptName = "AfState"                               'State
        SQLQuery = "Select sntName, sntPostalName, sntGroupName from snt"
    ElseIf optSortby(5).Value = True Then
        slRptName = "AfTimeZone"                    'Time Zone
        SQLQuery = "Select tztName, tztGroupName from tzt"
    ElseIf optSortby(0).Value = True Then
        slRptName = "AfVehicle"                             'vehicle
        SQLQuery = "Select vpfWegenerExport, vpfOLAExport, vefName, vefState, VffGroupName, VffWegenerExportID, VffOLAExportID from VPF_Vehicle_Options "
        SQLQuery = SQLQuery & "INNER JOIN VEF_VEhicles on vpfvefKCode = vefcode INNER JOIN VFF_Vehicle_FEatures on vefcode = vffvefcode "
        SQLQuery = SQLQuery & "where (vpfWegenerExport = 'Y' or vpfOLAExport = 'Y') and vefState = 'A'"
    ElseIf optSortby(6).Value = True Then
        slRptName = "AfOwner"                             'owner
        SQLQuery = "Select arttLastName, arttState from artt where arttType = 'O'"
    End If

    
    slExportName = Trim$(slRptName)
    slRptName = Trim$(slRptName) & ".rpt"
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    Screen.MousePointer = vbDefault

    
End Sub

Private Sub cmdReport_Click()
    mRunReportCode
End Sub

Private Sub cmdDone_Click()
    
    Unload frmGroupRpt


End Sub


Private Sub cmdReturn_Click()

    frmReports.Show
    Unload frmGroupRpt
    
End Sub
Private Sub Form_Load()

    frmGroupRpt.Caption = "Group Report - " & sgClientName

    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = True

End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmGroupRpt
    gCenterForm frmGroupRpt

End Sub

Private Sub Form_Initialize()

    mInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmGroupRpt = Nothing

End Sub

