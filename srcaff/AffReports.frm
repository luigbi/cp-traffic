VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmReports 
   Caption         =   "Report Selection"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AffReports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   7935
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ListBox lbcFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "AffReports.frx":08CA
      Left            =   120
      List            =   "AffReports.frx":08D1
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame frcGen 
      Caption         =   "Insert last-used report settings"
      Height          =   720
      Left            =   765
      TabIndex        =   4
      Top             =   6090
      Visible         =   0   'False
      Width           =   4680
      Begin VB.CommandButton cmdContinue 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3330
         TabIndex        =   7
         Top             =   285
         Width           =   1035
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "Yes-Except Dates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   6
         Top             =   285
         Width           =   1710
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "Yes"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   285
         Width           =   1035
      End
   End
   Begin VB.Timer tmcSetCntrls 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7530
      Top             =   5790
   End
   Begin VB.HScrollBar hbcRptSample 
      Height          =   240
      LargeChange     =   8490
      Left            =   120
      SmallChange     =   8490
      TabIndex        =   11
      Top             =   5775
      Visible         =   0   'False
      Width           =   6945
   End
   Begin VB.VScrollBar vbcRptSample 
      Height          =   2730
      LargeChange     =   1455
      Left            =   7070
      SmallChange     =   1455
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pbcRptSample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   2730
      Index           =   0
      Left            =   120
      ScaleHeight     =   2700
      ScaleWidth      =   6930
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   6960
      Begin VB.PictureBox pbcRptSample 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   165
         Index           =   1
         Left            =   15
         ScaleHeight     =   165
         ScaleWidth      =   6900
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   -15
         Visible         =   0   'False
         Width           =   6900
      End
   End
   Begin VB.ListBox lbcReports 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "AffReports.frx":08DE
      Left            =   120
      List            =   "AffReports.frx":08E0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtRepDesc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7440
      Top             =   4920
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6855
      FormDesignWidth =   7935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5715
      TabIndex        =   9
      Top             =   6330
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmReports - directory form for selecting reports to generate
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
'Const USEROPT_RPT = 20
'Const SITEOPT_RPT = 21
Private Sub cmdCancel_Click()
    Unload frmReports
End Sub

Private Sub cmdContinue_Click(Index As Integer)
    Dim iIndex As Integer
        
    If lbcReports.ListIndex < 0 Then
        Exit Sub
    End If
    
    '2/23/18: avoid error in report when it tries to resize itself
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    End If
    iIndex = lbcReports.ItemData(lbcReports.ListIndex)
    sgReportListName = Trim$(lbcReports.List(lbcReports.ListIndex))
    frmCrystal.igReportRnfCode = iIndex    'tmReportList(lbcRpt.ListIndex).tRnf.iCode
    
    Select Case iIndex
        Case 0
            frmCrystal.sgReportFormExe = "StationRpt"
            Set frmCrystal.fgReportForm = frmStationRpt
            frmStationRpt.Show
        Case 1
            frmCrystal.sgReportFormExe = "VehAffRpt"
            Set frmCrystal.fgReportForm = frmVehAffRpt
            frmVehAffRpt.Show
        Case 2
            frmCrystal.sgReportFormExe = "DelqRpt"
            Set frmCrystal.fgReportForm = frmDelqRpt
            frmDelqRpt.Show
        Case 3
            frmCrystal.sgReportFormExe = "LabelRpt"
            Set frmCrystal.fgReportForm = frmLabelRpt
            frmLabelRpt.Show
        Case 4
            frmCrystal.sgReportFormExe = "ClearRpt"
            Set frmCrystal.fgReportForm = frmClearRpt
            frmClearRpt.Show
        Case 5
            frmCrystal.sgReportFormExe = "PledgeRpt"
            Set frmCrystal.fgReportForm = frmPledgeRpt
            frmPledgeRpt.Show
        Case 6
            frmCrystal.sgReportFormExe = "AiredRpt"
            Set frmCrystal.fgReportForm = frmAiredRpt
            frmAiredRpt.Show
        Case 7
            frmCrystal.sgReportFormExe = "PostActivityRpt"
            Set frmCrystal.fgReportForm = frmPostActivityRpt
            frmPostActivityRpt.Show
        Case 8
            frmCrystal.sgReportFormExe = "LogActivityRpt"
            Set frmCrystal.fgReportForm = frmLogActivityRpt
            frmLogActivityRpt.Show
        Case 9
            frmCrystal.sgReportFormExe = "LogInactivityRpt"
            Set frmCrystal.fgReportForm = frmLogInactivityRpt
            frmLogInactivityRpt.Show
        Case 10
            frmCrystal.sgReportFormExe = "AlertRpt"
            Set frmCrystal.fgReportForm = frmAlertRpt
            frmAlertRpt.Show
        Case 11
            frmCrystal.sgReportFormExe = "AffiliateRpt"
            Set frmCrystal.fgReportForm = frmAffiliateRpt
            frmAffiliateRpt.Show
        Case 12
            frmCrystal.sgReportFormExe = "PgmClrRpt"
            Set frmCrystal.fgReportForm = frmPgmClrRpt
            frmPgmClrRpt.Show
        Case 13
            frmCrystal.sgReportFormExe = "PldgAirRpt"
            Set frmCrystal.fgReportForm = frmPldgAirRpt
            frmPldgAirRpt.Show
        Case 14
            frmCrystal.sgReportFormExe = "PldgAirRpt"
            Set frmCrystal.fgReportForm = frmPldgAirRpt
            frmPldgAirRpt.Show
        Case 15
            frmCrystal.sgReportFormExe = "VerifyRpt"
            Set frmCrystal.fgReportForm = frmVerifyRpt
            frmVerifyRpt.Show
        Case 16
            frmCrystal.sgReportFormExe = "JournalRpt"
            Set frmCrystal.fgReportForm = frmJournalRpt
            frmJournalRpt.Show
        Case 17
            frmCrystal.sgReportFormExe = "ExpMonRpt"
            Set frmCrystal.fgReportForm = frmExpMonRpt
            frmExpMonRpt.Show
        Case 18
            frmCrystal.sgReportFormExe = "MarkAssignRpt"
            Set frmCrystal.fgReportForm = frmMarkAssignRpt
            frmMarkAssignRpt.Show
        Case NCR_RPT
            frmCrystal.sgReportFormExe = "DelqRpt"
            Set frmCrystal.fgReportForm = frmDelqRpt
            frmDelqRpt.Show
        Case USEROPT_RPT
            frmCrystal.sgReportFormExe = "UserOptionsRpt"
            Set frmCrystal.fgReportForm = frmUserOptionsRpt
            frmUserOptionsRpt.Show
        Case SITEOPT_RPT
            frmCrystal.sgReportFormExe = "RptNoSel"
            Set frmCrystal.fgReportForm = frmRptNoSel
            frmRptNoSel.Show
        Case AFFSMISSINGWKS_RPT
            frmCrystal.sgReportFormExe = "DelqRpt"
            Set frmCrystal.fgReportForm = frmDelqRpt
            frmDelqRpt.Show
        Case REGIONASSIGN_RPT, REGIONASSIGNTRACE_RPT
            frmCrystal.sgReportFormExe = "RgAssignRpt"
            Set frmCrystal.fgReportForm = frmRgAssignRpt
            frmRgAssignRpt.Show
        Case GROUP_RPT
            frmCrystal.sgReportFormExe = "GroupRpt"
            Set frmCrystal.fgReportForm = frmGroupRpt
            frmGroupRpt.Show
        Case ADVFULFILL_RPT
            frmCrystal.sgReportFormExe = "AdvFulFillRpt"
            Set frmCrystal.fgReportForm = frmAdvFulFillRpt
            frmAdvFulFillRpt.Show
        Case CONTACTCOMMENTS_RPT
            frmCrystal.sgReportFormExe = "CommentRpt"
            Set frmCrystal.fgReportForm = frmCommentRpt
            frmCommentRpt.Show
        Case WEBLOGIMPORT_RPT
            frmCrystal.sgReportFormExe = "WebLogImportRpt"
            Set frmCrystal.fgReportForm = frmWebLogImportRpt
            frmWebLogImportRpt.Show
        Case LOGDELIVERY_RPT           '2-21-12
            frmCrystal.sgReportFormExe = "LogDeliveryRpt"
            Set frmCrystal.fgReportForm = frmLogDeliveryRpt
            frmLogDeliveryRpt.Show
        Case SPOTMGMT_Rpt               '3-9-12
            frmCrystal.sgReportFormExe = "SpotMgmtRpt"
            Set frmCrystal.fgReportForm = FrmSpotMgmtRpt
            FrmSpotMgmtRpt.Show
        Case EXPHISTORY_Rpt             '6-7-12
            frmCrystal.sgReportFormExe = "ExpHistoryRpt"
            Set frmCrystal.fgReportForm = frmExpHistoryRpt
            frmExpHistoryRpt.Show
        Case SPORTDECLARE_Rpt, SPORTCLEARANCE_Rpt           '10-9-12, 10-16-12
            frmCrystal.sgReportFormExe = "SportDeclareRpt"
            Set frmCrystal.fgReportForm = frmSportDeclareRpt
            frmSportDeclareRpt.Show
        Case RENEWALSTATUS_Rpt                              '11-8-12
            frmCrystal.sgReportFormExe = "RenewalRpt"
            Set frmCrystal.fgReportForm = frmRenewalRpt
            frmRenewalRpt.Show
        Case ADVCOMPLY_Rpt                                  '2-25-13
            frmCrystal.sgReportFormExe = "AdvComplyRpt"
            Set frmCrystal.fgReportForm = frmAdvComplyRpt
            frmAdvComplyRpt.Show
        Case RADARCLEAR_Rpt
            frmCrystal.sgReportFormExe = "RadarClrRpt"
            Set frmCrystal.fgReportForm = frmRadarClrRpt
            frmRadarClrRpt.Show
        Case ADVPLACEMENT_Rpt                   '2-10-14
            frmCrystal.sgReportFormExe = "AdvPlaceRpt"
            Set frmCrystal.fgReportForm = frmAdvPlaceRpt
            frmAdvPlaceRpt.Show
        Case MEASUREMENT_Rpt                    '12-17-14
            frmCrystal.sgReportFormExe = "MeasureRpt"
            Set frmCrystal.fgReportForm = frmMeasureRpt
            frmMeasureRpt.Show
        Case VEHICLE_VISUAL_RPT
            frmCrystal.sgReportFormExe = "VehVisualRpt"
            Set frmCrystal.fgReportForm = frmVehVisualRpt
            frmVehVisualRpt.Show
        Case WEB_VENDOR_RPT                     '2-9-17
            frmCrystal.sgReportFormExe = "WebVendorRpt"
            Set frmCrystal.fgReportForm = frmWebVendorRpt
            frmWebVendorRpt.Show
        Case DELIVERY_DETAIL_RPT                '5-22-18
            frmCrystal.sgReportFormExe = "DeliveryDetailRpt"
            Set frmCrystal.fgReportForm = frmDeliveryDetailRpt
            frmDeliveryDetailRpt.Show
        Case STATION_PERSONNEL_RPT              '8-22-18    FYM
            frmCrystal.sgReportFormExe = "StationPersonRpt"
            Set frmCrystal.fgReportForm = frmStationPersonRpt
            frmStationPersonRpt.Show
        Case AGREE_CLUSTER_RPT              '3-27-20  Affiliate agreement cluster report
            frmCrystal.sgReportFormExe = "ClusterRpt"
            Set frmCrystal.fgReportForm = frmClusterRpt
            frmClusterRpt.Show
        Case Else
            'frmReports.Show
    End Select
    'Unload frmReports
    frmCrystal.igReportButtonIndex = Index
    tmcSetCntrls.Enabled = True
End Sub

Private Sub Form_Activate()
    '7/7/21 - JW - prevents Aff and VB6 from Crashing when loading reports with that ReSize control -- having the controls all be invisible until form is done jerking around
    txtRepDesc.Height = lbcReports.Height
    'lbcReports.Visible = True
    txtRepDesc.Visible = True
    pbcRptSample(0).Visible = True
    pbcRptSample(1).Visible = True
    vbcRptSample.Visible = True
    lbcReports.Visible = False
    lbcFind.Visible = True
    txtFind.Visible = True
    Label1.Visible = True
    frcGen.Visible = True
    cmdCancel.Visible = True
    txtRepDesc.Height = lbcFind.Height + lbcFind.Top - 120
End Sub

Private Sub Form_Initialize()
    'D.S. If the window's state is max or min then resizing will cause an error
    If frmReports.WindowState = 1 Or frmReports.WindowState = 2 Then
        Exit Sub
    End If

    Me.Width = Screen.Width / 1.55
    Me.Height = Screen.Height / 1.5
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmReports
    gCenterForm frmReports
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    
    'Me.Width = Screen.Width / 1.55
    'Me.Height = Screen.Height / 1.5
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    'Add report names to the list box
    frmReports.Caption = "Report Selection - " & sgClientName
    If bgRemoteExport Then
        lbcReports.AddItem "Station Information"
        lbcReports.ItemData(lbcReports.NewIndex) = 0
        lbcReports.AddItem "User Options"
        lbcReports.ItemData(lbcReports.NewIndex) = USEROPT_RPT
        lbcReports.AddItem "Site Options"
        lbcReports.ItemData(lbcReports.NewIndex) = SITEOPT_RPT
    Else
        lbcReports.AddItem "Station Information"
        lbcReports.ItemData(lbcReports.NewIndex) = 0
        lbcReports.AddItem "Affiliate Agreements"
        lbcReports.ItemData(lbcReports.NewIndex) = 1
        lbcReports.AddItem "Overdue Affidavits"
        lbcReports.ItemData(lbcReports.NewIndex) = 2
        lbcReports.AddItem "Mailing Labels"
        lbcReports.ItemData(lbcReports.NewIndex) = 3
        lbcReports.AddItem "Advertiser Clearances"
        lbcReports.ItemData(lbcReports.NewIndex) = 4
        lbcReports.AddItem "Pledges"
        lbcReports.ItemData(lbcReports.NewIndex) = 5
        lbcReports.AddItem "Spot Clearance"
        lbcReports.ItemData(lbcReports.NewIndex) = 6
        'lbcReports.AddItem "Web Posting Activity"
        lbcReports.AddItem "Affiliate Affidavit Posting Activity"       'name change 6-28-12
        lbcReports.ItemData(lbcReports.NewIndex) = 7
        lbcReports.AddItem "Web Log Activity"
        lbcReports.ItemData(lbcReports.NewIndex) = 8
        lbcReports.AddItem "Web Log Inactivity"
        lbcReports.ItemData(lbcReports.NewIndex) = 9
        lbcReports.AddItem "Alert Status"
        lbcReports.ItemData(lbcReports.NewIndex) = 10
        lbcReports.AddItem "Affiliate Clearance Counts"
        lbcReports.ItemData(lbcReports.NewIndex) = 11
        lbcReports.AddItem "Program Clearance"              '10-24-05
        lbcReports.ItemData(lbcReports.NewIndex) = 12
        lbcReports.AddItem "Pledged vs Aired Clearance"
        lbcReports.ItemData(lbcReports.NewIndex) = PLEDGEVSAIR_RPT       '6-23-06
        lbcReports.AddItem "Fed vs Aired Clearance"
        lbcReports.ItemData(lbcReports.NewIndex) = FEDVSAIR_RPT       '07-10-06
        lbcReports.AddItem "Feed Verification"
        lbcReports.ItemData(lbcReports.NewIndex) = VERIFY_RPT   '8-18-08
        lbcReports.AddItem "Export Journal"
        lbcReports.ItemData(lbcReports.NewIndex) = EXPJOURNAL_RPT   '4-25-07
        lbcReports.AddItem "Export Monitoring"
        lbcReports.ItemData(lbcReports.NewIndex) = 17   '8-25-07
        '2-23-11 hide this report as well as the the Assignment import (was for USRN & not used)
        'lbcReports.AddItem "DMA Market Assignment"
        'lbcReports.ItemData(lbcReports.NewIndex) = 18   '8-25-07
        lbcReports.AddItem "Critically Overdue Report"      '10-5-14 report name change "Non-Compliant(NCR)"
        lbcReports.ItemData(lbcReports.NewIndex) = NCR_RPT   'index = 19 6-30-09
        'Dan M site options and user options reports added 1/4/09
        lbcReports.AddItem "User Options"
        lbcReports.ItemData(lbcReports.NewIndex) = USEROPT_RPT
        lbcReports.AddItem "Site Options"
        lbcReports.ItemData(lbcReports.NewIndex) = SITEOPT_RPT
        
        lbcReports.AddItem "Affiliates Missing Weeks"
        lbcReports.ItemData(lbcReports.NewIndex) = AFFSMISSINGWKS_RPT       '1-12-10
        
        lbcReports.AddItem "Regional Affiliate Copy Assignment"
        lbcReports.ItemData(lbcReports.NewIndex) = REGIONASSIGN_RPT        'index 23 as of 1-19-10
        lbcReports.AddItem "Regional Affiliate Copy Tracing"
        lbcReports.ItemData(lbcReports.NewIndex) = REGIONASSIGNTRACE_RPT        'index 23 as of 1-19-10
        lbcReports.AddItem "Groups"           '4-8-10
        lbcReports.ItemData(lbcReports.NewIndex) = GROUP_RPT
        lbcReports.AddItem "Advertiser Fulfillment"
        lbcReports.ItemData(lbcReports.NewIndex) = ADVFULFILL_RPT            '7-14-10
        lbcReports.AddItem "Contact Comments"
        lbcReports.ItemData(lbcReports.NewIndex) = CONTACTCOMMENTS_RPT         '12-20-10
        lbcReports.AddItem "Web Import Log"
        lbcReports.ItemData(lbcReports.NewIndex) = WEBLOGIMPORT_RPT             '12-15-11
        lbcReports.AddItem "Affiliate Delivery Summary"
        lbcReports.ItemData(lbcReports.NewIndex) = LOGDELIVERY_RPT             '2-21-12
        lbcReports.AddItem "Affiliate Spot Management"
        lbcReports.ItemData(lbcReports.NewIndex) = SPOTMGMT_Rpt             '3-9-12
        lbcReports.AddItem "Export History"
        lbcReports.ItemData(lbcReports.NewIndex) = EXPHISTORY_Rpt             '6-7-12
        lbcReports.AddItem "Station Sports Declaration"
        lbcReports.ItemData(lbcReports.NewIndex) = SPORTDECLARE_Rpt             '10-9-12
        lbcReports.AddItem "Sports Clearance"
        lbcReports.ItemData(lbcReports.NewIndex) = SPORTCLEARANCE_Rpt             '10-16-12
        lbcReports.AddItem "Agreement Renewal Status"
        lbcReports.ItemData(lbcReports.NewIndex) = RENEWALSTATUS_Rpt             '11-8-12
        lbcReports.AddItem "Advertiser Compliance"
        lbcReports.ItemData(lbcReports.NewIndex) = ADVCOMPLY_Rpt                 '2-25-13
        lbcReports.AddItem "Radar Clearance"
        lbcReports.ItemData(lbcReports.NewIndex) = RADARCLEAR_Rpt                 '9-24-13
        lbcReports.AddItem "Advertiser Placement"
        lbcReports.ItemData(lbcReports.NewIndex) = ADVPLACEMENT_Rpt                 '2-10-14
        lbcReports.AddItem "Affiliate Measurement"
        lbcReports.ItemData(lbcReports.NewIndex) = MEASUREMENT_Rpt                  '12-17-14
        lbcReports.AddItem "Vehicle Visual Summary"
        lbcReports.ItemData(lbcReports.NewIndex) = VEHICLE_VISUAL_RPT                '5-23-16
        lbcReports.AddItem "Web Vendor Export/Import"
        lbcReports.ItemData(lbcReports.NewIndex) = WEB_VENDOR_RPT
        lbcReports.AddItem "Affiliate Delivery Detail"                              '5-22-18
        lbcReports.ItemData(lbcReports.NewIndex) = DELIVERY_DETAIL_RPT
        lbcReports.AddItem "Station Personnel Detail"                           '8-22-18    FYM
        lbcReports.ItemData(lbcReports.NewIndex) = STATION_PERSONNEL_RPT
        lbcReports.AddItem "Affiliate Cluster Agreements"                           '3-27-20
        lbcReports.ItemData(lbcReports.NewIndex) = AGREE_CLUSTER_RPT
        'grdReports.AddItem "Dunning Letter"
        'grdReports.AddItem "Current Affiliates"
        'grdReports.AddItem "Certificates of Performance Past Due"
    End If
    
    lbcFind.Clear
    For ilLoop = 0 To lbcReports.ListCount - 1
        lbcFind.AddItem lbcReports.List(ilLoop)
    Next ilLoop
End Sub

Private Sub Form_Resize()
    txtRepDesc.Height = lbcFind.Height + lbcFind.Top - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReports = Nothing
End Sub

Private Sub hbcRptSample_Change()
    pbcRptSample(1).Left = -hbcRptSample.Value
End Sub

Private Sub lbcFind_Click()
    Dim ilLoop As Integer
    If lbcFind.Text = "" Then Exit Sub
    For ilLoop = 0 To lbcReports.ListCount - 1
        If lbcReports.List(ilLoop) = lbcFind.Text Then
            lbcReports.ListIndex = ilLoop
        End If
    Next ilLoop
End Sub

Private Sub lbcReports_Click()
    Dim iIndex As Integer
    Dim slPicture As String
    
    If lbcReports.ListIndex < 0 Then
        txtRepDesc.Text = ""
        Exit Sub
    End If
    slPicture = ""
    iIndex = lbcReports.ItemData(lbcReports.ListIndex)
    If iIndex = 0 Then      'Station INformation
        txtRepDesc.Text = "Lists stations in the system, including address, contact information, personnel, and web passwords (if selected). Station information can be sorted alphabetically by call letters, by owner, by DMA, by MSA, or by multicast. Selectivity options exist to show the affidavit contact, specific job titles, or all contacts. This report can also be run by person, as defined on the Station Information screen."
        slPicture = "AFFStation_Information_Owner.jpg"
    ElseIf iIndex = 1 Then      'Agreements
        slPicture = "AFFAffiliate_Agreements.jpg"
        txtRepDesc.Text = "Produces a list of the stations and vehicles they are affiliated with. Date options include active, entered, starting or ending agreement dates. Agreements by ending dates can be used for renewal calls. Agreements by entered dates can be used to proof data input. Additional station contact personnel and mailing information can be included for the affiliate sales or clearance department."
    ElseIf iIndex = 2 Then      'Overdue CPs
        txtRepDesc.Text = "Shows all stations that are delinquent in returning their affidavits, and lists their delinquent weeks. The affidavit contact and phone number are shown to assist you in affidavit retrieval. This report can be run by vehicle or station. Sort options include by vehicle, station, DMA Market Rank, Market Rep, or Producer."
        slPicture = "AFFOverdue_CP.jpg"
    ElseIf iIndex = 3 Then      'Labels
        txtRepDesc.Text = "CD Mailing Labels"
        slPicture = "AFFMailing_Labels.jpg"
        txtRepDesc.Text = "Generates mailing labels in three available label format options. The formats are: two column Avery 5163 (2 x 4) or (1 x 4) or three column (1 x 2 5/8). The labels are sorted alphabetically by Label ID as defined for the agreement."
    ElseIf iIndex = 4 Then      'advertiser clearance
        slPicture = "AffAdv_Clearances.jpg"
        'txtRepDesc.Text = "Times that each Affiliate carried an advertiser's spots"
        txtRepDesc.Text = "For an advertiser and contract, all spots aired in a given week are shown, on a station by station basis, with the stations listed alphabetically, or by DMA or MSA market name or rank. The spot times shown are those times supplied by the stations on their Certificates of Performance. Spots with times that have not yet been reported can be flagged, along with spots not aired, not carried, and spots aired outside of the pledged times."
    ElseIf iIndex = 5 Then      'Pledges
        txtRepDesc.Text = "Shows that portion of the affiliate agreement that specifies which spots will be aired, delayed or not aired per the station contract. This report can be run to show non-live pledges, missing pledges, or all pledges."
        slPicture = "AFFPledge_Report.jpg"
    ElseIf iIndex = 6 Then      'Spot Clearance
        txtRepDesc.Text = "Provides detailed information about when and where a given spot aired, with the aired date, time, copy, product, contract number, and length for each spot. This report can be generated for all or selective vehicles and affiliates, and by contract number. Selectivity options allow to you choose to include or exclude spots not aired, those not carried, or for stations that have not yet returned their affidavits. This report also has the option of printing out the certification statement that affiliates must agree to prior to posting an affidavit when using the Counterpoint Affidavit System. This report can be sorted by vehicle, advertiser, or DMA Market."
        slPicture = "AFFSpot_Clearance_Advertiser.jpg"
    ElseIf iIndex = 7 Then          'Affdavit posting activity
        txtRepDesc.Text = "Shows spots posted via the Counterpoint Affidavit system and station spots updated via Marketron Network Connect. The combination of the user name and IP address allows this report to serve as an electronic signature for reporting purposes. This report can be sorted by vehicle or station for all or selected vehicles and stations."
        slPicture = "AffAffPost_Activity.jpg"
    ElseIf iIndex = 8 Then
        txtRepDesc.Text = "Tracks the amount of time it takes for a station to view and/or download their logs from the Counterpoint Affidavit system. The report compares the export date to the date the station took action to calculate lag time. It can be generated for any selected log date span, and can be sorted by station or by vehicle."
        slPicture = "AFFWeb_Log_Activity.jpg"
    ElseIf iIndex = 9 Then
        txtRepDesc.Text = "Lists stations that have not printed their logs from the Counterpoint Affidavit system. Selectivity is for any log date span, with sorting by station or vehicle."
        slPicture = "AFFWeb_Log_Inactivity.jpg"
    ElseIf iIndex = 10 Then
        txtRepDesc.Text = "Provides a list of all alerts by user or by date from the Traffic and Affiliate Systems, including contract, log, and export alerts. The Alert Status report also includes the user that generated the alert, whether the alert is pending or has been cleared, alert descriptions, and the date and time of the alert."
        slPicture = "AFFAlert_Date.jpg"
    ElseIf iIndex = 11 Then         'advertiser clearance counts
        slPicture = "AffClearance_Counts.JPG"
        txtRepDesc.Text = "Compares the number of spots each affiliate was supposed to air to what they did air, to arrive at a percentage of spots aired. The report also totals those spots not yet reported, and any missed spots, for a specified date range."
    ElseIf iIndex = 12 Then         'program clearnace
        slPicture = "AFFProgram_Clearance_Minutes.jpg"
        txtRepDesc.Text = "This report uses named avails and shows the minutes or units aired for each one.  The total aired, not reported and not aired counts are listed.  Selectivity includes start and end dates and times as well as the spot statuses."
    ElseIf iIndex = EXPJOURNAL_RPT Then
        slPicture = "AFFExport_Journal.jpg"
        txtRepDesc.Text = "Used to verify all spots were successfully exported to the web. It can be generated for any desired starting log date in combination with any desired amount of days. An activity date range may be entered as well. The report can be generated to show all export activity, or just that activity that resulted in discrepancies. Individual stations can be broken out or suppressed."
    ElseIf iIndex = NCR_RPT Then
        slPicture = "AFFNon_Compliance_DMA.jpg"
        txtRepDesc.Text = "A list of stations that have not submitted affidavits. This is used to manage affidavit submissions.  Stations that appear are stations that have not submitted affidavits for a pre-defined consecutive number of weeks.  Any stations that appear retain the status ""Critically Overdue"" until the Affiliate Compliance Department changes it in the agreement.  Selectivity is by vehicle or station.  Sort options include vehicle, station, DMA Market Rank, Affiliate A/E or Producer."
        txtRepDesc.Text = txtRepDesc.Text & " When the report is created any stations that appear on the report are activated as non-compliant.  They retain this status until the option in the agreement is changed by the Affiliate Compliance Department."
    ElseIf iIndex = FEDVSAIR_RPT Then
        slPicture = "AFFFed_Vs_Aired_Advertiser.jpg"
        'txtRepDesc.Text = "This report tracks affiliate compliance.  It lists feed times from the network and station aired times.  A discrepancy only option compares and lists any spots outside the feed days and times. If you choose discrepancy only, it compares and lists any spots outside the feed date and time. Status Discrepancy will display any spots whose internal spot status is different from its internal pledged status. "
        'txtRepDesc.Text = txtRepDesc.Text & "If you show spot Status Codes, Network Non-Compliant or Station Non-Compliant spots are denoted with an N(etwork) or S(tation) in the far right column. Sorting options are by vehicle or advertiser and fed or air dates and times, and can show one avail type or all avails."
        txtRepDesc.Text = "This report tracks and displays affiliate compliance.  It displays feed times, and the aired times that were captured during the import process. A discrepancy only option lists any spots outside the feed days and times. Status Discrepancy will display any spots whose spot status is different from its pledged status. "
        txtRepDesc.Text = txtRepDesc.Text & "If you show spot Status Codes, Network Non-Compliant or Station Non-Compliant spots are denoted with an N(etwork) or S(tation) in the far right column. Sorting options are by vehicle or advertiser and fed or air dates and times, and can show one avail type or all avails."
    ElseIf iIndex = PLEDGEVSAIR_RPT Then
        slPicture = "AFFPledged_vs_Clearance.jpg"
        'txtRepDesc.Text = "This report shows the affiliate pledged times and day vs. the affiliate posted air times and dates for the selected date and times.  Status codes can be selected to display and separate for reporting.  The discrepancy option compares the pledged date and time vs. the aired date and time and shows spots that have aired outside of agreement pledge."
        txtRepDesc.Text = "This report lists the agreement pledged times and days to the aired times and dates for the selected time period that are captured at the time of the import. The report can be generated to show discrepancies only."
        txtRepDesc.Text = txtRepDesc.Text & "Showing spot status codes allows you to analyze how the spots were aired: live, delayed, not aired due to technical difficulties, or blackouts, and will also denote Network Non-Compliant and Station Non-Compliant spots with an N(etwork) or S(tation) in the far right column. "
        txtRepDesc.Text = txtRepDesc.Text & "If selected, Status Discrepancy will display any spots whose spot status is different from its pledged status."
    ElseIf iIndex = VERIFY_RPT Then
        slPicture = "AFFFeed_Verification.jpg"
        txtRepDesc.Text = "Tool to compare network programming to avails in the affiliate pledges for a selected span of weeks.  Network avails not in the pledge agreement are shown as Missing; avails in the agreement but not defined in the network programming are Extra."
    ElseIf iIndex = 17 Then             'export monitor
        slPicture = "AFFExport_Monitoring.jpg"
        txtRepDesc.Text = "A list of vehicles, affiliates and markets that were exported for a week."
    ElseIf iIndex = USEROPT_RPT Then
        slPicture = "AFFUser_Options.jpg"
        txtRepDesc.Text = "Lists user information and their security settings. Can be used to verify that your user security settings are correct."
    ElseIf iIndex = SITEOPT_RPT Then
        slPicture = "AFFSite_Options.jpg"
        txtRepDesc.Text = "Lists the system options being used, the email message templates, and the administrator information."
    ElseIf iIndex = AFFSMISSINGWKS_RPT Then
        'txtRepDesc.Text = "Lists the affiliates and the weeks that have not been posted for any vehicles they have agreements for. This report lists the station, the vehicle, the station contact, and the weeks reported as none aired. You can select the affiliates for all or specific dates."
        txtRepDesc.Text = "A listing of affiliates that submitted their affidavits as “NONE AIRED,” on the Counterpoint Affidavit System.  This report shows affiliates that are not delinquent because they have returned their affidavits, but they have not aired their spots. Only those affiliates that reported all of their spots as not having aired, while having an active agreement, will appear on this report."
        slPicture = "AFFMissing_Weeks.jpg"
    ElseIf iIndex = REGIONASSIGN_RPT Then
        txtRepDesc.Text = "Lists all regional copy that aired for the selected start and end date, sorted by vehicle or station. The selectivity options are to exclude spots with no regional copy, include not aired spots, or to show only the regional copy assigned to the spots. Specific contracts can be selected to show proof of regional copy assigned or aired."
        slPicture = "AFFRegional_Copy_Assignment_Vehicle.jpg"
    ElseIf iIndex = REGIONASSIGNTRACE_RPT Then
        slPicture = "AFFRegional_Copy_Tracing_Vehicle.jpg"
        txtRepDesc.Text = "A comprehensive list of all regional (and generic if selected) copy airing during a specified date range. The report is meant to be used as an internal tool to determine what regional copy is to air, and troubleshoot why a particular piece of regional copy did not air. This report can be run by vehicle or by station, with options to exclude spots lacking regional copy, include spots that did not air, or to show only the assigned regional copy."
    ElseIf iIndex = GROUP_RPT Then      'dump of group reports
        txtRepDesc.Text = "A comprehensive list of the defined DMA, MSA, State, Owner, Time Zone, Format and Vehicles and codes defined in the system."
        slPicture = "AFFGroup_Format.jpg"
    ElseIf iIndex = ADVFULFILL_RPT Then      'Spot report with regional copy
        slPicture = "Affiliate_Fulfillment.jpg"
        txtRepDesc.Text = "The Advertiser Fulfillment report gives your client a listing of where and when the affiliates aired their spots, showing the aired dates/times, along with the spot length and copy. Spot status codes can optionally be shown, as well as station counts, and stations not yet reported. "
        txtRepDesc.Text = txtRepDesc.Text & "If Status Codes are shown, Network and Station Non-Compliant Statuses will be denoted with an N(etwork) or an S(tation) in the far right column."
    ElseIf iIndex = CONTACTCOMMENTS_RPT Then
        txtRepDesc.Text = "Lists all comments associated with all or selective vehicles, stations, and personnel or departments. The report can be generated by a range of dates based on when comments were posted, and/or comment follow up dates."
        slPicture = "AFFContact_Comments.jpg"
    
    ElseIf iIndex = WEBLOGIMPORT_RPT Then
        txtRepDesc.Text = "Shows warnings or errors that occurred during the web import process, gathered from the web import log and reformatted into a user friendly format. If any errors are detected, spot information is printed that includes the vehicle, station, advertiser, product, date and time aired (if applicable), spot length, and ISCI. If there is no information to show for the dates selected, the report will be blank."
        slPicture = "AffWeb_Import_Log.jpg"
    ElseIf iIndex = LOGDELIVERY_RPT Then
        '9184
        txtRepDesc.Text = "Produces a list of vehicles that show the distribution of log exports for the Web, Stratus, Marketron, CBS, Clear Channel, Jelli, Wide Orbit or manual posting."
        slPicture = "AFFLog_Delivery_Vehicle.jpg"
    ElseIf iIndex = SPOTMGMT_Rpt Then
        txtRepDesc.Text = "Allows you to track unresolved missed spots and missed reasons, makegood spots, replacement spots, and bonus spots."
        slPicture = "AFFSpot_Management.jpg"
    ElseIf iIndex = EXPHISTORY_Rpt Then
        slPicture = "AFFExport_History.jpg"
        txtRepDesc.Text = "A seven day history of exports which details the user that initiated the export, the number of vehicles and date/time the export started and ended."
    ElseIf iIndex = SPORTDECLARE_Rpt Then       '10-9-12
        txtRepDesc.Text = "Report by station that lists all games in a schedule.  Station must declare which games they will carry, not carry or decide later."
        slPicture = "AFFStation_Sports_Declaration_Form.jpg"
    ElseIf iIndex = SPORTCLEARANCE_Rpt Then
        slPicture = "AFFSports_Clearance_Delinquent.jpg"
        txtRepDesc.Text = "Shows station clearance accountability for events aired vs. pledged for sports vehicles set to have pledges by event. Report selectivity is by date and time for delinquent clearances or by station potential clearances. Delinquent are those that have not been cleared or reported prior to today's date; potential lists all events in the agreement as an indicator of what the station's potential is in carrying all the events they have pledged to air."
    ElseIf iIndex = RENEWALSTATUS_Rpt Then         '11-8-12
        slPicture = "AFFAgreement_Renewal_Status.jpg"
        txtRepDesc.Text = "Report by vehicle or station indicating the agreements up for renewal.  User entered date span is tested against the agreement end or drop date."
    ElseIf iIndex = ADVCOMPLY_Rpt Then         '2-25-13
        slPicture = "Adv_Comp.jpg"
        'txtRepDesc.Text = "A spot clearance report by advertiser that assesses compliance using two different methods, either by pledge or by advertiser. Compliance by advertiser uses the ordered days, dates, and times to determine whether a spot is compliant or not."
        txtRepDesc.Text = "This report designates Network Non-Compliant spots with an N and Station Non-Compliant spots with an S in the far right Compliant Column. The report can be sorted by Advertiser, with an intermediate sort by station or vehicle, displaying one or multiple advertisers, vehicles, and stations, and can include/exclude non-reported stations. "
        txtRepDesc.Text = txtRepDesc.Text & "If you choose to Show Days/Dates/Times As Sold, you can select the Non-Compliant Only Network spots option, which will only show Network Non-Compliant spots and bypass all Network Compliant spot. Likewise, if you Show Days/Dates/Times As Pledged, only Station Non-Compliant spots will be displayed."
    ElseIf iIndex = RADARCLEAR_Rpt Then         '10-7-13
        slPicture = "AffRadar_Clearance.jpg"
        txtRepDesc.Text = "The Radar Clearance report indicates which minutes of inventory are parts of which Radar network for any given Radar vehicle, based on the Radar Table. The report can be generated for a specified date range, and can include all or selective Vehicles and Stations, Stations not yet reported, and/or Stations that did not air the spots."
    ElseIf iIndex = ADVPLACEMENT_Rpt Then         '2-10-14
        slPicture = "Adv_plc.jpg"
        txtRepDesc.Text = "This report offers a quick, detailed listing of aired information for a single advertiser, eliminating the need to alter the data after exporting. Aired dates/times (as opposed to feed dates/times), along with the spot length and copy are displayed. Regional copy can be highlighted, or spots without regional copy can be excluded. "
        txtRepDesc.Text = txtRepDesc & " Spot status codes can optionally be shown, as well as station counts, and stations not yet reported. If Status Codes are shown, Network Non-Compliant and Station Non-Compliant Statuses will also be denoted with an N(etwork) or S(tation) in the far right column."
    ElseIf iIndex = MEASUREMENT_Rpt Then            '12-17-14
        txtRepDesc.Text = "The Affiliate Measurement report is used to measure affiliate compliance and delinquent rates. The report displays measurement figures using data from when the Affiliate Measurement utility was last run."
        txtRepDesc.Text = txtRepDesc & "Information can be displayed by aired counts, or by year or aired percentages. All three variations show the number of Weeks Missing, Weeks Reported, Station Non-Compliant spots, Network Non-Compliant spots, and Responsiveness."
        slPicture = "Aff_Measurement.jpg"
    ElseIf iIndex = VEHICLE_VISUAL_RPT Then         '5-23-16
        txtRepDesc.Text = "The Vehicle Visual report is used to measure the value of each vehicle, by showing the amount of content and commercials, along with a count of the affiliates carrying it, and to list the vehicles in order of importance."
    ElseIf iIndex = WEB_VENDOR_RPT Then        '2-8-17
        txtRepDesc.Text = "The Web Vendor Export and Import reports show what was successfully sent to vendors and what was received back, including spot counts."
        slPicture = "AffWeb_Vendor_Export_Import.jpg"
    ElseIf iIndex = DELIVERY_DETAIL_RPT Then       '5-22-18
        txtRepDesc.Text = "Shows which delivery service(s) a station subscribes to for associated agreements.  A station can use 1 or more log and/or audio service, and will show a line for each different service vendor. Station call letters, time zone, vehicle name, agreement dates, market, format, delivery service type and vendor name is printed."
        slPicture = "AFFLog_Delivery_Detail.jpg"
    ElseIf iIndex = STATION_PERSONNEL_RPT Then      '8-22-18    FYM
        txtRepDesc.Text = "Print a report of station personnel information which includes station call Letters, DMA market, time zone, personnel name, title, phone #, and email address.  Also included are email, label, and export options. "
        txtRepDesc.Text = txtRepDesc.Text & "Report is sorted alphabetically by call letters and personnel name."
        slPicture = "AFFStation_Personnel.jpg"
    ElseIf iIndex = AGREE_CLUSTER_RPT Then      '3-27-20
        txtRepDesc.Text = "Lists agreements that are clustered together by one of more vehicles, then subsorted by one of the follow properties:  Station, Market, Format, Owner and Time Zone "
        slPicture = ""
        
    Else
        txtRepDesc.Text = ""
    'ElseIf grdReports.Row = 2 Then
    '    txtRepDesc.Text = "A List of all affiliated vehicles sorted by station."
    'ElseIf grdReports.Row = 3 Then
    '    txtRepDesc.Text = "Create a dunning letter"
    'ElseIf grdReports.Row = 4 Then
    '    txtRepDesc.Text = "A list of current affiliates"
    'ElseIf grdReports.Row = 5 Then
    '     txtRepDesc.Text = "A list of all Certificates of Performance that have not been received, by station."
    End If
    mSetRptSample slPicture

    igRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

End Sub

Private Sub lbcReports_DblClick()
    'cmdContinue_Click
    cmdContinue_Click 0
End Sub

Private Sub tmcSetCntrls_Timer()
    Dim slName As String
    Dim slStr As String
    
    tmcSetCntrls.Enabled = False
    '5/26/18: Reset the Form controls
    '5/24/18: Reset the Form controls
    slName = 10000 + frmCrystal.igReportRnfCode
    slStr = 10000 + igUstCode
    slName = slName & slStr
    frmCrystal.sgReportCtrlSaveName = Left$(frmCrystal.sgReportFormExe, 10) & slName
    If frmCrystal.igReportButtonIndex <> 2 Then
        frmCrystal.gSetReportCtrlsSetting
    End If
    Unload frmReports

End Sub

Private Sub txtFind_Change()
    Dim sName As String
    Dim lRow As Long
    Dim iLen As Integer
    Dim ilHightErr As Integer
    
    FilterList (txtFind.Text)
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyTab Or KeyCode = vbKeyReturn) Then
        txtFind.Text = lbcFind.Text
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lbcFind.ListIndex < lbcFind.ListCount - 1 Then
            lbcFind.ListIndex = lbcFind.ListIndex + 1
           KeyCode = 0
        End If
        If lbcFind.ListIndex = -1 Then
            If lbcFind.ListCount > 0 Then
                lbcFind.ListIndex = 0
            End If
        End If
    End If
    If KeyCode = vbKeyUp Then
        If lbcFind.ListIndex > 0 Then
            lbcFind.ListIndex = lbcFind.ListIndex - 1
        End If
        If lbcFind.ListIndex = -1 Then
            lbcFind.ListIndex = lbcFind.ListCount - 1
        End If
        KeyCode = 0
    End If
End Sub

Private Sub txtFind_LostFocus()
    'txtFind.Text = lbcFind.Text
End Sub

Private Sub vbcRptSample_Change()
   pbcRptSample(1).Top = -vbcRptSample.Value
End Sub
'
'           mSetRptSample = take the .gif report sample and place in picture box
'           <input> slPicture as string
'
Public Sub mSetRptSample(slPicture As String)

    On Error GoTo ErrHand
    If Trim$(slPicture) = "" Then
        pbcRptSample(1).Picture = LoadPicture()
        vbcRptSample.Max = vbcRptSample.Min
        vbcRptSample.Enabled = False
        hbcRptSample.Max = hbcRptSample.Min
        hbcRptSample.Enabled = False
    Else
        pbcRptSample(1).Picture = LoadPicture(sgReportDirectory & slPicture)
        vbcRptSample.Max = pbcRptSample(1).Height - pbcRptSample(0).Height
        vbcRptSample.Enabled = (pbcRptSample(0).Height < pbcRptSample(1).Height)
        If vbcRptSample.Enabled Then
            vbcRptSample.SmallChange = pbcRptSample(0).Height
            vbcRptSample.LargeChange = pbcRptSample(0).Height
        End If
        hbcRptSample.Max = pbcRptSample(1).Width - pbcRptSample(0).Width
        hbcRptSample.Enabled = (pbcRptSample(0).Width < pbcRptSample(1).Width)
        If hbcRptSample.Enabled Then
            hbcRptSample.SmallChange = pbcRptSample(0).Width
            hbcRptSample.LargeChange = pbcRptSample(0).Width
        End If
    End If
    Exit Sub
ErrHand:
    slPicture = ""
    pbcRptSample(1).Picture = LoadPicture()
End Sub

Private Sub FilterList(lsFilterText As String)
    lbcFind.Clear
    Dim ilLoop As Integer
    Dim ilTermsLoop As Integer
    Dim alTerms
    Dim blFound As Boolean
    blFound = False
    If Trim(lsFilterText) = "" Then
        ShowMasterList
        Exit Sub
    End If
    For ilLoop = 0 To lbcReports.ListCount - 1
        If InStr(1, lsFilterText, " ") = 0 Then
            'Single Term Search
            If InStr(1, LCase(lbcReports.List(ilLoop)), LCase(lsFilterText)) > 0 Then
                lbcFind.AddItem lbcReports.List(ilLoop)
                lbcFind.ItemData(lbcFind.NewIndex) = lbcReports.ItemData(ilLoop)
            End If
        Else
            'Multi Term Search
            alTerms = Split(lsFilterText, " ")
            For ilTermsLoop = 0 To UBound(alTerms)
                blFound = False
                If InStr(1, LCase(lbcReports.List(ilLoop)), LCase(alTerms(ilTermsLoop))) > 0 Then blFound = True
                If blFound = False Then Exit For
            Next ilTermsLoop
            If blFound = True Then
                lbcFind.AddItem lbcReports.List(ilLoop)
                lbcFind.ItemData(lbcFind.NewIndex) = lbcReports.ItemData(ilLoop)
            End If
        End If
    Next ilLoop
    ReselectFind
End Sub

Private Sub ShowMasterList()
    Dim ilLoop As Integer
    lbcFind.Clear
    For ilLoop = 0 To lbcReports.ListCount - 1
        lbcFind.AddItem lbcReports.List(ilLoop)
        lbcFind.ItemData(lbcFind.NewIndex) = lbcReports.ItemData(ilLoop)
    Next ilLoop
    ReselectFind
End Sub

Private Sub ReselectFind()
    Dim ilLoop As Integer
    If lbcReports.ListIndex = -1 Then Exit Sub
    For ilLoop = 0 To lbcFind.ListCount - 1
        If lbcFind.List(ilLoop) = lbcReports.Text Then
            lbcFind.ListIndex = ilLoop
        End If
    Next ilLoop
End Sub

