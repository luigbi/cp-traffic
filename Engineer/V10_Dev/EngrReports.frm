VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrReports 
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   7545
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EngrReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7545
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
      ItemData        =   "EngrReports.frx":030A
      Left            =   120
      List            =   "EngrReports.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
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
      Height          =   2415
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3960
      Top             =   4200
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4755
      FormDesignWidth =   7545
   End
   Begin VB.TextBox txtNotes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2640
      Width           =   7215
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
      Left            =   6000
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "EngrReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'*  EngrReports - directory form for selecting reports to generate
'*
'*  Created September 2004
'*
'*  Copyright Counterpoint Software, Inc.
'*******************************************************************
Option Explicit

Private Sub cmdCancel_Click()
    Unload EngrReports
End Sub

Private Sub cmdContinue_Click()
    Dim iIndex As Integer
    
    If lbcReports.ListIndex < 0 Then
        Exit Sub
    End If
    'iIndex = lbcReports.ItemData(lbcReports.ListIndex)
    igRptSource = vbModeless            'coming from rpt show the forms modeless;
                                        'coming from printer snapshot icon show forms modal
    Select Case igRptIndex
        Case USER_RPT, RELAY_RPT, MATTYPE_RPT, FOLLOW_RPT, SILENCE_RPT, TIMETYPE_RPT, AUDIONAME_RPT, AUDIOTYPE_RPT, AUDIOSOURCE_RPT, BUSGROUP_RPT, BUS_RPT
            EngrUserRpt.Show igRptSource
        Case NETCUE_RPT, CONTROL_RPT, COMMENT_RPT, EVENT_RPT, AUTOMATION_RPT
            EngrUserRpt.Show igRptSource
        Case SITE_RPT
            EngrSiteRpt.Show igRptSource
        Case ACTIVITY_RPT
            EngrActivityRpt.Show igRptSource
        Case LIBRARY_RPT
            EngrLibRpt.Show igRptSource
        Case LIBRARYEVENT_RPT, TEMPLATEAIR_RPT
            EngrLibEvtRpt.Show igRptSource
        Case AUDIOINUSE_RPT
            EngrSourceInUseRpt.Show igRptSource
        Case TEMPLATE_RPT
            EngrTemplateRpt.Show igRptSource
        Case TEMPLATEEVENT_RPT
            EngrTempEvtRpt.Show igRptSource
    End Select
    Unload EngrReports
End Sub

Private Sub Form_Activate()
    txtRepDesc.Height = lbcReports.Height
End Sub

Private Sub Form_Initialize()

    'D.S. If the window's state is max or min then resizing will cause an error
    If EngrReports.WindowState = 1 Or EngrReports.WindowState = 2 Then
        Exit Sub
    End If
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    
    gSetFonts EngrReports
    gCenterForm EngrReports

End Sub

Private Sub Form_Load()
Dim ilLoop As Integer

    'Add report names to the list box
    For ilLoop = LBound(tgReportNames) To UBound(tgReportNames) - 1
        lbcReports.AddItem Trim$(tgReportNames(ilLoop).sRptName)
        lbcReports.ItemData(lbcReports.NewIndex) = tgReportNames(ilLoop).iRptIndex
    Next ilLoop
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrReports = Nothing
End Sub
Private Sub lbcReports_Click()
    Dim iIndex As Integer
    Dim ilLoop As Integer
    If lbcReports.ListIndex < 0 Then
        txtRepDesc.text = ""
        Exit Sub
    End If
    igRptIndex = lbcReports.ItemData(lbcReports.ListIndex)
    'txtRepDesc.text = Trim$(tgReportNames(igRptIndex).sRptDesc)
    For ilLoop = LBound(tgReportNames) To UBound(tgReportNames) - 1
        If igRptIndex = tgReportNames(ilLoop).iRptIndex Then
            txtRepDesc.text = Trim$(tgReportNames(ilLoop).sRptDesc)
        End If
    Next ilLoop
   
End Sub

Private Sub lbcReports_DblClick()
    cmdContinue_Click
End Sub
