VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStationSearch 
   Caption         =   "Affiliate CRM"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   FillColor       =   &H00FFFFFF&
   Icon            =   "AffStationSearch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11580
   Visible         =   0   'False
   Begin VB.Frame frcCompliant 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   465
      Left            =   7335
      TabIndex        =   39
      Top             =   6480
      Width           =   1620
      Begin VB.OptionButton rbcCompliant 
         Caption         =   "Agency"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   42
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton rbcCompliant 
         Caption         =   "Station"
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   41
         Top             =   45
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.Label lacCompliant 
         Caption         =   "Compliant"
         Height          =   225
         Left            =   0
         TabIndex        =   40
         Top             =   30
         Width           =   750
      End
   End
   Begin VB.VScrollBar vbcStation 
      Height          =   720
      Left            =   10995
      TabIndex        =   38
      Top             =   3390
      Width           =   270
   End
   Begin VB.ListBox lbcUserOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   1
      ItemData        =   "AffStationSearch.frx":08CA
      Left            =   1620
      List            =   "AffStationSearch.frx":08CC
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcUserOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   0
      ItemData        =   "AffStationSearch.frx":08CE
      Left            =   870
      List            =   "AffStationSearch.frx":08D0
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5700
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcStationType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000005&
      Height          =   180
      Left            =   10080
      ScaleHeight     =   150
      ScaleWidth      =   810
      TabIndex        =   32
      Top             =   2115
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmcGen 
      Caption         =   "Generate"
      Height          =   375
      Left            =   10965
      TabIndex        =   35
      Top             =   60
      Width           =   885
   End
   Begin VB.ListBox lbcDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearch.frx":08D2
      Left            =   7560
      List            =   "AffStationSearch.frx":08D4
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearch.frx":08D6
      Left            =   8115
      List            =   "AffStationSearch.frx":08D8
      Sorted          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10500
      Picture         =   "AffStationSearch.frx":08DA
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   9555
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcAdvertiser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffStationSearch.frx":09D4
      Left            =   9135
      List            =   "AffStationSearch.frx":09D6
      Sorted          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   735
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   27
      Top             =   75
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   34
      Top             =   255
      Width           =   60
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   9585
      Top             =   5385
   End
   Begin VB.PictureBox pbcCommentType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000005&
      Height          =   180
      Left            =   10170
      ScaleHeight     =   150
      ScaleWidth      =   450
      TabIndex        =   15
      Top             =   2985
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.OptionButton rbcComments 
      Caption         =   "Follow-Up Only"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   5760
      TabIndex        =   25
      Top             =   6720
      Width           =   1455
   End
   Begin VB.OptionButton rbcComments 
      Caption         =   "All Comments"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   5760
      TabIndex        =   24
      Top             =   6495
      Value           =   -1  'True
      Width           =   1470
   End
   Begin VB.TextBox edcCommentTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   540
      HideSelection   =   0   'False
      Left            =   4590
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   3990
      Visible         =   0   'False
      Width           =   2280
   End
   Begin V81Affiliate.AffCommentGrid udcCommentGrid 
      Height          =   390
      Left            =   165
      TabIndex        =   21
      Top             =   1170
      Visible         =   0   'False
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   767
      CommentGridForm =   0
   End
   Begin VB.CommandButton cmcEMail 
      Caption         =   "Station E-Mail"
      Height          =   225
      Left            =   9390
      TabIndex        =   20
      Top             =   6795
      Width           =   1320
   End
   Begin V81Affiliate.AffContactGrid udcContactGrid 
      Height          =   390
      Left            =   210
      TabIndex        =   19
      Top             =   5310
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   767
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9600
      Top             =   4695
   End
   Begin VB.Timer tmcStation 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   10125
      Top             =   4035
   End
   Begin VB.ListBox lbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      ItemData        =   "AffStationSearch.frx":09D8
      Left            =   180
      List            =   "AffStationSearch.frx":09DA
      TabIndex        =   18
      Top             =   6210
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.ListBox lbcOwner 
      Height          =   255
      ItemData        =   "AffStationSearch.frx":09DC
      Left            =   10455
      List            =   "AffStationSearch.frx":09DE
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   5850
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmcAddComment 
      Caption         =   "Add"
      Height          =   195
      Left            =   9390
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmcRefresh 
      Caption         =   "Refresh"
      Height          =   225
      Left            =   9105
      TabIndex        =   13
      Top             =   6540
      Width           =   885
   End
   Begin VB.TextBox edcCallLetters 
      Height          =   315
      Left            =   2925
      TabIndex        =   11
      Top             =   6525
      Width           =   1005
   End
   Begin VB.CommandButton cmcCloseComment 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9225
      TabIndex        =   9
      Top             =   3015
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton cmcFilter 
      Caption         =   "Filter"
      Enabled         =   0   'False
      Height          =   225
      Left            =   4110
      TabIndex        =   5
      Top             =   6585
      Width           =   885
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AffStationSearch.frx":09E0
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   15
      Width           =   60
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11040
      Top             =   5295
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7050
      FormDesignWidth =   11580
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      Height          =   225
      Left            =   10080
      TabIndex        =   4
      Top             =   6540
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStations 
      Height          =   1005
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   29
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   29
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAgreementInfo 
      Height          =   1005
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   62
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   62
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpotInfo 
      Height          =   1005
      Left            =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4245
      Visible         =   0   'False
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPostedInfo 
      Height          =   1005
      Left            =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3180
      Visible         =   0   'False
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPostBuy 
      Height          =   540
      Left            =   7875
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   953
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcPrinter 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   11070
      Picture         =   "AffStationSearch.frx":0CEA
      Top             =   6450
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lacCommentTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8055
      TabIndex        =   23
      Top             =   2610
      Visible         =   0   'False
      Width           =   2910
      WordWrap        =   -1  'True
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Picture         =   "AffStationSearch.frx":15B4
      Top             =   6570
      Width           =   480
   End
   Begin VB.Label lacUserOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Assigned"
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
      Height          =   300
      Left            =   870
      TabIndex        =   14
      Top             =   6525
      Width           =   1095
   End
   Begin VB.Label lacFilter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Left            =   5055
      TabIndex        =   12
      Top             =   6540
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lacCallLetters 
      Caption         =   "Call Letter Search"
      Height          =   405
      Left            =   2085
      TabIndex        =   10
      Top             =   6495
      Width           =   870
   End
End
Attribute VB_Name = "frmStationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmStationSearch - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'Missing
'
'Using User Initials
'Station:  If filter is "Contains" and call letters selected, include stations with matching call letters defined in History
'Station:  Show station if agreement assigned to you or looking at Service or Market Rep
'Agreement: Show Pledge date range or start date plus days?
'Agreement: Test if Receipt only or Count only and show in Red (Not Yet Posted) instead of Brown (Not Exported)
'Weeks: Test if Receipt only or Count only and show in Red (Not Yet Posted) instead of Brown (Not Exported)
'Weeks:  Maybe show letter of each color.
'Follow-Up filter on user rights and which comments to view (only user own; all comments assign to Market rep, service rep or maybe department)
'Comment:  Include ones assigned to you via E-Mail "To".  Allow to see those thru next Sunday.
'Comment:  Update record when Ok set or cleared
'Comment expand button
'Comment sort.  Add station and other fields
'Comment in Follow up mode only allow viewing of comments not Ok'd
'Comment:  Have a toggle to see comments created by other users, might be by department
'Comment:  Add a delete column or button or allow comment to be blanked out or add trash can
'Comment:  Test site option to see if comment can be altered
'Spots:  New color for MG and Bonus
'Print Icon:  Add Print Icon
'E-Mail send operation

'Write a utility program to match Contact and E-Mail by station

'Follow-Up Comment rules:
'1.       Comment must be entered by the signed-in user,
'2.       Comment must have a follow-up date with no OK flag,
'3.       Follow-up date must be dated today or prior if the comment is a manually-entered comment or an email to a station,
'4.       Follow-up date must be next Sunday or prior if the comment is an email sent to a fellow employee.


Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private lmRowSelected As Long
Private bmInMouseUp As Boolean
Private bmInInit As Boolean

Private smMissedFeedDate As String
Private smMissedFeedTime As String
Private smMGAirDate As String
Private smMGAirTime As String

Private smUserType As String

Private smWeek1 As String
Private smWeek54 As String
Private imShttCode As Integer
Private lmAttCode As Long
Private imVefCode As Integer
Private lmStationMaxRow As Long

Private imLastStationColSorted As Integer
Private imLastStationSort As Integer

Private imLastAgreementInfoColSorted As Integer
Private imLastAgreementInfoSort As Integer

Private imLastPostedInfoColSorted As Integer
Private imLastPostedInfoSort As Integer

Private imLastSpotInfoColSorted As Integer
Private imLastSpotInfoSort As Integer

Private bmStationScrollAllowed As Boolean
Private lmStationTopRow As Long

Private bmAgreementScrollAllowed As Boolean
Private lmAgreementTopRow As Long

Private bmPostedScrollAllowed As Boolean
Private lmPostedTopRow As Long

Private bmInScroll As Integer
Private imDelaySource As Integer    '0=Station; 1=Agreement(Vehicle); 2= Posted Inofrmation (Weeks)

Private bmContactViaCmtCol As Boolean

Private imCtrlVisible As Integer
Private lmEnableRow As Long
Private lmEnableCol As Long
Private lmTopRow As Long

Private rst_Shtt As ADODB.Recordset
Private rst_att As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset
Private rst_Webl As ADODB.Recordset
Private rst_Ast As ADODB.Recordset
Private rst_Lst As ADODB.Recordset
Private rst_cct As ADODB.Recordset
Private rst_artt As ADODB.Recordset
Private rst_dnt As ADODB.Recordset
Private rst_mnt As ADODB.Recordset
Private rst_clt As ADODB.Recordset
Private rst_chf As ADODB.Recordset
Private rst_clf As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset

Private imPostBuyVef() As Integer

Private hmAst As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmGameAstInfo() As ASTINFO

Private tmFilterLink() As FILTERLINK
Private tmAndFilterLink() As FILTERLINK
Private imFilterShttCode() As Integer

Private smCommentType As String

Private smStationType As String

Private lmCommentStartLoc As Long

'Grid Controls

'10/31/18: Replace how station grid is populated
Private Type STATIONGRIDKEY
    sKey As String * 100
    lRow As Long
End Type
Dim tmStationGridKey() As STATIONGRIDKEY
Dim smStationGridData() As String

'Station Grid-grdStations
Const SSELECTINDEX = 0
Const SCALLLETTERINDEX = 1     'If changed, change frmStation and frmContactEMail as this constant is defined within those modules.
Const SDUEINDEX = 2
Const SDMARANKINDEX = 3
Const SDMAMARKETINDEX = 4      'If changed, change frmStation as this constant is defined within the module.
Const SMSARANKINDEX = 5
Const SMSAMARKETINDEX = 6
Const SSTATEINDEX = 7
Const SZONEINDEX = 8
Const SZIPINDEX = 9
Const SFORMATINDEX = 10
Const SOWNERINDEX = 11
Const SOPERATORINDEX = 12
Const SCLUSTERINDEX = 13
Const SMCASTINDEX = 14
Const SPWINDEX = 15
Const SWEBSITEINDEX = 16
Const SAGRMNTINDEX = 17
Const SCOMMENTINDEX = 18
Const SSHTTCODEINDEX = 19   'If changed, change frmStation and frmContactEMail as this constant is defined within those modules.
Const SFREQMONIKERINDEX = 20
Const SCLUSTERNAMESINDEX = 21
Const SMCASTNAMESINDEX = 22
Const SAUDP12PLUSINDEX = 23
Const SWATTSINDEX = 24
Const SPASSWORDINDEX = 25
Const SWEBADDRESSINDEX = 26
Const SOPERATORNAMEINDEX = 27
Const SSORTINDEX = 28

Const PSSTATIONTYPEINDEX = 0
Const PSADVERTISERINDEX = 1
Const PSWEEKOFINDEX = 2
Const PSCONTRACTINDEX = 3


'Week Info Grid- grdAgreementInfo
Const ASELECTINDEX = 0
Const AVEHICLEINDEX = 1
Const ADUEINDEX = 2
Const ADAYTIMEINDEX = 3
Const AWEEK1INDEX = 4
Const AAGRMNTINDEX = 58
Const ATIMERANGEINDEX = 59
Const AATTCODEINDEX = 60
Const ASORTINDEX = 61

'Posted Info Grid- grdPostedInfo
Const PSELECTINDEX = 0
Const PWEEKINDEX = 1
Const PNOSCHDINDEX = 2
Const PNOAIREDINDEX = 3
Const PNOCMPLINDEX = 4
Const PPERCENTAINDEX = 5
Const PPERCENTCINDEX = 6
Const PPOSTDATEINDEX = 7
Const PBYINDEX = 8
Const PIPINDEX = 9
Const PSTATUSINDEX = 10
Const PCPTTINDEX = 11
Const PSORTINDEX = 12


'Spot Info Grid- grdSpotInfo
Const DDATEINDEX = 0
Const DFEDINDEX = 1
Const DPLEGDEDDAYINDEX = 2
Const DPLEGDEDTIMEINDEX = 3
Const DAIREDDATEINDEX = 4
Const DAIREDTIMEINDEX = 5
Const DADVTINDEX = 6
Const DPRODINDEX = 7
Const DLENGTHINDEX = 8
Const DISCIINDEX = 9
Const DCARTINDEX = 10
Const DCOMMENTINDEX = 11
Const DSTATUSINDEX = 12
Const DCNTRNOINDEX = 13
Const DGAMEINFOINDEX = 14
Const DMMRINFOINDEX = 15
Const DASTCODEINDEX = 16
Const DSORTINDEX = 17



Private Sub mClearGrid(grdCtrl As MSHFlexGrid)
    Dim llRow As Long
    Dim llCol As Long
    
    grdCtrl.Redraw = False
    'Set color within cells
    For llRow = grdCtrl.FixedRows To grdCtrl.Rows - 1 Step 1
        For llCol = 0 To grdCtrl.Cols - 1 Step 1
            grdCtrl.Row = llRow
            grdCtrl.Col = llCol
            grdCtrl.Text = ""
            grdCtrl.CellBackColor = vbWhite
        Next llCol
    Next llRow
    grdCtrl.Redraw = True
End Sub



Private Sub cmcAddComment_Click()
    If udcCommentGrid.Visible Then
        mSaveCommentsAndContacts
        udcCommentGrid.Action 1 'Init
        udcCommentGrid.Action 6 'Add Row
    End If
End Sub

Private Sub cmcAddComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    edcCommentTip.Visible = False
End Sub

Private Sub cmcCloseComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    edcCommentTip.Visible = False
End Sub

Private Sub cmcDone_GotFocus()
    mSaveCommentsAndContacts
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdPostBuy.Col
        Case PSWEEKOFINDEX
            lbcDate.Visible = Not lbcDate.Visible
        Case PSADVERTISERINDEX
            lbcAdvertiser.Visible = Not lbcAdvertiser.Visible
        Case PSCONTRACTINDEX
            lbcContract.Visible = Not lbcContract.Visible
    End Select
End Sub

Private Sub cmcEMail_Click()
    'If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) = AFFILIATECRM) Then
        If (grdAgreementInfo.Visible) Or (udcContactGrid.Visible) Or udcCommentGrid.Visible Then
            igContactEmailShttCode = imShttCode
        Else
            igContactEmailShttCode = -1
        End If
        frmContactEMail.Show vbModal
        If udcCommentGrid.Visible Then
            mSaveCommentsAndContacts
            udcCommentGrid.StationCode = imShttCode
            udcCommentGrid.Action 3 'Populate
            lmCommentStartLoc = udcCommentGrid.ColumnStartLocation(6)
        End If
    'End If
End Sub

Private Sub cmcEMail_GotFocus()
    lbcUserOption(0).Visible = False
    mSaveCommentsAndContacts
End Sub

Private Sub cmcFilter_GotFocus()
    lbcUserOption(0).Visible = False
    mSaveCommentsAndContacts
End Sub

Private Sub cmcCloseComment_Click()
    mMousePointer vbHourglass
    mHideCommentButtons
    'Save any changes
    mSaveCommentsAndContacts
    udcContactGrid.Visible = False
    udcCommentGrid.Visible = False
    mSetGridPosition
    mSetCommands
    mMousePointer vbDefault
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim imEnabled As Integer
    
    imEnabled = tmcStation.Enabled
    tmcStation.Enabled = False
    If Not udcContactGrid.VerifyRights("M") Then
        Exit Sub
    End If
    If grdStations.Visible Then
        If sgStationSearchCallSource = "P" Then
            ilRet = MsgBox("OK to Exit Post-Buy Planning?", vbQuestion + vbYesNo, "Exit")
        Else
            ilRet = MsgBox("OK to Exit Affiliate Management?", vbQuestion + vbYesNo, "Exit")
        End If
        If ilRet = vbNo Then
            tmcStation.Enabled = imEnabled
            Exit Sub
        End If
    End If
    Unload frmStationSearch
    Exit Sub
   
End Sub


Private Sub cmcFilter_Click()
    tmcStation.Enabled = False
    mSaveCommentsAndContacts
    frmStationSearchFilter.Show vbModal
    If Not igFilterReturn Then
        'If grdStations.Visible = False Then
        '    mMousePointer vbHourglass
        '    mPopStationsGrid
        '    grdStations.Visible = True
        '    rbcComments(0).Enabled = True
        '    rbcComments(1).Enabled = True
        '    mMousePointer vbDefault
        'End If
        tmcStation.Enabled = False
        Unload frmStationSearch
        Exit Sub
    End If
    'Create Filter test
    mMousePointer vbHourglass
    If UBound(tgFilterDef) > 0 Then
        lacFilter.Caption = "On"
    Else
        lacFilter.Caption = "Off"
    End If
    mBuildFilter
    mPopStationsGrid
    grdStations.Visible = True
    If UBound(tgFilterDef) > 0 Then
        lacFilter.Visible = True
        udcCommentGrid.FilterStatus = True  'vbTrue
    Else
        lacFilter.Visible = False
        udcCommentGrid.FilterStatus = False 'vbFalse
    End If
    rbcComments(0).Enabled = True
    rbcComments(1).Enabled = True
    If udcCommentGrid.Visible Then
        lmCommentStartLoc = udcCommentGrid.ColumnStartLocation(6)
        udcCommentGrid.Action 3 'Populate
    End If
    mMousePointer vbDefault
End Sub

Private Sub cmcGen_Click()
    Dim slSDate As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim blRet As Boolean
    
    If lbcDate.ListIndex < 0 Then
        MsgBox "Week Of must be defined"
        Exit Sub
    End If
    slSDate = Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX))
    If slSDate <> "[All]" Then
        If Not gIsDate(slSDate) Then
            MsgBox "'Week Of' not a valid date"
            Exit Sub
        End If
    End If
    If lbcAdvertiser.ListIndex < 0 Then
        MsgBox "'Advertiser' must be selected"
        Exit Sub
    End If
    If lbcContract.ListIndex < 0 Then
        MsgBox "Contract or [All] must be selected"
        Exit Sub
    End If
    ilIndex = -1
    For ilLoop = 0 To UBound(tgFilterTypes) - 1 Step 1
        If Trim$(tgFilterTypes(ilLoop).sFieldName) = "Call Letters" Then
            ilIndex = ilLoop
        End If
    Next ilLoop
    If ilIndex = -1 Then
        Exit Sub
    End If
    mMousePointer vbHourglass
    blRet = mCreatePostBuyStationFilter()
    mMousePointer vbDefault
    Exit Sub
End Sub

Private Sub cmcGen_GotFocus()
    mSetShowCommentsAndContacts
    mSetShow
End Sub

Private Sub cmcRefresh_Click()
    Dim blRet As Boolean
    mMousePointer vbHourglass
    If sgStationSearchCallSource = "P" Then
        blRet = mCreatePostBuyStationFilter()
    Else
        mPopStationsGrid
    End If
    mSetCommands
    mMousePointer vbDefault
End Sub

Private Sub cmcRefresh_GotFocus()
    lbcUserOption(0).Visible = False
End Sub

Private Sub edcCallLetters_Change()
    Dim slStr As String
    Dim slInputStr As String
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    Dim llStationRow As Long
    
    slInputStr = UCase$(Trim$(edcCallLetters.Text))
    Do While Len(slInputStr) < 4
        slInputStr = slInputStr & "!"
    Loop
'    llMin = grdStations.FixedRows
'    llMax = lmStationMaxRow 'grdStations.Rows - 1
'    Do While llMin <= llMax
'        llMiddle = (llMin + llMax) \ 2
'        slStr = UCase$(Trim$(grdStations.TextMatrix(llMiddle, SCALLLETTERINDEX)))
'        'If InStr(1, slStr, slInputStr, vbBinaryCompare) = 1 Then
'        '    grdStations.TopRow = llMiddle
'        '    Exit Sub
'        'End If
'        ilResult = StrComp(slStr, slInputStr, vbBinaryCompare)
'        Select Case ilResult
'            Case 0:
'                grdStations.TopRow = llMiddle
'                Exit Sub
'            Case 1:
'                llMax = llMiddle - 1
'            Case -1:
'                llMin = llMiddle + 1
'        End Select
'    Loop
'    grdStations.TopRow = llMin
    llMin = LBound(tmStationGridKey)
    llMax = UBound(tmStationGridKey) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        llStationRow = tmStationGridKey(llMiddle).lRow
        slStr = UCase$(Trim$(smStationGridData(llStationRow + SCALLLETTERINDEX)))
        'If InStr(1, slStr, slInputStr, vbBinaryCompare) = 1 Then
        '    grdStations.TopRow = llMiddle
        '    Exit Sub
        'End If
        ilResult = StrComp(slStr, slInputStr, vbBinaryCompare)
        Select Case ilResult
            Case 0:
                vbcStation.Value = llMiddle
                mFillStationGrid
                Exit Sub
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    'grdStations.TopRow = llMin
    vbcStation.Value = llMin
    mFillStationGrid
    Exit Sub
End Sub

Private Sub edcCallLetters_GotFocus()
    lbcUserOption(0).Visible = False
End Sub

Private Sub edcDropdown_Change()
    Dim slStr As String
    
    Select Case lmEnableCol
        Case PSWEEKOFINDEX
            mDropdownChangeEvent lbcDate
        Case PSADVERTISERINDEX
            mDropdownChangeEvent lbcAdvertiser
        Case PSCONTRACTINDEX
            mDropdownChangeEvent lbcContract
    End Select
    mSetPSCommands
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If edcDropDown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmEnableCol
            Case PSWEEKOFINDEX
                gProcessArrowKey Shift, KeyCode, lbcDate, True
            Case PSADVERTISERINDEX
                gProcessArrowKey Shift, KeyCode, lbcAdvertiser, True
            Case PSCONTRACTINDEX
                gProcessArrowKey Shift, KeyCode, lbcContract, True
        End Select
    End If
End Sub

Private Sub grdPostBuy_EnterCell()
    mSetShow
End Sub

Private Sub grdPostBuy_GotFocus()
    mSetShow
End Sub

Private Sub grdPostBuy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim llCode As Long
    Dim ilRet As Integer

    If bmInMouseUp Then
        Exit Sub
    End If
    bmInMouseUp = True
    If Y < grdPostBuy.RowHeight(0) Then
        bmInMouseUp = False
        Exit Sub
    End If
    ilCol = grdPostBuy.MouseCol
    ilRow = grdPostBuy.MouseRow
    If ilCol < grdPostBuy.FixedCols Then
        grdPostBuy.Redraw = True
        bmInMouseUp = False
        Exit Sub
    End If
    If ilRow < grdPostBuy.FixedRows Then
        grdPostBuy.Redraw = True
        bmInMouseUp = False
        Exit Sub
    End If
    grdPostBuy.Col = ilCol
    grdPostBuy.Row = ilRow
    'If Not mColOk() Then
    '    grdPostBuy.Redraw = True
    '    Exit Sub
    'End If
    grdPostBuy.Redraw = True
    mEnableBox
    bmInMouseUp = False
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub grdStations_GotFocus()
    lbcUserOption(0).Visible = False
    mSaveCommentsAndContacts
End Sub

Private Sub imcKey_Click()
    lbcKey.Visible = Not lbcKey.Visible
End Sub

'
'           User has clicked Printer icon for list of stations to be printed.
'           The list printed is in the same sort order as the screen, with
'           filtering included.
'
Private Sub imcPrinter_Click()
    Dim llRow As Long
    Dim slStr As String
    Dim ilFN As Integer
    Dim slExportName As String
    Dim ilRptDest As Integer
    Dim ilExportType As Integer
    Dim SQLQuery As String
    Dim slRptName As String
    Dim sGenDate As String
    Dim sGenTime As String
    Dim slCallLetters As String
    Dim slDMA As String
    Dim slDMARank As String
    Dim slMSA As String
    Dim slMSARank As String
    Dim slState As String
    Dim slZone As String
    Dim slZip As String
    Dim slFormat As String
    Dim slOwner As String
    Dim slOper As String * 1
    Dim slSister As String * 1
    Dim slMulticast As String * 1
    Dim slPW As String * 1
    Dim slWeb As String * 1
    Dim slAgree As String * 1
    Dim slComment As String * 1
    Dim slSelect As String
    Dim slSelectCommand As String
    Dim slSelectFrom As String
    Dim slSelectTo As String
    Dim ilRet As Integer
    Dim llShttCode As Long
    Dim ilOption As Integer
    Dim llKey As Long
    Dim llStationRow As Long
    
    On Error GoTo ErrHand
    lbcUserOption(0).Visible = False
    'ilRet = MsgBox("Include Contact Information?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Station Contact")
    sgGenMsg = "Affiliate Management Station Filter Selectivity  "
    sgCMCTitle(0) = "OK"
    sgCMCTitle(1) = "Cancel"
    sgCMCTitle(2) = ""
    sgCMCTitle(3) = ""
    sgRadioTitle(0) = "Contact Information"
    sgRadioTitle(1) = "Vehicle List"
    sgRadioTitle(2) = "Contact and Vehicles"
    sgRadioTitle(3) = "None"
    igDefCMC = 0
    igEditBox = 2
    sgEditValue = "2"
    frmGenMsg.Show vbModal

'    If ilRet = vbCancel Then
'        Exit Sub
'    End If

    If igAnsCMC <> 0 Then
        Exit Sub
    End If
    
    ilOption = Val(sgEditValue)
    If ilOption = 0 Or ilOption = 2 Then        'include contacts or both contacts and vehicles
        sgCrystlFormula1 = "'Y'"
    Else
        sgCrystlFormula1 = "'N'"
    End If
    If ilOption = 1 Or ilOption = 2 Then    'include vehicles Or both
        sgCrystlFormula2 = "'Y'"
    Else
        sgCrystlFormula2 = "'N'"
    End If
    
    Screen.MousePointer = vbHourglass
    sgReportListName = "Station Filter"
    gUserActivityLog "S", sgReportListName & ": Prepass"
    'get generation date and time for crystal report filter of records
    sGenDate = Format$(gNow(), "m/d/yyyy")
    sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
'    If ilRet = vbYes Then       'Include personnel info
'         sgCrystlFormula1 = "'Y'"
'    Else
'        sgCrystlFormula1 = "'N'"
'    End If
    
    If sgStationSearchCallSource <> "P" Then
        'Filter information is in tgFilterInfo
        For llRow = 0 To UBound(tgFilterDef) - 1 Step 1
            If tgFilterDef(llRow).iSelect <> 0 Then
                'tgFilterDef(llRow).iSelect is used to find the field name from tgFilterTypes()
                For ilFN = 0 To UBound(tgFilterTypes) - 1 Step 1
                    If tgFilterDef(llRow).iSelect = tgFilterTypes(ilFN).iSelect Then
                        'Name obtained from tgFilterTypes(ilFN).sFieldName
                        slSelect = Trim$(tgFilterTypes(ilFN).sFieldName)
                        Exit For
                    End If
                Next ilFN
                'tgFilterDef(llRow).iOperator: 0=Contains; 1=Equal; 2=Not Equal; 3=Range
                If tgFilterDef(llRow).iOperator = 0 Then
                    slSelectCommand = "Contains"
                ElseIf tgFilterDef(llRow).iOperator = 1 Then
                    slSelectCommand = "Equal"
                ElseIf tgFilterDef(llRow).iOperator = 2 Then
                    slSelectCommand = "Not Equal"
                ElseIf tgFilterDef(llRow).iOperator = 3 Then
                    slSelectCommand = "Range"
                Else
                    slSelectCommand = "Greater or Equal"
                End If
                'tgFilterDef(llRow).sFromValue and tgFilterDef(llRow).sToValue contain the values
                'to print along with the Operator.  The ToValue will be blank except for Range.
                slSelectFrom = Trim$(tgFilterDef(llRow).sFromValue)
                slSelectTo = Trim$(tgFilterDef(llRow).sToValue)
                SQLQuery = "INSERT INTO SMR "
                'smrstring5 and smrstring3 are the from/to spans
                'smrstring5 could be very long, it is 100 bytes long
                SQLQuery = SQLQuery & " (smrString1, smrString2 , smrstring5, smrstring3, "
                SQLQuery = SQLQuery & "smrType, smrSeqNo, smrGenDate, smrGenTime) "
                
                SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$(slSelect) & "', '" & Trim$(slSelectCommand) & "', '" & Trim$(slSelectFrom) & "', '" & Trim$(slSelectTo) & "', "
                SQLQuery = SQLQuery & "'A', " & llRow & ", '" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
                cnn.BeginTrans
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
                    cnn.RollbackTrans
                    Exit Sub
                End If
                cnn.CommitTrans
                llRow = llRow
            Else
                Exit For
            End If
        Next llRow
    Else
        SQLQuery = "INSERT INTO SMR "
        SQLQuery = SQLQuery & " (smrString1, smrString2 , smrstring5, smrstring3, "
        SQLQuery = SQLQuery & "smrType, smrSeqNo, smrGenDate, smrGenTime) "
        
        SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$("Station Type") & "', '" & Trim$("Equal") & "', '" & Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSSTATIONTYPEINDEX)) & "', '" & Trim$("") & "', "
        SQLQuery = SQLQuery & "'A', " & 0 & ", '" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        SQLQuery = "INSERT INTO SMR "
        SQLQuery = SQLQuery & " (smrString1, smrString2 , smrstring5, smrstring3, "
        SQLQuery = SQLQuery & "smrType, smrSeqNo, smrGenDate, smrGenTime) "
        
        SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$("Advertiser") & "', '" & Trim$("Equal") & "', '" & gFixQuote(Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSADVERTISERINDEX))) & "', '" & Trim$("") & "', "
        SQLQuery = SQLQuery & "'A', " & 1 & ", '" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        SQLQuery = "INSERT INTO SMR "
        SQLQuery = SQLQuery & " (smrString1, smrString2 , smrstring5, smrstring3, "
        SQLQuery = SQLQuery & "smrType, smrSeqNo, smrGenDate, smrGenTime) "
        
        SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$("Week") & "', '" & Trim$("Equal") & "', '" & Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX)) & "', '" & Trim$("") & "', "
        SQLQuery = SQLQuery & "'A', " & 2 & ", '" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        SQLQuery = "INSERT INTO SMR "
        SQLQuery = SQLQuery & " (smrString1, smrString2 , smrstring5, smrstring3, "
        SQLQuery = SQLQuery & "smrType, smrSeqNo, smrGenDate, smrGenTime) "
        
        SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$("Contract") & "', '" & Trim$("Equal") & "', '" & Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSCONTRACTINDEX)) & "', '" & Trim$("") & "', "
        SQLQuery = SQLQuery & "'A', " & 3 & ", '" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
    End If
    ''Station grid information is in grdStation.
    ''Index into grdStations.TextMatrix is found in the declare area up at top of the module
    ''SCALLLETTERSINDEX - SCOMMENTINDEX
    'For llRow = grdStations.FixedRows To grdStations.Rows - 1 Step 1
    For llKey = 0 To UBound(tmStationGridKey) - 1 Step 1
        llStationRow = tmStationGridKey(llKey).lRow
        'slCallLetters = Trim$(grdStations.TextMatrix(llRow, SCALLLETTERINDEX))
        slCallLetters = Trim$(smStationGridData(llStationRow + SCALLLETTERINDEX))
        If slCallLetters <> "" Then
        '    slDMA = gFixQuote(Trim$(grdStations.TextMatrix(llRow, SDMAMARKETINDEX)))
        '    slDMARank = Trim$(grdStations.TextMatrix(llRow, SDMARANKINDEX))
        '    slMSA = gFixQuote(Trim$(grdStations.TextMatrix(llRow, SMSAMARKETINDEX)))
        '    slMSARank = Trim$(grdStations.TextMatrix(llRow, SMSARANKINDEX))
        '    slState = Trim$(grdStations.TextMatrix(llRow, SSTATEINDEX))
        '    slZone = Trim$(grdStations.TextMatrix(llRow, SZONEINDEX))
        '    slZip = Trim$(grdStations.TextMatrix(llRow, SZIPINDEX))
        '    slFormat = gFixQuote(Trim$(grdStations.TextMatrix(llRow, SFORMATINDEX)))
        '    slOwner = gFixQuote(Trim$(grdStations.TextMatrix(llRow, SOWNERINDEX)))
        '    slOper = (grdStations.TextMatrix(llRow, SOPERATORINDEX))
        '    slSister = (grdStations.TextMatrix(llRow, SCLUSTERINDEX))
        '    slMulticast = (grdStations.TextMatrix(llRow, SMCASTINDEX))
        '    slPW = (grdStations.TextMatrix(llRow, SPWINDEX))
        '    slWeb = (grdStations.TextMatrix(llRow, SWEBSITEINDEX))
        '    slAgree = (grdStations.TextMatrix(llRow, SAGRMNTINDEX))
        '    slComment = (grdStations.TextMatrix(llRow, SCOMMENTINDEX))
        '    llShttCode = (grdStations.TextMatrix(llRow, SSHTTCODEINDEX))
            slDMA = gFixQuote(Trim$(smStationGridData(llStationRow + SDMAMARKETINDEX)))
            slDMARank = Trim$(smStationGridData(llStationRow + SDMARANKINDEX))
            slMSA = gFixQuote(Trim$(smStationGridData(llStationRow + SMSAMARKETINDEX)))
            slMSARank = Trim$(smStationGridData(llStationRow + SMSARANKINDEX))
            slState = Trim$(smStationGridData(llStationRow + SSTATEINDEX))
            slZone = Trim$(smStationGridData(llStationRow + SZONEINDEX))
            slZip = Trim$(smStationGridData(llStationRow + SZIPINDEX))
            slFormat = gFixQuote(Trim$(smStationGridData(llStationRow + SFORMATINDEX)))
            slOwner = gFixQuote(Trim$(smStationGridData(llStationRow + SOWNERINDEX)))
            slOper = (smStationGridData(llStationRow + SOPERATORINDEX))
            slSister = (smStationGridData(llStationRow + SCLUSTERINDEX))
            slMulticast = (smStationGridData(llStationRow + SMCASTINDEX))
            slPW = (smStationGridData(llStationRow + SPWINDEX))
            slWeb = (smStationGridData(llStationRow + SWEBSITEINDEX))
            slAgree = (smStationGridData(llStationRow + SAGRMNTINDEX))
            slComment = (smStationGridData(llStationRow + SCOMMENTINDEX))
            llShttCode = (smStationGridData(llStationRow + SSHTTCODEINDEX))
            
            SQLQuery = "INSERT INTO SMR "
            SQLQuery = SQLQuery & " (smrCallLetters, smrDMAMarket , smrMSAMarket, smrDMARank, smrMSARank, "
            SQLQuery = SQLQuery & "smrState, smrZone, smrZipCode, smrFormat, "
            SQLQuery = SQLQuery & "smrOwner, smrOperator, smrSister, "
            SQLQuery = SQLQuery & "smrPassword, smrWeb, smrAGree, smrComment, "
            SQLQuery = SQLQuery & "smrType, smrSeqNo, smrShttCode, smrGenDate, smrGenTime) "
    
            SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$(slCallLetters) & "', '" & slDMA & "', '" & slMSA & "', '" & slDMARank & "', '" & slMSARank & "', "
            SQLQuery = SQLQuery & "'" & Trim$(slState) & " ', '" & Trim$(slZone) & "', '" & Trim$(slZip) & "', '" & Trim$(slFormat) & "', "
            SQLQuery = SQLQuery & "'" & Trim$(slOwner) & "  ', '" & Trim$(slOper) & "', '" & slSister & "', "
            SQLQuery = SQLQuery & "'" & Trim$(slPW) & " ', '" & slWeb & "', '" & slAgree & "', '" & slComment & "', "

            SQLQuery = SQLQuery & "'B', " & llKey & ", " & llShttCode & ", '" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
            cnn.BeginTrans
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
                cnn.RollbackTrans
                Exit Sub
            End If
            cnn.CommitTrans
            llRow = llRow
        Else                'blank call letters, all done
            Exit For
        End If
    'Next llRow
    Next llKey
    
    ilExportType = 0        'no export
    ilRptDest = 0           'force to printer
    slRptName = "afStnFilter.rpt"
    slExportName = "StationFilter"
    SQLQuery = "SELECT * FROM smr inner join shtt on smrshttcode = shttcode "
    SQLQuery = SQLQuery + " WHERE  smrGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND smrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'" & " and smrshttcode = shttcode"
    'QLQuery = "SELECT * FROM smr  "
    'SQLQuery = SQLQuery + " WHERE  smrGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND smrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
    Screen.MousePointer = vbHourglass
    gUserActivityLog "E", sgReportListName & ": Prepass"

    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
    Screen.MousePointer = vbDefault
    
    ' Delete the info we stored in the temporary SMR table
    SQLQuery = "DELETE FROM SMR "
    SQLQuery = SQLQuery & " WHERE (smrGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and smrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-imcPrint_Click"
    Exit Sub
End Sub

Private Sub lbcAdvertiser_Click()
    edcDropDown.Text = lbcAdvertiser.List(lbcAdvertiser.ListIndex)
    mSetPSCommands
End Sub

Private Sub lbcContract_Click()
    edcDropDown.Text = lbcContract.List(lbcContract.ListIndex)
    mSetPSCommands
End Sub

Private Sub lbcDate_Click()
    edcDropDown.Text = lbcDate.List(lbcDate.ListIndex)
    mSetPSCommands
End Sub

Private Sub lbcUserOption_Click(Index As Integer)
    tmcStation.Enabled = False
    If sgStationSearchCallSource <> "P" Then
        lacUserOption.Caption = lbcUserOption(0).List(lbcUserOption(0).ListIndex)
        lbcUserOption(0).Visible = False
    Else
        lacUserOption.Caption = lbcUserOption(1).List(lbcUserOption(1).ListIndex)
        lbcUserOption(1).Visible = False
    End If
    If Not bmInInit Then
        tmcStation.Enabled = True
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    lbcUserOption(0).Visible = False
End Sub

Private Sub pbcCommentType_Click()
    If Trim$(smCommentType) = "All" Then
        smCommentType = "   Dept"
    ElseIf Trim$(smCommentType) = "Mine" Then
        smCommentType = "    All"
    ElseIf Trim$(smCommentType) = "Dept" Then
        smCommentType = "   Mine"
    End If
    pbcCommentType.Cls
    pbcCommentType_Paint
    udcCommentGrid.Action 3 'Populate
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        Me.Visible = True
        Me.Width = Screen.Width / 1.1   '1.05
        Me.Height = Screen.Height / 1.2 '1.15
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
        gSetFonts frmStationSearch
        pbcArrow.Width = 90
        gCenterForm frmStationSearch
        
        If sgStationSearchCallSource = "P" Then
            bgPostBuyVisible = True
        Else
            bgManagementVisible = True
        End If
        mMousePointer vbHourglass
        tmcStart.Enabled = True
        'mSetGridColumns
        'mSetGridTitles
        'gGrid_IntegralHeight grdStations
        'gGrid_FillWithRows grdStations
        'mPopStationsGrid
        'lbcKey.FontBold = False
        'lbcKey.FontName = "Arial"
        'lbcKey.FontBold = False
        'lbcKey.FontSize = 8
        'lbcKey.Height = (lbcKey.ListCount - 1) * 225
        'lbcKey.Move imcKey.Left, imcKey.Top - lbcKey.Height
        'imFirstTime = False
        'Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    mSetShowCommentsAndContacts
    mSetShow
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.1   '1.05
'    Me.Height = Screen.Height / 1.2 '1.15
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
'    gSetFonts frmStationSearch
'    pbcArrow.Width = 90
'    gCenterForm frmStationSearch
End Sub

Private Sub Form_Load()
    mMousePointer vbHourglass
    
    mInit
    mMousePointer vbDefault
    Exit Sub
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    edcCommentTip.Visible = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    mSetGridColumns
    mSetGridPosition
    gSetListBoxHeight lbcUserOption(0), 4
    gSetListBoxHeight lbcUserOption(1), 2
    lbcUserOption(0).Left = lacUserOption.Left
    lbcUserOption(0).Top = lacUserOption.Top - lbcUserOption(0).Height
    lbcUserOption(1).Left = lacUserOption.Left
    lbcUserOption(1).Top = lacUserOption.Top - lbcUserOption(1).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    On Error Resume Next
    If sgStationSearchCallSource = "P" Then
        bgPostBuyVisible = False
    Else
        bgManagementVisible = False
    End If
    
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmGameAstInfo
    Erase tmFilterLink
    Erase tmAndFilterLink
    Erase imPostBuyVef
    Erase imFilterShttCode
    
    rst_Shtt.Close
    rst_att.Close
    rst_Cptt.Close
    rst_Webl.Close
    rst_Ast.Close
    rst_Lst.Close
    rst_cct.Close
    rst_artt.Close
    rst_dnt.Close
    rst_mnt.Close
    rst_clt.Close
    rst_chf.Close
    rst_clf.Close
    rst_Gsf.Close
    On Error GoTo 0
    
    Set frmStationSearch = Nothing
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long
    frmStationSearch.Visible = False
    Screen.MousePointer = vbHourglass
    bmInInit = True
    If sgStationSearchCallSource = "P" Then
        frmStationSearch.Caption = "Post-Buy Planning- " & Trim$(sgUserName)
    Else
        frmStationSearch.Caption = "Management- " & Trim$(sgUserName)
    End If
    
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    bmInScroll = False
    bmInMouseUp = False
    imDelaySource = -1
    smCommentType = "    All"
    pbcFocus.Move -100, -100
    pbcClickFocus.Move -100, -100
    imcPrinter.Picture = frmDirectory!imcPrinter.Picture
    mPopFilterTypes
    ReDim tgFilterDef(0 To 0) As FILTERDEF
    ReDim tgNotFilterDef(0 To 0) As FILTERDEF
    ReDim tmFilterLink(0 To 0) As FILTERLINK
    ReDim tmAndFilterLink(0 To 0) As FILTERLINK
    ReDim igCommentShttCode(0 To 0) As Integer
    ReDim tmStationGridKey(0 To 0) As STATIONGRIDKEY
    ReDim smStationGridData(0 To 0) As String

    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    imLastStationColSorted = -1
    imLastStationSort = -1
    imLastAgreementInfoColSorted = -1
    imLastAgreementInfoSort = -1
    imLastPostedInfoColSorted = -1
    imLastPostedInfoSort = -1
    imLastSpotInfoColSorted = -1
    imLastSpotInfoSort = -1
    
    bmStationScrollAllowed = True
    lmStationTopRow = -1
    bmAgreementScrollAllowed = True
    lmAgreementTopRow = -1
    bmPostedScrollAllowed = True
    lmPostedTopRow = -1
    
    bmContactViaCmtCol = False
    mGetUserType
    mPopUserOption
    If (smUserType <> "M") And (smUserType <> "S") Then
        lbcUserOption(0).ListIndex = 3
        lbcUserOption(1).ListIndex = 1
        lacUserOption.Caption = "All Stations"
    Else
        If (smUserType = "M") And (bgMarketRepDefinedByStation = False) Then
            lbcUserOption(0).ListIndex = 3
            lbcUserOption(1).ListIndex = 1
            lacUserOption.Caption = "All Stations"
        End If
        If (smUserType = "S") And (bgServiceRepDefinedByStation = False) Then
            lbcUserOption(0).ListIndex = 3
            lbcUserOption(1).ListIndex = 1
            lacUserOption.Caption = "All Stations"
        End If
    End If
    ilRet = gPopFormats()
    ilRet = gPopStates()
    ilRet = gPopTimeZones()
    ilRet = gPopMarkets()
    ilRet = gPopMSAMarkets()
    ilRet = gPopOwnerNames()
    ilRet = gPopTeams()
    ilRet = gPopLangs()
    mPopOwnerList
    gPopMntInfo "O", tgOperatorInfo()
    gPopMntInfo "A", tgAreaInfo()
    gPopMntInfo "C", tgCityInfo()
    gPopMntInfo "Y", tgCountyInfo()
    gPopRepInfo "M", tgMarketRepInfo()
    gPopMntInfo "M", tgMonikerInfo()
    gPopRepInfo "S", tgServiceRepInfo()
    ilRet = gPopStates()
    ilRet = gPopTimeZones()
   
    mClearGrid grdStations
    mClearGrid grdAgreementInfo
    mClearGrid grdPostedInfo
    mClearGrid grdSpotInfo

    smWeek1 = Format$(gNow(), sgShowDateForm)   'sgSQLDateForm)
    smWeek1 = gObtainNextSunday(gObtainNextMonday(gObtainNextSunday(smWeek1)))
    smWeek54 = DateAdd("d", -(54 * 7) + 1, smWeek1)

    mPopListKey
    
    gBuildStationCount smWeek1, smWeek54
    
    bmInInit = False
    mSetCommands
    'If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) <> AFFILIATECRM) Then
    '    rbcComments(0).Visible = False
    '    rbcComments(1).Visible = False
    'End If
    
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mPopStationsGridSv()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilShtt As Integer
    Dim llRet As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llCell As Long
    Dim llNext As Long
    Dim llCellColor As Long
    Dim ilIncludeStation As Integer
    Dim llFilterDefIndex As Long
    Dim llNotFilterDefIndex As Long
    Dim llYellowRow As Long
    Dim slMoniker As String
    Dim ilMnt As Integer
    Dim ilNoStations As Integer
    Dim slFilter As String
    Dim ilSingle As Integer
    Dim blMatch As Boolean
    
    On Error GoTo ErrHand:
    '9/17/11: Reload Stations if required
    '         Note: Changes my this user will not cause reloading of the station table
    gPopStations
    ReDim igCommentShttCode(0 To 0) As Integer
    mBuildSingleFilter
    ilNoStations = 0
    grdStations.Rows = 2
    mClearGrid grdStations
    grdStations.Redraw = False
    grdStations.Row = 0
    For llCol = SCALLLETTERINDEX To SMCASTINDEX Step 1
        grdStations.Col = llCol
        grdStations.CellBackColor = LIGHTBLUE
    Next llCol
    grdStations.Col = SSELECTINDEX
    grdStations.CellBackColor = 16711808    'BROWN
    grdStations.CellForeColor = vbWhite
    grdStations.CellFontName = "Arial Narrow"
    llRow = grdStations.FixedRows
    For ilShtt = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        blMatch = False
        If UBound(imFilterShttCode) > LBound(imFilterShttCode) Then
            For ilSingle = 0 To UBound(imFilterShttCode) - 1 Step 1
                If tgStationInfo(ilShtt).iCode = imFilterShttCode(ilSingle) Then
                    blMatch = True
                    Exit For
                End If
            Next ilSingle
        Else
            blMatch = True
        End If
        If (tgStationInfo(ilShtt).iType = 0) And (blMatch) Then
            If (UBound(tmFilterLink) > 0) And (lacFilter.Caption = "On") Then
                For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
                    ilIncludeStation = True
                    'Test if station matches all conditions
                    'Select: 0=DMA; 1=Format; 2=MSA; 3=Owner; 4=Vehicle; 5=Zip
                    'Operator: 0=Contains; 1=Equal; 2=Greater than; 3=Less Than; 4=Not Equal
                    llFilterDefIndex = tmFilterLink(llCell).lFilterDefIndex
                    If llFilterDefIndex >= 0 Then
                        mTestFilter ilShtt, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            llNext = tmFilterLink(llCell).lNextAnd
                            Do While llNext <> -1
                                llFilterDefIndex = tmAndFilterLink(llNext).lFilterDefIndex
                                If llFilterDefIndex >= 0 Then
                                    mTestFilter ilShtt, llFilterDefIndex, ilIncludeStation
                                    If Not ilIncludeStation Then
                                        Exit Do
                                    End If
                                End If
                                llNext = tmAndFilterLink(llNext).lNextAnd
                            Loop
                        End If
                    End If
                    'Test Not array
                    If ilIncludeStation Then
                        llNotFilterDefIndex = tmFilterLink(llCell).lNotFilterDefIndex
                        If llNotFilterDefIndex >= 0 Then
                            mTestFilter ilShtt, llNotFilterDefIndex, ilIncludeStation
                            If ilIncludeStation Then
                                llNext = tmFilterLink(llCell).lNextAnd
                                Do While llNext <> -1
                                    llNotFilterDefIndex = tmAndFilterLink(llNext).lNotFilterDefIndex
                                    If llNotFilterDefIndex >= 0 Then
                                        mTestFilter ilShtt, llNotFilterDefIndex, ilIncludeStation
                                        If Not ilIncludeStation Then
                                            Exit Do
                                        End If
                                    End If
                                    llNext = tmAndFilterLink(llNext).lNextAnd
                                Loop
                            End If
                        End If
                    End If
                    
                    If ilIncludeStation Then
                        Exit For
                    End If
                Next llCell
                
            Else
                ilIncludeStation = True
            End If
            If ilIncludeStation And (sgStationSearchCallSource <> "P") Then
                If lacUserOption.Caption = "Assigned" Then
                    If smUserType = "M" Then
                        If igUstCode <> tgStationInfo(ilShtt).iMktRepUstCode Then
                            ilIncludeStation = False
                        End If
                    ElseIf smUserType = "S" Then
                        If igUstCode <> tgStationInfo(ilShtt).iServRepUstCode Then
                            ilIncludeStation = False
                        End If
                    End If
                ElseIf lacUserOption.Caption = "All Affiliates" Then
                    'SQLQuery = "SELECT attCode FROM att"
                    'SQLQuery = SQLQuery + " WHERE ("
                    'SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                    'Set rst_att = gSQLSelectCall(SQLQuery)
                    'If rst_att.EOF Then
                    '    ilIncludeStation = False
                    'End If
                    If tgStationInfo(ilShtt).sAgreementExist <> "Y" Then
                        ilIncludeStation = False
                    End If
                ElseIf lacUserOption.Caption = "Non-Affiliates" Then
                    'SQLQuery = "SELECT attCode FROM att"
                    'SQLQuery = SQLQuery + " WHERE ("
                    'SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                    'Set rst_att = gSQLSelectCall(SQLQuery)
                    'If Not rst_att.EOF Then
                    '    ilIncludeStation = False
                    'End If
                    If tgStationInfo(ilShtt).sAgreementExist = "Y" Then
                        ilIncludeStation = False
                    End If
                ElseIf lacUserOption.Caption = "All Stations" Then
                End If
            End If
            If ilIncludeStation Then
                ilNoStations = ilNoStations + 1
                igCommentShttCode(UBound(igCommentShttCode)) = tgStationInfo(ilShtt).iCode
                ReDim Preserve igCommentShttCode(0 To UBound(igCommentShttCode) + 1) As Integer
                
                If llRow >= grdStations.Rows Then
                    grdStations.AddItem ""
                End If
                grdStations.Row = llRow
                grdStations.Col = SSELECTINDEX
                grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
                For llCol = SCALLLETTERINDEX To SPWINDEX Step 1
                    grdStations.Col = llCol
                    grdStations.CellBackColor = LIGHTYELLOW
                Next llCol
                grdStations.TextMatrix(llRow, SCALLLETTERINDEX) = Trim$(tgStationInfo(ilShtt).sCallLetters)
                grdStations.TextMatrix(llRow, SDUEINDEX) = gBinarySearchStationCount(tgStationInfo(ilShtt).iCode)  'mGetDueCountForStation(tgStationInfo(ilShtt).icode)
                grdStations.TextMatrix(llRow, SDMARANKINDEX) = ""
                llRet = gBinarySearchMkt(CLng(tgStationInfo(ilShtt).iMktCode))
                If llRet <> -1 Then
                    If tgMarketInfo(llRet).iRank <> 0 Then
                        grdStations.TextMatrix(llRow, SDMARANKINDEX) = tgMarketInfo(llRet).iRank
                    End If
                End If
                grdStations.TextMatrix(llRow, SDMAMARKETINDEX) = Trim$(tgStationInfo(ilShtt).sMarket)
                grdStations.TextMatrix(llRow, SAUDP12PLUSINDEX) = Format(tgStationInfo(ilShtt).lAudP12Plus, "##,###,###")
                grdStations.TextMatrix(llRow, SWATTSINDEX) = Format(tgStationInfo(ilShtt).lWatts, "##,###,###")
                llRet = gBinarySearchMSAMkt(CLng(tgStationInfo(ilShtt).iMSAMktCode))
                If llRet <> -1 Then
                    If tgMSAMarketInfo(llRet).iRank <> 0 Then
                        grdStations.TextMatrix(llRow, SMSARANKINDEX) = tgMSAMarketInfo(llRet).iRank
                    End If
                    grdStations.TextMatrix(llRow, SMSAMARKETINDEX) = Trim$(tgMSAMarketInfo(llRet).sName)
                Else
                    grdStations.TextMatrix(llRow, SMSARANKINDEX) = ""
                    grdStations.TextMatrix(llRow, SMSAMARKETINDEX) = ""
                End If
                grdStations.TextMatrix(llRow, SSTATEINDEX) = Trim$(tgStationInfo(ilShtt).sPostalName)
                If tgStationInfo(ilShtt).iAckDaylight = 1 Then
                    grdStations.TextMatrix(llRow, SZONEINDEX) = Trim$(tgStationInfo(ilShtt).sZone) & "*"
                Else
                    grdStations.TextMatrix(llRow, SZONEINDEX) = Trim$(tgStationInfo(ilShtt).sZone)
                End If
                grdStations.TextMatrix(llRow, SZIPINDEX) = Trim$(tgStationInfo(ilShtt).sZip)
                ilRet = gBinarySearchFmt(CLng(tgStationInfo(ilShtt).iFormatCode))
                If ilRet <> -1 Then
                    grdStations.TextMatrix(llRow, SFORMATINDEX) = Trim$(tgFormatInfo(ilRet).sName)
                Else
                    grdStations.TextMatrix(llRow, SFORMATINDEX) = ""
                End If
                grdStations.TextMatrix(llRow, SOWNERINDEX) = ""
                'For ilLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
                '    If tgStationInfo(ilShtt).lOwnerCode = tgOwnerInfo(ilLoop).lCode Then
                '        grdStations.TextMatrix(llRow, SOWNERINDEX) = Trim$(tgOwnerInfo(ilLoop).sName)
                '        Exit For
                '    End If
                'Next ilLoop
                llRet = mBinarySearchOwner(tgStationInfo(ilShtt).lOwnerCode)
                If llRet <> -1 Then
                    grdStations.TextMatrix(llRow, SOWNERINDEX) = Trim$(tgOwnerInfo(llRet).sName)
                End If
                grdStations.TextMatrix(llRow, SOPERATORINDEX) = ""
                grdStations.TextMatrix(llRow, SOPERATORNAMEINDEX) = ""
                If tgStationInfo(ilShtt).lOperatorMntCode > 0 Then
                    grdStations.TextMatrix(llRow, SOPERATORINDEX) = "*"
                    ilMnt = gBinarySearchMnt(tgStationInfo(ilShtt).lOperatorMntCode, tgOperatorInfo())
                    If ilMnt <> -1 Then
                        grdStations.TextMatrix(llRow, SOPERATORNAMEINDEX) = UCase$(Trim$(tgOperatorInfo(ilMnt).sName))
                    End If
                End If
                grdStations.TextMatrix(llRow, SCLUSTERINDEX) = ""
                If tgStationInfo(ilShtt).lMarketClusterGroupID <= 0 Then
                    grdStations.TextMatrix(llRow, SCLUSTERINDEX) = ""
                Else
                    grdStations.TextMatrix(llRow, SCLUSTERINDEX) = "*"
                    SQLQuery = "SELECT shttCallLetters FROM shtt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " shttClusterGroupID = " & tgStationInfo(ilShtt).lMarketClusterGroupID & ")"
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    Do While Not rst_Shtt.EOF
                        grdStations.TextMatrix(llRow, SCLUSTERNAMESINDEX) = grdStations.TextMatrix(llRow, SCLUSTERNAMESINDEX) & " " & Trim$(rst_Shtt!shttCallLetters)
                        rst_Shtt.MoveNext
                    Loop
                    grdStations.TextMatrix(llRow, SCLUSTERNAMESINDEX) = Trim$(grdStations.TextMatrix(llRow, SCLUSTERNAMESINDEX))
                End If
                
                If Trim$(tgStationInfo(ilShtt).sWebPW) <> "" Then
                    grdStations.TextMatrix(llRow, SPASSWORDINDEX) = Trim$(tgStationInfo(ilShtt).sWebPW)
                    grdStations.TextMatrix(llRow, SPWINDEX) = "*"
                Else
                    grdStations.TextMatrix(llRow, SPASSWORDINDEX) = ""
                    grdStations.TextMatrix(llRow, SPWINDEX) = ""
                End If
                
                'SQLQuery = "SELECT mgtCode FROM mgt"
                'SQLQuery = SQLQuery + " WHERE ("
                'SQLQuery = SQLQuery & " mgtShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                'Set rst_mgt = gSQLSelectCall(SQLQuery)
                'If rst_mgt.EOF Then
                If tgStationInfo(ilShtt).lMultiCastGroupID <= 0 Then
                    grdStations.TextMatrix(llRow, SMCASTINDEX) = ""
                Else
                    grdStations.TextMatrix(llRow, SMCASTINDEX) = "*"
                    SQLQuery = "SELECT shttCallLetters FROM shtt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " shttMultiCastGroupID = " & tgStationInfo(ilShtt).lMultiCastGroupID & ")"
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    Do While Not rst_Shtt.EOF
                        grdStations.TextMatrix(llRow, SMCASTNAMESINDEX) = grdStations.TextMatrix(llRow, SMCASTNAMESINDEX) & " " & Trim$(rst_Shtt!shttCallLetters)
                        rst_Shtt.MoveNext
                    Loop
                    grdStations.TextMatrix(llRow, SMCASTNAMESINDEX) = Trim$(grdStations.TextMatrix(llRow, SMCASTNAMESINDEX))
                End If
                grdStations.Col = SWEBSITEINDEX
                If Trim$(tgStationInfo(ilShtt).sWebAddress) = "" Then
                    grdStations.CellBackColor = GRAY    'vbWhite
                    grdStations.TextMatrix(llRow, SWEBSITEINDEX) = ""
                    grdStations.TextMatrix(llRow, SWEBADDRESSINDEX) = ""
                Else
                    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                    grdStations.TextMatrix(llRow, SWEBSITEINDEX) = "W"
                    grdStations.TextMatrix(llRow, SWEBADDRESSINDEX) = Trim$(tgStationInfo(ilShtt).sWebAddress)
                End If
                grdStations.Col = SCALLLETTERINDEX
                grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                
                grdStations.Col = SAGRMNTINDEX
                grdStations.CellAlignment = flexAlignCenterCenter
                'SQLQuery = "SELECT attCode FROM att"
                'SQLQuery = SQLQuery + " WHERE ("
                'SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                'Set rst_att = gSQLSelectCall(SQLQuery)
                'If Not rst_att.EOF Then
                If tgStationInfo(ilShtt).sAgreementExist = "Y" Then
                    grdStations.TextMatrix(llRow, SAGRMNTINDEX) = "A"
                    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                    grdStations.Col = SSELECTINDEX
                    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
                    grdStations.CellFontName = "Monotype Sorts"
                    grdStations.TextMatrix(llRow, SSELECTINDEX) = "t"
                Else
                    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY    'vbWhite
                    grdStations.Col = SSELECTINDEX
                    grdStations.CellBackColor = GRAY    'vbWhite
                    grdStations.TextMatrix(llRow, SSELECTINDEX) = ""
                End If
                
                grdStations.Col = SCOMMENTINDEX
                grdStations.CellAlignment = flexAlignCenterCenter
                'SQLQuery = "SELECT cctCode FROM cct"
                'SQLQuery = SQLQuery + " WHERE ("
                'SQLQuery = SQLQuery & " cctShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                'Set rst_cct = gSQLSelectCall(SQLQuery)
                'If Not rst_cct.EOF Then
                If tgStationInfo(ilShtt).sCommentExist = "Y" Then
                    grdStations.TextMatrix(llRow, SCOMMENTINDEX) = "C"
                    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                Else
                    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY    'vbWhite
                End If
                            
                grdStations.TextMatrix(llRow, SSHTTCODEINDEX) = tgStationInfo(ilShtt).iCode
                slMoniker = ""
                If tgStationInfo(ilShtt).lMonikerMntCode > 0 Then
                    SQLQuery = "SELECT mntName FROM mnt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " mntCode = " & tgStationInfo(ilShtt).lMonikerMntCode & ")"
                    Set rst_mnt = gSQLSelectCall(SQLQuery)
                    If Not rst_mnt.EOF Then
                        slMoniker = Trim$(rst_mnt!mntName)
                    End If
                End If
                grdStations.TextMatrix(llRow, SFREQMONIKERINDEX) = Trim$(Trim$(tgStationInfo(ilShtt).sFrequency) & " " & slMoniker)
                llRow = llRow + 1
            End If
        End If
    Next ilShtt
    If llRow > grdStations.FixedRows Then
        lmStationMaxRow = llRow - 1
    Else
        lmStationMaxRow = llRow
    End If
    'If grdStations.Rows < grdStations.Height \ grdStations.RowHeight(1) Then
    '    grdStations.Rows = grdStations.Rows + 2 * (grdStations.Height \ grdStations.RowHeight(1)) - grdStations.Rows + 1
    'Else
    '    grdStations.Rows = grdStations.Rows + grdStations.Height \ grdStations.RowHeight(1) + 1
    'End If
    grdStations.Rows = grdStations.Rows + ((cmcDone.Top - grdStations.Top) \ grdStations.RowHeight(1))
    For llYellowRow = llRow To grdStations.Rows - 1 Step 1
        grdStations.Row = llYellowRow
        For llCol = SSELECTINDEX To SCOMMENTINDEX Step 1
            grdStations.Col = llCol
            grdStations.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llYellowRow
    imLastStationSort = -1
    imLastStationColSorted = -1
    mStationSortCol SCALLLETTERINDEX
    grdStations.Row = 0
    grdStations.Col = SSHTTCODEINDEX
    lmRowSelected = -1
    If sgFilterName = "" Then
        If igFilterChgd Then
            slFilter = "Filter: Custom"
        Else
            slFilter = "Filter: None"
        End If
    ElseIf igFilterChgd Then
        slFilter = "Filter: " & sgFilterName & " modified"
    Else
        slFilter = "Filter: " & sgFilterName
    End If
    If sgStationSearchCallSource = "P" Then
        frmStationSearch.Caption = "Post-Buy Planning- " & Trim$(sgUserName) & ", " & ilNoStations & " Stations" & " " & slFilter
    Else
        frmStationSearch.Caption = "Management- " & Trim$(sgUserName) & ", " & ilNoStations & " Stations" & " " & slFilter
    End If
    grdStations.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopStationsGrid"
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdPostBuy.Width = 0.6 * grdStations.Width
    grdPostBuy.ColWidth(PSSTATIONTYPEINDEX) = grdPostBuy.Width * 0.15
    grdPostBuy.ColWidth(PSWEEKOFINDEX) = grdPostBuy.Width * 0.2
    grdPostBuy.ColWidth(PSCONTRACTINDEX) = grdPostBuy.Width * 0.2
    grdPostBuy.ColWidth(PSADVERTISERINDEX) = grdPostBuy.Width - grdPostBuy.ColWidth(PSSTATIONTYPEINDEX) - grdPostBuy.ColWidth(PSWEEKOFINDEX) - grdPostBuy.ColWidth(PSCONTRACTINDEX) - 15
    gGrid_IntegralHeight grdPostBuy
    grdPostBuy.Height = 2 * grdPostBuy.RowHeight(0) + 15
    cmcGen.Top = grdPostBuy.Top
    cmcGen.Left = grdPostBuy.Left + grdPostBuy.Width + 30
    
    grdStations.Width = frmStationSearch.Width - 300 - GRIDSCROLLWIDTH
    grdStations.Height = cmcDone.Top - 180
    If sgStationSearchCallSource = "P" Then
        grdStations.Height = grdStations.Height - 2 * grdStations.RowHeight(0)
    End If
    gGrid_IntegralHeight grdStations
    grdStations.Height = grdStations.Height + 30
    gGrid_FillWithRows grdStations
    grdStations.Move (frmStationSearch.Width - grdStations.Width) \ 2, (cmcDone.Top - grdStations.Height) \ 2
    If sgStationSearchCallSource = "P" Then
        grdPostBuy.Move grdStations.Left, 60
        grdStations.Top = grdPostBuy.Top + grdPostBuy.Height + 60
    End If
    vbcStation.Height = grdStations.Height
    
    grdStations.ColWidth(SSHTTCODEINDEX) = 0
    grdStations.ColWidth(SSORTINDEX) = 0
    grdStations.ColWidth(SFREQMONIKERINDEX) = 0
    grdStations.ColWidth(SCLUSTERNAMESINDEX) = 0
    grdStations.ColWidth(SMCASTNAMESINDEX) = 0
    grdStations.ColWidth(SPASSWORDINDEX) = 0
    grdStations.ColWidth(SAUDP12PLUSINDEX) = 0
    grdStations.ColWidth(SWATTSINDEX) = 0
    grdStations.ColWidth(SWEBADDRESSINDEX) = 0
    grdStations.ColWidth(SOPERATORNAMEINDEX) = 0
    grdStations.ColWidth(SSELECTINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SCALLLETTERINDEX) = grdStations.Width * 0.08
    grdStations.ColWidth(SDUEINDEX) = grdStations.Width * 0.03
    grdStations.ColWidth(SDMARANKINDEX) = grdStations.Width * 0.035
    grdStations.ColWidth(SDMAMARKETINDEX) = grdStations.Width * 0.07
    grdStations.ColWidth(SMSARANKINDEX) = grdStations.Width * 0.035
    grdStations.ColWidth(SMSAMARKETINDEX) = grdStations.Width * 0.07
    grdStations.ColWidth(SSTATEINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SZONEINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SZIPINDEX) = grdStations.Width * 0.05
    grdStations.ColWidth(SFORMATINDEX) = grdStations.Width * 0.08
    grdStations.ColWidth(SOWNERINDEX) = grdStations.Width * 0.1
    grdStations.ColWidth(SOPERATORINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SCLUSTERINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SMCASTINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SPWINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SWEBSITEINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SAGRMNTINDEX) = grdStations.Width * 0.04
    grdStations.ColWidth(SCOMMENTINDEX) = grdStations.Width * 0.04
   
    If gIsUsingNovelty Then
        grdStations.ColWidth(SPWINDEX) = 0
    End If
    
    
    grdStations.ColWidth(SDMAMARKETINDEX) = grdStations.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To SCOMMENTINDEX Step 1
        If ilCol <> SDMAMARKETINDEX Then
            grdStations.ColWidth(SDMAMARKETINDEX) = grdStations.ColWidth(SDMAMARKETINDEX) - grdStations.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStations
    grdStations.ColAlignment(SOPERATORINDEX) = flexAlignCenterCenter
    grdStations.ColAlignment(SCLUSTERINDEX) = flexAlignCenterCenter
    grdStations.ColAlignment(SMCASTINDEX) = flexAlignCenterCenter
    grdStations.ColAlignment(SPWINDEX) = flexAlignCenterCenter
    grdStations.ColAlignment(SWEBSITEINDEX) = flexAlignCenterCenter
    grdStations.ColAlignment(SDUEINDEX) = flexAlignRightCenter
    grdStations.ColAlignment(SDMARANKINDEX) = flexAlignRightCenter
    grdStations.ColAlignment(SMSARANKINDEX) = flexAlignRightCenter
    
    vbcStation.Move grdStations.Left + grdStations.Width - vbcStation.Width, grdStations.Top, GRIDSCROLLWIDTH, grdStations.Height
    
    udcCommentGrid.Width = grdStations.Width + pbcArrow.Width
    udcCommentGrid.Left = grdStations.Left - pbcArrow.Width
    udcCommentGrid.Height = grdStations.Height - grdStations.RowHeight(0) - grdStations.RowHeight(1)
    
    
    udcContactGrid.Width = grdStations.Width + pbcArrow.Width
    udcContactGrid.Left = grdStations.Left - pbcArrow.Width
    udcContactGrid.Height = grdStations.Height - grdStations.RowHeight(0) - grdStations.RowHeight(1)
    
    grdAgreementInfo.Width = grdStations.Width
    grdAgreementInfo.Height = grdStations.Height - grdStations.RowHeight(0) - grdStations.RowHeight(1)
    grdAgreementInfo.Move grdStations.Left, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1)
    grdAgreementInfo.ColWidth(ATIMERANGEINDEX) = 0
    grdAgreementInfo.ColWidth(AATTCODEINDEX) = 0
    grdAgreementInfo.ColWidth(ASORTINDEX) = 0
    grdAgreementInfo.ColWidth(ASELECTINDEX) = grdAgreementInfo.Width * 0.04
    grdAgreementInfo.ColWidth(AVEHICLEINDEX) = grdAgreementInfo.Width * 0.12
    grdAgreementInfo.ColWidth(ADUEINDEX) = grdAgreementInfo.Width * 0.03
    grdAgreementInfo.ColWidth(ADAYTIMEINDEX) = grdAgreementInfo.Width * 0.09
    grdAgreementInfo.Row = 0
    For ilCol = AWEEK1INDEX To AWEEK1INDEX + 53 Step 1
        grdAgreementInfo.ColWidth(ilCol) = grdAgreementInfo.Width * 0.0125
        grdAgreementInfo.Col = ilCol
        grdAgreementInfo.CellFontBold = False
    Next ilCol
    grdAgreementInfo.ColWidth(AAGRMNTINDEX) = grdAgreementInfo.Width * 0.04
   
    
    
    grdAgreementInfo.ColWidth(AVEHICLEINDEX) = grdAgreementInfo.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To AAGRMNTINDEX Step 1
        If ilCol <> AVEHICLEINDEX Then
            grdAgreementInfo.ColWidth(AVEHICLEINDEX) = grdAgreementInfo.ColWidth(AVEHICLEINDEX) - grdAgreementInfo.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdAgreementInfo
    grdAgreementInfo.ColAlignment(ADUEINDEX) = flexAlignRightCenter
    
    
    grdPostedInfo.Width = grdStations.Width
    grdPostedInfo.Height = grdStations.Height - grdStations.RowHeight(0) - grdStations.RowHeight(1)
    grdPostedInfo.Move grdStations.Left, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1)
    grdPostedInfo.ColWidth(PCPTTINDEX) = 0
    grdPostedInfo.ColWidth(PSORTINDEX) = 0
    grdPostedInfo.ColWidth(PSELECTINDEX) = grdPostedInfo.Width * 0.04
    grdPostedInfo.ColWidth(PWEEKINDEX) = grdPostedInfo.Width * 0.08
    grdPostedInfo.ColWidth(PNOSCHDINDEX) = grdPostedInfo.Width * 0.05
    grdPostedInfo.ColWidth(PNOAIREDINDEX) = grdPostedInfo.Width * 0.05
    grdPostedInfo.ColWidth(PNOCMPLINDEX) = grdPostedInfo.Width * 0.06
    grdPostedInfo.ColWidth(PPERCENTAINDEX) = grdPostedInfo.Width * 0.05
    grdPostedInfo.ColWidth(PPERCENTCINDEX) = grdPostedInfo.Width * 0.05
    grdPostedInfo.ColWidth(PPOSTDATEINDEX) = grdPostedInfo.Width * 0.08
    grdPostedInfo.ColWidth(PBYINDEX) = grdPostedInfo.Width * 0.12
    grdPostedInfo.ColWidth(PIPINDEX) = grdPostedInfo.Width * 0.12
    grdPostedInfo.ColWidth(PSTATUSINDEX) = grdPostedInfo.Width * 0.04
   
    
    
    grdPostedInfo.ColWidth(PBYINDEX) = grdPostedInfo.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To PSTATUSINDEX Step 1
        If ilCol <> PBYINDEX Then
            grdPostedInfo.ColWidth(PBYINDEX) = grdPostedInfo.ColWidth(PBYINDEX) - grdPostedInfo.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdPostedInfo
    grdPostedInfo.ColAlignment(PWEEKINDEX) = flexAlignRightCenter
    grdPostedInfo.ColAlignment(PNOSCHDINDEX) = flexAlignRightCenter
    grdPostedInfo.ColAlignment(PNOAIREDINDEX) = flexAlignRightCenter
    grdPostedInfo.ColAlignment(PNOCMPLINDEX) = flexAlignRightCenter
    grdPostedInfo.ColAlignment(PPERCENTAINDEX) = flexAlignRightCenter
    grdPostedInfo.ColAlignment(PPERCENTCINDEX) = flexAlignRightCenter
    grdPostedInfo.ColAlignment(PPOSTDATEINDEX) = flexAlignRightCenter
    
    
    grdSpotInfo.Width = grdStations.Width
    grdSpotInfo.Height = grdStations.Height - grdStations.RowHeight(0) - grdStations.RowHeight(1)
    grdSpotInfo.Move grdStations.Left, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1)
    grdSpotInfo.ColWidth(DGAMEINFOINDEX) = 0
    grdSpotInfo.ColWidth(DMMRINFOINDEX) = 0
    grdSpotInfo.ColWidth(DCNTRNOINDEX) = 0
    grdSpotInfo.ColWidth(DASTCODEINDEX) = 0
    grdSpotInfo.ColWidth(DSORTINDEX) = 0
    grdSpotInfo.ColWidth(DDATEINDEX) = grdSpotInfo.Width * 0.06
    grdSpotInfo.ColWidth(DFEDINDEX) = grdSpotInfo.Width * 0.08
    grdSpotInfo.ColWidth(DPLEGDEDDAYINDEX) = grdSpotInfo.Width * 0.08
    grdSpotInfo.ColWidth(DPLEGDEDTIMEINDEX) = grdSpotInfo.Width * 0.12
    grdSpotInfo.ColWidth(DAIREDDATEINDEX) = grdSpotInfo.Width * 0.07
    grdSpotInfo.ColWidth(DAIREDTIMEINDEX) = grdSpotInfo.Width * 0.08
    grdSpotInfo.ColWidth(DADVTINDEX) = grdSpotInfo.Width * 0.08
    grdSpotInfo.ColWidth(DPRODINDEX) = grdSpotInfo.Width * 0.08
    grdSpotInfo.ColWidth(DLENGTHINDEX) = grdSpotInfo.Width * 0.03
    grdSpotInfo.ColWidth(DISCIINDEX) = grdSpotInfo.Width * 0.1
    grdSpotInfo.ColWidth(DCARTINDEX) = grdSpotInfo.Width * 0.06
    grdSpotInfo.ColWidth(DSTATUSINDEX) = grdSpotInfo.Width * 0.04
   
    
    
    grdSpotInfo.ColWidth(DCOMMENTINDEX) = grdSpotInfo.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To DSTATUSINDEX Step 1
        If ilCol <> DCOMMENTINDEX Then
            grdSpotInfo.ColWidth(DCOMMENTINDEX) = grdSpotInfo.ColWidth(DCOMMENTINDEX) - grdSpotInfo.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdSpotInfo
    grdSpotInfo.ColAlignment(DDATEINDEX) = flexAlignRightCenter
    grdSpotInfo.ColAlignment(DFEDINDEX) = flexAlignRightCenter
    grdSpotInfo.ColAlignment(DAIREDDATEINDEX) = flexAlignRightCenter
    grdSpotInfo.ColAlignment(DAIREDTIMEINDEX) = flexAlignRightCenter
    
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    Dim ilWk As Integer
    Dim ilNumber As Integer
    
    grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Veh"
    grdStations.TextMatrix(0, SCALLLETTERINDEX) = "Call Letters"
    grdStations.TextMatrix(0, SDUEINDEX) = "DUE"
    grdStations.TextMatrix(0, SDMARANKINDEX) = "D#"
    grdStations.TextMatrix(0, SDMAMARKETINDEX) = "DMA"
    grdStations.TextMatrix(0, SMSARANKINDEX) = "M#"
    grdStations.TextMatrix(0, SMSAMARKETINDEX) = "MSA"
    grdStations.TextMatrix(0, SSTATEINDEX) = "State"
    grdStations.TextMatrix(0, SZONEINDEX) = "Zone"
    grdStations.TextMatrix(0, SZIPINDEX) = "Zip"
    grdStations.TextMatrix(0, SFORMATINDEX) = "Format"
    grdStations.TextMatrix(0, SOWNERINDEX) = "Owner"
    grdStations.TextMatrix(0, SOPERATORINDEX) = "Opr"
    grdStations.TextMatrix(0, SCLUSTERINDEX) = "Sist"
    grdStations.TextMatrix(0, SMCASTINDEX) = "Cast"
    grdStations.TextMatrix(0, SPWINDEX) = "P/W"
    grdStations.TextMatrix(0, SWEBSITEINDEX) = "Web"
    grdStations.TextMatrix(0, SAGRMNTINDEX) = "Agr"
    grdStations.TextMatrix(0, SCOMMENTINDEX) = "Cmt"

    grdPostBuy.TextMatrix(0, PSSTATIONTYPEINDEX) = "Stations"    '"Veh"
    grdPostBuy.TextMatrix(0, PSWEEKOFINDEX) = "Week Of"    '"Veh"
    grdPostBuy.TextMatrix(0, PSADVERTISERINDEX) = "Advertiser"
    grdPostBuy.TextMatrix(0, PSCONTRACTINDEX) = "Contract"


    grdAgreementInfo.TextMatrix(0, ASELECTINDEX) = "Wks"
    grdAgreementInfo.TextMatrix(0, AVEHICLEINDEX) = "Vehicle"
    grdAgreementInfo.TextMatrix(0, ADUEINDEX) = "Due"
    grdAgreementInfo.TextMatrix(0, ADAYTIMEINDEX) = "Date Range" '"Day/Time"
    ilNumber = -2
    For ilWk = 1 To 54 Step 1
        grdAgreementInfo.Row = 0
        grdAgreementInfo.Col = AWEEK1INDEX + ilWk - 1
        grdAgreementInfo.CellFontName = "Arial Narrow"
        grdAgreementInfo.CellFontBold = False
        grdAgreementInfo.CellFontSize = 8
        grdAgreementInfo.CellAlignment = flexAlignCenterCenter
        If ilNumber = -2 Then
            grdAgreementInfo.TextMatrix(0, AWEEK1INDEX + ilWk - 1) = "N"
        ElseIf ilNumber = -1 Then
            grdAgreementInfo.TextMatrix(0, AWEEK1INDEX + ilWk - 1) = "C"
            ilNumber = 0
        ElseIf ilNumber = 0 Then
            grdAgreementInfo.TextMatrix(0, AWEEK1INDEX + ilWk - 1) = ""
        Else
            grdAgreementInfo.TextMatrix(0, AWEEK1INDEX + ilWk - 1) = ilNumber
        End If
        ilNumber = ilNumber + 1
        If ilNumber = 10 Then
            ilNumber = 0
        End If
    Next ilWk
    grdAgreementInfo.TextMatrix(0, AAGRMNTINDEX) = "Agr"

    grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Spts"
    grdPostedInfo.TextMatrix(0, PWEEKINDEX) = "W/O"
    grdPostedInfo.TextMatrix(0, PNOSCHDINDEX) = "# Sch"
    grdPostedInfo.TextMatrix(0, PNOAIREDINDEX) = "# Air"
    grdPostedInfo.TextMatrix(0, PNOCMPLINDEX) = "# Cmpl"
    grdPostedInfo.TextMatrix(0, PPERCENTAINDEX) = "%A"
    grdPostedInfo.TextMatrix(0, PPERCENTCINDEX) = "%C"
    grdPostedInfo.TextMatrix(0, PPOSTDATEINDEX) = "Post Date"
    grdPostedInfo.TextMatrix(0, PBYINDEX) = "By"
    grdPostedInfo.TextMatrix(0, PIPINDEX) = "IP Address"
    grdPostedInfo.TextMatrix(0, PSTATUSINDEX) = "Status"

    grdSpotInfo.TextMatrix(0, DDATEINDEX) = "Feed Date"
    grdSpotInfo.TextMatrix(0, DFEDINDEX) = "Feed Time"
    grdSpotInfo.TextMatrix(0, DPLEGDEDDAYINDEX) = "Pledge Days"
    grdSpotInfo.TextMatrix(0, DPLEGDEDTIMEINDEX) = "Pledge Times"
    grdSpotInfo.TextMatrix(0, DAIREDDATEINDEX) = "Air Date"
    grdSpotInfo.TextMatrix(0, DAIREDTIMEINDEX) = "Aired Time"
    grdSpotInfo.TextMatrix(0, DADVTINDEX) = "Advt"
    grdSpotInfo.TextMatrix(0, DPRODINDEX) = "Prod"
    grdSpotInfo.TextMatrix(0, DLENGTHINDEX) = "Len"
    grdSpotInfo.TextMatrix(0, DISCIINDEX) = "ISCI"
    grdSpotInfo.TextMatrix(0, DCARTINDEX) = "Cart"
    grdSpotInfo.TextMatrix(0, DCOMMENTINDEX) = "Comment"
    grdSpotInfo.TextMatrix(0, DSTATUSINDEX) = "Status"

End Sub

Private Sub mStationSortColSv(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdStations.FixedRows To grdStations.Rows - 1 Step 1
        slStr = Trim$(grdStations.TextMatrix(llRow, SCALLLETTERINDEX))
        If slStr <> "" Then
            If (ilCol = SDMARANKINDEX) Or (ilCol = SMSARANKINDEX) Then
                slSort = UCase$(Trim$(grdStations.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = "999"
                End If
                Do While Len(slSort) < 3
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = SDUEINDEX Then
                slSort = UCase$(Trim$(grdStations.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    If (ilCol = imLastStationColSorted) Then
                        If imLastStationSort = flexSortStringNoCaseAscending Then
                            slSort = "000"
                        Else
                            slSort = "999"
                        End If
                    Else
                        If (ilCol <> imLastStationColSorted) Or (imLastStationSort = -1) Then
                            slSort = "000"
                        Else
                            slSort = "999"
                        End If
                    End If
                End If
                
                Do While Len(slSort) < 3
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdStations.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slStr = grdStations.TextMatrix(llRow, SSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastStationColSorted) Or ((ilCol = imLastStationColSorted) And (imLastStationSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 5
                    slRow = "0" & slRow
                Loop
                grdStations.TextMatrix(llRow, SSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 5
                    slRow = "0" & slRow
                Loop
                grdStations.TextMatrix(llRow, SSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If (ilCol = SDUEINDEX) And ((ilCol <> imLastStationColSorted) Or (imLastStationSort = -1)) Then
        imLastStationSort = flexSortStringNoCaseAscending
        imLastStationColSorted = ilCol
    End If
    If ilCol = imLastStationColSorted Then
        imLastStationColSorted = SSORTINDEX
    Else
        imLastStationColSorted = -1
        imLastStationColSorted = -1
    End If
    gGrid_SortByCol grdStations, SCALLLETTERINDEX, SSORTINDEX, imLastStationColSorted, imLastStationSort
    imLastStationColSorted = ilCol
    mSetCommands
End Sub
Private Sub mAgreementInfoSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdAgreementInfo.FixedRows To grdAgreementInfo.Rows - 1 Step 1
        slStr = Trim$(grdAgreementInfo.TextMatrix(llRow, AVEHICLEINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdAgreementInfo.TextMatrix(llRow, ilCol)))
            If slSort = "" Then
                slSort = Chr(32)
            End If
            slStr = grdAgreementInfo.TextMatrix(llRow, SSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastAgreementInfoColSorted) Or ((ilCol = imLastAgreementInfoColSorted) And (imLastAgreementInfoSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAgreementInfo.TextMatrix(llRow, ASORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdAgreementInfo.TextMatrix(llRow, ASORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastAgreementInfoColSorted Then
        imLastAgreementInfoColSorted = ASORTINDEX
    Else
        imLastAgreementInfoColSorted = -1
        imLastAgreementInfoColSorted = -1
    End If
    gGrid_SortByCol grdAgreementInfo, AVEHICLEINDEX, ASORTINDEX, imLastAgreementInfoColSorted, imLastAgreementInfoSort
    imLastAgreementInfoColSorted = ilCol
End Sub

Private Sub mPostedInfoSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdPostedInfo.FixedRows To grdPostedInfo.Rows - 1 Step 1
        slStr = Trim$(grdPostedInfo.TextMatrix(llRow, AVEHICLEINDEX))
        If slStr <> "" Then
            If ilCol = PWEEKINDEX Then
                slSort = Trim$(Str$(gDateValue(grdPostedInfo.TextMatrix(llRow, PWEEKINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = PPOSTDATEINDEX Then
                slSort = Trim$(Str$(gDateValue(grdPostedInfo.TextMatrix(llRow, PPOSTDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdPostedInfo.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slStr = grdPostedInfo.TextMatrix(llRow, SSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastPostedInfoColSorted) Or ((ilCol = imLastPostedInfoColSorted) And (imLastPostedInfoSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPostedInfo.TextMatrix(llRow, SSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPostedInfo.TextMatrix(llRow, SSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastPostedInfoColSorted Then
        imLastPostedInfoColSorted = SSORTINDEX
    Else
        imLastPostedInfoColSorted = -1
        imLastPostedInfoColSorted = -1
    End If
    gGrid_SortByCol grdPostedInfo, SCALLLETTERINDEX, SSORTINDEX, imLastPostedInfoColSorted, imLastPostedInfoSort
    imLastPostedInfoColSorted = ilCol
End Sub

Private Sub mSpotInfoSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdSpotInfo.FixedRows To grdSpotInfo.Rows - 1 Step 1
        slStr = Trim$(grdSpotInfo.TextMatrix(llRow, DDATEINDEX))
        If slStr <> "" Then
            If ilCol = DDATEINDEX Then
                slSort = Trim$(Str$(gDateValue(grdSpotInfo.TextMatrix(llRow, DDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
                slStr = Trim$(Str$(gTimeToLong(grdSpotInfo.TextMatrix(llRow, DFEDINDEX), False)))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
                slSort = slSort & slStr
            ElseIf (ilCol = DAIREDDATEINDEX) Then
                slSort = Trim$(Str$(gDateValue(grdSpotInfo.TextMatrix(llRow, DAIREDDATEINDEX))))
                If InStr(1, slSort, "MG:", vbTextCompare) > 0 Then
                    slSort = Trim$(Mid$(slSort, 3))
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = DAIREDTIMEINDEX) Then
                slSort = Trim$(Str$(gTimeToLong(grdSpotInfo.TextMatrix(llRow, DAIREDTIMEINDEX), False)))
                If InStr(1, slSort, "@", vbTextCompare) > 0 Then
                    slSort = Trim$(Mid$(slSort, 1))
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdSpotInfo.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slStr = grdSpotInfo.TextMatrix(llRow, DSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastSpotInfoColSorted) Or ((ilCol = imLastSpotInfoColSorted) And (imLastSpotInfoSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdSpotInfo.TextMatrix(llRow, DSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdSpotInfo.TextMatrix(llRow, DSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastSpotInfoColSorted Then
        imLastSpotInfoColSorted = DSORTINDEX
    Else
        imLastSpotInfoColSorted = -1
        imLastSpotInfoColSorted = -1
    End If
    gGrid_SortByCol grdSpotInfo, DDATEINDEX, DSORTINDEX, imLastSpotInfoColSorted, imLastSpotInfoSort
    imLastSpotInfoColSorted = ilCol
End Sub


Private Sub grdAgreementInfo_GotFocus()
    mSetShowCommentsAndContacts
End Sub

Private Sub grdAgreementInfo_Scroll()
    'If Not bmAgreementScrollAllowed Then
    '    grdAgreementInfo.TopRow = lmAgreementTopRow
    'End If
    Dim slVehicleName As String
    Dim llAttCode As Long
    If bmInScroll Then
        Exit Sub
    End If
'    If Not bmAgreementScrollAllowed Then
'        bmInScroll = True
'        If grdAgreementInfo.TopRow > lmAgreementTopRow Then
'            If grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, AATTCODEINDEX) = "" Then
'                grdAgreementInfo.TopRow = lmAgreementTopRow
'            Else
'                grdAgreementInfo.TopRow = lmAgreementTopRow + 1
'                If bmAgreementScrollAllowed Then
'                    slVehicleName = ""
'                    llAttCode = 0
'                Else
'                    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, AVEHICLEINDEX))
'                    llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, AATTCODEINDEX)))
'                End If
'                grdAgreementInfo.TextMatrix(lmAgreementTopRow, ASELECTINDEX) = "t"
'                grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, ASELECTINDEX) = "s"
'                mMousePointer vbHourglass
'                mRepopGrids slVehicleName, llAttCode, 0
'                mSetGridPosition
'                mMousePointer vbDefault
'            End If
'        ElseIf grdAgreementInfo.TopRow < lmAgreementTopRow Then
'            If lmAgreementTopRow - 1 < grdAgreementInfo.FixedRows Then
'                grdAgreementInfo.TopRow = lmAgreementTopRow
'            Else
'                grdAgreementInfo.TopRow = lmAgreementTopRow - 1
'                If bmAgreementScrollAllowed Then
'                    slVehicleName = ""
'                    llAttCode = 0
'                Else
'                    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow - 1, AVEHICLEINDEX))
'                    llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow - 1, AATTCODEINDEX)))
'                End If
'                grdAgreementInfo.TextMatrix(lmAgreementTopRow, ASELECTINDEX) = "t"
'                grdAgreementInfo.TextMatrix(lmAgreementTopRow - 1, ASELECTINDEX) = "s"
'                mMousePointer vbHourglass
'                mRepopGrids slVehicleName, llAttCode, 0
'                mSetGridPosition
'                mMousePointer vbDefault
'            End If
'        End If
'        bmInScroll = False
'    End If
    If Not bmAgreementScrollAllowed Then
        bmInScroll = True
        If grdAgreementInfo.TopRow > lmAgreementTopRow Then
            If grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, AATTCODEINDEX) = "" Then
                grdAgreementInfo.TopRow = lmAgreementTopRow
            Else
                grdAgreementInfo.TopRow = lmAgreementTopRow + 1
                If grdAgreementInfo.TextMatrix(lmAgreementTopRow, ASELECTINDEX) <> "" Then
                    grdAgreementInfo.TextMatrix(lmAgreementTopRow, ASELECTINDEX) = "t"
                End If
                If bmAgreementScrollAllowed Then
                    slVehicleName = ""
                    llAttCode = 0
                    If grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, ASELECTINDEX) <> "" Then
                        grdAgreementInfo.TextMatrix(lmAgreementTopRow + 1, ASELECTINDEX) = "s"
                    End If
                    mMousePointer vbHourglass
                    mRepopGrids slVehicleName, llAttCode, 0
                    mSetGridPosition
                    mMousePointer vbDefault
                Else
                    imDelaySource = 1
                    tmcDelay.Enabled = True
                End If
            End If
        ElseIf grdAgreementInfo.TopRow < lmAgreementTopRow Then
            If lmAgreementTopRow - 1 < grdAgreementInfo.FixedRows Then
                grdAgreementInfo.TopRow = lmAgreementTopRow
            Else
                grdAgreementInfo.TopRow = lmAgreementTopRow - 1
                If grdAgreementInfo.TextMatrix(lmAgreementTopRow, ASELECTINDEX) <> "" Then
                    grdAgreementInfo.TextMatrix(lmAgreementTopRow, ASELECTINDEX) = "t"
                End If
                If bmAgreementScrollAllowed Then
                    slVehicleName = ""
                    llAttCode = 0
                    If grdAgreementInfo.TextMatrix(lmAgreementTopRow - 1, ASELECTINDEX) <> "" Then
                        grdAgreementInfo.TextMatrix(lmAgreementTopRow - 1, ASELECTINDEX) = "s"
                    End If
                    mMousePointer vbHourglass
                    mRepopGrids slVehicleName, llAttCode, 0
                    mSetGridPosition
                    mMousePointer vbDefault
                Else
                    imDelaySource = 1
                    tmcDelay.Enabled = True
                End If
            End If
        End If
        bmInScroll = False
    End If
End Sub


Private Sub grdPostedInfo_GotFocus()
    mSetShowCommentsAndContacts
End Sub

Private Sub grdPostedInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sLWeek As String
    Dim blNoToolTip As Boolean
    
    'grdPostedInfo.ToolTipText = ""
    blNoToolTip = True
    If (grdPostedInfo.MouseRow >= grdPostedInfo.FixedRows) And (grdPostedInfo.MouseCol > PSELECTINDEX) And (grdPostedInfo.TextMatrix(grdPostedInfo.MouseRow, grdPostedInfo.MouseCol)) <> "" Then
        grdPostedInfo.ToolTipText = grdPostedInfo.TextMatrix(grdPostedInfo.MouseRow, grdPostedInfo.MouseCol)
        blNoToolTip = False
    End If
    If (grdPostedInfo.MouseRow >= grdPostedInfo.FixedRows) And (grdPostedInfo.MouseCol = PSTATUSINDEX) Then
        sLWeek = grdPostedInfo.TextMatrix(grdPostedInfo.MouseRow, PWEEKINDEX)
        grdPostedInfo.Row = grdPostedInfo.MouseRow
        grdPostedInfo.Col = grdPostedInfo.MouseCol
        blNoToolTip = False
        Select Case grdPostedInfo.CellBackColor
            Case LIGHTBLUECOLOR
                grdPostedInfo.ToolTipText = sLWeek & ": " & "Not Compliant"
            Case vbMagenta
                grdPostedInfo.ToolTipText = sLWeek & ": " & "Partially Posted"
            Case vbBlue
                grdPostedInfo.ToolTipText = sLWeek & ": " & "Not Aired"
            Case MIDGREENCOLOR
                grdPostedInfo.ToolTipText = sLWeek & ": " & "Compliant"
            Case vbRed
                grdPostedInfo.ToolTipText = sLWeek & ": " & "Not Yet Posted"
            Case BROWN
                grdPostedInfo.ToolTipText = sLWeek & ": " & "Not Yet Exported"
            Case Else
                blNoToolTip = True
        End Select
    End If
    If blNoToolTip Then
        grdPostedInfo.ToolTipText = ""
    End If
End Sub

Private Sub grdPostedInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilRet As Integer
    
    
    If bmInMouseUp Then
        Exit Sub
    End If
    bmInMouseUp = True
    If Y < grdPostedInfo.RowHeight(0) Then
        If (grdSpotInfo.Visible) Then
            bmInMouseUp = False
            Exit Sub
        End If
        grdPostedInfo.Row = grdPostedInfo.MouseRow
        grdPostedInfo.Col = grdPostedInfo.MouseCol
        If grdPostedInfo.CellBackColor = LIGHTBLUE Then
            mMousePointer vbHourglass
            'llCol = mGetCol(grdPostedInfo, X)
            'If llCol <> -1 Then
            '    grdPostedInfo.Col = llCol
                mPostedInfoSortCol grdPostedInfo.Col
                grdPostedInfo.Row = 0
                grdPostedInfo.Col = PCPTTINDEX
            'End If
        End If
        mMousePointer vbDefault
        bmInMouseUp = False
        Exit Sub
    Else
        If Not gGrid_DetermineRowCol(grdPostedInfo, X, Y) Then
            bmInMouseUp = False
            Exit Sub
        End If
        If Trim$(grdPostedInfo.TextMatrix(grdPostedInfo.Row, PWEEKINDEX)) = "" Then
            bmInMouseUp = False
            Exit Sub
        End If
        grdPostedInfo.TopRow = grdPostedInfo.Row
    End If

    If grdPostedInfo.Col = PSELECTINDEX Then
        If grdSpotInfo.Visible Then
            grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Spts"
            grdSpotInfo.Visible = False
            mSetGridPosition
            grdPostedInfo.TextMatrix(grdPostedInfo.Row, PSELECTINDEX) = "t"
            bmPostedScrollAllowed = True
            lmPostedTopRow = -1
        Else
            mMousePointer vbHourglass
            grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Wks"
            ilRet = mPopSpotInfoGrid()
            If ilRet Then
                grdSpotInfo.Visible = True
                mSetGridPosition
                grdPostedInfo.TextMatrix(grdPostedInfo.Row, PSELECTINDEX) = "s"
                bmPostedScrollAllowed = False
                lmPostedTopRow = grdPostedInfo.TopRow
            End If
            mMousePointer vbDefault
        End If
    End If
    bmInMouseUp = False
End Sub

Private Sub grdPostedInfo_Scroll()
    'If Not bmPostedScrollAllowed Then
    '    grdPostedInfo.TopRow = lmPostedTopRow
    'End If
    Dim slVehicleName As String
    Dim llAttCode As Long
    Dim llCpttCode As Long
    If bmInScroll Then
        Exit Sub
    End If
'    If Not bmPostedScrollAllowed Then
'        bmInScroll = True
'        If grdPostedInfo.TopRow > lmPostedTopRow Then
'            If grdPostedInfo.TextMatrix(lmPostedTopRow + 1, PCPTTINDEX) = "" Then
'                grdPostedInfo.TopRow = lmPostedTopRow
'            Else
'                grdPostedInfo.TopRow = lmPostedTopRow + 1
'                If bmPostedScrollAllowed Then
'                    slVehicleName = ""
'                    llAttCode = 0
'                    llCpttCode = 0
'                Else
'                    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow, AVEHICLEINDEX))
'                    llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow, AATTCODEINDEX)))
'                    llCpttCode = Val(Trim$(grdPostedInfo.TextMatrix(lmPostedTopRow + 1, PCPTTINDEX)))
'                End If
'                grdPostedInfo.TextMatrix(lmPostedTopRow, PSELECTINDEX) = "t"
'                grdPostedInfo.TextMatrix(lmPostedTopRow + 1, PSELECTINDEX) = "s"
'                mMousePointer vbHourglass
'                mRepopGrids slVehicleName, llAttCode, llCpttCode
'                mSetGridPosition
'                mMousePointer vbDefault
'            End If
'        ElseIf grdPostedInfo.TopRow < lmPostedTopRow Then
'            If lmPostedTopRow - 1 < grdPostedInfo.FixedRows Then
'                grdPostedInfo.TopRow = lmPostedTopRow
'            Else
'                grdPostedInfo.TopRow = lmPostedTopRow - 1
'                If bmPostedScrollAllowed Then
'                    slVehicleName = ""
'                    llAttCode = 0
'                    llCpttCode = 0
'                Else
'                    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow, AVEHICLEINDEX))
'                    llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow, AATTCODEINDEX)))
'                    llCpttCode = Val(Trim$(grdPostedInfo.TextMatrix(lmPostedTopRow - 1, PCPTTINDEX)))
'                End If
'                grdPostedInfo.TextMatrix(lmPostedTopRow, PSELECTINDEX) = "t"
'                grdPostedInfo.TextMatrix(lmPostedTopRow - 1, PSELECTINDEX) = "s"
'                mMousePointer vbHourglass
'                mRepopGrids slVehicleName, llAttCode, llCpttCode
'                mSetGridPosition
'                mMousePointer vbDefault
'            End If
'        End If
'        bmInScroll = False
'    End If
    If Not bmPostedScrollAllowed Then
        bmInScroll = True
        If grdPostedInfo.TopRow > lmPostedTopRow Then
            If grdPostedInfo.TextMatrix(lmPostedTopRow + 1, PCPTTINDEX) = "" Then
                grdPostedInfo.TopRow = lmPostedTopRow
            Else
                grdPostedInfo.TopRow = lmPostedTopRow + 1
                If grdPostedInfo.TextMatrix(lmPostedTopRow, PSELECTINDEX) <> "" Then
                    grdPostedInfo.TextMatrix(lmPostedTopRow, PSELECTINDEX) = "t"
                End If
                If bmPostedScrollAllowed Then
                    slVehicleName = ""
                    llAttCode = 0
                    llCpttCode = 0
                    If grdPostedInfo.TextMatrix(lmPostedTopRow + 1, PSELECTINDEX) <> "" Then
                        grdPostedInfo.TextMatrix(lmPostedTopRow + 1, PSELECTINDEX) = "s"
                    End If
                    mMousePointer vbHourglass
                    mRepopGrids slVehicleName, llAttCode, llCpttCode
                    mSetGridPosition
                    mMousePointer vbDefault
                Else
                    imDelaySource = 2
                    tmcDelay.Enabled = True
                End If
            End If
        ElseIf grdPostedInfo.TopRow < lmPostedTopRow Then
            If lmPostedTopRow - 1 < grdPostedInfo.FixedRows Then
                grdPostedInfo.TopRow = lmPostedTopRow
            Else
                grdPostedInfo.TopRow = lmPostedTopRow - 1
                If grdPostedInfo.TextMatrix(lmPostedTopRow, PSELECTINDEX) <> "" Then
                    grdPostedInfo.TextMatrix(lmPostedTopRow, PSELECTINDEX) = "t"
                End If
                If bmPostedScrollAllowed Then
                    slVehicleName = ""
                    llAttCode = 0
                    llCpttCode = 0
                    If grdPostedInfo.TextMatrix(lmPostedTopRow - 1, PSELECTINDEX) <> "" Then
                        grdPostedInfo.TextMatrix(lmPostedTopRow - 1, PSELECTINDEX) = "s"
                    End If
                    mMousePointer vbHourglass
                    mRepopGrids slVehicleName, llAttCode, llCpttCode
                    mSetGridPosition
                    mMousePointer vbDefault
                Else
                    imDelaySource = 2
                    tmcDelay.Enabled = True
                End If
            End If
        End If
        bmInScroll = False
    End If

End Sub

Private Sub grdSpotInfo_GotFocus()
    mSetShowCommentsAndContacts
End Sub

Private Sub grdSpotInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim slDate As String
    Dim blNoToolTip As Boolean
    
    'grdSpotInfo.ToolTipText = ""
    blNoToolTip = True
    If (grdSpotInfo.MouseRow >= grdSpotInfo.FixedRows) And (grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, grdSpotInfo.MouseCol)) <> "" Then
        If (grdSpotInfo.MouseCol = DADVTINDEX) Then
            If Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DCNTRNOINDEX)) <> "" Then
                grdSpotInfo.ToolTipText = "# " & Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DCNTRNOINDEX)) & " " & Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, grdSpotInfo.MouseCol))
            Else
                grdSpotInfo.ToolTipText = Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, grdSpotInfo.MouseCol))
            End If
        ElseIf (grdSpotInfo.MouseCol = DDATEINDEX) Then
            If Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DGAMEINFOINDEX)) <> "" Then
                grdSpotInfo.ToolTipText = "Game " & Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DGAMEINFOINDEX))
            Else
                grdSpotInfo.ToolTipText = Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, grdSpotInfo.MouseCol))
            End If
        ElseIf (grdSpotInfo.MouseCol = DCOMMENTINDEX) Then
            If Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DMMRINFOINDEX)) <> "" Then
                grdSpotInfo.ToolTipText = Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DMMRINFOINDEX))
            Else
                grdSpotInfo.ToolTipText = Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, grdSpotInfo.MouseCol))
            End If
        Else
            grdSpotInfo.ToolTipText = Trim$(grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, grdSpotInfo.MouseCol))
        End If
        blNoToolTip = False
    End If
    If (grdSpotInfo.MouseRow >= grdSpotInfo.FixedRows) And (grdSpotInfo.MouseCol = DSTATUSINDEX) Then
        'slDate = grdSpotInfo.TextMatrix(grdSpotInfo.MouseRow, DDATEINDEX)
        grdSpotInfo.Row = grdSpotInfo.MouseRow
        grdSpotInfo.Col = grdSpotInfo.MouseCol
        blNoToolTip = False
        Select Case grdSpotInfo.CellBackColor
            Case vbMagenta
                grdSpotInfo.ToolTipText = "Partially Posted"    'slDate & ": " & "Partially Posted"
            Case vbBlue
                grdSpotInfo.ToolTipText = "Not Aired"   'slDate & ": " & "Not Aired"
            'Case GRAY
            '    grdSpotInfo.ToolTipText = "Not Aired, Compliant"   'slDate & ": " & "Not Aired"
            'Case LIGHTGRAY
            '    grdSpotInfo.ToolTipText = "Not Aired, Not Compliant"   'slDate & ": " & "Not Aired"
            Case MIDGREENCOLOR
                grdSpotInfo.ToolTipText = "Compliant" 'slDate & ": " & "Aired and Compliant"
            Case vbRed
                grdSpotInfo.ToolTipText = "Not Yet Posted"  'slDate & ": " & "Not Yet Posted"
            Case BROWN
                grdSpotInfo.ToolTipText = "Not Yet Exported"    'slDate & ": " & "Not Yet Exported"
            Case LIGHTBLUECOLOR
                grdSpotInfo.ToolTipText = "Not Compliant"   'slDate & ": " & "Not Compliant"
            Case Else
                blNoToolTip = True
        End Select
    End If
    If blNoToolTip Then
        grdSpotInfo.ToolTipText = ""
    End If
End Sub

Private Sub grdSpotInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llCol As Long
    
    If bmInMouseUp Then
        Exit Sub
    End If
    bmInMouseUp = True
    If Y < grdSpotInfo.RowHeight(0) Then
        grdSpotInfo.Row = grdSpotInfo.MouseRow
        grdSpotInfo.Col = grdSpotInfo.MouseCol
        If grdSpotInfo.CellBackColor = LIGHTBLUE Then
            mMousePointer vbHourglass
            'llCol = mGetCol(grdAgreementInfo, X)
            'If llCol <> -1 Then
            '    grdSpotInfo.Col = llCol
                mSpotInfoSortCol grdSpotInfo.Col
                grdSpotInfo.Row = 0
                grdSpotInfo.Col = DASTCODEINDEX
            'End If
            mMousePointer vbDefault
        End If
        bmInMouseUp = False
        Exit Sub
    End If
    'If Not gGrid_DetermineRowCol(grdSpotInfo, X, Y) Then
    '    Exit Sub
    'End If
    bmInMouseUp = False
End Sub

Private Sub grdStations_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilPos As Integer
    Dim blNoToolTip As Boolean
    
    'grdStations.ToolTipText = ""
    blNoToolTip = True
    If (grdStations.MouseRow >= grdStations.FixedRows) And (grdStations.MouseCol > SSELECTINDEX) And ((grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) <> "")) Then
        If grdStations.MouseCol = SOPERATORINDEX Then
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "*" Then
                grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, SOPERATORNAMEINDEX)
                blNoToolTip = False
            End If
        ElseIf grdStations.MouseCol = SCLUSTERINDEX Then
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "*" Then
                grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, SCLUSTERNAMESINDEX)
                blNoToolTip = False
            End If
        ElseIf grdStations.MouseCol = SMCASTINDEX Then
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "*" Then
                grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, SMCASTNAMESINDEX)
                blNoToolTip = False
            End If
        ElseIf grdStations.MouseCol = SPWINDEX Then
            grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, SPASSWORDINDEX)
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SWEBSITEINDEX Then
            grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, SWEBADDRESSINDEX)
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SAGRMNTINDEX Then
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "A" Then
                grdStations.ToolTipText = "Agreement Exist(s)" 'grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
                blNoToolTip = False
            End If
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "SA" Then
                grdStations.ToolTipText = "Service Agreement Exist(s)" 'grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
                blNoToolTip = False
            End If
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "A+SA" Then
                grdStations.ToolTipText = "Agreement + Service Agreement Exist(s)" 'grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
                blNoToolTip = False
            End If
        ElseIf grdStations.MouseCol = SCOMMENTINDEX Then
            If grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) = "C" Then
                grdStations.ToolTipText = "Comment Exist(s)" 'grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
                blNoToolTip = False
            End If
        ElseIf grdStations.MouseCol = SCALLLETTERINDEX Then
            grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) & " " & grdStations.TextMatrix(grdStations.MouseRow, SFREQMONIKERINDEX)
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SOWNERINDEX Then
            grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SZONEINDEX Then
            ilPos = InStr(1, grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol), "*", vbTextCompare)
            If ilPos <= 0 Then
                grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
            Else
                grdStations.ToolTipText = Left$(grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol), ilPos - 1) & " No Daylight Savings"
            End If
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SFORMATINDEX Then
            If Trim$(grdStations.TextMatrix(grdStations.MouseRow, SWATTSINDEX)) <> "" Then
                grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol) & ", " & Trim$(grdStations.TextMatrix(grdStations.MouseRow, SWATTSINDEX)) & " Watts"
            Else
                grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
            End If
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SDMARANKINDEX Then
            If Trim$(grdStations.TextMatrix(grdStations.MouseRow, SAUDP12PLUSINDEX)) <> "" Then
                grdStations.ToolTipText = "DMA Rank, P12+ " & Trim$(grdStations.TextMatrix(grdStations.MouseRow, SAUDP12PLUSINDEX))
            Else
                grdStations.ToolTipText = "DMA Rank"
            End If
            blNoToolTip = False
        ElseIf grdStations.MouseCol = SMSARANKINDEX Then
            grdStations.ToolTipText = "MSA Rank"
            blNoToolTip = False
        Else
            grdStations.ToolTipText = grdStations.TextMatrix(grdStations.MouseRow, grdStations.MouseCol)
            blNoToolTip = False
        End If
    End If
    If blNoToolTip Then
        grdStations.ToolTipText = ""
    End If
End Sub

Private Sub grdStations_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilRet As Integer
    
    If bmInMouseUp Then
        Exit Sub
    End If
    bmInMouseUp = True
    llRow = grdStations.MouseRow
    If Y < grdStations.RowHeight(0) Then
        If (grdStations.MouseCol = SSELECTINDEX) Then
            grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Veh"
            If (Not bmAgreementScrollAllowed) And (grdAgreementInfo.Visible) Then
                grdAgreementInfo.TextMatrix(0, ASELECTINDEX) = "Wks"
                grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, ASELECTINDEX) = "t"
            End If
            If (Not bmPostedScrollAllowed) And (grdPostedInfo.Visible) Then
                grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Spts"
                grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) = "t"
            End If
            grdSpotInfo.Visible = False
            grdPostedInfo.Visible = False
            grdAgreementInfo.Visible = False
            If rbcComments(0).Value Then
                udcContactGrid.Visible = False
                udcCommentGrid.Visible = False
                mHideCommentButtons
            Else
                udcContactGrid.Visible = False
                udcCommentGrid.VehicleCode = 0
            End If
            mSetGridPosition
            If grdStations.Row > llRow Then
                If grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) <> "" Then
                    grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "t"
                    mSetStationGridData grdStations.Row, "t"
                End If
            End If
            'grdStations.TopRow = grdStations.FixedRows
            vbcStation.Value = vbcStation.Min
            'grdStations.Redraw = False
            'mDisplayStationInfo
            'grdStations.Redraw = True
            mFillStationGrid
            bmStationScrollAllowed = True
            lmStationTopRow = -1
            bmAgreementScrollAllowed = True
            lmAgreementTopRow = -1
            bmPostedScrollAllowed = True
            lmPostedTopRow = -1
            mSetCommands
            bmInMouseUp = False
            Exit Sub
        End If
        If (grdAgreementInfo.Visible) Or (udcContactGrid.Visible) Or (udcCommentGrid.Visible) Then
            mSetCommands
            bmInMouseUp = False
            Exit Sub
        End If
        grdStations.Row = grdStations.MouseRow
        grdStations.Col = grdStations.MouseCol
        If grdStations.CellBackColor = LIGHTBLUE Then
            mMousePointer vbHourglass
            'llCol = mGetCol(grdStations, X)
            'If llCol <> -1 Then
            '    grdStations.Col = llCol
                mStationSortCol grdStations.Col
                grdStations.Row = 0
                grdStations.Col = SSHTTCODEINDEX
                vbcStation.Value = vbcStation.Min
                'grdStations.Redraw = False
                'mDisplayStationInfo
                'grdStations.Redraw = True
                mFillStationGrid
            'End If
        End If
        mSetCommands
        mMousePointer vbDefault
        bmInMouseUp = False
        Exit Sub
    Else
        If Not gGrid_DetermineRowCol(grdStations, X, Y) Then
            mSetCommands
            bmInMouseUp = False
            Exit Sub
        End If
        If Trim$(grdStations.TextMatrix(grdStations.Row, SCALLLETTERINDEX)) = "" Then
            mSetCommands
            bmInMouseUp = False
            Exit Sub
        End If
        grdStations.TopRow = grdStations.Row
    End If
    If grdStations.Col = SSELECTINDEX Then
        If grdStations.CellBackColor = LIGHTGREENCOLOR Then   'LIGHTBLUE Then
            If grdAgreementInfo.Visible Then
                grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Veh"
                If (Not bmAgreementScrollAllowed) And (grdAgreementInfo.Visible) Then
                    grdAgreementInfo.TextMatrix(0, ASELECTINDEX) = "Wks"
                    grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, ASELECTINDEX) = "t"
                End If
                If (Not bmPostedScrollAllowed) And (grdPostedInfo.Visible) Then
                    grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Spts"
                    grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) = "t"
                End If
                grdSpotInfo.Visible = False
                grdPostedInfo.Visible = False
                grdAgreementInfo.Visible = False
                If rbcComments(0).Value Then
                    udcContactGrid.Visible = False
                    udcCommentGrid.Visible = False
                    mHideCommentButtons
                Else
                    udcContactGrid.Visible = False
                    udcCommentGrid.VehicleCode = 0
                End If
                mSetGridPosition
                If grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) <> "" Then
                    grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "t"
                    mSetStationGridData grdStations.Row, "t"
                End If
                bmStationScrollAllowed = True
                lmStationTopRow = -1
                bmAgreementScrollAllowed = True
                lmAgreementTopRow = -1
                bmPostedScrollAllowed = True
                lmPostedTopRow = -1
            Else
                grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Stns"
                ilRet = mPopAgreementInfoGrid()
                mSaveCommentsAndContacts
                If udcContactGrid.Visible Then
                    udcContactGrid.StationCode = imShttCode
                    udcContactGrid.Action 3 'Populate
                End If
                If udcCommentGrid.Visible Then
                    udcCommentGrid.StationCode = imShttCode
                    udcCommentGrid.Action 3 'Populate
                    lmCommentStartLoc = udcCommentGrid.ColumnStartLocation(6)
                End If
                grdAgreementInfo.Visible = True
                mSetGridPosition
                grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "s"
                mSetStationGridData grdStations.Row, "s"
                bmStationScrollAllowed = False
                lmStationTopRow = grdStations.TopRow
                mFindAndDisplay grdStations.Row
            End If
        End If
    End If
    If (grdStations.Col = SCALLLETTERINDEX) And (frmDirectory!cmdStation.Enabled = True) Then
        sgStationCallSource = "S"
        igTCShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
        sgTCCallLetters = Trim$(grdStations.TextMatrix(grdStations.Row, SCALLLETTERINDEX))
        If bgStationVisible Then
            frmStation.SetFocus
        Else
            'Get all of the latest passwords and email addresses from the web
            gRemoteTestForNewEmail
            gRemoteTestForNewWebPW
            frmStation.Show
        End If
    End If
    If (grdStations.Col = SAGRMNTINDEX) And (frmDirectory!cmdAgreements.Enabled = True) Then
        sgAgreementCallSource = "S"
        igTCShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
        sgTCCallLetters = Trim$(grdStations.TextMatrix(grdStations.Row, SCALLLETTERINDEX))
        lgTCAttCode = 0
        If bgAgreementVisible Then
            frmAgmnt.SetFocus
        Else
            frmAgmnt.Show
        End If
    End If

    If grdStations.Col = SCOMMENTINDEX Then
        'If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) = AFFILIATECRM) Then
            If udcCommentGrid.Visible Then
                If udcContactGrid.Visible Then
                    If rbcComments(0).Value Then
                        cmcCloseComment_Click
                    Else
                        udcContactGrid.Visible = False
                    End If
                Else
                    imShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
                    udcContactGrid.StationCode = imShttCode
                    udcContactGrid.Action 3 'populate
                    udcContactGrid.Visible = True
                End If
                edcCommentTip.Visible = False
                mSetGridPosition
            Else
                'ilRet = mPopCommentGrid()
                imShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
                udcCommentGrid.StationCode = imShttCode
                udcCommentGrid.Action 3 'populate
                lmCommentStartLoc = udcCommentGrid.ColumnStartLocation(6)
                udcCommentGrid.Visible = True
                udcContactGrid.StationCode = imShttCode
                udcContactGrid.Action 3 'populate
                udcContactGrid.Visible = True
                bmStationScrollAllowed = False
                mSetGridPosition
            End If
        'Else
        '    MsgBox "Call Counterpoint as the Comment feature is not activated"
        'End If
    End If
    If grdStations.Col = SWEBSITEINDEX Then
        If grdStations.CellBackColor = LIGHTGREENCOLOR Then   'LIGHTBLUE Then
            mBranchToWebAddress
        End If
    End If
    mSetCommands
    bmInMouseUp = False
End Sub

Private Sub grdAgreementInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilFound As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim sLWeek As String
    Dim llColLeftPos As Long
    Dim blNoToolTip As Boolean
    
    'grdAgreementInfo.ToolTipText = ""
    blNoToolTip = True
    If (grdAgreementInfo.MouseRow = 0) And (grdAgreementInfo.MouseCol >= AWEEK1INDEX) And (grdAgreementInfo.MouseCol <= AWEEK1INDEX + 53) Then
        sLWeek = DateAdd("d", -((grdAgreementInfo.MouseCol - AWEEK1INDEX) * 7 + 6), smWeek1)
        If grdAgreementInfo.MouseCol = AWEEK1INDEX Then
            grdAgreementInfo.ToolTipText = "Next Week: " & sLWeek
        ElseIf grdAgreementInfo.MouseCol = AWEEK1INDEX + 1 Then
            grdAgreementInfo.ToolTipText = "Current Week: " & sLWeek
        Else
            grdAgreementInfo.ToolTipText = sLWeek
        End If
        blNoToolTip = False
    End If
    If (grdAgreementInfo.MouseRow >= grdAgreementInfo.FixedRows) And (grdAgreementInfo.MouseCol > ASELECTINDEX) And (grdAgreementInfo.TextMatrix(grdAgreementInfo.MouseRow, grdAgreementInfo.MouseCol)) <> "" Then
        If (grdAgreementInfo.MouseCol = AAGRMNTINDEX) And (grdAgreementInfo.TextMatrix(grdAgreementInfo.MouseRow, grdAgreementInfo.MouseCol) = "*") Then
            grdAgreementInfo.ToolTipText = "Multi-Cast"
        ElseIf (grdAgreementInfo.MouseCol = AAGRMNTINDEX) And (grdAgreementInfo.TextMatrix(grdAgreementInfo.MouseRow, grdAgreementInfo.MouseCol) = "SA") Then
            grdAgreementInfo.ToolTipText = "Service Agreement"
        ElseIf (grdAgreementInfo.MouseCol = AAGRMNTINDEX) And (grdAgreementInfo.TextMatrix(grdAgreementInfo.MouseRow, grdAgreementInfo.MouseCol) = "*SA") Then
            grdAgreementInfo.ToolTipText = "Multi-Cast + Service Agreement"
        ElseIf (grdAgreementInfo.MouseCol = ADAYTIMEINDEX) Then
            grdAgreementInfo.ToolTipText = grdAgreementInfo.TextMatrix(grdAgreementInfo.MouseRow, ATIMERANGEINDEX)
        Else
            grdAgreementInfo.ToolTipText = grdAgreementInfo.TextMatrix(grdAgreementInfo.MouseRow, grdAgreementInfo.MouseCol)
        End If
        blNoToolTip = False
    End If
    If (grdAgreementInfo.MouseRow >= grdAgreementInfo.FixedRows) And (grdAgreementInfo.MouseCol >= AWEEK1INDEX) And (grdAgreementInfo.MouseCol <= AWEEK1INDEX + 53) Then
        sLWeek = DateAdd("d", -((grdAgreementInfo.MouseCol - AWEEK1INDEX) * 7 + 6), smWeek1)
        grdAgreementInfo.Row = grdAgreementInfo.MouseRow
        grdAgreementInfo.Col = grdAgreementInfo.MouseCol
        blNoToolTip = False
        Select Case grdAgreementInfo.CellBackColor
            Case LIGHTBLUECOLOR
                grdAgreementInfo.ToolTipText = sLWeek & ": " & "Not Compliant"
            Case vbMagenta
                grdAgreementInfo.ToolTipText = sLWeek & ": " & "Partially Posted"
            Case vbBlue
                grdAgreementInfo.ToolTipText = sLWeek & ": " & "Not Aired"
            Case MIDGREENCOLOR
                grdAgreementInfo.ToolTipText = sLWeek & ": " & "Compliant"
            Case vbRed
                grdAgreementInfo.ToolTipText = sLWeek & ": " & "Not Yet Posted"
            Case BROWN
                grdAgreementInfo.ToolTipText = sLWeek & ": " & "Not Yet Exported"
            Case Else
                blNoToolTip = True
        End Select
    End If
    If blNoToolTip Then
        grdAgreementInfo.ToolTipText = ""
    End If
End Sub

Private Sub grdAgreementInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilRet As Integer
    Dim sLWeek As String
    
    
    If bmInMouseUp Then
        Exit Sub
    End If
    bmInMouseUp = True
    If Y < grdAgreementInfo.RowHeight(0) Then
        If (grdPostedInfo.Visible) Then
            bmInMouseUp = False
            Exit Sub
        End If
        grdAgreementInfo.Row = grdAgreementInfo.MouseRow
        grdAgreementInfo.Col = grdAgreementInfo.MouseCol
        If grdAgreementInfo.CellBackColor = LIGHTBLUE Then
            mMousePointer vbHourglass
            'llCol = mGetCol(grdAgreementInfo, X)
            'If llCol <> -1 Then
            '    grdAgreementInfo.Col = llCol
                mAgreementInfoSortCol grdAgreementInfo.Col
                grdAgreementInfo.Row = 0
                grdAgreementInfo.Col = AATTCODEINDEX
            'End If
        End If
        mMousePointer vbDefault
        bmInMouseUp = False
        Exit Sub
    Else
        If Not gGrid_DetermineRowCol(grdAgreementInfo, X, Y) Then
            bmInMouseUp = False
            Exit Sub
        End If
        If Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, AVEHICLEINDEX)) = "" Then
            bmInMouseUp = False
            Exit Sub
        End If
        grdAgreementInfo.TopRow = grdAgreementInfo.Row
    End If

    If grdAgreementInfo.Col = ASELECTINDEX Then
        If grdAgreementInfo.CellBackColor = LIGHTGREENCOLOR Then 'LIGHTBLUE Then
            mMousePointer vbHourglass
            If grdPostedInfo.Visible Then
                grdAgreementInfo.TextMatrix(0, ASELECTINDEX) = "Wks"
                If (Not bmPostedScrollAllowed) And (grdPostedInfo.Visible) Then
                    grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) = "t"
                End If
                grdSpotInfo.Visible = False
                grdPostedInfo.Visible = False
                mSetGridPosition
                grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, ASELECTINDEX) = "t"
                bmAgreementScrollAllowed = True
                lmAgreementTopRow = -1
                bmPostedScrollAllowed = True
                lmPostedTopRow = -1
            Else
                grdAgreementInfo.TextMatrix(0, PSELECTINDEX) = "Veh"
                ilRet = mPopPostedInfoGrid()
                If udcCommentGrid.Visible Then
                    udcCommentGrid.VehicleCode = imVefCode
                End If
                grdPostedInfo.Visible = True
                mSetGridPosition
                grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, ASELECTINDEX) = "s"
                bmAgreementScrollAllowed = False
                lmAgreementTopRow = grdAgreementInfo.TopRow
            End If
            mMousePointer vbDefault
        End If
    End If
    If (grdAgreementInfo.Col >= AWEEK1INDEX) And (grdAgreementInfo.Col <= AWEEK1INDEX + 53) Then
        If grdAgreementInfo.CellBackColor <> GRAY Then 'LIGHTBLUE Then
            mMousePointer vbHourglass
            sLWeek = DateAdd("d", -((grdAgreementInfo.Col - AWEEK1INDEX) * 7 + 6), smWeek1)
            If grdPostedInfo.Visible = False Then
                grdAgreementInfo.TextMatrix(0, PSELECTINDEX) = "Veh"
                ilRet = mPopPostedInfoGrid()
                grdPostedInfo.Visible = True
                mSetGridPosition
                grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, ASELECTINDEX) = "s"
                bmAgreementScrollAllowed = False
                lmAgreementTopRow = grdAgreementInfo.TopRow
            Else
                grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) = "t"
            End If
            For llRow = grdPostedInfo.FixedRows To grdPostedInfo.Rows - 1 Step 1
                If grdPostedInfo.TextMatrix(llRow, PWEEKINDEX) <> "" Then
                    If gDateValue(grdPostedInfo.TextMatrix(llRow, PWEEKINDEX)) = gDateValue(sLWeek) Then
                        bmPostedScrollAllowed = True
                        grdPostedInfo.Row = llRow
                        grdPostedInfo.TopRow = grdPostedInfo.Row
                        grdPostedInfo.Col = PSELECTINDEX
                        grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Wks"
                        ilRet = mPopSpotInfoGrid()
                        If ilRet Then
                            grdSpotInfo.Visible = True
                            mSetGridPosition
                            grdPostedInfo.TextMatrix(grdPostedInfo.Row, PSELECTINDEX) = "s"
                            bmPostedScrollAllowed = False
                            lmPostedTopRow = grdPostedInfo.TopRow
                        End If
                    End If
                End If
            Next llRow
            mMousePointer vbDefault
        End If
    End If
    If (grdAgreementInfo.Col = AAGRMNTINDEX) And (frmDirectory!cmdAgreements.Enabled = True) Then
        sgAgreementCallSource = "S"
        igTCShttCode = Val(grdStations.TextMatrix(grdStations.TopRow, SSHTTCODEINDEX))
        sgTCCallLetters = Trim$(grdStations.TextMatrix(grdStations.TopRow, SCALLLETTERINDEX))
        lgTCAttCode = grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, AATTCODEINDEX)
        If bgAgreementVisible Then
            frmAgmnt.SetFocus
        Else
            frmAgmnt.Show
        End If
    End If
    bmInMouseUp = False
End Sub

Private Function mPopAgreementInfoGrid() As Integer
    Dim llRow As Long
    Dim llWeek1 As Long
    Dim llWeek54 As Long
    Dim llCpttDate As Long
    Dim llWeekDate As Long
    Dim llVef As Long
    Dim ilCol As Integer
    Dim slWeek1 As String
    Dim slWeek54 As String
    Dim llYellowRow As Long
    Dim llCol As Long
    Dim ilAnyNotCompliant As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilCompliantCount As Integer
    Dim sLWeek As String
    Dim slDate As String
    Dim llColor As Long
    
    Dim slPSDate As String
    Dim slPEDate As String
    Dim ilPAdfCode As Integer
    Dim blAdvtInWk As Boolean
    Dim blCpttDelinq As Boolean
    
    mPopAgreementInfoGrid = False
    On Error GoTo ErrHand:
    grdAgreementInfo.Rows = 2
    mClearGrid grdAgreementInfo
    gGrid_FillWithRows grdAgreementInfo
    If (grdStations.Row < grdStations.FixedRows) Or (grdStations.Row > grdStations.Rows) Then
        Exit Function
    End If
    If (grdStations.Col < SSELECTINDEX) Or (grdStations.Col > SCOMMENTINDEX) Then
        Exit Function
    End If
    If Trim$(grdStations.TextMatrix(grdStations.Row, SCALLLETTERINDEX)) = "" Then
        Exit Function
    End If
    grdAgreementInfo.Redraw = False
    If (sgStationSearchCallSource = "P") Then
        slPSDate = Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX))
        If slPSDate <> "[All]" Then
            slPEDate = DateAdd("d", 6, slPSDate)
        Else
            slPSDate = smWeek54
            slPEDate = smWeek1
        End If
        ilPAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
    End If
    llRow = grdAgreementInfo.FixedRows
    llWeek1 = gDateValue(smWeek1)
    llWeek54 = gDateValue(smWeek54)
    slWeek1 = Format$(smWeek1, sgSQLDateForm)
    slWeek54 = Format$(smWeek54, sgSQLDateForm)
    If sgStationSearchCallSource = "P" Then
        slDate = grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX)
        grdAgreementInfo.Row = 0
        grdAgreementInfo.Col = AWEEK1INDEX
        llColor = grdAgreementInfo.CellBackColor
        For llCol = AWEEK1INDEX + 1 To AWEEK1INDEX + 53 Step 1
            sLWeek = DateAdd("d", -((llCol - AWEEK1INDEX) * 7 + 6), smWeek1)
            grdAgreementInfo.Row = 0
            If slDate <> "[All]" Then
                If gDateValue(sLWeek) = gDateValue(slDate) Then
                    grdAgreementInfo.Col = llCol
                    grdAgreementInfo.CellBackColor = vbGreen
                Else
                    grdAgreementInfo.Col = llCol
                    grdAgreementInfo.CellBackColor = llColor
                End If
            Else
                grdAgreementInfo.Col = llCol
                grdAgreementInfo.CellBackColor = llColor
            End If
        Next llCol
    End If
    imShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
    SQLQuery = "SELECT * FROM att "
    SQLQuery = SQLQuery + " WHERE ("
    If sgStationSearchCallSource = "P" Then
        SQLQuery = SQLQuery & "(attOnAir <= '" & slWeek1 & "')"
        SQLQuery = SQLQuery & " AND (attOffAir >= '" & slWeek54 & "')"
        SQLQuery = SQLQuery & " AND (attDropDate >= '" & slWeek54 & "')"
        SQLQuery = SQLQuery & " AND attServiceAgreement <> 'Y'"
        SQLQuery = SQLQuery & " AND (attShfCode = " & imShttCode & "))"
    Else
        SQLQuery = SQLQuery & "(((attOnAir <= '" & slWeek1 & "')"
        SQLQuery = SQLQuery & " AND (attOffAir >= '" & slWeek54 & "'))"
        SQLQuery = SQLQuery & " OR ((attDropDate >= '" & slWeek54 & "')"
        SQLQuery = SQLQuery & "AND (attOnAir <= '" & slWeek1 & "')))"
        SQLQuery = SQLQuery & " AND (attShfCode = " & imShttCode & "))"
    End If
    Set rst_att = gSQLSelectCall(SQLQuery)
    Do While Not rst_att.EOF
        If (gDateValue(rst_att!attOnAir) <= gDateValue(rst_att!attOffAir)) And (gDateValue(rst_att!attOnAir) <= gDateValue(rst_att!attDropDate)) Then
            
            blAdvtInWk = False
            blCpttDelinq = False
            If (sgStationSearchCallSource = "P") Then
                SQLQuery = "SELECT Count(cpttCode)"
                SQLQuery = SQLQuery & " FROM cptt WHERE "
                SQLQuery = SQLQuery & " (cpttAtfCode = " & rst_att!attCode
                SQLQuery = SQLQuery & " AND cpttStatus = 0 AND cpttPostingStatus < 2"
                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slPSDate, sgSQLDateForm) & "'" & ")"
                Set rst_Cptt = gSQLSelectCall(SQLQuery)
                If Not rst_Cptt.EOF Then
                    If (rst_Cptt(0).Value <> 0) Then
                        blCpttDelinq = True
                    End If
                End If
                SQLQuery = "SELECT Count(lstCode)"
                SQLQuery = SQLQuery & " FROM lst"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery + " lstLogVefCode = " & rst_att!attvefCode
                SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format$(slPSDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(slPEDate, sgSQLDateForm) & "'"
                If lbcContract.ListIndex > 0 Then
                    SQLQuery = SQLQuery & " AND lstCntrNo = " & lbcContract.List(lbcContract.ListIndex)
                Else
                    SQLQuery = SQLQuery + " AND lstCntrNo <> 0"
                End If
                SQLQuery = SQLQuery + " AND lstAdfCode = " & ilPAdfCode & ")"
                Set rst_Lst = gSQLSelectCall(SQLQuery)
                If Not (rst_Lst.EOF) Then
                    If (rst_Lst(0).Value <> 0) Then
                        blAdvtInWk = True
                    End If
                End If
            End If
            If llRow >= grdAgreementInfo.Rows Then
                grdAgreementInfo.AddItem ""
            End If
            grdAgreementInfo.Row = llRow
            grdAgreementInfo.Col = ASELECTINDEX
            grdAgreementInfo.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
            llVef = gBinarySearchVef(rst_att!attvefCode)
            If llVef <> -1 Then
                grdAgreementInfo.TextMatrix(llRow, AVEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
                grdAgreementInfo.Col = AVEHICLEINDEX
                If blAdvtInWk Then
                    grdAgreementInfo.CellBackColor = vbGreen
                Else
                    grdAgreementInfo.CellBackColor = LIGHTYELLOW
                End If
            End If
            grdAgreementInfo.TextMatrix(llRow, ADUEINDEX) = mGetDueCountForAgreement(rst_att!attCode)
            'Missing: Showing date range instead
            grdAgreementInfo.TextMatrix(llRow, ADAYTIMEINDEX) = ""
            If gDateValue(rst_att!attDropDate) < gDateValue(rst_att!attOffAir) Then
                If gDateValue(rst_att!attDropDate) <> gDateValue("12/31/2069") Then
                    grdAgreementInfo.TextMatrix(llRow, ADAYTIMEINDEX) = Format$(rst_att!attOnAir, "m/d/yy") & "-" & Format$(rst_att!attDropDate, "m/d/yy")
                Else
                    grdAgreementInfo.TextMatrix(llRow, ADAYTIMEINDEX) = Format$(rst_att!attOnAir, "m/d/yy") & "-" & "TFN"
                End If
            Else
                If gDateValue(rst_att!attOffAir) <> gDateValue("12/31/2069") Then
                    grdAgreementInfo.TextMatrix(llRow, ADAYTIMEINDEX) = Format$(rst_att!attOnAir, "m/d/yy") & "-" & Format$(rst_att!attOffAir, "m/d/yy")
                Else
                    grdAgreementInfo.TextMatrix(llRow, ADAYTIMEINDEX) = Format$(rst_att!attOnAir, "m/d/yy") & "-" & "TFN"
                End If
            End If
            grdAgreementInfo.Col = ADAYTIMEINDEX
            grdAgreementInfo.CellBackColor = LIGHTYELLOW
            If IsNull(rst_att!attVehProgStartTime) Then
                grdAgreementInfo.TextMatrix(llRow, ATIMERANGEINDEX) = ""
            Else
                grdAgreementInfo.TextMatrix(llRow, ATIMERANGEINDEX) = Format$(rst_att!attVehProgStartTime, "hh:mm:ssA/P")
                If Not IsNull(rst_att!attVehProgEndTime) Then
                    grdAgreementInfo.TextMatrix(llRow, ATIMERANGEINDEX) = grdAgreementInfo.TextMatrix(llRow, ATIMERANGEINDEX) & "-" & Format$(rst_att!attVehProgEndTime, "hh:mm:ssA/P")
                End If
            End If

            For ilCol = AWEEK1INDEX To AWEEK1INDEX + 53 Step 1
                grdAgreementInfo.Col = ilCol
                grdAgreementInfo.CellBackColor = GRAY   'vbWhite
                grdAgreementInfo.TextMatrix(llRow, ilCol) = ""
            Next ilCol
            grdAgreementInfo.Col = AAGRMNTINDEX
            grdAgreementInfo.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY   'vbWhite
            grdAgreementInfo.CellAlignment = flexAlignCenterCenter
            grdAgreementInfo.TextMatrix(llRow, AAGRMNTINDEX) = ""
            If rst_att!attMulticast = "Y" Then
                grdAgreementInfo.TextMatrix(llRow, AAGRMNTINDEX) = "*"
            End If
            If rst_att!attServiceAgreement = "Y" Then
                grdAgreementInfo.TextMatrix(llRow, AAGRMNTINDEX) = grdAgreementInfo.TextMatrix(llRow, AAGRMNTINDEX) & "SA"
            End If
            grdAgreementInfo.Col = ASELECTINDEX
            grdAgreementInfo.CellBackColor = GRAY   'vbWhite
            SQLQuery = "SELECT * FROM cptt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " cpttatfCode = " & rst_att!attCode & ")"
            Set rst_Cptt = gSQLSelectCall(SQLQuery)
            Do While Not rst_Cptt.EOF
                llCpttDate = gDateValue(rst_Cptt!CpttStartDate)
                If (llCpttDate >= llWeek54) And (llCpttDate <= llWeek1) Then
                    grdAgreementInfo.Col = ASELECTINDEX
                    grdAgreementInfo.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
                    grdAgreementInfo.CellFontName = "Monotype Sorts"
                    grdAgreementInfo.TextMatrix(llRow, ASELECTINDEX) = "t"
                    ilCol = AWEEK1INDEX + (DateDiff("d", Format(llCpttDate, sgShowDateForm), smWeek1, vbMonday)) \ 7
                    grdAgreementInfo.Col = ilCol
                    If rst_Cptt!cpttStatus = 2 Then
                        grdAgreementInfo.CellBackColor = vbBlue
                    ElseIf rst_Cptt!cpttStatus = 0 Then
                        If rst_Cptt!cpttPostingStatus = 1 Then
                            grdAgreementInfo.CellBackColor = vbMagenta
                        Else
                            'Posting Type: 0=Receipt Only; 1=Spot Count; 2=Post Dates/Times
                            If rst_att!attPostingType > 1 Then
                                llCpttDate = gDateValue(rst_Cptt!CpttStartDate)
                                llWeekDate = gDateValue(gObtainPrevMonday(gAdjYear(Format$(llCpttDate, "m/d/yy"))))
                                SQLQuery = "SELECT * FROM ast"
                                SQLQuery = SQLQuery + " WHERE ("
                                SQLQuery = SQLQuery + " astFeedDate >= '" & Format(llWeekDate, sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery + " AND astFeedDate <= '" & Format(llWeekDate + 6, sgSQLDateForm) & "'"
                                SQLQuery = SQLQuery & " AND astatfCode = " & rst_att!attCode & ")"
                                Set rst_Ast = gSQLSelectCall(SQLQuery)
                                If Not rst_Ast.EOF Then
                                    grdAgreementInfo.CellBackColor = vbRed
                                Else
                                    grdAgreementInfo.CellBackColor = BROWN
                                End If
                            Else
                                grdAgreementInfo.CellBackColor = vbRed
                            End If
                        End If
                    Else
                        'Missing:  Compliant and one not aired
                        ilAnyNotCompliant = False
                        If (rst_Cptt!cpttNoSpotsGen > 0) Or (rst_Cptt!cpttNoSpotsAired > 0) Or (rst_Cptt!cpttNoCompliant > 0) Then
                            ilSchdCount = rst_Cptt!cpttNoSpotsGen
                            ilAiredCount = rst_Cptt!cpttNoSpotsAired
                            If rbcCompliant(1).Value Then
                                ilCompliantCount = rst_Cptt!cpttAgyCompliant
                            Else
                                ilCompliantCount = rst_Cptt!cpttNoCompliant
                            End If
                            'If ilAiredCount <> ilCompliantCount Then
                            If ilSchdCount > ilCompliantCount Then
                                ilAnyNotCompliant = True
                            End If
                        End If
                        If rst_Cptt!cpttPostingStatus = 2 Then
                            If ilAnyNotCompliant Then
                                grdAgreementInfo.CellBackColor = LIGHTBLUECOLOR    'LIGHTGREEN
                            Else
                                grdAgreementInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                            End If
                        ElseIf rst_Cptt!cpttPostingStatus = 0 Then
                            If rst_att!attPostingType > 1 Then
                                grdAgreementInfo.CellBackColor = vbRed
                            Else
                                grdAgreementInfo.CellBackColor = MIDGREENCOLOR
                            End If
                        ElseIf (rst_Cptt!cpttPostingStatus = 1) Then
                            grdAgreementInfo.CellBackColor = vbMagenta
                        End If
                    End If
                    grdAgreementInfo.TextMatrix(llRow, ilCol) = ""
                End If
                rst_Cptt.MoveNext
            Loop
            grdAgreementInfo.TextMatrix(llRow, AATTCODEINDEX) = rst_att!attCode
            llRow = llRow + 1
        End If
        rst_att.MoveNext
    Loop
    'If grdAgreementInfo.Rows < grdAgreementInfo.Height \ grdAgreementInfo.RowHeight(1) Then
    '    grdAgreementInfo.Rows = grdAgreementInfo.Rows + 2 * (grdAgreementInfo.Height \ grdAgreementInfo.RowHeight(1)) - grdAgreementInfo.Rows + 1
    'Else
    '    grdAgreementInfo.Rows = grdAgreementInfo.Rows + grdAgreementInfo.Height \ grdAgreementInfo.RowHeight(1) + 1
    'End If
    grdAgreementInfo.Rows = grdAgreementInfo.Rows + ((cmcDone.Top - grdAgreementInfo.Top) \ grdAgreementInfo.RowHeight(1))
    For llYellowRow = llRow To grdAgreementInfo.Rows - 1 Step 1
        grdAgreementInfo.Row = llYellowRow
        For llCol = ASELECTINDEX To AAGRMNTINDEX Step 1
            grdAgreementInfo.Col = llCol
            grdAgreementInfo.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llYellowRow
    imLastAgreementInfoSort = -1
    imLastAgreementInfoColSorted = -1
    mAgreementInfoSortCol AVEHICLEINDEX
    grdAgreementInfo.Row = 0
    grdAgreementInfo.Col = AATTCODEINDEX
    grdAgreementInfo.Redraw = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopAgreementInfoGrid"
    grdAgreementInfo.Redraw = True
End Function

Private Function mPopPostedInfoGrid() As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llWeek1 As Long
    Dim llWeek54 As Long
    Dim llCpttDate As Long
    Dim llVef As Long
    Dim ilCol As Integer
    Dim slWeek1 As String
    Dim slWeek54 As String
    Dim llWeekDate As Long
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim llYellowRow As Long
    Dim ilAnyAstExist As Integer
    Dim ilAnyNotCompliant As Integer
    Dim slLatestDate As String
    Dim ilAdfCode As Integer
    Dim ilShtt As Integer
    Dim ilAst As Integer
    Dim slMoDate As String
    Dim ilRet As Integer
    'dan 7701 add to avoid repetition
    Dim blIsMarketron As Boolean
    
    Dim blTaskBlocked As Boolean
    
    mPopPostedInfoGrid = False
    On Error GoTo ErrHand:
    grdPostedInfo.Rows = 2
    mClearGrid grdPostedInfo
    If (grdAgreementInfo.Row < grdAgreementInfo.FixedRows) Or (grdAgreementInfo.Row > grdAgreementInfo.Rows) Then
        Exit Function
    End If
    If (grdAgreementInfo.Col < ASELECTINDEX) Or (grdAgreementInfo.Col > AAGRMNTINDEX) Then
        Exit Function
    End If
    If Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, AVEHICLEINDEX)) = "" Then
        Exit Function
    End If
    grdPostedInfo.Redraw = False
    llRow = grdPostedInfo.FixedRows
    llWeek1 = gDateValue(smWeek1)
    llWeek54 = gDateValue(smWeek54)
    lmAttCode = Val(grdAgreementInfo.TextMatrix(grdAgreementInfo.Row, AATTCODEINDEX))
    '7701 Dan M moved out of loop to here
    If gIsVendorWithAgreement(lmAttCode, Vendors.NetworkConnect) Then
        blIsMarketron = True
    Else
        blIsMarketron = False
    End If
    
    blTaskBlocked = False
    bgTaskBlocked = False
    sgTaskBlockedName = "Affiliate Management"
    
    'dan M 12/2/15 moved from loop
    SQLQuery = "SELECT * FROM att "
    SQLQuery = SQLQuery + " WHERE ("
    SQLQuery = SQLQuery & " attCode = " & lmAttCode & ")"
    Set rst_att = gSQLSelectCall(SQLQuery)

    SQLQuery = "SELECT * FROM cptt"
    SQLQuery = SQLQuery + " WHERE ("
    SQLQuery = SQLQuery & " cpttatfCode = " & lmAttCode & ")"
    SQLQuery = SQLQuery & " ORDER BY cpttStartDate DESC"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    Do While Not rst_Cptt.EOF
        llCpttDate = gDateValue(rst_Cptt!CpttStartDate)
        If (llCpttDate >= llWeek54) And (llCpttDate <= llWeek1) Then
            imVefCode = rst_Cptt!cpttvefcode
            'dan M 12/2/15 moved out of loop
'            SQLQuery = "SELECT * FROM att "
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " attCode = " & lmAttCode & ")"
'            Set rst_att = gSQLSelectCall(SQLQuery)
            If Not rst_att.EOF Then
                If llRow >= grdPostedInfo.Rows Then
                    grdPostedInfo.AddItem ""
                End If
                grdPostedInfo.Row = llRow
                grdPostedInfo.Col = PSELECTINDEX
                grdPostedInfo.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
                grdPostedInfo.CellFontName = "Monotype Sorts"
                grdPostedInfo.TextMatrix(llRow, PSELECTINDEX) = "t"
                For llCol = PWEEKINDEX To PIPINDEX Step 1
                    grdPostedInfo.Col = llCol
                    grdPostedInfo.CellBackColor = LIGHTYELLOW
                Next llCol
                grdPostedInfo.TextMatrix(llRow, PWEEKINDEX) = Format$(Trim$(rst_Cptt!CpttStartDate), "m/d/yy")
                ilSchdCount = 0
                ilAiredCount = 0
                ilPledgeCompliantCount = 0
                ilAgyCompliantCount = 0
                ilAnyAstExist = False
                ilAnyNotCompliant = False
                grdPostedInfo.TextMatrix(llRow, PPERCENTAINDEX) = ""
                grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = ""
                If rst_att!attPostingType = 0 Then  'Receipt
                    grdPostedInfo.TextMatrix(llRow, PNOSCHDINDEX) = ""
                    grdPostedInfo.TextMatrix(llRow, PNOAIREDINDEX) = ""
                    grdPostedInfo.TextMatrix(llRow, PNOCMPLINDEX) = ""
                ElseIf rst_att!attPostingType = 1 Then  'Count
                    ilSchdCount = rst_Cptt!cpttNoSpotsGen
                    ilAiredCount = rst_Cptt!cpttNoSpotsAired
                    grdPostedInfo.TextMatrix(llRow, PNOSCHDINDEX) = rst_Cptt!cpttNoSpotsGen
                    grdPostedInfo.TextMatrix(llRow, PNOAIREDINDEX) = rst_Cptt!cpttNoSpotsAired
                    grdPostedInfo.TextMatrix(llRow, PNOCMPLINDEX) = ""
                Else
                    'Spots by date and spots by advertiser
                    If (rst_Cptt!cpttNoSpotsGen > 0) Or (rst_Cptt!cpttNoSpotsAired > 0) Or (rst_Cptt!cpttNoCompliant > 0) Then
                        ilAnyAstExist = True
                        ilSchdCount = rst_Cptt!cpttNoSpotsGen
                        ilAiredCount = rst_Cptt!cpttNoSpotsAired
                        ilPledgeCompliantCount = rst_Cptt!cpttNoCompliant
                        ilAgyCompliantCount = rst_Cptt!cpttAgyCompliant
                        If rbcCompliant(1).Value Then
                            'If ilAiredCount <> ilAgyCompliantCount Then
                            If ilSchdCount > ilAgyCompliantCount Then
                                ilAnyNotCompliant = True
                            End If
                        Else
                            'If ilAiredCount <> ilPledgeCompliantCount Then
                            If ilSchdCount > ilPledgeCompliantCount Then
                                ilAnyNotCompliant = True
                            End If
                        End If
                    Else
                        llWeekDate = gDateValue(gObtainPrevMonday(gAdjYear(Format$(llCpttDate, "m/d/yy"))))
                        'SQLQuery = "SELECT * FROM ast"
                        'SQLQuery = SQLQuery + " WHERE ("
                        'SQLQuery = SQLQuery + " astFeedDate >= '" & Format(llWeekDate, sgSQLDateForm) & "'"
                        'SQLQuery = SQLQuery + " AND astFeedDate <= '" & Format(llWeekDate + 6, sgSQLDateForm) & "'"
                        'SQLQuery = SQLQuery & " AND astatfCode = " & lmAttCode & ")"
                        'Set rst_Ast = gSQLSelectCall(SQLQuery)
                        'Do While Not rst_Ast.EOF
                        '    ilAnyAstExist = True
                        '    gIncSpotCounts rst_Ast!astPledgeStatus, rst_Ast!astStatus, rst_Ast!astCPStatus, Format$(rst_Ast!astPledgeDate, "m/d/yy"), Format$(rst_Ast!astAirDate, "m/d/yy"), Format$(rst_Ast!astPledgeStartTime, "h:mm:ssAM/PM"), Format$(rst_Ast!astPledgeEndTime, "h:mm:ssAM/PM"), Format$(rst_Ast!astAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                        '    rst_Ast.MoveNext
                        'Loop
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = rst_Cptt!cpttCode
                        tgCPPosting(0).iStatus = rst_Cptt!cpttStatus
                        tgCPPosting(0).iPostingStatus = rst_Cptt!cpttPostingStatus
                        tgCPPosting(0).lAttCode = rst_Cptt!cpttatfCode
                        tgCPPosting(0).iAttTimeType = 0 'Not used
                        tgCPPosting(0).iVefCode = rst_Cptt!cpttvefcode  'imVefCode
                        tgCPPosting(0).iShttCode = rst_Cptt!cpttshfcode
                        ilShtt = gBinarySearchStationInfoByCode(tgCPPosting(0).iShttCode)
                        If ilShtt <> -1 Then
                            tgCPPosting(0).sZone = tgStationInfoByCode(ilShtt).sZone
                        Else
                            tgCPPosting(0).sZone = ""
                        End If
                        slMoDate = Format$(llWeekDate, "m/d/yy")
                        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
                        tgCPPosting(0).sAstStatus = rst_Cptt!cpttAstStatus
                        igTimes = 1 'By Week
                        ilAdfCode = -1
                        'Dan 9/26/13  6442
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, False, True, , , , , , True)
                        'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True)
                        If (bgTaskBlocked) Then blTaskBlocked = True
                        bgTaskBlocked = False
                        For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                            ilAnyAstExist = True
                            'gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, tmAstInfo(ilAst).iStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                            gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
                        Next ilAst
                        If rbcCompliant(1).Value Then
                            'If ilAiredCount <> ilAgyCompliantCount Then
                            If ilSchdCount > ilAgyCompliantCount Then
                                ilAnyNotCompliant = True
                            End If
                        Else
                            'If ilAiredCount <> ilPledgeCompliantCount Then
                            If ilSchdCount > ilPledgeCompliantCount Then
                                ilAnyNotCompliant = True
                            End If
                        End If
                        '1/11/13: Where clause is being dropped, maybe caused by result set changing
                        ilRet = 0
                        SQLQuery = "Update cptt Set "
                        SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
                        SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
                        SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
                        SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
                        '1/11/13: Where clause is being dropped, maybe caused by result set changing
                        SQLQuery = SQLQuery & " Where cpttCode = " & tgCPPosting(0).lCpttCode   'rst_Cptt!cpttCode
                        'cnn.Execute SQLQuery, rdExecDirect
                        '1/11/13: Where clause is being dropped, maybe caused by result set changing
                        If ilRet = 0 Then
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "frmStationSearch-mPopPostedInfoGrid"
                                grdPostedInfo.Redraw = True
                                mPopPostedInfoGrid = False
                                Exit Function
                            End If
                        End If
                    End If
                    grdPostedInfo.TextMatrix(llRow, PNOSCHDINDEX) = ilSchdCount
                    grdPostedInfo.TextMatrix(llRow, PNOAIREDINDEX) = ilAiredCount
                    If rbcCompliant(1).Value Then
                        grdPostedInfo.TextMatrix(llRow, PNOCMPLINDEX) = ilAgyCompliantCount
                    Else
                        grdPostedInfo.TextMatrix(llRow, PNOCMPLINDEX) = ilPledgeCompliantCount
                    End If
                End If
                
                If ilSchdCount > 0 Then
                    grdPostedInfo.TextMatrix(llRow, PPERCENTAINDEX) = (100 * CLng(ilAiredCount)) \ ilSchdCount
                End If
                grdPostedInfo.TextMatrix(llRow, PPOSTDATEINDEX) = ""
                grdPostedInfo.TextMatrix(llRow, PBYINDEX) = ""
                grdPostedInfo.TextMatrix(llRow, PIPINDEX) = ""
                'dan M added for Marketron
                slLatestDate = "1999-01-01"
                If rst_att!attPostingType > 1 Then
                    If ilSchdCount > 0 Then
                        If rbcCompliant(1).Value Then
                            grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = (100 * CLng(ilAgyCompliantCount)) \ ilSchdCount
                        Else
                            grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = (100 * CLng(ilPledgeCompliantCount)) \ ilSchdCount
                        End If
                    End If
                    SQLQuery = "SELECT * FROM webl"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " weblType = 1"
                    SQLQuery = SQLQuery & " AND weblPostDay >= '" & Format$(llCpttDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND weblPostDay <= '" & Format$(llCpttDate + 6, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND weblAttCode = " & rst_Cptt!cpttatfCode & ")"
                    SQLQuery = SQLQuery & " Order by weblDate desc, weblTime desc"
                    Set rst_Webl = gSQLSelectCall(SQLQuery)
                    If Not rst_Webl.EOF Then
                        grdPostedInfo.TextMatrix(llRow, PPOSTDATEINDEX) = Format$(Trim$(rst_Webl!webldate), sgShowDateForm)
                        grdPostedInfo.TextMatrix(llRow, PBYINDEX) = Trim$(rst_Webl!weblUserName)
                        grdPostedInfo.TextMatrix(llRow, PIPINDEX) = Trim$(rst_Webl!weblIP)
                    End If
                End If
                'Dan M 11/04/10 add test for Marketron--use latest date
                '7701
                If blIsMarketron Then
'                If rst_att!attExportToMarketron = "Y" Then
                    If ilSchdCount > 0 Then
                        'grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = (100 * CLng(ilPledgeCompliantCount)) \ ilSchdCount
                        If rbcCompliant(1).Value Then
                            grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = (100 * CLng(ilAgyCompliantCount)) \ ilSchdCount
                        Else
                            grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = (100 * CLng(ilPledgeCompliantCount)) \ ilSchdCount
                        End If
                    End If
                    SQLQuery = "SELECT * FROM webl"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " weblType = 3"
                    SQLQuery = SQLQuery & " AND weblPostDay >= '" & Format$(llCpttDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND weblPostDay <= '" & Format$(llCpttDate + 6, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND weblAttCode = " & rst_Cptt!cpttatfCode & ")"
                    SQLQuery = SQLQuery & " Order by weblDate desc, weblTime desc"
                    Set rst_Webl = gSQLSelectCall(SQLQuery)
                    If Not rst_Webl.EOF Then
                        If DateDiff("d", slLatestDate, rst_Webl!webldate) > 0 Then
                            grdPostedInfo.TextMatrix(llRow, PPOSTDATEINDEX) = Format$(Trim$(rst_Webl!webldate), sgShowDateForm)
                            grdPostedInfo.TextMatrix(llRow, PBYINDEX) = Trim$(rst_Webl!weblUserName)
                            grdPostedInfo.TextMatrix(llRow, PIPINDEX) = Trim$(rst_Webl!weblIP)
                        End If
                    End If
                End If
                grdPostedInfo.Row = llRow
                grdPostedInfo.Col = PSTATUSINDEX
                grdPostedInfo.CellBackColor = GRAY
                If rst_Cptt!cpttStatus = 2 Then
                    grdPostedInfo.CellBackColor = vbBlue
                ElseIf rst_Cptt!cpttStatus = 0 Then
                    If rst_Cptt!cpttPostingStatus = 1 Then
                        grdPostedInfo.CellBackColor = vbMagenta
                    Else
                        'Posting Type: 0=Receipt Only; 1=Spot Count; 2=Post Dates/Times
                        If rst_att!attPostingType > 1 Then
                            grdPostedInfo.TextMatrix(llRow, PNOAIREDINDEX) = ""
                            grdPostedInfo.TextMatrix(llRow, PNOCMPLINDEX) = ""
                            grdPostedInfo.TextMatrix(llRow, PPERCENTAINDEX) = ""
                            grdPostedInfo.TextMatrix(llRow, PPERCENTCINDEX) = ""
                            llWeekDate = gDateValue(gObtainPrevMonday(gAdjYear(Format$(llCpttDate, "m/d/yy"))))
                            SQLQuery = "SELECT * FROM ast"
                            SQLQuery = SQLQuery + " WHERE ("
                            SQLQuery = SQLQuery + " astFeedDate >= '" & Format(llWeekDate, sgSQLDateForm) & "'"
                            SQLQuery = SQLQuery + " AND astFeedDate <= '" & Format(llWeekDate + 6, sgSQLDateForm) & "'"
                            SQLQuery = SQLQuery & " AND astatfCode = " & lmAttCode & ")"
                            Set rst_Ast = gSQLSelectCall(SQLQuery)
                            If Not rst_Ast.EOF Then
                                grdPostedInfo.CellBackColor = vbRed
                            Else
                                grdPostedInfo.CellBackColor = BROWN
                            End If
                        Else
                            grdPostedInfo.CellBackColor = vbRed
                        End If
                    End If
                Else
                    'Missing:  Compliant and one not aired
                    If rst_Cptt!cpttPostingStatus = 2 Then
                        If ilAnyNotCompliant Then
                            grdPostedInfo.CellBackColor = LIGHTBLUECOLOR
                        Else
                            grdPostedInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                        End If
                    ElseIf rst_Cptt!cpttPostingStatus = 0 Then
                        If rst_att!attPostingType > 1 Then
                            grdPostedInfo.CellBackColor = vbRed
                        Else
                            grdPostedInfo.CellBackColor = MIDGREENCOLOR
                        End If
                    ElseIf rst_Cptt!cpttPostingStatus = 1 Then
                        grdPostedInfo.CellBackColor = vbMagenta
                    End If
                End If
                grdPostedInfo.TextMatrix(llRow, PSTATUSINDEX) = ""
                grdPostedInfo.TextMatrix(llRow, PCPTTINDEX) = rst_Cptt!cpttCode
                llRow = llRow + 1
            End If
        End If
        rst_Cptt.MoveNext
    Loop
    If blTaskBlocked Then
        gMsgBox "*** Some Station(s) were blocked during the generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""

    'If grdPostedInfo.Rows < 2 * ((cmcDone.Top - grdPostedInfo.Top) \ grdPostedInfo.RowHeight(1)) Then
    '    grdPostedInfo.Rows = 2 * grdPostedInfo.Rows - 2 * ((cmcDone.Top - grdPostedInfo.Top) \ grdPostedInfo.RowHeight(1))
    'Else
    '    grdPostedInfo.Rows = grdPostedInfo.Rows + grdPostedInfo.Height \ grdPostedInfo.RowHeight(1) + 1
    'End If
    grdPostedInfo.Rows = grdPostedInfo.Rows + ((cmcDone.Top - grdPostedInfo.Top) \ grdPostedInfo.RowHeight(1))
    For llYellowRow = llRow To grdPostedInfo.Rows - 1 Step 1
        grdPostedInfo.Row = llYellowRow
        For llCol = PSELECTINDEX To PSTATUSINDEX Step 1
            grdPostedInfo.Col = llCol
            grdPostedInfo.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llYellowRow
    imLastPostedInfoSort = -1
    imLastPostedInfoColSorted = -1
    grdPostedInfo.Redraw = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopPostedInfoGrid"
    grdPostedInfo.Redraw = True
    '1/11/13: bypass Update CPTT SQL call if error
End Function

Private Function mPopSpotInfoGrid() As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim llCpttDate As Long
    Dim llWeekDate As Long
    Dim slMoDate As String
    Dim llAdf As Long
    Dim ilAdfCode As Integer
    Dim ilRet As Integer
    Dim ilAst As Integer
    Dim llVeh As Long
    Dim ilShtt As Integer
    Dim ilCompliant As Integer
    Dim ilAnyAired As Integer
    Dim llYellowRow As Long
    Dim blProcessAst As Boolean
    Dim ilAstStatus As Integer
    Dim slAstPledgeDate As String
    Dim ilAstPledgeStatus As Integer
    Dim slPledgeDays As String
    Dim slAstAirDate As String
    Dim slAstAirTime As String
    Dim slAstPledgeStartTime As String
    Dim slAstTruePledgeEndTime As String
    Dim llMonPdDate As Long
    Dim ilUpper As Integer
    Dim ilStatus As Integer
    Dim slCallLetters As String
    
    mPopSpotInfoGrid = False
    On Error GoTo ErrHand:
    grdSpotInfo.Rows = 2
    mClearGrid grdSpotInfo
    If (grdPostedInfo.Row < grdPostedInfo.FixedRows) Or (grdPostedInfo.Row > grdPostedInfo.Rows) Then
        Exit Function
    End If
    If (grdPostedInfo.Col < PSELECTINDEX) Or (grdPostedInfo.Col > PSTATUSINDEX) Then
        Exit Function
    End If
    If Trim$(grdPostedInfo.TextMatrix(grdPostedInfo.Row, PWEEKINDEX)) = "" Then
        Exit Function
    End If
    grdSpotInfo.Redraw = False
    lgSelGameGsfCode = 0
    grdSpotInfo.Row = 0
    grdSpotInfo.Col = DDATEINDEX
    grdSpotInfo.CellBackColor = LIGHTBLUE
    grdSpotInfo.Col = DADVTINDEX
    grdSpotInfo.CellBackColor = LIGHTBLUE
    grdSpotInfo.Col = DISCIINDEX
    grdSpotInfo.CellBackColor = LIGHTBLUE
    llRow = grdSpotInfo.FixedRows
    llWeekDate = gDateValue(Trim$(grdPostedInfo.TextMatrix(grdPostedInfo.Row, PWEEKINDEX)))
    slMoDate = gObtainPrevMonday(Format$(llWeekDate, "m/d/yy"))
    SQLQuery = "SELECT * FROM cptt"
    SQLQuery = SQLQuery + " WHERE ("
    SQLQuery = SQLQuery & " cpttCode = " & grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PCPTTINDEX) & ")"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If Not rst_Cptt.EOF Then
        llVeh = gBinarySearchVef(CLng(rst_Cptt!cpttvefcode))
        If llVeh <> -1 Then
            If tgVehicleInfo(llVeh).sVehType = "G" Then
                igGameVefCode = rst_Cptt!cpttvefcode
                sgGameStartDate = slMoDate
                sgGameEndDate = DateAdd("d", 6, sgGameStartDate)
                lgGameAttCode = rst_Cptt!cpttatfCode
                frmGetGame.Show vbModal
                If lgSelGameGsfCode <= 0 Then
                    grdSpotInfo.Redraw = True
                    Exit Function
                End If
            End If
        End If
        ReDim tgCPPosting(0 To 1) As CPPOSTING
        tgCPPosting(0).lCpttCode = rst_Cptt!cpttCode
        tgCPPosting(0).iStatus = rst_Cptt!cpttStatus
        tgCPPosting(0).iPostingStatus = rst_Cptt!cpttPostingStatus
        tgCPPosting(0).lAttCode = rst_Cptt!cpttatfCode
        tgCPPosting(0).iAttTimeType = 0 'Not used
        tgCPPosting(0).iVefCode = rst_Cptt!cpttvefcode  'imVefCode
        tgCPPosting(0).iShttCode = rst_Cptt!cpttshfcode
        ilShtt = gBinarySearchStationInfoByCode(imShttCode)
        If ilShtt <> -1 Then
            tgCPPosting(0).sZone = tgStationInfoByCode(ilShtt).sZone
            slCallLetters = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
        Else
            tgCPPosting(0).sZone = ""
            slCallLetters = "Missing: " & imShttCode
        End If
        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
        tgCPPosting(0).sAstStatus = rst_Cptt!cpttAstStatus
        bgTaskBlocked = False
        sgTaskBlockedName = "Affiliate Management"
        igTimes = 1 'By Week
        ilAdfCode = -1
        ''Dan M 6442 9/26/13
        'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, False, True)
        ''ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True)
        If lgSelGameGsfCode <= 0 Then
            'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, False, False, True, , , , , , True)
            ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, False, True, , , , , , True)
        Else
            'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmGameAstInfo(), -1, False, False, True, , , , , , True)
            ilRet = gGetAstInfo(hmAst, tmCPDat(), tmGameAstInfo(), -1, True, False, True, , , , , , True)
            ReDim tmAstInfo(0 To UBound(tmGameAstInfo)) As ASTINFO
            ilUpper = 0
            For ilAst = 0 To UBound(tmGameAstInfo) - 1 Step 1
                If tmGameAstInfo(ilAst).lgsfCode = lgSelGameGsfCode Then
                    tmAstInfo(ilUpper) = tmGameAstInfo(ilAst)
                    ilUpper = ilUpper + 1
                End If
            Next ilAst
            ReDim Preserve tmAstInfo(0 To ilUpper) As ASTINFO
        End If
        '2/5/18: add block message
        If bgTaskBlocked Then
            bgTaskBlocked = False
            sgTaskBlockedName = ""
            gMsgBox "*** " & slCallLetters & ": Spots not obtained as blocked", vbCritical
            grdSpotInfo.Redraw = True
            Exit Function
        End If
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        
        If rst_Cptt!cpttStatus = 2 Then
            ilAnyAired = False
            For ilAst = 0 To UBound(tmAstInfo) - 1 Step 1
                'If (tmAstInfo(ilAst).iStatus = 0) Or (tmAstInfo(ilAst).iStatus = 1) Or (tmAstInfo(ilAst).iStatus = 6) Or (tmAstInfo(ilAst).iStatus = 7) Or (tmAstInfo(ilAst).iStatus = 9) Or (tmAstInfo(ilAst).iStatus = 10) Or (tmAstInfo(ilAst).iStatus = 20) Or (tmAstInfo(ilAst).iStatus = 21) Or (tmAstInfo(ilAst).iStatus = 22) Then
                ilStatus = gGetAirStatus(tmAstInfo(ilAst).iStatus)
                If (ilStatus = 0) Or (ilStatus = 1) Or (ilStatus = 6) Or (ilStatus = 7) Or (ilStatus = 9) Or (ilStatus = 10) Or (ilStatus = ASTEXTENDED_MG) Or (ilStatus = ASTEXTENDED_BONUS) Or (ilStatus = ASTEXTENDED_REPLACEMENT) Then
                    ilAnyAired = True
                    Exit For
                End If
            Next ilAst
        Else
            ilAnyAired = True
        End If
        If sgStationSearchCallSource = "P" Then
            If lbcAdvertiser.ListIndex >= 0 Then
                ilAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
            End If
        End If
        For ilAst = 0 To UBound(tmAstInfo) - 1 Step 1
            blProcessAst = True
            If sgStationSearchCallSource = "P" Then
                If ilAdfCode <> tmAstInfo(ilAst).iAdfCode Then
                    blProcessAst = False
                End If
            End If
            If (tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iPledgeStatus)).iPledged <> 2) And blProcessAst Then
                If llRow >= grdSpotInfo.Rows Then
                    grdSpotInfo.AddItem ""
                End If
                grdSpotInfo.Row = llRow
                For llCol = DDATEINDEX To DCOMMENTINDEX Step 1
                    grdSpotInfo.Col = llCol
                    grdSpotInfo.CellBackColor = LIGHTYELLOW
                Next llCol
                grdSpotInfo.TextMatrix(llRow, DDATEINDEX) = Format$(Trim$(tmAstInfo(ilAst).sFeedDate), "m/d/yy")
                grdSpotInfo.TextMatrix(llRow, DFEDINDEX) = Trim$(tmAstInfo(ilAst).sFeedTime)
                'Missing
                grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = ""
                If tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iPledgeStatus)).iPledged = 0 Then   'Live
                    grdSpotInfo.TextMatrix(llRow, DPLEGDEDDAYINDEX) = Trim$(tmAstInfo(ilAst).sPdDays)
                    If Trim$(tmAstInfo(ilAst).sPledgeEndTime) = "" Then
                        grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sPledgeStartTime) '& "-" & Trim$(tmAstInfo(ilAst).sPledgeEndTime)
                    Else
                        grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sPledgeStartTime) & "-" & Trim$(tmAstInfo(ilAst).sPledgeEndTime)
                    End If
                ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iPledgeStatus)).iPledged = 1 Then    'Delayed
                    grdSpotInfo.TextMatrix(llRow, DPLEGDEDDAYINDEX) = Trim$(tmAstInfo(ilAst).sPdDays)
                    grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sPledgeStartTime) & "-" & Trim$(tmAstInfo(ilAst).sPledgeEndTime)
                ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iPledgeStatus)).iPledged = 2 Then    'Not Carried
                    grdSpotInfo.TextMatrix(llRow, DPLEGDEDDAYINDEX) = ""
                    grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = ""
                ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iPledgeStatus)).iPledged = 3 Then    'Not Pledged
                    If tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iPledgeStatus)).iStatus = 6 Then
                        grdSpotInfo.TextMatrix(llRow, DPLEGDEDDAYINDEX) = Trim$(tmAstInfo(ilAst).sPdDays)
                        If Trim$(tmAstInfo(ilAst).sPledgeEndTime) = "" Then
                            grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sPledgeStartTime) '& "-" & Trim$(tmAstInfo(ilAst).sPledgeEndTime)
                        Else
                            grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sPledgeStartTime) & "-" & Trim$(tmAstInfo(ilAst).sPledgeEndTime)
                        End If
                    Else
                        grdSpotInfo.TextMatrix(llRow, DPLEGDEDDAYINDEX) = ""
                        grdSpotInfo.TextMatrix(llRow, DPLEGDEDTIMEINDEX) = "None"
                    End If
                End If
                'grdSpotInfo.TextMatrix(llRow, DDATEINDEX) = Format$(Trim$(rst_Ast!astFeedDate), sgShowDateForm)
                grdSpotInfo.TextMatrix(llRow, DAIREDDATEINDEX) = ""
                grdSpotInfo.TextMatrix(llRow, DAIREDTIMEINDEX) = ""
                ilStatus = gGetAirStatus(tmAstInfo(ilAst).iStatus)
                If (ilStatus <= 10) Or (ilStatus = ASTAIR_MISSED_MG_BYPASS) Then
                    If (tmAstInfo(ilAst).iCPStatus = 1) And ((tgStatusTypes(ilStatus).iPledged <= 1) Or (ilStatus = 6) Or (ilStatus = 7)) Then
                        'If DateValue(Trim$(tmAstInfo(ilAst).sFeedDate)) <> DateValue(Trim$(tmAstInfo(ilAst).sAirDate)) Then
                            grdSpotInfo.TextMatrix(llRow, DAIREDDATEINDEX) = Format$(Trim$(tmAstInfo(ilAst).sAirDate), "m/d/yy")
                        'End If
                        grdSpotInfo.TextMatrix(llRow, DAIREDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sAirTime)
                    End If
                Else
                    grdSpotInfo.TextMatrix(llRow, DAIREDDATEINDEX) = Format$(Trim$(tmAstInfo(ilAst).sAirDate), "m/d/yy")
                    grdSpotInfo.TextMatrix(llRow, DAIREDTIMEINDEX) = Trim$(tmAstInfo(ilAst).sAirTime)
                End If
                llAdf = gBinarySearchAdf(CLng(tmAstInfo(ilAst).iAdfCode))
                If llAdf <> -1 Then
                    grdSpotInfo.TextMatrix(llRow, DADVTINDEX) = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                Else
                    grdSpotInfo.TextMatrix(llRow, DADVTINDEX) = ""
                End If
                grdSpotInfo.TextMatrix(llRow, DLENGTHINDEX) = Trim$(Str(tmAstInfo(ilAst).iLen))
                If tmAstInfo(ilAst).iRegionType <= 0 Then
                    grdSpotInfo.TextMatrix(llRow, DPRODINDEX) = Trim$(tmAstInfo(ilAst).sProd)
                    grdSpotInfo.TextMatrix(llRow, DISCIINDEX) = Trim$(tmAstInfo(ilAst).sISCI)
                    grdSpotInfo.TextMatrix(llRow, DCARTINDEX) = Trim$(tmAstInfo(ilAst).sCart)
                Else
                    grdSpotInfo.TextMatrix(llRow, DPRODINDEX) = Trim$(tmAstInfo(ilAst).sRProduct)
                    grdSpotInfo.Col = DISCIINDEX
                    grdSpotInfo.CellBackColor = LIGHTMAGENTACOLOR   'RGB(247, 209, 255)
                    grdSpotInfo.TextMatrix(llRow, DISCIINDEX) = Trim$(tmAstInfo(ilAst).sRISCI)
                    grdSpotInfo.TextMatrix(llRow, DCARTINDEX) = Trim$(tmAstInfo(ilAst).sRCart)
                End If
                'If tmAstInfo(ilAst).iStatus = 20 Then
                If ilStatus = ASTEXTENDED_MG Then
                    grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Makegood"
                'ElseIf tmAstInfo(ilAst).iStatus = 21 Then
                ElseIf ilStatus = ASTEXTENDED_BONUS Then
                    grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Bonus"
                ElseIf ilStatus = ASTEXTENDED_REPLACEMENT Then
                    grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Replacement"
                'D.S. 4/5/12
                ElseIf ilStatus = ASTAIR_MISSED_MG_BYPASS Then
                    grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Missed MG Bypassed"
                ElseIf (ilStatus = 2) Or (ilStatus = 3) Or (ilStatus = 4) Or (ilStatus = 5) Or (ilStatus = 8) Then
                    'Missing: Missed Reason from Web Site
                    If tmAstInfo(ilAst).iCPStatus = 1 Then
                        If tmAstInfo(ilAst).iRegionType = 2 Then   'Blackout
                            grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Missed: Blackout"   '"Missed"
                        ElseIf tmAstInfo(ilAst).iSpotType = 2 Then 'Fill
                            grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Missed: Fill"   '"Missed"
                        ElseIf tmAstInfo(ilAst).iRegionType = 1 Then 'Region Copy
                            grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Missed: Region Copy"   '"Missed"
                        Else
                            grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = "Missed"   '"Missed"
                        End If
                    Else
                        grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = ""   '"Missed"
                    End If
                Else
                    grdSpotInfo.TextMatrix(llRow, DCOMMENTINDEX) = ""
                End If
                grdSpotInfo.Col = DSTATUSINDEX
                If ilStatus <= 10 Then
                    If tmAstInfo(ilAst).iCPStatus = 1 Then
                        'If (tmAstInfo(ilAst).iStatus = 2) Or (tmAstInfo(ilAst).iStatus = 3) Or (tmAstInfo(ilAst).iStatus = 4) Or (tmAstInfo(ilAst).iStatus = 5) Or (tmAstInfo(ilAst).iStatus = 8) Then
                        '    'Non-Compliant
                        '    ilCompliant = False
                        'ElseIf Weekday(Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy")) <> Weekday(Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy")) Then
                        '    'Non-Compliant
                        '    ilCompliant = False
                        'ElseIf (gTimeToLong(Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), False) < gTimeToLong(Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), False)) Or (gTimeToLong(Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), False) > gTimeToLong(Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), True)) Then
                        '    If (tmAstInfo(ilAst).iStatus = 0) Or (tmAstInfo(ilAst).iStatus = 1) Or (tmAstInfo(ilAst).iStatus = 6) Or (tmAstInfo(ilAst).iStatus = 7) Or (tmAstInfo(ilAst).iStatus = 9) Or (tmAstInfo(ilAst).iStatus = 10) Then
                        '        'Non-Compliant
                        '        ilCompliant = False
                        '    Else
                        '        ilCompliant = True
                        '    End If
                        'Else
                        '    ilCompliant = True
                        'End If
                        ilAstStatus = ilStatus  'tmAstInfo(ilAst).iStatus
                        slAstPledgeDate = tmAstInfo(ilAst).sPledgeDate
                        ilAstPledgeStatus = tmAstInfo(ilAst).iPledgeStatus
                        slPledgeDays = tmAstInfo(ilAst).sTruePledgeDays
                        slAstAirDate = tmAstInfo(ilAst).sAirDate
                        slAstAirTime = tmAstInfo(ilAst).sAirTime
                        slAstPledgeStartTime = tmAstInfo(ilAst).sPledgeStartTime
                        slAstTruePledgeEndTime = tmAstInfo(ilAst).sTruePledgeEndTime
                        If (gGetAirStatus(ilAstStatus) = 6) Or (gGetAirStatus(ilAstStatus) = 7) Then
                            ilAstStatus = 1
                        End If
                        'If (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2) Or (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3) Then
                        '    'Non-Compliant
                        '    ilCompliant = False
                        'ElseIf (Weekday(Format$(slAstPledgeDate, "m/d/yy")) <> Weekday(Format$(slAstAirDate, "m/d/yy"))) And (tgStatusTypes(gGetAirStatus(ilAstPledgeStatus)).iPledged = 0) Then
                        ''Test if aired on Pledge day
                        '    ilCompliant = False
                        'ElseIf (Mid$(slPledgeDays, Weekday(Format$(slAstAirDate, "m/d/yy"), vbMonday), 1) <> "Y") And (tgStatusTypes(gGetAirStatus(ilAstPledgeStatus)).iPledged = 1) Then
                        '    'Non-Compliant
                        '    ilCompliant = False
                        ''Test if aired within pledge times
                        'ElseIf (gTimeToLong(Format$(slAstAirTime, "h:mm:ssAM/PM"), False) < gTimeToLong(Format$(slAstPledgeStartTime, "h:mm:ssAM/PM"), False)) Or (gTimeToLong(Format$(slAstAirTime, "h:mm:ssAM/PM"), False) > gTimeToLong(Format$(slAstTruePledgeEndTime, "h:mm:ssAM/PM"), True)) Then
                        '    'If (ilInAstStatus = 0) Or (ilInAstStatus = 1) Or (ilInAstStatus = 6) Or (ilInAstStatus = 7) Or (ilInAstStatus = 9) Or (ilInAstStatus = 10) Then
                        '    '    'Non-Compliant
                        '    'Else
                        '    '    ilOutCompliantCount = ilOutCompliantCount + 1
                        '    'End If
                        '    ilCompliant = False
                        'Else
                        '    'Test if in correct week
                        '    llMonPdDate = DateValue(gObtainPrevMonday(slAstPledgeDate))
                        '    If (DateValue(slAstAirDate) >= llMonPdDate) And (DateValue(slAstAirDate) <= llMonPdDate + 6) Then
                        '        ilCompliant = True
                        '    Else
                        '        ilCompliant = False
                        '    End If
                        'End If
                        If rbcCompliant(1).Value Then
                            If (tmAstInfo(ilAst).sAgencyCompliant = "N") Then
                                ilCompliant = False
                            Else
                                ilCompliant = True
                            End If
                        Else
                            If (tmAstInfo(ilAst).sStationCompliant = "N") Then
                                ilCompliant = False
                            Else
                                ilCompliant = True
                            End If
                        End If
                    Else
                        ilCompliant = False
                    End If
                Else
                    ilCompliant = True
                End If
                If ilAnyAired Then
                    If tmAstInfo(ilAst).iCPStatus = 1 Then
                        If ilStatus <= 10 Or ilStatus = ASTAIR_MISSED_MG_BYPASS Then
                            If tgStatusTypes(ilStatus).iPledged = 0 Then   'Live
                                If ilCompliant Then
                                    grdSpotInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                                Else
                                    grdSpotInfo.CellBackColor = LIGHTBLUECOLOR    'Orange
                                End If
                            ElseIf tgStatusTypes(ilStatus).iPledged = 1 Then    'Delayed
                                If ilCompliant Then
                                    grdSpotInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                                Else
                                    grdSpotInfo.CellBackColor = LIGHTBLUECOLOR    'Orange
                                End If
                            ElseIf tgStatusTypes(ilStatus).iPledged = 2 Then    'Not Carried
                                'grdSpotInfo.CellBackColor = vbBlue
                                If ilCompliant Then
                                    grdSpotInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                                Else
                                    grdSpotInfo.CellBackColor = LIGHTBLUECOLOR    'Orange
                                End If
                                'grdSpotInfo.CellBackColor = vbBlue 'ORANGECOLOR    'vbRed
                            ElseIf tgStatusTypes(ilStatus).iPledged = 3 Then    'Not Pledged
                                If ilStatus = 6 Then
                                    grdSpotInfo.CellBackColor = LIGHTBLUECOLOR    'Orange
                                ElseIf ilStatus = 7 Then
                                    grdSpotInfo.CellBackColor = LIGHTBLUECOLOR    'Orange
                                Else
                                    'grdSpotInfo.CellBackColor = vbBlue 'ORANGECOLOR    'vbRed
                                    If ilCompliant Then
                                        grdSpotInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                                    Else
                                        grdSpotInfo.CellBackColor = LIGHTBLUECOLOR    'Orange
                                    End If
                                End If
                            End If
                        Else
                            'Missing:  Color for MG and Bonus
                            grdSpotInfo.CellBackColor = MIDGREENCOLOR    'LIGHTGREEN
                        End If
                    ElseIf tmAstInfo(ilAst).iCPStatus = 0 Then
                        llCpttDate = gDateValue(rst_Cptt!CpttStartDate)
                        llWeekDate = gDateValue(gObtainPrevMonday(gAdjYear(Format$(llCpttDate, "m/d/yy"))))
                        SQLQuery = "SELECT * FROM ast"
                        SQLQuery = SQLQuery + " WHERE ("
                        SQLQuery = SQLQuery + " astFeedDate >= '" & Format(llWeekDate, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery + " AND astFeedDate <= '" & Format(llWeekDate + 6, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery & " AND astatfCode = " & rst_Cptt!cpttatfCode & ")"
                        Set rst_Ast = gSQLSelectCall(SQLQuery)
                        If Not rst_Ast.EOF Then
                            grdSpotInfo.CellBackColor = vbRed   'GRAY
                        Else
                            grdSpotInfo.CellBackColor = BROWN   'GRAY
                        End If
                    Else
                        grdSpotInfo.CellBackColor = vbBlue
                    End If
                Else
                    grdSpotInfo.CellBackColor = vbBlue
                End If
                grdSpotInfo.TextMatrix(llRow, DCNTRNOINDEX) = tmAstInfo(ilAst).lCntrNo
                If tmAstInfo(ilAst).lgsfCode > 0 Then
                    grdSpotInfo.TextMatrix(llRow, DGAMEINFOINDEX) = mGetGameInfo(tmAstInfo(ilAst).lgsfCode)
                Else
                    grdSpotInfo.TextMatrix(llRow, DGAMEINFOINDEX) = ""
                End If
                smMissedFeedDate = grdSpotInfo.TextMatrix(llRow, DDATEINDEX)
                smMissedFeedTime = grdSpotInfo.TextMatrix(llRow, DFEDINDEX)
                
                grdSpotInfo.TextMatrix(llRow, DMMRINFOINDEX) = mGetMGAndMissedInfo(ilAst)
                
                If (ilStatus = ASTEXTENDED_MG) Or (ilStatus = ASTEXTENDED_REPLACEMENT) Then
                    grdSpotInfo.TextMatrix(llRow, DDATEINDEX) = smMissedFeedDate
                    grdSpotInfo.TextMatrix(llRow, DFEDINDEX) = smMissedFeedTime
                ElseIf (tmAstInfo(ilAst).lLkAstCode > 0) And (smMGAirDate <> "") Then
                    grdSpotInfo.TextMatrix(llRow, DAIREDDATEINDEX) = "MG:" & smMGAirDate
                    grdSpotInfo.TextMatrix(llRow, DAIREDTIMEINDEX) = "@ " & smMGAirTime
                End If
                grdSpotInfo.TextMatrix(llRow, DASTCODEINDEX) = tmAstInfo(ilAst).lCode
                llRow = llRow + 1
            End If
        Next ilAst
    End If
    'If grdSpotInfo.Rows < grdSpotInfo.Height \ grdSpotInfo.RowHeight(1) Then
    '    grdSpotInfo.Rows = grdSpotInfo.Rows + 2 * (grdSpotInfo.Height \ grdSpotInfo.RowHeight(1)) - grdSpotInfo.Rows + 1
    'Else
    '    grdSpotInfo.Rows = grdSpotInfo.Rows + grdSpotInfo.Height \ grdSpotInfo.RowHeight(1) + 1
    'End If
    grdSpotInfo.Rows = grdSpotInfo.Rows + ((cmcDone.Top - grdSpotInfo.Top) \ grdSpotInfo.RowHeight(1))
    For llYellowRow = llRow To grdSpotInfo.Rows - 1 Step 1
        grdSpotInfo.Row = llYellowRow
        For llCol = DDATEINDEX To DSTATUSINDEX Step 1
            grdSpotInfo.Col = llCol
            grdSpotInfo.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llYellowRow
    imLastSpotInfoSort = -1
    imLastSpotInfoColSorted = -1
    mPopSpotInfoGrid = True
    grdSpotInfo.Redraw = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopSpotInfoGrid"
End Function


Private Function mGetTitle(ilCode As Integer) As String
    Dim rst_tnt As ADODB.Recordset

    mGetTitle = ""
    SQLQuery = "Select tntCode, tntTitle From Tnt where tntCode = " & ilCode
    Set rst_tnt = gSQLSelectCall(SQLQuery)
    If Not rst_tnt.EOF Then
        mGetTitle = Trim$(rst_tnt!tntTitle)
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mGetTitle"
End Function


Private Sub mSetGridPosition()
    Dim llRow As Long
    Dim llKey As Long
    Dim llStationRow As Long
    
    lmStationTopRow = grdStations.TopRow
    lmAgreementTopRow = grdAgreementInfo.TopRow
    lmPostedTopRow = grdPostedInfo.TopRow
    If grdSpotInfo.Visible Then
        grdStations.Height = grdStations.RowHeight(0) + grdStations.RowHeight(1)
        vbcStation.Height = grdStations.Height
        grdAgreementInfo.Height = grdAgreementInfo.RowHeight(0) + grdAgreementInfo.RowHeight(1)
        grdPostedInfo.Height = grdPostedInfo.RowHeight(0) + grdPostedInfo.RowHeight(1)
        If udcContactGrid.Visible And udcCommentGrid.Visible Then
            mShareThird grdPostedInfo, grdSpotInfo
        ElseIf udcCommentGrid.Visible Then
            mShareHalfComment grdPostedInfo, grdSpotInfo
        Else
            mShareNone grdPostedInfo, grdSpotInfo
        End If
        bmStationScrollAllowed = False
        bmAgreementScrollAllowed = False
        bmPostedScrollAllowed = False
        mFindAndDisplay grdStations.Row
        Exit Sub
    End If
    If grdPostedInfo.Visible Then
        grdStations.Height = grdStations.RowHeight(0) + grdStations.RowHeight(1)
        vbcStation.Height = grdStations.Height
        grdAgreementInfo.Height = grdAgreementInfo.RowHeight(0) + grdAgreementInfo.RowHeight(1)
        If udcContactGrid.Visible And udcCommentGrid.Visible Then
            mShareThird grdAgreementInfo, grdPostedInfo
        ElseIf udcCommentGrid.Visible Then
            mShareHalfComment grdAgreementInfo, grdPostedInfo
        Else
            mShareNone grdAgreementInfo, grdPostedInfo
        End If
        bmStationScrollAllowed = False
        bmAgreementScrollAllowed = False
        bmPostedScrollAllowed = True
        mFindAndDisplay grdStations.Row
        Exit Sub
    End If
    If grdAgreementInfo.Visible Then
        grdStations.Height = grdStations.RowHeight(0) + grdStations.RowHeight(1)
        vbcStation.Height = grdStations.Height
        If udcContactGrid.Visible And udcCommentGrid.Visible Then
            mShareThird grdStations, grdAgreementInfo
        ElseIf udcCommentGrid.Visible Then
            mShareHalfComment grdStations, grdAgreementInfo
        Else
            mShareNone grdStations, grdAgreementInfo
        End If
        bmStationScrollAllowed = False
        bmAgreementScrollAllowed = True
        bmPostedScrollAllowed = True
        mFindAndDisplay grdStations.Row
        Exit Sub
    End If
    If udcContactGrid.Visible And udcCommentGrid.Visible Then
        grdStations.Height = grdStations.RowHeight(0) + grdStations.RowHeight(1)
        vbcStation.Height = grdStations.Height
        udcContactGrid.Move grdStations.Left - pbcArrow.Width, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1), grdStations.Width + pbcArrow.Width, (cmcDone.Top - (grdStations.Top + grdStations.Height)) / 2 - 75
        udcContactGrid.Visible = True
        udcContactGrid.ZOrder
        udcCommentGrid.Move grdStations.Left - pbcArrow.Width, udcContactGrid.Top + udcContactGrid.Height, grdStations.Width + pbcArrow.Width, cmcDone.Top - (udcContactGrid.Top + udcContactGrid.Height) - 75
        udcCommentGrid.Visible = True
        udcCommentGrid.ZOrder
        mSetCommentButtons
        bmStationScrollAllowed = False
        bmAgreementScrollAllowed = True
        bmPostedScrollAllowed = True
    ElseIf udcContactGrid.Visible Then
        mHideCommentButtons
        grdStations.Height = grdStations.RowHeight(0) + grdStations.RowHeight(1)
        gGrid_IntegralHeight grdStations
        vbcStation.Height = grdStations.Height
        udcContactGrid.Move grdStations.Left - pbcArrow.Width, grdStations.Top + grdStations.RowHeight(0) + grdStations.RowHeight(1), grdStations.Width + pbcArrow.Width, cmcDone.Top - (grdStations.Top + grdStations.Height) - 75
        udcContactGrid.Visible = True
        udcContactGrid.ZOrder
        bmStationScrollAllowed = False
        bmAgreementScrollAllowed = True
        bmPostedScrollAllowed = True
    ElseIf udcCommentGrid.Visible Then
        grdStations.Height = (cmcDone.Top - grdStations.Top) / 2 - 75
        gGrid_IntegralHeight grdStations
        vbcStation.Height = grdStations.Height
        udcCommentGrid.Move grdStations.Left - pbcArrow.Width, grdStations.Top + grdStations.Height, grdStations.Width + pbcArrow.Width, cmcDone.Top - (grdStations.Top + grdStations.Height) - 75
        udcCommentGrid.Visible = True
        udcCommentGrid.ZOrder
        mSetCommentButtons
        bmStationScrollAllowed = True
        bmAgreementScrollAllowed = True
        bmPostedScrollAllowed = True
    Else
        mHideCommentButtons
        grdStations.Height = cmcDone.Top - 180
        If sgStationSearchCallSource = "P" Then
            grdStations.Height = grdStations.Height - 2 * grdStations.RowHeight(0)
        End If
        gGrid_IntegralHeight grdStations
        grdStations.Height = grdStations.Height + 30
        vbcStation.Height = grdStations.Height
        For llRow = grdStations.FixedRows To grdStations.Rows - 1 Step 1
            If grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "s" Then
                grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "t"
            End If
        Next llRow
        For llKey = LBound(tmStationGridKey) To UBound(tmStationGridKey) - 1 Step 1
            llStationRow = tmStationGridKey(llKey).lRow
            If smStationGridData(llStationRow + SSELECTINDEX) = "s" Then
                smStationGridData(llStationRow + SSELECTINDEX) = "t"
            End If
        Next llKey
        bmStationScrollAllowed = True
        bmAgreementScrollAllowed = True
        bmPostedScrollAllowed = True
    End If
    mFindAndDisplay grdStations.Row

End Sub

Private Sub mShareNone(grdCtrlAbove As MSHFlexGrid, grdCtrl As MSHFlexGrid)
    mHideCommentButtons
    grdCtrl.Move grdStations.Left, grdCtrlAbove.Top + grdCtrlAbove.RowHeight(0) + grdCtrlAbove.RowHeight(1), grdStations.Width, cmcDone.Top - (grdCtrlAbove.Top + grdCtrlAbove.Height) - 75
    gGrid_IntegralHeight grdCtrl
    grdCtrl.Height = grdCtrl.Height + 30
    grdCtrl.Visible = True
    grdCtrl.ZOrder
End Sub

Private Sub mShareHalfComment(grdCtrlAbove As MSHFlexGrid, grdCtrlTopShare As MSHFlexGrid)
    grdCtrlTopShare.Move grdStations.Left, grdCtrlAbove.Top + grdCtrlAbove.RowHeight(0) + grdCtrlAbove.RowHeight(1), grdStations.Width, (cmcDone.Top - (grdCtrlAbove.Top + grdCtrlAbove.Height)) / 2 - 75
    gGrid_IntegralHeight grdCtrlTopShare
    grdCtrlTopShare.Height = grdCtrlTopShare.Height + 30
    grdCtrlTopShare.Visible = True
    grdCtrlTopShare.ZOrder

    'grdCtrlButtonShare.Move grdStations.Left, grdCtrlTopShare.Top + grdCtrlTopShare.Height, grdStations.Width, cmcDone.Top - (grdCtrlTopShare.Top + grdCtrlTopShare.Height) - 75
    'gGrid_IntegralHeight grdCtrlButtonShare
    'grdCtrlButtonShare.Height = grdCtrlButtonShare.Height + 30
    'grdCtrlButtonShare.Visible = True
    'grdCtrlButtonShare.ZOrder
    'cmdButton.Move grdCtrlButtonShare.Left + grdCtrlButtonShare.Width - cmdButton.Width - GRIDSCROLLWIDTH, grdCtrlButtonShare.Top + 15, cmdButton.Width, grdCtrlButtonShare.RowHeight(0) - 15
    'cmdButton.Visible = True
    'cmdButton.ZOrder
    udcCommentGrid.Move grdStations.Left - pbcArrow.Width, grdCtrlTopShare.Top + grdCtrlTopShare.Height, grdStations.Width + pbcArrow.Width, cmcDone.Top - (grdCtrlTopShare.Top + grdCtrlTopShare.Height) - 75
    udcCommentGrid.Visible = True
    udcCommentGrid.ZOrder
    mSetCommentButtons

End Sub

Private Sub mShareThird(grdCtrlAbove As MSHFlexGrid, grdCtrlTopShare As MSHFlexGrid)
    grdCtrlTopShare.Move grdStations.Left, grdCtrlAbove.Top + grdCtrlAbove.RowHeight(0) + grdCtrlAbove.RowHeight(1), grdStations.Width, (cmcDone.Top - (grdCtrlAbove.Top + grdCtrlAbove.Height)) / 3 - 75
    gGrid_IntegralHeight grdCtrlTopShare
    grdCtrlTopShare.Height = grdCtrlTopShare.Height + 30
    grdCtrlTopShare.Visible = True
    grdCtrlTopShare.ZOrder
    udcContactGrid.Move grdStations.Left - pbcArrow.Width, grdCtrlTopShare.Top + grdCtrlTopShare.Height, grdStations.Width + pbcArrow.Width, (cmcDone.Top - (grdCtrlTopShare.Top + grdCtrlTopShare.Height)) / 2 - 75
    udcContactGrid.Visible = True
    udcContactGrid.ZOrder
    
    udcCommentGrid.Move grdStations.Left - pbcArrow.Width, udcContactGrid.Top + udcContactGrid.Height, grdStations.Width + pbcArrow.Width, cmcDone.Top - (udcContactGrid.Top + udcContactGrid.Height) - 75
    udcCommentGrid.Visible = True
    udcCommentGrid.ZOrder
    mSetCommentButtons

End Sub


Private Function mGetCol(grdCtrl As MSHFlexGrid, X As Single) As Long
    Dim llColLeftPos As Long
    Dim llCol As Long
    
    mGetCol = -1
    llColLeftPos = grdCtrl.ColPos(0)
    For llCol = 0 To grdCtrl.Cols - 1 Step 1
        If grdCtrl.ColWidth(llCol) > 0 Then
            If (X >= llColLeftPos) And (X <= llColLeftPos + grdCtrl.ColWidth(llCol)) Then
                mGetCol = llCol
                Exit Function
            End If
            llColLeftPos = llColLeftPos + grdCtrl.ColWidth(llCol)
        End If
    Next llCol
End Function

Private Sub grdStations_Scroll()
    'If Not bmStationScrollAllowed Then
    '    grdStations.TopRow = lmStationTopRow
    'End If
    Dim slVehicleName As String
    If bmInScroll Then
        Exit Sub
    End If
'    If Not bmStationScrollAllowed Then
'        bmInScroll = True
'        If grdStations.TopRow > lmStationTopRow Then
'            If grdStations.TextMatrix(lmStationTopRow + 1, SSHTTCODEINDEX) = "" Then
'                grdStations.TopRow = lmStationTopRow
'            Else
'                grdStations.TopRow = lmStationTopRow + 1
'                If bmAgreementScrollAllowed Then
'                    slVehicleName = ""
'                Else
'                    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow, AVEHICLEINDEX))
'                End If
'                grdAgreementInfo.TextMatrix(lmStationTopRow, ASELECTINDEX) = "t"
'                mMousePointer vbHourglass
'                mRepopGrids slVehicleName, 0, 0
'                mSetGridPosition
'                mMousePointer vbDefault
'            End If
'        ElseIf grdStations.TopRow < lmStationTopRow Then
'            If lmStationTopRow - 1 < grdStations.FixedRows Then
'                grdStations.TopRow = lmStationTopRow
'            Else
'                grdStations.TopRow = lmStationTopRow - 1
'                If bmAgreementScrollAllowed Then
'                    slVehicleName = ""
'                Else
'                    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(lmAgreementTopRow, AVEHICLEINDEX))
'                End If
'                grdAgreementInfo.TextMatrix(lmStationTopRow, ASELECTINDEX) = "t"
'                mMousePointer vbHourglass
'                mRepopGrids slVehicleName, 0, 0
'                mSetGridPosition
'                mMousePointer vbDefault
'            End If
'        End If
'        bmInScroll = False
'    End If
    If Not bmStationScrollAllowed Then
        bmInScroll = True
        If grdStations.TopRow > lmStationTopRow Then
            If grdStations.TextMatrix(lmStationTopRow + 1, SSHTTCODEINDEX) = "" Then
                grdStations.TopRow = lmStationTopRow
            Else
                grdStations.TopRow = lmStationTopRow + 1
                If grdStations.TextMatrix(lmStationTopRow, SSELECTINDEX) <> "" Then
                    grdStations.TextMatrix(lmStationTopRow, SSELECTINDEX) = "t"
                End If
                If bmAgreementScrollAllowed Then
                    slVehicleName = ""
                    mMousePointer vbHourglass
                    mRepopGrids slVehicleName, 0, 0
                    mSetGridPosition
                    mMousePointer vbDefault
                Else
                    imDelaySource = 0
                    tmcDelay.Enabled = True
                End If
            End If
        ElseIf grdStations.TopRow < lmStationTopRow Then
            If lmStationTopRow - 1 < grdStations.FixedRows Then
                grdStations.TopRow = lmStationTopRow
            Else
                grdStations.TopRow = lmStationTopRow - 1
                If grdStations.TextMatrix(lmStationTopRow, SSELECTINDEX) <> "" Then
                    grdStations.TextMatrix(lmStationTopRow, SSELECTINDEX) = "t"
                End If
                If bmAgreementScrollAllowed Then
                    slVehicleName = ""
                    mMousePointer vbHourglass
                    mRepopGrids slVehicleName, 0, 0
                    mSetGridPosition
                    mMousePointer vbDefault
                Else
                    imDelaySource = 0
                    tmcDelay.Enabled = True
                End If
            End If
        End If
        bmInScroll = False
    End If
End Sub

Private Sub mSetCommands()
    If (grdAgreementInfo.Visible = False) And (udcCommentGrid.Visible = False) And (udcContactGrid.Visible = False) Then
        If (imLastStationColSorted = SCALLLETTERINDEX) And (imLastStationSort = flexSortStringNoCaseAscending) And (grdStations.TextMatrix(grdStations.FixedRows, SCALLLETTERINDEX) <> "") Then
            edcCallLetters.Visible = True
            lacCallLetters.Visible = True
        Else
            edcCallLetters.Visible = False
            lacCallLetters.Visible = False
        End If
        If (sgShttTimeStamp <> gFileDateTime(sgDBPath & "Shtt.mkd")) Then
            cmcRefresh.Enabled = True
        Else
            cmcRefresh.Enabled = False
        End If
        cmcFilter.Enabled = True
        lacFilter.Enabled = True
        If sgStationSearchCallSource <> "P" Then
            lacUserOption.Visible = True
        End If
        'cmcEMail.Caption = "E-Mails"
        'If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) = AFFILIATECRM) Then
            cmcEMail.Enabled = True
        'Else
        '    cmcEMail.Enabled = False
        'End If
        imcPrinter.Visible = True
    Else
        edcCallLetters.Visible = False
        lacCallLetters.Visible = False
        cmcRefresh.Enabled = False
        cmcFilter.Enabled = False
        lacFilter.Enabled = False
        'lacUserOption.Enabled = False
        lacUserOption.Visible = False
        lbcUserOption(0).Visible = False
        'cmcEMail.Caption = "E-Mail"
        If (sgEMail <> "") Or (StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) Or (StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) Then
            'If ((Asc(sgSpfUsingFeatures9) And AFFILIATECRM) = AFFILIATECRM) Then
                cmcEMail.Enabled = True
            'Else
            '    cmcEMail.Enabled = False
            'End If
        Else
            cmcEMail.Enabled = False
        End If
        imcPrinter.Visible = False
    End If
    If (udcCommentGrid.Visible = False) Then
        cmcAddComment.Enabled = False
        pbcCommentType.Enabled = False
    Else
        If rbcComments(0).Value Then
            cmcAddComment.Enabled = True
        ElseIf rbcComments(1).Value Then
            'If no follow-up, then disable add
            cmcAddComment.Enabled = udcCommentGrid.FollowUpAddAllowed()
        End If
        pbcCommentType.Enabled = True
    End If
    
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbcUserOption(0).Visible = False
    'lbcKey.Visible = True
    lbcKey.ZOrder
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lbcKey.Visible = False
End Sub

Private Sub lacFilter_Click()
    mMousePointer vbHourglass
    lbcUserOption(0).Visible = False
    mSaveCommentsAndContacts
    If lacFilter.Caption = "On" Then
        lacFilter.Caption = "Off"
    Else
        lacFilter.Caption = "On"
    End If
    mPopStationsGrid
    If lacFilter.Caption = "On" Then
        udcCommentGrid.FilterStatus = True  'vbTrue
    Else
        udcCommentGrid.FilterStatus = False 'vbFalse
    End If
    If rbcComments(1).Value Then
        'Follow-up comments showing
        udcCommentGrid.Action 3 'Population
    Else
        'All Comments, only show if station selected
        If grdAgreementInfo.Visible Then
            udcCommentGrid.Action 3 'Population
        Else
            udcContactGrid.Visible = False
            udcCommentGrid.Visible = False
            mHideCommentButtons
        End If
    End If
    mSetGridPosition
    mSetCommands
    mMousePointer vbDefault
End Sub

Private Sub lacUserOption_Click()
    tmcStation.Enabled = False
    
    'sgStationSearchCallSource = "P" is handled with pbcStationType instead of here
    'see mEnableBox.  lcaUserOption is hidden for P
    'P=Post By Advertiser
    
    If sgStationSearchCallSource <> "P" Then
    '    If lacUserOption.Caption = "Assigned" Then
    '        lacUserOption.Caption = "All Affiliates"
    '    ElseIf lacUserOption.Caption = "All Affiliates" Then
    '        lacUserOption.Caption = "Non-Affiliates"
    '    ElseIf lacUserOption.Caption = "Non-Affiliates" Then
    '        lacUserOption.Caption = "All Stations"
    '    ElseIf lacUserOption.Caption = "All Stations" Then
    '        lacUserOption.Caption = "Assigned"
    '    End If
        If lbcUserOption(0).Visible Then
            lbcUserOption(0).Visible = False
        Else
            gSetListBoxHeight lbcUserOption(0), 4
            lbcUserOption(0).Left = lacUserOption.Left
            lbcUserOption(0).Top = lacUserOption.Top - lbcUserOption(0).Height
            If lacUserOption.Caption = "Assigned" Then
                lbcUserOption(0).ListIndex = 0
            ElseIf lacUserOption.Caption = "All Affiliates" Then
                lbcUserOption(0).ListIndex = 1
            ElseIf lacUserOption.Caption = "Non-Affiliates" Then
                lbcUserOption(0).ListIndex = 2
            ElseIf lacUserOption.Caption = "All Stations" Then
                lbcUserOption(0).ListIndex = 3
            End If
            lbcUserOption(0).Visible = True
        End If
    Else
    '    If lacUserOption.Caption = "Delinquent" Then
    '        lacUserOption.Caption = "All Stations"
    '    ElseIf lacUserOption.Caption = "All Stations" Then
    '        lacUserOption.Caption = "Delinquent"
    '    End If
        If lacUserOption.Caption = "Delinquent" Then
            lbcUserOption(1).ListIndex = 0
        ElseIf lacUserOption.Caption = "All Stations" Then
            lbcUserOption(1).ListIndex = 1
        End If
        lbcUserOption(1).Visible = True
    End If
    'tmcStation.Enabled = True
End Sub

Public Function mBinarySearchOwner(llCode As Long) As Long
    
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llListCode As Long
    
    llMin = 0
    llMax = lbcOwner.ListCount - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        llListCode = Val(lbcOwner.List(llMiddle))
        If llCode = llListCode Then
            'found the match
            mBinarySearchOwner = lbcOwner.ItemData(llMiddle)
            Exit Function
        ElseIf llCode < llListCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchOwner = -1
    Exit Function
    
End Function


Private Sub mPopOwnerList()
    Dim ilLoop As Integer
    Dim slStr As String
    
    lbcOwner.Clear
    For ilLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
        slStr = Trim$(Str$(tgOwnerInfo(ilLoop).lCode))
        Do While Len(slStr) < 9
            slStr = "0" & slStr
        Loop
        lbcOwner.AddItem slStr
        lbcOwner.ItemData(lbcOwner.NewIndex) = ilLoop
    Next ilLoop
End Sub

Private Sub mPopListKey()
    Dim slStr As String
    lbcKey.Clear
    lbcKey.AddItem "Action"
    lbcKey.AddItem "     Light Green: Go To Button"
    lbcKey.AddItem "     Light Blue: Sortable Column"
    lbcKey.AddItem "     Yellow: Information Only"
    lbcKey.AddItem ""
    lbcKey.AddItem "Status"
    lbcKey.AddItem "     Brown: Not Yet Exported"
    lbcKey.AddItem "     Dark Green: Compliant"
    lbcKey.AddItem "     Blue: Not Aired"
    lbcKey.AddItem "     Light Blue: Not Compliant"
    lbcKey.AddItem "     Magenta: Partially Posted"
    lbcKey.AddItem "     Light Magenta: Regional Copy"
    lbcKey.AddItem "     Red: Not Yet Posted"
    lbcKey.AddItem ""
    lbcKey.AddItem "Titles"
    lbcKey.AddItem "     Opr: Operator; Sist: Sister Station;"
    lbcKey.AddItem "     Cast: Multi-Cast; Agr: Agreement"
    lbcKey.AddItem "     Cmt: Contact/Comment; Cmpl: Compliant"
    'lbcKey.FontBold = False
    'lbcKey.FontName = "Arial"
    'lbcKey.FontBold = False
    'lbcKey.FontSize = 8
    'lbcKey.Height = lbcKey.ListCount * 120
    'lbcKey.Move imcKey.Left, imcKey.Top - lbcKey.Height
End Sub

Private Sub mBuildFilter()
    Dim llRow As Long
    Dim llTotalCells As Long
    Dim ilLoop As Integer
    Dim llCell As Long
    Dim ilFilter As Integer
    Dim llRepeat As Long
    Dim llCycleLoop As Long
    Dim llCycleCount As Long
    Dim llRepeatCount As Long
    Dim ilMatch As Integer
    Dim blSelectionExist As Boolean
    
    '5/6/18: Set to max GroupID
    Dim tlCount(0 To 31) As FILTERCOUNT  '0=Format(1); 1= Owner(3); 2=Vehicle(4); 3=DMA(0) or MSA(2) or Zip(5) or Territory(6)
    
    '
    'Design:  Build array of OR items.  Each OR item to reference a link list of AND items with that OR item
    '
    'To Build the OR array
    'Step  Description
    '  1   Count the number of items that require an OR item
    '      Each Format requires a separate OR item (Group 0)
    '      Each Owner requires a separate OR item (Group 1)
    '      Each Vehicle requires a separate OR item (Group 2)
    '      Each DMA, MSA, Zip, Territory and Station requires separate OR item (Group 3)
    '      These are treated as one group because the each represent a geographic area
    '  2   Sort number of OR items for each group above from largest to smallest counts
    '  3   Create an UDT to hold each OR item (product of the counts)
    '  4   Distribute the groups into the UDT from the largest to smallest
    '      The number of times that an item should be repeated is determined by taking the
    '      number of times item repeated divided by the number of items in the group
    '      The first repeat is determined from the total number of OR divided by the group count
    '
    
    For ilLoop = 0 To UBound(tlCount) Step 1
        tlCount(ilLoop).iCount = 0
        tlCount(ilLoop).iType = ilLoop
    Next ilLoop
    'Compute Counts for each group
    blSelectionExist = False
    For llRow = 0 To UBound(tgFilterDef) - 1 Step 1
        'Operator: 0=Contains; 1=Equal; 2=Not Equal; 3=Range; 4=Greater or Equal
        '''If ((tgFilterDef(llRow).iOperator <> 0) And (tgFilterDef(llRow).iOperator <> 2)) Then
        ''If ((tgFilterDef(llRow).iOperator <> 0) And (tgFilterDef(llRow).iOperator <> 2)) Or ((tgFilterDef(llRow).iOperator = 0) And (tgFilterDef(llRow).iSelect = SFCALLLETTERS)) Then
        '4/16/20: Bypass Service Agreement filter here as it needs to be an AND operation. Moved to mPopStationsGrid
        'If (tgFilterDef(llRow).iOperator <> 2) Then
        If (tgFilterDef(llRow).iOperator <> 2) And (tgFilterDef(llRow).iSelect <> 51) Then
            'Select Case tgFilterDef(llRow).iSelect
            '    'Case 0, 2, 5, 6, 7
            '    Case 8, 17, 41, 34, 3
            '        tlCount(3).iCount = tlCount(3).iCount + 1
            '        blSelectionExist = True
            '    'Case 1
            '    Case 10
            '        tlCount(0).iCount = tlCount(0).iCount + 1
            '        blSelectionExist = True
            '    'Case 3
            '    Case 20
            '        tlCount(1).iCount = tlCount(1).iCount + 1
            '        blSelectionExist = True
            '    'Case 4
            '    Case 37
            '        tlCount(2).iCount = tlCount(2).iCount + 1
            '        blSelectionExist = True
            'End Select
            tlCount(tgFilterDef(llRow).iCountGroup).iCount = tlCount(tgFilterDef(llRow).iCountGroup).iCount + 1
            blSelectionExist = True
        End If
    Next llRow
    'Determine total number of OR required
    llTotalCells = 1
    For ilLoop = 0 To UBound(tgFilterDef) - 1 Step 1
        'If tlCount(ilLoop).iCount > 0 Then
        '    llTotalCells = llTotalCells * tlCount(ilLoop).iCount
        'End If
        
        ilMatch = False
        For llRow = 0 To ilLoop - 1 Step 1
            If tgFilterDef(llRow).iCountGroup = tgFilterDef(ilLoop).iCountGroup Then
            ilMatch = True
        End If
        Next llRow
    
        If tlCount(tgFilterDef(ilLoop).iCountGroup).iCount > 0 And (Not ilMatch) Then
            llTotalCells = llTotalCells * tlCount(tgFilterDef(ilLoop).iCountGroup).iCount
        End If
    Next ilLoop
    ReDim tmFilterLink(0 To llTotalCells) As FILTERLINK
    ReDim tmAndFilterLink(0 To 0) As FILTERLINK
    If blSelectionExist Then
        For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
            tmFilterLink(llCell).lFilterDefIndex = -1
            tmFilterLink(llCell).lNotFilterDefIndex = -1
            tmFilterLink(llCell).lNextAnd = -1
        Next llCell
        'Sort Counts from larest to smallest
        ArraySortTyp fnAV(tlCount(), 0), UBound(tlCount) + 1, 1, LenB(tlCount(0)), 0, -1, 0
        llRepeatCount = llTotalCells / tlCount(0).iCount
        llCycleCount = 1
        For ilLoop = 0 To UBound(tlCount) Step 1
            llCell = 0
            For llCycleLoop = 1 To llCycleCount Step 1
                For ilFilter = 0 To UBound(tgFilterDef) - 1 Step 1
                    'Operator: 0=Contains; 1=Equal; 2=Not Equal; 3=Range; 4=Greater or Equal    '2=Greater Than; 3=Less Than; 4=Not Equal
                    '''If (tgFilterDef(ilFilter).iOperator <> 0) And (tgFilterDef(ilFilter).iOperator <> 2) Then
                    ''If ((tgFilterDef(ilFilter).iOperator <> 0) And (tgFilterDef(ilFilter).iOperator <> 2)) Or ((tgFilterDef(ilFilter).iOperator = 0) And (tgFilterDef(ilFilter).iSelect = SFCALLLETTERS)) Then
                    '4/16/20: Bypass Service Agreement filter here as it needs to be an AND operation. Moved to mPopStationsGrid
                    'If (tgFilterDef(ilFilter).iOperator <> 2) Then
                    If (tgFilterDef(ilFilter).iOperator <> 2) And (tgFilterDef(ilFilter).iSelect <> 51) Then
                        ilMatch = False
'                        Select Case tlCount(ilLoop).iType
'                            Case 0  'Format
'                                If tgFilterDef(ilFilter).iSelect = 10 Then
'                                    ilMatch = True
'                                End If
'                            Case 1  'Owner
'                                If tgFilterDef(ilFilter).iSelect = 20 Then
'                                    ilMatch = True
'                                End If
'                            Case 2  'Vehicle
'                                If tgFilterDef(ilFilter).iSelect = 37 Then
'                                    ilMatch = True
'                                End If
'                            Case 3  'DMA, MSA, ZIP, Territory and Call letters
'                                If (tgFilterDef(ilFilter).iSelect = 8) Or (tgFilterDef(ilFilter).iSelect = 17) Or (tgFilterDef(ilFilter).iSelect = 41) Or (tgFilterDef(ilFilter).iSelect = 34) Or (tgFilterDef(ilFilter).iSelect = 3) Then
'                                    ilMatch = True
'                                End If
                                
'                        End Select
                        If tlCount(ilLoop).iType = tgFilterDef(ilFilter).iCountGroup Then
                            ilMatch = True
                        End If
                        If ilMatch Then
                            For llRepeat = 1 To llRepeatCount Step 1
                                If tmFilterLink(llCell).lFilterDefIndex < 0 Then
                                    tmFilterLink(llCell).lFilterDefIndex = ilFilter
                                    tmFilterLink(llCell).lNextAnd = -1
                                Else
                                    tmAndFilterLink(UBound(tmAndFilterLink)).lFilterDefIndex = ilFilter
                                    tmAndFilterLink(UBound(tmAndFilterLink)).lNextAnd = tmFilterLink(llCell).lNextAnd
                                    tmFilterLink(llCell).lNextAnd = UBound(tmAndFilterLink)
                                    ReDim Preserve tmAndFilterLink(0 To UBound(tmAndFilterLink) + 1) As FILTERLINK
                                End If
                                llCell = llCell + 1
                            Next llRepeat
                        End If
                    End If
                Next ilFilter
            Next llCycleLoop
            If ilLoop = UBound(tlCount) Then
                Exit For
            End If
            If tlCount(ilLoop + 1).iCount = 0 Then
                Exit For
            End If
            
            llRepeatCount = llRepeatCount / tlCount(ilLoop + 1).iCount
            llCycleCount = llTotalCells / (llRepeatCount * tlCount(ilLoop + 1).iCount)
        Next ilLoop
    Else
        ReDim tmFilterLink(0 To 1) As FILTERLINK
        ReDim tmAndFilterLink(0 To 0) As FILTERLINK
        For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
            tmFilterLink(llCell).lFilterDefIndex = -1
            tmFilterLink(llCell).lNotFilterDefIndex = -1
            tmFilterLink(llCell).lNextAnd = -1
        Next llCell
    End If
    'Add Contains and Not's
    For ilFilter = 0 To UBound(tgFilterDef) - 1 Step 1
        'Operator: 0=Contains; 1=Equal; 2=Not Equal; 3=Range; 4=Greater or Equal    '2=Greater Than; 3=Less Than; 4=Not Equal
        'If ((tgFilterDef(ilFilter).iOperator = 0) And (tgFilterDef(ilFilter).iSelect <> SFCALLLETTERS)) Or (tgFilterDef(ilFilter).iOperator = 2) Then
        If tgFilterDef(ilFilter).iOperator = 2 Then
            For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
                If tmFilterLink(llCell).lNotFilterDefIndex < 0 Then
                    tmFilterLink(llCell).lNotFilterDefIndex = ilFilter
                    'tmFilterLink(llCell).lNextAnd = -1
                Else
                    tmAndFilterLink(UBound(tmAndFilterLink)).lNotFilterDefIndex = ilFilter
                    tmAndFilterLink(UBound(tmAndFilterLink)).lNextAnd = tmFilterLink(llCell).lNextAnd
                    tmFilterLink(llCell).lNextAnd = UBound(tmAndFilterLink)
                    ReDim Preserve tmAndFilterLink(0 To UBound(tmAndFilterLink) + 1) As FILTERLINK
                End If
            Next llCell
        End If
    Next ilFilter
    
End Sub

Private Sub pbcCommentType_GotFocus()
    mSaveCommentsAndContacts
End Sub

Private Sub pbcCommentType_Paint()
    pbcCommentType.CurrentX = 0
    pbcCommentType.CurrentY = 0
    pbcCommentType.Print smCommentType
End Sub

Private Sub pbcFocus_GotFocus()
    lbcUserOption(0).Visible = False
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer
    Dim ilIndex As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        Do
            ilNext = False
            Select Case grdPostBuy.Col
                Case PSSTATIONTYPEINDEX
                    mSetShow
                    pbcClickFocus.SetFocus
                    Exit Sub
                Case Else
                    grdPostBuy.Col = grdPostBuy.Col - 1
            End Select
            'If mColOk() Then
            '    Exit Do
            'Else
            '    ilNext = True
            'End If
        Loop While ilNext
        mSetShow
    Else
        grdPostBuy.Row = grdPostBuy.FixedRows
        grdPostBuy.Col = grdPostBuy.FixedCols
        'Do
        '    If mColOk() Then
        '        Exit Do
        '    Else
        '        grdPostBuy.Col = grdPostBuy.Col + 1
        '    End If
        'Loop
    End If
    mEnableBox
End Sub

Private Sub pbcStationType_Click()
    If smStationType = "Delinquent" Then
        smStationType = "All Stations"
        If lbcDate.List(0) <> "[All]" Then
            lbcDate.AddItem "[All]", 0
        End If
    ElseIf smStationType = "All Stations" Then
        smStationType = "Delinquent"
        If lbcDate.List(0) = "[All]" Then
            lbcDate.RemoveItem 0
        End If
        If grdPostBuy.TextMatrix(lmEnableRow, lmEnableCol) = "[All]" Then
            grdPostBuy.TextMatrix(lmEnableRow, lmEnableCol) = ""
        End If
    End If
    pbcStationType.Cls
    pbcStationType_Paint
End Sub

Private Sub pbcStationType_Paint()
    pbcStationType.CurrentX = 0
    pbcStationType.CurrentY = 0
    pbcStationType.Print smStationType
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilNext As Integer
    Dim ilIndex As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim ilLoop As Integer

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        llEnableCol = lmEnableCol
        mSetShow
        grdPostBuy.Row = llEnableRow
        grdPostBuy.Col = llEnableCol
        Do
            ilNext = False
            Select Case grdPostBuy.Col
                Case PSCONTRACTINDEX
                    pbcClickFocus.SetFocus
                    Exit Sub
                Case Else
                    grdPostBuy.Col = grdPostBuy.Col + 1
            End Select
            'If mColOk() Then
            '    Exit Do
            'Else
            '    ilNext = True
            'End If
        Loop While ilNext
    Else
        grdPostBuy.Row = grdPostBuy.FixedRows
        grdPostBuy.Col = grdPostBuy.FixedCols
        'Do
        '    If mColOk() Then
        '        Exit Do
        '    Else
        '        grdPostBuy.Col = grdPostBuy.Col + 1
        '    End If
        'Loop
    End If
    mEnableBox
End Sub

Private Sub rbcComments_Click(Index As Integer)
    If rbcComments(Index).Value Then
        tmcStation.Enabled = False
        udcCommentGrid.Action 5 'Save comments
        If Index = 1 Then
            If (smUserType = "M") Or (smUserType = "S") Then
                smCommentType = "   Mine"
                pbcCommentType.Cls
                pbcCommentType_Paint
                pbcCommentType.Enabled = False
            End If
            udcCommentGrid.CommentGridForm = 1  'See Follow-up
        ElseIf Index = 0 Then
            If (smUserType = "M") Or (smUserType = "S") Then
                smCommentType = "    All"
                pbcCommentType.Cls
                pbcCommentType_Paint
            End If
            pbcCommentType.Enabled = True
            udcCommentGrid.CommentGridForm = 0  'See All Comments
            If udcContactGrid.Visible Then
                If (imShttCode > 0) Then
                    If udcContactGrid.VerifyRights("M") Then
                        udcContactGrid.StationCode = imShttCode
                        udcContactGrid.Action 5 'Save
                    End If
                End If
            End If
        End If
        lmCommentStartLoc = udcCommentGrid.ColumnStartLocation(6)
        If ((Index = 0) And (imShttCode > 0)) Or (Index = 1) Then
            If (imShttCode > 0) And (bmStationScrollAllowed = False) Then
                udcContactGrid.StationCode = imShttCode
                udcContactGrid.Action 3 'populate
                udcContactGrid.Visible = True
            End If
            udcCommentGrid.StationCode = imShttCode
            udcCommentGrid.Action 3 'Populate
            udcCommentGrid.Visible = True
        Else
            mHideCommentButtons
            udcCommentGrid.Visible = False
        End If
        mSetGridPosition
    End If
    mSetCommands
End Sub

Private Sub rbcComments_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    edcCommentTip.Visible = False
End Sub

Private Sub rbcComments_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Index = 1) And (udcCommentGrid.Visible = False) And (rbcComments(1).Value) Then
        rbcComments_Click 1
    End If
End Sub

Private Sub rbcCompliant_Click(Index As Integer)
    If rbcCompliant(Index).Value Then
        mCompliantTypeChg
    End If
End Sub

Private Sub rbcCompliant_GotFocus(Index As Integer)
    mSaveCommentsAndContacts
    mSetShow
End Sub

Private Sub tmcDelay_Timer()
    Dim slVehicleName As String
    Dim llAttCode As Long
    Dim llCpttCode As Long
    
    tmcDelay.Enabled = False
    bmInScroll = True
    Select Case imDelaySource
        Case 0  'Station
            slVehicleName = Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, AVEHICLEINDEX))
            mMousePointer vbHourglass
            mRepopGrids slVehicleName, 0, 0
            mSetGridPosition
            mMousePointer vbDefault
        Case 1  'Agreement
            slVehicleName = Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, AVEHICLEINDEX))
            llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, AATTCODEINDEX)))
            If grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, ASELECTINDEX) <> "" Then
                grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, ASELECTINDEX) = "s"
            End If
            mMousePointer vbHourglass
            mRepopGrids slVehicleName, llAttCode, 0
            mSetGridPosition
            mMousePointer vbDefault
        Case 2  'Posted Info (CPTT)
            slVehicleName = Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, AVEHICLEINDEX))
            llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(grdAgreementInfo.TopRow, AATTCODEINDEX)))
            llCpttCode = Val(Trim$(grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PCPTTINDEX)))
            If grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) <> "" Then
                grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) = "s"
            End If
            mMousePointer vbHourglass
            mRepopGrids slVehicleName, llAttCode, llCpttCode
            mSetGridPosition
            mMousePointer vbDefault
    End Select
    bmInScroll = False
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mMousePointer vbHourglass
    lacUserOption.FontBold = True
    lacFilter.FontBold = True
    mSetGridColumns
    mSetGridTitles
    'udcContactGrid.Font.Name = grdStations.Font.Name
    'udcContactGrid.Font.Size = grdStations.Font.Size
    'udcContactGrid.Font.Bold = grdStations.Font.Bold
    udcContactGrid.Action 2 'Init
    udcContactGrid.Source = "M"
    udcCommentGrid.Action 2 'Init

    'gGrid_IntegralHeight grdStations
    'gGrid_FillWithRows grdStations
    lbcKey.FontBold = False
    lbcKey.FontName = "Arial"
    lbcKey.FontBold = False
    lbcKey.FontSize = 8
    lbcKey.Height = (lbcKey.ListCount - 1) * 225
    lbcKey.Move imcKey.Left, imcKey.Top - lbcKey.Height
    imFirstTime = False
    If sgStationSearchCallSource = "P" Then
        smStationType = "Delinquent"
        mPopDate
        mPopAdvt
        grdPostBuy.Visible = True
        grdStations.Visible = True
        cmcFilter.Visible = False
        lacFilter.Visible = False
        lbcUserOption(1).ListIndex = 0
        lacUserOption.Caption = "Delinquent"
        lacUserOption.Visible = False
        grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSSTATIONTYPEINDEX) = smStationType
    Else
        frmStationSearchFilter.Show vbModal
        If igFilterReturn Then
            'Create Filter test
            mMousePointer vbHourglass
            If UBound(tgFilterDef) > 0 Then
                lacFilter.Caption = "On"
            Else
                lacFilter.Caption = "Off"
            End If
            mBuildFilter
            mPopStationsGrid
            If (grdStations.TextMatrix(grdStations.FixedRows, SCALLLETTERINDEX) = "") And (UBound(tgFilterDef) <= 0) Then
                If lacUserOption.Caption <> "All Stations" Then
                    lbcUserOption(0).ListIndex = 3
                    lacUserOption.Caption = "All Stations"
                    mPopStationsGrid
                End If
            End If
            grdStations.Visible = True
            If UBound(tgFilterDef) > 0 Then
                lacFilter.Visible = True
                udcCommentGrid.FilterStatus = True  'vbTrue
            Else
                lacFilter.Visible = False
                udcCommentGrid.FilterStatus = False 'vbFalse
            End If
            rbcComments(0).Enabled = True
            rbcComments(1).Enabled = True
            cmcFilter.Enabled = True
        Else
            tmcStation.Enabled = False
            Unload frmStationSearch
            Exit Sub
        End If
    End If
    mMousePointer vbDefault
End Sub

Private Sub tmcStation_Timer()
    Dim slSDate As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim blRet As Boolean
    
    tmcStation.Enabled = False
    mMousePointer vbHourglass
    If sgStationSearchCallSource <> "P" Then
        grdStations.Visible = True
        mPopStationsGrid
        lmCommentStartLoc = udcCommentGrid.ColumnStartLocation(6)
        rbcComments(0).Enabled = True
        rbcComments(1).Enabled = True
    Else
        If lbcDate.ListIndex < 0 Then
            mMousePointer vbDefault
            Exit Sub
        End If
        slSDate = Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX))
        If Not gIsDate(slSDate) Then
            mMousePointer vbDefault
            Exit Sub
        End If
        If lbcAdvertiser.ListIndex < 0 Then
            mMousePointer vbDefault
            Exit Sub
        End If
        If lbcContract.ListIndex < 0 Then
            mMousePointer vbDefault
            Exit Sub
        End If
        ilIndex = -1
        For ilLoop = 0 To UBound(tgFilterTypes) - 1 Step 1
            If Trim$(tgFilterTypes(ilLoop).sFieldName) = "Call Letters" Then
                ilIndex = ilLoop
            End If
        Next ilLoop
        If ilIndex = -1 Then
            mMousePointer vbDefault
            Exit Sub
        End If
        blRet = mCreatePostBuyStationFilter()
    End If
    mMousePointer vbDefault
End Sub

Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
    gSetMousePointer grdStations, grdAgreementInfo, ilMousepointer
    gSetMousePointer grdPostedInfo, grdSpotInfo, ilMousepointer
End Sub

Private Sub mFilterCompareStr(slInStr As String, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim slStr As String
    Dim slFrom As String
    slStr = UCase(Trim$(slInStr))
    slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
    'Operator: 0=Contains; 1=Equal; 2=Greater Than; 3=Less Than; 4=Not Equal
    Select Case tgFilterDef(llFilterDefIndex).iOperator
        Case 0  'Contains
            If InStr(1, slStr, slFrom, vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 1  'Equal
            If (slFrom = "") And (tgFilterDef(llFilterDefIndex).sCntrlType = "E") Then
                If slStr <> "" Then
                    ilIncludeStation = False
                End If
            Else
                If slFrom = "[DEFINED]" Then
                    If slStr = "" Then
                        ilIncludeStation = False
                    End If
                Else
                    If (StrComp(slStr, slFrom, vbBinaryCompare) <> 0) Or (slStr = "") Then
                        ilIncludeStation = False
                    End If
                End If
            End If
        Case 2  'Not Equal
            If slFrom = "[DEFINED]" Then
                If slStr <> "" Then
                    ilIncludeStation = False
                End If
            Else
                If (StrComp(slStr, slFrom, vbBinaryCompare) = 0) Or (slStr = "") Then
                    ilIncludeStation = False
                End If
            End If
        Case 3  'Range
            If StrComp(slStr, slFrom, vbBinaryCompare) < 0 Then
                ilIncludeStation = False
            End If
            If StrComp(slStr, UCase(Trim$(tgFilterDef(llFilterDefIndex).sToValue)), vbBinaryCompare) > 0 Then
                ilIncludeStation = False
            End If
        Case 4  'Greater or Equal
            If (StrComp(slStr, slFrom, vbBinaryCompare) < 0) Or (slStr = "") Then
                ilIncludeStation = False
            End If
    End Select

End Sub

Private Sub mFilterComparePhone(slInStr As String, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim slStr As String
    Dim slFrom As String
    slStr = UCase(Trim$(mConvertPhone(slInStr)))
    If UCase$(Trim(tgFilterDef(llFilterDefIndex).sFromValue)) <> "[DEFINED]" Then
        slFrom = UCase(Trim$(mConvertPhone(tgFilterDef(llFilterDefIndex).sFromValue)))
    Else
        slFrom = UCase$(Trim(tgFilterDef(llFilterDefIndex).sFromValue))
    End If
    'Operator: 0=Contains; 1=Equal; 2=Greater Than; 3=Less Than; 4=Not Equal
    Select Case tgFilterDef(llFilterDefIndex).iOperator
        Case 0  'Contains
            If InStr(1, slStr, slFrom, vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 1  'Equal
            If slFrom = "[DEFINED]" Then
                If slStr = "" Then
                    ilIncludeStation = False
                End If
            Else
                If StrComp(slStr, slFrom, vbBinaryCompare) <> 0 Then
                    ilIncludeStation = False
                End If
            End If
        Case 2  'Not Equal
            If slFrom = "[DEFINED]" Then
                If slStr <> "" Then
                    ilIncludeStation = False
                End If
            Else
                If StrComp(slStr, slFrom, vbBinaryCompare) = 0 Then
                    ilIncludeStation = False
                End If
            End If
        Case 3  'Range
            If StrComp(slStr, slFrom, vbBinaryCompare) < 0 Then
                ilIncludeStation = False
            End If
            If StrComp(slStr, UCase(Trim$(tgFilterDef(llFilterDefIndex).sToValue)), vbBinaryCompare) > 0 Then
                ilIncludeStation = False
            End If
        Case 4  'Greater or Equal
            If StrComp(slStr, slFrom, vbBinaryCompare) < 0 Then
                ilIncludeStation = False
            End If
    End Select

End Sub

Private Sub mFilterCompareLong(llLong As Long, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim slStr As String
    'Operator: 0=Contains; 1=Equal; 2=Greater Than; 3=Less Than; 4=Not Equal
    Select Case tgFilterDef(llFilterDefIndex).iOperator
        Case 0  'Contains
            slStr = Trim$(Str(llLong))
            If InStr(1, slStr, UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue)), vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 1  'Equal
            If llLong <> tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
        Case 2  'Not Equal
            If llLong = tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
        Case 3  'Range
            If llLong < tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
            If llLong > tgFilterDef(llFilterDefIndex).lToValue Then
                ilIncludeStation = False
            End If
        Case 4  'Greater or Equal
            If llLong < tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
    End Select

End Sub
Private Sub mFilterCompareInteger(ilInteger As Integer, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim slStr As String
    'Operator: 0=Contains; 1=Equal; 2=Greater Than; 3=Less Than; 4=Not Equal
    Select Case tgFilterDef(llFilterDefIndex).iOperator
        Case 0  'Contains
            slStr = Trim$(Str(ilInteger))
            If InStr(1, slStr, UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue)), vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 1  'Equal
            If ilInteger <> tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
        Case 2  'Not Equal
            If ilInteger = tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
        Case 3  'Range
            If ilInteger < tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
            If ilInteger > tgFilterDef(llFilterDefIndex).lToValue Then
                ilIncludeStation = False
            End If
        Case 4  'Greater or Equal
            If ilInteger < tgFilterDef(llFilterDefIndex).lFromValue Then
                ilIncludeStation = False
            End If
    End Select

End Sub

Private Sub mFilterCompareCurrency(slStr As String, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim clStr As Currency
    Dim clFrom As Currency
    Dim clTo As Currency
    If Trim$(slStr) <> "" Then
        clStr = CCur(Trim$(slStr))
    Else
        clStr = 0
    End If
    If Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "" Then
        clFrom = CCur(tgFilterDef(llFilterDefIndex).sFromValue)
    Else
        clFrom = 0
    End If
    If Trim$(tgFilterDef(llFilterDefIndex).sToValue) <> "" Then
        clTo = CCur(tgFilterDef(llFilterDefIndex).sToValue)
    Else
        clTo = 0
    End If
    'Operator: 0=Contains; 1=Equal; 2=Greater Than; 3=Less Than; 4=Not Equal
    Select Case tgFilterDef(llFilterDefIndex).iOperator
        Case 0  'Contains
            If InStr(1, UCase(slStr), UCase(tgFilterDef(llFilterDefIndex).sFromValue), vbBinaryCompare) <= 0 Then
                ilIncludeStation = False
            End If
        Case 1  'Equal
            If clStr <> clFrom Then
                ilIncludeStation = False
            End If
        Case 2  'Not Equal
            If clStr = clFrom Then
                ilIncludeStation = False
            End If
        Case 3  'Range
            If clStr < clFrom Then
                ilIncludeStation = False
            End If
            If clStr > clTo Then
                ilIncludeStation = False
            End If
        Case 4  'Greater or Equal
            If clStr < clFrom Then
                ilIncludeStation = False
            End If
    End Select

End Sub

Private Sub mTestFilter(ilShtt As Integer, llFilterDefIndex As Long, ilIncludeStation As Integer)
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String
    Dim ilFmt As Integer
    Dim llMSA As Long
    Dim llOwner As Long
    Dim ilVef As Integer
    Dim ilMnt As Integer
    Dim ilTzt As Integer
    Dim llMkt As Long
    Dim llLoop As Long
    Dim ilPos As Integer
    Dim ilSvRank As Integer
    Dim blRecExist As Boolean
    Dim slFrom As String
    
    Select Case tgFilterDef(llFilterDefIndex).iSelect
        Case SFAREA
            llValue = tgStationInfo(ilShtt).lAreaMntCode
            ilMnt = gBinarySearchMnt(llValue, tgAreaInfo())
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgAreaInfo(ilMnt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFCALLLETTERSCHGDATE
            'Get date from clt
            SQLQuery = "SELECT cltEndDate FROM clt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " cltShfCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_clt = gSQLSelectCall(SQLQuery)
            If Not rst_clt.EOF Then
                Do While Not rst_clt.EOF
                    ilIncludeStation = True
                    mFilterCompareLong gDateValue(Format(rst_clt!cltEndDate, sgShowDateForm)), llFilterDefIndex, ilIncludeStation
                    If ilIncludeStation Then
                        Exit Do
                    End If
                    rst_clt.MoveNext
                Loop
            Else
                ilIncludeStation = False
            End If
        Case SFCALLLETTERS  'Station
            If sgStationSearchCallSource <> "P" Then
                If (tgFilterDef(llFilterDefIndex).iOperator = 1) Then
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters)) & ", " & Trim$(tgStationInfo(ilShtt).sMarket)
                Else
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters))
                End If
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                If (Not ilIncludeStation) And (tgFilterDef(llFilterDefIndex).iOperator = 1) Then
                    slStr = Trim$(UCase(tgFilterDef(llFilterDefIndex).sFromValue))
                    ilPos = InStr(1, slStr, ",", vbTextCompare)
                    If ilPos > 0 Then
                        slStr = Trim$(Left(slStr, ilPos - 1))
                    End If
                    SQLQuery = "SELECT cltCallLetters FROM clt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " cltShfCode = " & tgStationInfo(ilShtt).iCode & " AND"
                    SQLQuery = SQLQuery & " cltCallLetters = '" & slStr & "')"
                    Set rst_clt = gSQLSelectCall(SQLQuery)
                    If Not rst_clt.EOF Then
                        ilIncludeStation = True
                    End If
                End If
                If (Not ilIncludeStation) And (tgFilterDef(llFilterDefIndex).iOperator = 0) Then
                    SQLQuery = "SELECT cltCallLetters FROM clt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " cltShfCode = " & tgStationInfo(ilShtt).iCode & " AND"
                    SQLQuery = SQLQuery & " cltCallLetters LIKE '%" & Trim$(UCase(tgFilterDef(llFilterDefIndex).sFromValue)) & "%')"
                    Set rst_clt = gSQLSelectCall(SQLQuery)
                    If Not rst_clt.EOF Then
                        ilIncludeStation = True
                    End If
                End If
            Else
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            End If
        Case SFCITYLIC
            llValue = tgStationInfo(ilShtt).lCityLicMntCode
            ilMnt = gBinarySearchMnt(llValue, tgCityInfo())
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgCityInfo(ilMnt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFCOMMERCIAL
            slStr = tgStationInfo(ilShtt).sStationType
            If slStr = "N" Then
                slStr = "Non-Commercial"
            Else
                slStr = "Commercial"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFCOUNTYLIC
            llValue = tgStationInfo(ilShtt).lCountyLicMntCode
            ilMnt = gBinarySearchMnt(llValue, tgCountyInfo())
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgCountyInfo(ilMnt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFDAYLIGHT
            ilValue = tgStationInfo(ilShtt).iAckDaylight
            If ilValue = 1 Then
                slStr = "Ignore Daylight Savings"
            Else
                slStr = "Honor Daylight Savings"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFDMA  'DMA
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMarket))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFDMARANK
            llMkt = gBinarySearchMkt(CLng(tgStationInfo(ilShtt).iMktCode))
            If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (Trim$(tgFilterDef(llFilterDefIndex).sToValue) <> "") Then
                If llMkt <> -1 Then
                    If tgMarketInfo(llMkt).iRank = 0 Then
                        mFilterCompareInteger 999, llFilterDefIndex, ilIncludeStation
                    Else
                        mFilterCompareInteger tgMarketInfo(llMkt).iRank, llFilterDefIndex, ilIncludeStation
                    End If
                Else
                    ilIncludeStation = False
                End If
            Else
                If llMkt <> -1 Then
                    If tgMarketInfo(llMkt).iRank <> 0 Then
                        ilIncludeStation = False
                    End If
                Else
                    ilIncludeStation = False
                End If
            End If
        Case SFFORMAT  'Format
            llValue = tgStationInfo(ilShtt).iFormatCode
            ilFmt = gBinarySearchFmt(CInt(llValue))
            If ilFmt <> -1 Then
                slStr = UCase$(Trim$(tgFormatInfo(ilFmt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFFREQ
            mFilterCompareCurrency tgStationInfo(ilShtt).sFrequency, llFilterDefIndex, ilIncludeStation
        Case SFHISTSTARTDATE
            mFilterCompareLong tgStationInfo(ilShtt).lHistStartDate, llFilterDefIndex, ilIncludeStation
        Case SFPERMID
            mFilterCompareLong tgStationInfo(ilShtt).lPermStationID, llFilterDefIndex, ilIncludeStation
        Case SFMAILADDRESS
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMailAddress1))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            If Not ilIncludeStation Then
                ilIncludeStation = True
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMailAddress2))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            End If
            If Not ilIncludeStation Then
                ilIncludeStation = True
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMailState))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            End If
            If Not ilIncludeStation Then
                ilIncludeStation = True
                llValue = tgStationInfo(ilShtt).lMailCityMntCode
                ilMnt = gBinarySearchMnt(llValue, tgCityInfo())
                If ilMnt <> -1 Then
                    slStr = UCase$(Trim$(tgCityInfo(ilMnt).sName))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                Else
                    ilIncludeStation = False
                End If
            End If
        Case SFMARKETREP
            slStr = ""
            ilValue = tgStationInfo(ilShtt).iMktRepUstCode
            If ilValue <> 0 Then
                For llLoop = 0 To UBound(tgMarketRepInfo) - 1 Step 1
                    If tgMarketRepInfo(llLoop).iUstCode = ilValue Then
                        slStr = UCase$(Trim$(tgMarketRepInfo(llLoop).sName))
                        Exit For
                    End If
                Next llLoop
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFMONIKER
            llValue = tgStationInfo(ilShtt).lMonikerMntCode
            ilMnt = gBinarySearchMnt(llValue, tgMonikerInfo())
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgMonikerInfo(ilMnt).sName))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case SFMSA  'MSA
            llValue = tgStationInfo(ilShtt).iMSAMktCode
            llMSA = gBinarySearchMSAMkt(llValue)
            If llMSA <> -1 Then
                slStr = UCase$(Trim$(tgMSAMarketInfo(llMSA).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFMSARANK
            llMkt = gBinarySearchMSAMkt(CLng(tgStationInfo(ilShtt).iMSAMktCode))
            If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (Trim$(tgFilterDef(llFilterDefIndex).sToValue) <> "") Then
                If llMkt <> -1 Then
                    If tgMSAMarketInfo(llMkt).iRank = 0 Then
                        mFilterCompareInteger 999, llFilterDefIndex, ilIncludeStation
                    Else
                        mFilterCompareInteger tgMSAMarketInfo(llMkt).iRank, llFilterDefIndex, ilIncludeStation
                    End If
                Else
                    ilIncludeStation = False
                End If
            Else
                If llMkt <> -1 Then
                    If tgMSAMarketInfo(llMkt).iRank <> 0 Then
                        ilIncludeStation = False
                    End If
                Else
                    ilIncludeStation = False
                End If
            End If
        Case SFONAIR
            If tgStationInfo(ilShtt).sOnAir = "N" Then
                slStr = "Off Air"
            Else
                slStr = "On Air"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFOPERATOR
            llValue = tgStationInfo(ilShtt).lOperatorMntCode
            ilMnt = gBinarySearchMnt(llValue, tgOperatorInfo())
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgOperatorInfo(ilMnt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFOWNER  'Owner
            llValue = tgStationInfo(ilShtt).lOwnerCode
            llOwner = mBinarySearchOwner(llValue)
            If llOwner <> -1 Then
                slStr = UCase$(Trim$(tgOwnerInfo(llOwner).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFP12PLUS
            mFilterCompareLong tgStationInfo(ilShtt).lAudP12Plus, llFilterDefIndex, ilIncludeStation
        Case SFWATTS
            mFilterCompareLong tgStationInfo(ilShtt).lWatts, llFilterDefIndex, ilIncludeStation
        Case SFPHONE
            slStr = tgStationInfo(ilShtt).sPhone
            If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
                mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
                If Not ilIncludeStation Then
                    ilIncludeStation = True
                    slStr = tgStationInfo(ilShtt).sFax
                    mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
                End If
                If Not ilIncludeStation Then
                    SQLQuery = "SELECT arttPhone, arttFax FROM artt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
                    Set rst_artt = gSQLSelectCall(SQLQuery)
                    Do While Not rst_artt.EOF
                        ilIncludeStation = True
                        slStr = rst_artt!arttPhone
                        mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            Exit Do
                        End If
                        ilIncludeStation = True
                        slStr = rst_artt!arttFax
                        mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            Exit Do
                        End If
                        rst_artt.MoveNext
                    Loop
                End If
            Else
                If Trim$(slStr) <> "" Then
                    ilIncludeStation = False
                Else
                    SQLQuery = "SELECT arttPhone FROM artt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
                    Set rst_artt = gSQLSelectCall(SQLQuery)
                    Do While Not rst_artt.EOF
                        slStr = rst_artt!arttPhone
                        If Trim$(slStr) <> "" Then
                            ilIncludeStation = False
                            Exit Do
                        End If
                        rst_artt.MoveNext
                    Loop
                End If
            End If
        Case SFPHYADDRESS
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyAddress1))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            If Not ilIncludeStation Then
                ilIncludeStation = True
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyAddress2))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            End If
            If Not ilIncludeStation Then
                ilIncludeStation = True
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyState))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            End If
            If Not ilIncludeStation Then
                ilIncludeStation = True
                llValue = tgStationInfo(ilShtt).lPhyCityMntCode
                ilMnt = gBinarySearchMnt(llValue, tgCityInfo())
                If ilMnt <> -1 Then
                    slStr = UCase$(Trim$(tgCityInfo(ilMnt).sName))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                Else
                    ilIncludeStation = False
                End If
            End If
        Case SFSERIAL
            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
                If (tgFilterDef(llFilterDefIndex).iOperator = 1) And (Trim(tgFilterDef(llFilterDefIndex).sFromValue) = "") Then
                    'treat with AND operator
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo1))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    If ilIncludeStation Then
                        slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo2))
                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    End If
                Else
                    'Equal, treat with OR operator
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo1))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    If Not ilIncludeStation Then
                        ilIncludeStation = True
                        slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo2))
                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    End If
                End If
            Else
                'Not Equal, treat with AND operator
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo1))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                If ilIncludeStation Then
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo2))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                End If
            End If
        Case SFSERVICEREP
            slStr = ""
            ilValue = tgStationInfo(ilShtt).iServRepUstCode
            If ilValue <> 0 Then
                For llLoop = 0 To UBound(tgServiceRepInfo) - 1 Step 1
                    If tgServiceRepInfo(llLoop).iUstCode = ilValue Then
                        slStr = UCase$(Trim$(tgServiceRepInfo(llLoop).sName))
                        Exit For
                    End If
                Next llLoop
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFSTATELIC
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sStateLic))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFEMAIL
            blRecExist = False
            slStr = "Not Checked"
            SQLQuery = "SELECT arttWebEMail FROM artt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            Do While Not rst_artt.EOF
                blRecExist = True
                If rst_artt!arttWebEMail = "Y" Then
                    slStr = "Checked"
                    Exit Do
                End If
                rst_artt.MoveNext
            Loop
            If blRecExist Then
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case SFISCI
            blRecExist = False
            slStr = "Not Checked"
            SQLQuery = "SELECT arttISCI2Contact FROM artt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            Do While Not rst_artt.EOF
                blRecExist = True
                If rst_artt!arttISCI2Contact = "1" Then
                    slStr = "Checked"
                    Exit Do
                End If
                rst_artt.MoveNext
            Loop
            If blRecExist Then
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case SFLABEL
            blRecExist = False
            slStr = "Not Checked"
            SQLQuery = "SELECT arttAffContact FROM artt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            Do While Not rst_artt.EOF
                blRecExist = True
                If rst_artt!arttAffContact = "1" Then
                    slStr = "Checked"
                    Exit Do
                End If
                rst_artt.MoveNext
            Loop
            If blRecExist Then
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = False
            End If
        Case SFPERSONNEL
            SQLQuery = "SELECT arttFirstName, arttLastName FROM artt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            If Not rst_artt.EOF Then
                Do While Not rst_artt.EOF
                    If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
                        ilIncludeStation = True
                        slStr = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            Exit Do
                        End If
                    Else
                        slStr = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
                        If Trim$(slStr) <> "" Then
                            ilIncludeStation = False
                            Exit Do
                        End If
                    End If
                    rst_artt.MoveNext
                Loop
            Else
                If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
                    ilIncludeStation = False
                End If
            End If
        Case SFAGREEMENT
            If tgStationInfo(ilShtt).sAgreementExist = "Y" Then
                SQLQuery = "SELECT attVefCode FROM att"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' AND"
                SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                Set rst_att = gSQLSelectCall(SQLQuery)
                If Not rst_att.EOF Then
                    slStr = "Active"
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    If Not ilIncludeStation Then
                        ilIncludeStation = True
                        slStr = "All"
                    End If
                Else
                    SQLQuery = "SELECT attVefCode FROM att"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
                    Set rst_att = gSQLSelectCall(SQLQuery)
                    If Not rst_att.EOF Then
                        slStr = "All"
                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                        If Not ilIncludeStation Then
                            ilIncludeStation = True
                            slStr = "None Active"
                        End If
                    Else
                        slStr = "None"
                    End If
                End If
            Else
                slStr = "None"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFWEGENER
            If tgStationInfo(ilShtt).sUsedForWegener = "Y" Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFXDS
            If tgStationInfo(ilShtt).sUsedForXDigital = "Y" Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFZONE
            ilValue = tgStationInfo(ilShtt).iTztCode
            ilTzt = gBinarySearchTzt(ilValue)
            If ilTzt <> -1 Then
                slStr = UCase$(Trim$(tgTimeZoneInfo(ilTzt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFENTERPRISEID
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sEnterpriseID))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFVEHICLEACTIVE  'Vehicle
            SQLQuery = "SELECT attVefCode FROM att"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' AND"
            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_att = gSQLSelectCall(SQLQuery)
            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        ilIncludeStation = True
                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
                        If ilVef <> -1 Then
                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                        Else
                            ilIncludeStation = False
                        End If
                        If ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = False
                End If
            Else
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
                        If ilVef <> -1 Then
                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                        Else
                            ilIncludeStation = True
                        End If
                        If Not ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = True
                End If
            End If
        Case SFVEHICLEALL  'Vehicle
            SQLQuery = "SELECT attVefCode FROM att"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_att = gSQLSelectCall(SQLQuery)
            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        ilIncludeStation = True
                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
                        If ilVef <> -1 Then
                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                        Else
                            ilIncludeStation = False
                        End If
                        If ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = False
                End If
            Else
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
                        If ilVef <> -1 Then
                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                        Else
                            ilIncludeStation = True
                        End If
                        If Not ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = True
                End If
            End If
        Case SFWEBADDRESS
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sWebAddress))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFWEBPW
            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sWebPW))
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFXDSID
            mFilterCompareLong tgStationInfo(ilShtt).lXDSStationID, llFilterDefIndex, ilIncludeStation
        Case SFZIP  'Zip
            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
                If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) = "") And (tgFilterDef(llFilterDefIndex).iOperator = 1) Then
                    ilIncludeStation = False
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
                    If slStr = "" Then
                        If (Trim$(tgStationInfo(ilShtt).sMailAddress1) <> "") Or (Trim$(tgStationInfo(ilShtt).sMailAddress2) <> "") Then
                            ilIncludeStation = True
                        End If
                    End If
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyZip))
                    If slStr = "" Then
                        If (Trim$(tgStationInfo(ilShtt).sPhyAddress1) <> "") Or (Trim$(tgStationInfo(ilShtt).sPhyAddress2) <> "") Then
                            ilIncludeStation = True
                        End If
                    End If
                Else
                    'Equal, treat with OR operator
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    If Not ilIncludeStation Then
                        ilIncludeStation = True
                        slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyZip))
                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    End If
                End If
            Else
                'Not Equal, treat with AND operator
                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                If ilIncludeStation Then
                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyZip))
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                End If
            End If
        Case SFTERRITORY  'Territory
            llValue = tgStationInfo(ilShtt).lMntCode
            ilMnt = gBinarySearchMnt(llValue, tgTerritoryInfo())
            If ilMnt <> -1 Then
                slStr = UCase$(Trim$(tgTerritoryInfo(ilMnt).sName))
            Else
                slStr = ""
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFMULTICAST
            If tgStationInfo(ilShtt).lMultiCastGroupID > 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFSISTER
            If tgStationInfo(ilShtt).lMarketClusterGroupID > 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        '6048
        Case SFEMAILADDRESS
            SQLQuery = "select arttEmail FROM artt WHERE arttShttCode = " & tgStationInfo(ilShtt).iCode
            Set rst_artt = gSQLSelectCall(SQLQuery)
            Do While Not rst_artt.EOF
                If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
                    ilIncludeStation = True
                    slStr = Trim$(rst_artt!arttEmail)
                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
                    If ilIncludeStation Then
                        Exit Do
                    End If
                Else
                    slStr = Trim$(rst_artt!arttEmail)
                    If Trim$(slStr) <> "" Then
                        ilIncludeStation = False
                        Exit Do
                    End If
                End If
                rst_artt.MoveNext
            Loop
        Case SFDUE  'Affidavits Due
            slStr = gBinarySearchStationCount(tgStationInfo(ilShtt).iCode)
            If slStr = "" Then
                ilValue = 0
            Else
                ilValue = Val(slStr)
            End If
            mFilterCompareInteger ilValue, llFilterDefIndex, ilIncludeStation
        '5/6/18: Test for Vendor ID
        Case SFLOGDELIVERY
            SQLQuery = "SELECT vatWvtVendorID FROM vat_Vendor_Agreement Left Outer Join att on vatAttCode = attCode "
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' AND"
            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_att = gSQLSelectCall(SQLQuery)
            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        ilIncludeStation = True
                        llValue = rst_att!vatwvtvendorid
                        mFilterCompareLong llValue, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = False
                End If
            Else
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        llValue = rst_att!vatwvtvendorid
                        mFilterCompareLong llValue, llFilterDefIndex, ilIncludeStation
                        If Not ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = True
                End If
            End If
        Case SFAUDIODELIVERY
            SQLQuery = "SELECT vatWvtVendorID FROM vat_Vendor_Agreement Left Outer Join att on vatAttCode = attCode "
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' AND"
            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
            Set rst_att = gSQLSelectCall(SQLQuery)
            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        ilIncludeStation = True
                        llValue = rst_att!vatwvtvendorid
                        mFilterCompareLong llValue, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = False
                End If
            Else
                If Not rst_att.EOF Then
                    Do While Not rst_att.EOF
                        llValue = rst_att!vatwvtvendorid
                        mFilterCompareLong llValue, llFilterDefIndex, ilIncludeStation
                        If Not ilIncludeStation Then
                            Exit Do
                        End If
                        rst_att.MoveNext
                    Loop
                Else
                    ilIncludeStation = True
                End If
            End If
        Case SFSERVICEAGREEMENT
            slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
            If slFrom <> "BOTH" Then
                SQLQuery = "SELECT Count(1) as SACount FROM att"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode
                SQLQuery = SQLQuery & " And attServiceAgreement = '" & "Y" & "')"
                Set rst_att = gSQLSelectCall(SQLQuery)
                If Not rst_att.EOF Then
                    If rst_att!SACount > 0 Then
                        slStr = "ONLY"
                    Else
                        slStr = "NONE"
                    End If
                Else
                    slStr = "NONE"
                End If
                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            Else
                ilIncludeStation = True
            End If
        End Select

End Sub


Private Sub mBranchToWebAddress()
    On Error GoTo ErrHand:

    imShttCode = Val(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
    SQLQuery = "SELECT shttWebAddress FROM shtt"
    SQLQuery = SQLQuery + " WHERE ("
    SQLQuery = SQLQuery & " ShttCode = " & imShttCode & ")"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If Not rst_Shtt.EOF Then
        ShellExecute 0&, vbNullString, rst_Shtt!shttWebAddress, vbNullString, vbNullString, vbNormalFocus
    End If
    grdStations.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mBranchToWebAddress"
    grdStations.Redraw = True
End Sub

Private Sub mSetShowCommentsAndContacts()
    If udcCommentGrid.Visible Then
        udcCommentGrid.Action 1 'Clear focus
    End If
    If udcContactGrid.Visible Then
        udcContactGrid.Action 1 'Clear Focus
    End If
End Sub

Private Sub mSaveCommentsAndContacts()
    If udcCommentGrid.Visible Then
        udcCommentGrid.Action 5 'Save
    End If
    If udcContactGrid.Visible Then
        If imShttCode > 0 Then
            If udcContactGrid.VerifyRights("M") Then
                udcContactGrid.Action 5 'Save
            End If
        End If
    End If
End Sub

Private Sub udcCommentGrid_CommentFocus()
    If udcContactGrid.Visible Then
        If imShttCode > 0 Then
            If udcContactGrid.VerifyRights("M") Then
                udcContactGrid.Action 5 'Save
            End If
        End If
    End If
End Sub

Private Sub udcCommentGrid_CommentStatus(slStatus As String)
    Dim slSQLQuery As String
    Dim ilShtt As Integer
    '11/26/17
    Dim blRepopRequired As Boolean
    
    On Error GoTo ErrHand:
    
    grdStations.TextMatrix(grdStations.Row, SCOMMENTINDEX) = slStatus
    If slStatus = "C" Then
        slSQLQuery = "UPDATE shtt SET shttCommentExist = 'Y'"
        slSQLQuery = slSQLQuery & " WHERE shttCode = " & grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX)
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "frmStationSearch-CommentStatus"
            Exit Sub
        End If
        '11/26/17
        blRepopRequired = False
        ilShtt = gBinarySearchStationInfoByCode(grdStations.TextMatrix(grdStations.Row, SSHTTCODEINDEX))
        If ilShtt <> -1 Then
            tgStationInfoByCode(ilShtt).sCommentExist = "Y"
            ilShtt = gBinarySearchStation(Trim$(tgStationInfoByCode(ilShtt).sCallLetters))
            If ilShtt <> -1 Then
                tgStationInfo(ilShtt).sCommentExist = "Y"
            Else
                blRepopRequired = True
            End If
        Else
            blRepopRequired = True
        End If
        '11/26/17
        gFileChgdUpdate "shtt.mkd", blRepopRequired
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-CommentStatus"
End Sub

Private Sub udcCommentGrid_GetCommentType()
    udcCommentGrid.WhichComment = Left$(Trim$(smCommentType), 1)
End Sub

Private Sub udcCommentGrid_StationSelected(ilShttCode As Integer, slVehicleName As String)
    Dim llRow As Long
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilNoAgreementInfo As Integer
    Dim llAgreementRow As Long
    Dim llStationRow As Long
    Dim llKey As Long
    
    mMousePointer vbHourglass
    If udcContactGrid.VerifyRights("M") Then
        udcContactGrid.Action 5 'Save
    End If
    grdSpotInfo.Visible = False
    grdPostedInfo.TextMatrix(0, PSELECTINDEX) = "Spts"
    bmPostedScrollAllowed = True
    lmPostedTopRow = -1
    If Trim$(slVehicleName) = "" Then
        grdSpotInfo.Visible = False
        grdPostedInfo.Visible = False
        grdAgreementInfo.TextMatrix(0, ASELECTINDEX) = "Wks"
        bmAgreementScrollAllowed = True
        lmAgreementTopRow = -1
        bmPostedScrollAllowed = True
        lmPostedTopRow = -1
    End If
    'mClearGrid grdStations
    mClearGrid grdAgreementInfo
    mClearGrid grdPostedInfo
    mClearGrid grdSpotInfo
'    For llRow = grdStations.FixedRows To grdStations.Rows - 1 Step 1
'        If ilShttCode = Val(grdStations.TextMatrix(llRow, SSHTTCODEINDEX)) Then
    For llKey = LBound(tmStationGridKey) To UBound(tmStationGridKey) - 1 Step 1
        llStationRow = tmStationGridKey(llKey).lRow
        If ilShttCode = Val(smStationGridData(llStationRow + SSHTTCODEINDEX)) Then
            bmStationScrollAllowed = True
            'grdStations.TopRow = llRow
            vbcStation.Value = llStationRow
            bmStationScrollAllowed = False
            mRepopGrids slVehicleName, 0, 0
            Exit For
        End If
    Next llKey
    mSetGridPosition
    mMousePointer vbDefault
End Sub

Private Sub udcCommentGrid_Tip(slTip As String)
    If slTip <> "" Then
        edcCommentTip.Visible = False
        lacCommentTip.Caption = slTip
        'lacCommentTip.Move (udcCommentGrid.Width - edcCommentTip.Width) \ 2, udcCommentGrid.Top - edcCommentTip.Height
        'lacCommentTip.Visible = True
        'lacCommentTip.ZOrder
        edcCommentTip.Text = slTip
        edcCommentTip.Width = lacCommentTip.Width
        edcCommentTip.Height = lacCommentTip.Height + 60
        'edcCommentTip.Move (udcCommentGrid.Width - edcCommentTip.Width) \ 2, udcCommentGrid.Top - edcCommentTip.Height, lacCommentTip.Width, lacCommentTip.Height + 60
        edcCommentTip.Move udcCommentGrid.Left + lmCommentStartLoc, udcCommentGrid.Top - edcCommentTip.Height
        edcCommentTip.Visible = True
        edcCommentTip.ZOrder
    Else
        'lacCommentTip.Visible = False
        edcCommentTip.Visible = False
    End If
End Sub

Private Sub mSetCommentButtons()
    Dim llStartLoc As Long
    cmcCloseComment.Move udcCommentGrid.Left + udcCommentGrid.Width - cmcCloseComment.Width - GRIDSCROLLWIDTH, udcCommentGrid.Top + 15, cmcCloseComment.Width, grdStations.RowHeight(0) - 15
    cmcCloseComment.Visible = True
    cmcCloseComment.ZOrder
    cmcAddComment.Move cmcCloseComment.Left - cmcAddComment.Width, cmcCloseComment.Top, cmcAddComment.Width, cmcCloseComment.Height
    cmcAddComment.Visible = True
    cmcAddComment.ZOrder
    pbcCommentType.Move cmcAddComment.Left - pbcCommentType.Width, cmcCloseComment.Top, pbcCommentType.Width, cmcCloseComment.Height
    pbcCommentType.Visible = True
    pbcCommentType.ZOrder
End Sub

Private Sub mHideCommentButtons()
    cmcCloseComment.Visible = False
    cmcAddComment.Visible = False
    pbcCommentType.Visible = False
End Sub


Private Sub mRepopGrids(slVehicleName As String, llAttCode As Long, llCpttCode As Long)
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilNoAgreementInfo As Integer
    Dim llAgreementRow As Long
    Dim llCptt As Long
    Dim llCpttRow As Long
    
    lmStationTopRow = grdStations.TopRow
    grdStations.Row = lmStationTopRow
    grdStations.Col = SSELECTINDEX
    imShttCode = Val(grdStations.TextMatrix(lmStationTopRow, SSHTTCODEINDEX))
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 3 'Populate
    udcCommentGrid.StationCode = imShttCode 'retain station so that correct comments show when Follow-up changed back to All Comments
    If rbcComments(0).Value Then
        udcCommentGrid.Action 3 'Populate
    End If
    If (udcCommentGrid.Visible) And (udcContactGrid.Visible = False) Then
        udcContactGrid.Visible = True
    End If
    If grdAgreementInfo.Visible Then
        mMousePointer vbHourglass
        grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Stns"
        If llAttCode <= 0 Then
            ilRet = mPopAgreementInfoGrid()
        End If
        If grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) <> "" Then
            grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "s"
            mSetStationGridData grdStations.Row, "s"
        End If
        bmStationScrollAllowed = False
        lmStationTopRow = grdStations.TopRow
        If grdPostedInfo.Visible Then
            ilNoAgreementInfo = 0
            If llAttCode > 0 Then
                For ilVef = grdAgreementInfo.FixedRows To grdAgreementInfo.Rows - 1 Step 1
                    If llAttCode = Val(grdAgreementInfo.TextMatrix(ilVef, AATTCODEINDEX)) Then
                        ilNoAgreementInfo = 1
                        llAgreementRow = ilVef
                        Exit For
                    End If
                Next ilVef
            End If
            If ilNoAgreementInfo <= 0 Then
                For ilVef = grdAgreementInfo.FixedRows To grdAgreementInfo.Rows - 1 Step 1
                    If StrComp(slVehicleName, grdAgreementInfo.TextMatrix(ilVef, AVEHICLEINDEX), vbTextCompare) = 0 Then
                        ilNoAgreementInfo = ilNoAgreementInfo + 1
                        llAgreementRow = ilVef
                        Exit For
                    End If
                Next ilVef
            End If
            If llCpttCode <= 0 Then
                grdSpotInfo.Visible = False
            End If
            If ilNoAgreementInfo = 1 Then
                bmAgreementScrollAllowed = True
                grdAgreementInfo.TopRow = llAgreementRow
                grdAgreementInfo.Row = llAgreementRow
                grdAgreementInfo.Col = ASELECTINDEX
                bmAgreementScrollAllowed = False
                lmAgreementTopRow = grdAgreementInfo.TopRow
                If llCpttCode <= 0 Then
                    ilRet = mPopPostedInfoGrid()
                Else
                    If grdSpotInfo.Visible Then
                        llCpttRow = -1
                        For llCptt = grdPostedInfo.FixedRows To grdPostedInfo.Rows - 1 Step 1
                            If llCpttCode = Val(grdPostedInfo.TextMatrix(llCptt, PCPTTINDEX)) Then
                                llCpttRow = llCptt
                                Exit For
                            End If
                        Next llCptt
                        If llCpttRow >= grdPostedInfo.FixedRows Then
                            bmPostedScrollAllowed = True
                            grdPostedInfo.TopRow = llCpttRow
                            grdPostedInfo.Row = llCpttRow
                            grdPostedInfo.Col = PSELECTINDEX
                            bmPostedScrollAllowed = False
                            lmPostedTopRow = grdPostedInfo.TopRow
                            ilRet = mPopSpotInfoGrid()
                        Else
                            grdSpotInfo.Visible = False
                        End If
                    End If
                End If
                grdPostedInfo.Visible = True
                grdAgreementInfo.TextMatrix(0, PSELECTINDEX) = "Veh"
            ElseIf ilNoAgreementInfo > 1 Then
                grdPostedInfo.Visible = False
            ElseIf ilNoAgreementInfo = 0 Then
                grdPostedInfo.Visible = False
                grdAgreementInfo.Visible = False
                bmAgreementScrollAllowed = True
                lmAgreementTopRow = -1
                grdAgreementInfo.TextMatrix(0, ASELECTINDEX) = "Wks"
                grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Veh"
            End If
        End If
        mMousePointer vbDefault
    End If
End Sub

Private Sub mGetUserType()
    On Error GoTo ErrHand:
    smUserType = "O"
    SQLQuery = "SELECT ustDntCode FROM Ust Where ustCode = " & igUstCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!ustDntCode > 0 Then
            SQLQuery = "SELECT dntType FROM Dnt Where dntCode = " & rst!ustDntCode
            Set rst_dnt = gSQLSelectCall(SQLQuery)
            If Not rst_dnt.EOF Then
                smUserType = rst_dnt!dntType
            End If
        End If
    End If
    If smUserType = "" Then
        smUserType = "O"
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-GetUserType"
End Sub

Private Function mConvertPhone(slInStr As String) As String
    Dim slOutStr As String
    Dim slChar As String
    Dim ilLoop As Integer
    Dim slStr As String
    
    slOutStr = ""
    slStr = Trim$(slInStr)
    If slStr <> "" Then
        For ilLoop = 1 To Len(slStr) Step 1
            slChar = Mid$(slStr, ilLoop, 1)
            If (slChar >= "0") And (slChar <= "9") Then
                slOutStr = slOutStr & slChar
            End If
        Next ilLoop
    End If
    mConvertPhone = slOutStr

End Function

Private Sub mPopFilterTypes()
    tgFilterTypes(0).sFieldName = "Area"
    tgFilterTypes(0).iSelect = SFAREA
    tgFilterTypes(0).sContainAllowed = "Y"
    tgFilterTypes(0).sEqualAllowed = "Y"
    tgFilterTypes(0).sRangeAllowed = "N"
    tgFilterTypes(0).sNoEqualAllowed = "Y"
    tgFilterTypes(0).sGreaterOrEqual = "N"
    tgFilterTypes(0).sCntrlType = "L"
    tgFilterTypes(0).iCountGroup = 0
    
    tgFilterTypes(1).sFieldName = "Call Letters Change Date"
    tgFilterTypes(1).iSelect = SFCALLLETTERSCHGDATE
    tgFilterTypes(1).sContainAllowed = "N"
    tgFilterTypes(1).sEqualAllowed = "Y"
    tgFilterTypes(1).sRangeAllowed = "Y"
    tgFilterTypes(1).sNoEqualAllowed = "N"
    tgFilterTypes(1).sGreaterOrEqual = "N"
    tgFilterTypes(1).sCntrlType = "E"
    tgFilterTypes(1).iCountGroup = 1

    tgFilterTypes(2).sFieldName = "Call Letters"
    tgFilterTypes(2).iSelect = SFCALLLETTERS
    tgFilterTypes(2).sContainAllowed = "Y"
    tgFilterTypes(2).sEqualAllowed = "N"
    tgFilterTypes(2).sRangeAllowed = "N"
    tgFilterTypes(2).sNoEqualAllowed = "N"
    tgFilterTypes(2).sGreaterOrEqual = "N"
    tgFilterTypes(2).sCntrlType = "L"
    tgFilterTypes(2).iCountGroup = 0

    tgFilterTypes(3).sFieldName = "City of License"
    tgFilterTypes(3).iSelect = SFCITYLIC
    tgFilterTypes(3).sContainAllowed = "Y"
    tgFilterTypes(3).sEqualAllowed = "Y"
    tgFilterTypes(3).sRangeAllowed = "N"
    tgFilterTypes(3).sNoEqualAllowed = "Y"
    tgFilterTypes(3).sGreaterOrEqual = "N"
    tgFilterTypes(3).sCntrlType = "L"
    tgFilterTypes(3).iCountGroup = 0

    tgFilterTypes(4).sFieldName = "Commercial"
    tgFilterTypes(4).iSelect = SFCOMMERCIAL
    tgFilterTypes(4).sContainAllowed = "N"
    tgFilterTypes(4).sEqualAllowed = "Y"
    tgFilterTypes(4).sRangeAllowed = "N"
    tgFilterTypes(4).sNoEqualAllowed = "N"
    tgFilterTypes(4).sGreaterOrEqual = "N"
    tgFilterTypes(4).sCntrlType = "T"
    tgFilterTypes(4).iCountGroup = 2

    tgFilterTypes(5).sFieldName = "County of License"
    tgFilterTypes(5).iSelect = SFCOUNTYLIC
    tgFilterTypes(5).sContainAllowed = "Y"
    tgFilterTypes(5).sEqualAllowed = "Y"
    tgFilterTypes(5).sRangeAllowed = "N"
    tgFilterTypes(5).sNoEqualAllowed = "Y"
    tgFilterTypes(5).sGreaterOrEqual = "N"
    tgFilterTypes(5).sCntrlType = "L"
    tgFilterTypes(5).iCountGroup = 0

    tgFilterTypes(6).sFieldName = "Daylight Savings"
    tgFilterTypes(6).iSelect = SFDAYLIGHT
    tgFilterTypes(6).sContainAllowed = "N"
    tgFilterTypes(6).sEqualAllowed = "Y"
    tgFilterTypes(6).sRangeAllowed = "N"
    tgFilterTypes(6).sNoEqualAllowed = "N"
    tgFilterTypes(6).sGreaterOrEqual = "N"
    tgFilterTypes(6).sCntrlType = "T"
    tgFilterTypes(6).iCountGroup = 3

    tgFilterTypes(7).sFieldName = "DMA Market"
    tgFilterTypes(7).iSelect = SFDMA
    tgFilterTypes(7).sContainAllowed = "Y"
    tgFilterTypes(7).sEqualAllowed = "Y"
    tgFilterTypes(7).sRangeAllowed = "N"
    tgFilterTypes(7).sNoEqualAllowed = "Y"
    tgFilterTypes(7).sGreaterOrEqual = "N"
    tgFilterTypes(7).sCntrlType = "L"
    tgFilterTypes(7).iCountGroup = 0
    
    tgFilterTypes(8).sFieldName = "Format"
    tgFilterTypes(8).iSelect = SFFORMAT
    tgFilterTypes(8).sContainAllowed = "Y"
    tgFilterTypes(8).sEqualAllowed = "Y"
    tgFilterTypes(8).sRangeAllowed = "N"
    tgFilterTypes(8).sNoEqualAllowed = "Y"
    tgFilterTypes(8).sGreaterOrEqual = "N"
    tgFilterTypes(8).sCntrlType = "L"
    tgFilterTypes(8).iCountGroup = 4

    tgFilterTypes(9).sFieldName = "Frequency"
    tgFilterTypes(9).iSelect = SFFREQ
    tgFilterTypes(9).sContainAllowed = "Y"
    tgFilterTypes(9).sEqualAllowed = "Y"
    tgFilterTypes(9).sRangeAllowed = "Y"
    tgFilterTypes(9).sNoEqualAllowed = "Y"
    tgFilterTypes(9).sGreaterOrEqual = "N"
    tgFilterTypes(9).sCntrlType = "E"
    tgFilterTypes(9).iCountGroup = 5

    tgFilterTypes(10).sFieldName = "Historical Start Date"
    tgFilterTypes(10).iSelect = SFHISTSTARTDATE
    tgFilterTypes(10).sContainAllowed = "Y"
    tgFilterTypes(10).sEqualAllowed = "Y"
    tgFilterTypes(10).sRangeAllowed = "Y"
    tgFilterTypes(10).sNoEqualAllowed = "N"
    tgFilterTypes(10).sGreaterOrEqual = "N"
    tgFilterTypes(10).sCntrlType = "E"
    tgFilterTypes(10).iCountGroup = 6

    tgFilterTypes(11).sFieldName = "ID"
    tgFilterTypes(11).iSelect = SFPERMID
    tgFilterTypes(11).sContainAllowed = "Y"
    tgFilterTypes(11).sEqualAllowed = "Y"
    tgFilterTypes(11).sRangeAllowed = "Y"
    tgFilterTypes(11).sNoEqualAllowed = "N"
    tgFilterTypes(11).sGreaterOrEqual = "N"
    tgFilterTypes(11).sCntrlType = "E"
    tgFilterTypes(11).iCountGroup = 0

    tgFilterTypes(12).sFieldName = "Mailing Address"
    tgFilterTypes(12).iSelect = SFMAILADDRESS
    tgFilterTypes(12).sContainAllowed = "Y"
    tgFilterTypes(12).sEqualAllowed = "N"
    tgFilterTypes(12).sRangeAllowed = "N"
    tgFilterTypes(12).sNoEqualAllowed = "N"
    tgFilterTypes(12).sGreaterOrEqual = "N"
    tgFilterTypes(12).sCntrlType = "E"
    tgFilterTypes(12).iCountGroup = 0

    tgFilterTypes(13).sFieldName = "Market Rep"
    tgFilterTypes(13).iSelect = SFMARKETREP
    tgFilterTypes(13).sContainAllowed = "Y"
    tgFilterTypes(13).sEqualAllowed = "Y"
    tgFilterTypes(13).sRangeAllowed = "N"
    tgFilterTypes(13).sNoEqualAllowed = "Y"
    tgFilterTypes(13).sGreaterOrEqual = "N"
    tgFilterTypes(13).sCntrlType = "L"
    tgFilterTypes(13).iCountGroup = 7

    tgFilterTypes(14).sFieldName = "Moniker"
    tgFilterTypes(14).iSelect = SFMONIKER
    tgFilterTypes(14).sContainAllowed = "Y"
    tgFilterTypes(14).sEqualAllowed = "Y"
    tgFilterTypes(14).sRangeAllowed = "N"
    tgFilterTypes(14).sNoEqualAllowed = "N"
    tgFilterTypes(14).sGreaterOrEqual = "N"
    tgFilterTypes(14).sCntrlType = "L"
    tgFilterTypes(14).iCountGroup = 8

    tgFilterTypes(15).sFieldName = "MSA Market"
    tgFilterTypes(15).iSelect = SFMSA
    tgFilterTypes(15).sContainAllowed = "Y"
    tgFilterTypes(15).sEqualAllowed = "Y"
    tgFilterTypes(15).sRangeAllowed = "N"
    tgFilterTypes(15).sNoEqualAllowed = "Y"
    tgFilterTypes(15).sGreaterOrEqual = "N"
    tgFilterTypes(15).sCntrlType = "L"
    tgFilterTypes(15).iCountGroup = 0

    tgFilterTypes(16).sFieldName = "On Air"
    tgFilterTypes(16).iSelect = SFONAIR
    tgFilterTypes(16).sContainAllowed = "N"
    tgFilterTypes(16).sEqualAllowed = "Y"
    tgFilterTypes(16).sRangeAllowed = "N"
    tgFilterTypes(16).sNoEqualAllowed = "N"
    tgFilterTypes(16).sGreaterOrEqual = "N"
    tgFilterTypes(16).sCntrlType = "T"
    tgFilterTypes(16).iCountGroup = 9

    tgFilterTypes(17).sFieldName = "Operator"
    tgFilterTypes(17).iSelect = SFOPERATOR
    tgFilterTypes(17).sContainAllowed = "Y"
    tgFilterTypes(17).sEqualAllowed = "Y"
    tgFilterTypes(17).sRangeAllowed = "N"
    tgFilterTypes(17).sNoEqualAllowed = "Y"
    tgFilterTypes(17).sGreaterOrEqual = "N"
    tgFilterTypes(17).iCountGroup = 10

    tgFilterTypes(18).sFieldName = "Owner"
    tgFilterTypes(18).iSelect = SFOWNER
    tgFilterTypes(18).sContainAllowed = "Y"
    tgFilterTypes(18).sEqualAllowed = "Y"
    tgFilterTypes(18).sRangeAllowed = "N"
    tgFilterTypes(18).sNoEqualAllowed = "Y"
    tgFilterTypes(18).sGreaterOrEqual = "N"
    tgFilterTypes(18).sCntrlType = "L"
    tgFilterTypes(18).iCountGroup = 11
  
    tgFilterTypes(19).sFieldName = "P12+"
    tgFilterTypes(19).iSelect = SFP12PLUS
    tgFilterTypes(19).sContainAllowed = "Y"
    tgFilterTypes(19).sEqualAllowed = "Y"
    tgFilterTypes(19).sRangeAllowed = "Y"
    tgFilterTypes(19).sNoEqualAllowed = "Y"
    tgFilterTypes(10).sGreaterOrEqual = "N"
    tgFilterTypes(19).sCntrlType = "E"
    tgFilterTypes(19).iCountGroup = 12
  
    tgFilterTypes(20).sFieldName = "Phone"
    tgFilterTypes(20).iSelect = SFPHONE
    tgFilterTypes(20).sContainAllowed = "Y"
    tgFilterTypes(20).sEqualAllowed = "Y"
    tgFilterTypes(20).sRangeAllowed = "N"
    tgFilterTypes(20).sNoEqualAllowed = "N"
    tgFilterTypes(20).sGreaterOrEqual = "N"
    tgFilterTypes(20).sCntrlType = "E"
    tgFilterTypes(20).iCountGroup = 13

    tgFilterTypes(21).sFieldName = "Physical Address"
    tgFilterTypes(21).iSelect = SFPHYADDRESS
    tgFilterTypes(21).sContainAllowed = "Y"
    tgFilterTypes(21).sEqualAllowed = "N"
    tgFilterTypes(21).sRangeAllowed = "N"
    tgFilterTypes(21).sNoEqualAllowed = "N"
    tgFilterTypes(21).sGreaterOrEqual = "N"
    tgFilterTypes(21).sCntrlType = "E"
    tgFilterTypes(21).iCountGroup = 0

    tgFilterTypes(22).sFieldName = "Serial #"
    tgFilterTypes(22).iSelect = SFSERIAL
    tgFilterTypes(22).sContainAllowed = "N"
    tgFilterTypes(22).sEqualAllowed = "Y"
    tgFilterTypes(22).sRangeAllowed = "N"
    tgFilterTypes(22).sNoEqualAllowed = "N"
    tgFilterTypes(22).sGreaterOrEqual = "N"
    tgFilterTypes(22).sCntrlType = "E"
    tgFilterTypes(22).iCountGroup = 0

    tgFilterTypes(23).sFieldName = "Service Rep"
    tgFilterTypes(23).iSelect = SFSERVICEREP
    tgFilterTypes(23).sContainAllowed = "Y"
    tgFilterTypes(23).sEqualAllowed = "Y"
    tgFilterTypes(23).sRangeAllowed = "N"
    tgFilterTypes(23).sNoEqualAllowed = "Y"
    tgFilterTypes(23).sGreaterOrEqual = "N"
    tgFilterTypes(23).sCntrlType = "L"
    tgFilterTypes(23).iCountGroup = 14

    tgFilterTypes(24).sFieldName = "State of License"
    tgFilterTypes(24).iSelect = SFSTATELIC
    tgFilterTypes(24).sContainAllowed = "Y"
    tgFilterTypes(24).sEqualAllowed = "Y"
    tgFilterTypes(24).sRangeAllowed = "N"
    tgFilterTypes(24).sNoEqualAllowed = "Y"
    tgFilterTypes(24).sGreaterOrEqual = "N"
    tgFilterTypes(24).sCntrlType = "L"
    tgFilterTypes(24).iCountGroup = 0

    tgFilterTypes(25).sFieldName = "Station E-Mail Check"
    tgFilterTypes(25).iSelect = SFEMAIL
    tgFilterTypes(25).sContainAllowed = "N"
    tgFilterTypes(25).sEqualAllowed = "Y"
    tgFilterTypes(25).sRangeAllowed = "N"
    tgFilterTypes(25).sNoEqualAllowed = "N"
    tgFilterTypes(25).sGreaterOrEqual = "N"
    tgFilterTypes(25).sCntrlType = "T"
    tgFilterTypes(25).iCountGroup = 15

    tgFilterTypes(26).sFieldName = "Station ISCI Mark"
    tgFilterTypes(26).iSelect = SFISCI
    tgFilterTypes(26).sContainAllowed = "N"
    tgFilterTypes(26).sEqualAllowed = "Y"
    tgFilterTypes(26).sRangeAllowed = "N"
    tgFilterTypes(26).sNoEqualAllowed = "N"
    tgFilterTypes(26).sGreaterOrEqual = "N"
    tgFilterTypes(26).sCntrlType = "T"
    tgFilterTypes(26).iCountGroup = 16

    tgFilterTypes(27).sFieldName = "Station Label Mark"
    tgFilterTypes(27).iSelect = SFLABEL
    tgFilterTypes(27).sContainAllowed = "N"
    tgFilterTypes(27).sEqualAllowed = "Y"
    tgFilterTypes(27).sRangeAllowed = "N"
    tgFilterTypes(27).sNoEqualAllowed = "N"
    tgFilterTypes(27).sGreaterOrEqual = "N"
    tgFilterTypes(27).sCntrlType = "T"
    tgFilterTypes(27).iCountGroup = 17
    
    tgFilterTypes(28).sFieldName = "Personnel Name"
    tgFilterTypes(28).iSelect = SFPERSONNEL
    tgFilterTypes(28).sContainAllowed = "Y"
    tgFilterTypes(28).sEqualAllowed = "Y"
    tgFilterTypes(28).sRangeAllowed = "N"
    tgFilterTypes(28).sNoEqualAllowed = "N"
    tgFilterTypes(28).sGreaterOrEqual = "N"
    tgFilterTypes(28).sCntrlType = "E"
    tgFilterTypes(28).iCountGroup = 18

    tgFilterTypes(29).sFieldName = "Station with Agreements"
    tgFilterTypes(29).iSelect = SFAGREEMENT
    tgFilterTypes(29).sContainAllowed = "N"
    tgFilterTypes(29).sEqualAllowed = "Y"
    tgFilterTypes(29).sRangeAllowed = "N"
    tgFilterTypes(29).sNoEqualAllowed = "N"
    tgFilterTypes(29).sGreaterOrEqual = "N"
    tgFilterTypes(29).sCntrlType = "T"
    tgFilterTypes(29).iCountGroup = 19

    tgFilterTypes(30).sFieldName = "Stations with Wegener"
    tgFilterTypes(30).iSelect = SFWEGENER
    tgFilterTypes(30).sContainAllowed = "N"
    tgFilterTypes(30).sEqualAllowed = "Y"
    tgFilterTypes(30).sRangeAllowed = "N"
    tgFilterTypes(30).sNoEqualAllowed = "N"
    tgFilterTypes(30).sGreaterOrEqual = "N"
    tgFilterTypes(30).sCntrlType = "T"
    tgFilterTypes(30).iCountGroup = 20

    tgFilterTypes(31).sFieldName = "Stations with XDS"
    tgFilterTypes(31).iSelect = SFXDS
    tgFilterTypes(31).sContainAllowed = "N"
    tgFilterTypes(31).sEqualAllowed = "Y"
    tgFilterTypes(31).sRangeAllowed = "N"
    tgFilterTypes(31).sNoEqualAllowed = "N"
    tgFilterTypes(31).sGreaterOrEqual = "N"
    tgFilterTypes(31).sCntrlType = "T"
    tgFilterTypes(31).iCountGroup = 21

    tgFilterTypes(32).sFieldName = "Territory"
    tgFilterTypes(32).iSelect = SFTERRITORY
    tgFilterTypes(32).sContainAllowed = "Y"
    tgFilterTypes(32).sEqualAllowed = "Y"
    tgFilterTypes(32).sRangeAllowed = "N"
    tgFilterTypes(32).sNoEqualAllowed = "Y"
    tgFilterTypes(32).sGreaterOrEqual = "N"
    tgFilterTypes(32).sCntrlType = "L"
    tgFilterTypes(32).iCountGroup = 0

    tgFilterTypes(33).sFieldName = "Time Zone"
    tgFilterTypes(33).iSelect = SFZONE
    tgFilterTypes(33).sContainAllowed = "N"
    tgFilterTypes(33).sEqualAllowed = "Y"
    tgFilterTypes(33).sRangeAllowed = "N"
    tgFilterTypes(33).sNoEqualAllowed = "Y"
    tgFilterTypes(33).sGreaterOrEqual = "N"
    tgFilterTypes(33).sCntrlType = "L"
    tgFilterTypes(33).iCountGroup = 0

    tgFilterTypes(34).sFieldName = "Transact Enterprise ID"
    tgFilterTypes(34).iSelect = SFENTERPRISEID
    tgFilterTypes(34).sContainAllowed = "Y"
    tgFilterTypes(34).sEqualAllowed = "Y"
    tgFilterTypes(34).sRangeAllowed = "N"
    tgFilterTypes(34).sNoEqualAllowed = "N"
    tgFilterTypes(34).sGreaterOrEqual = "N"
    tgFilterTypes(34).sCntrlType = "E"
    tgFilterTypes(34).iCountGroup = 0

    tgFilterTypes(35).sFieldName = "Vehicle-Active"
    tgFilterTypes(35).iSelect = SFVEHICLEACTIVE
    tgFilterTypes(35).sContainAllowed = "Y"
    tgFilterTypes(35).sEqualAllowed = "Y"
    tgFilterTypes(35).sRangeAllowed = "N"
    tgFilterTypes(35).sNoEqualAllowed = "Y"
    tgFilterTypes(35).sGreaterOrEqual = "N"
    tgFilterTypes(35).sCntrlType = "L"
    tgFilterTypes(35).iCountGroup = 22

    tgFilterTypes(45).sFieldName = "Vehicle-All"
    tgFilterTypes(45).iSelect = SFVEHICLEALL
    tgFilterTypes(45).sContainAllowed = "Y"
    tgFilterTypes(45).sEqualAllowed = "Y"
    tgFilterTypes(45).sRangeAllowed = "N"
    tgFilterTypes(45).sNoEqualAllowed = "Y"
    tgFilterTypes(45).sGreaterOrEqual = "N"
    tgFilterTypes(45).sCntrlType = "L"
    tgFilterTypes(45).iCountGroup = 22

    tgFilterTypes(36).sFieldName = "Web Address"
    tgFilterTypes(36).iSelect = SFWEBADDRESS
    tgFilterTypes(36).sContainAllowed = "Y"
    tgFilterTypes(36).sEqualAllowed = "Y"
    tgFilterTypes(36).sRangeAllowed = "N"
    tgFilterTypes(36).sNoEqualAllowed = "N"
    tgFilterTypes(36).sGreaterOrEqual = "N"
    tgFilterTypes(36).sCntrlType = "E"
    tgFilterTypes(36).iCountGroup = 23

    tgFilterTypes(37).sFieldName = "Web Password"
    tgFilterTypes(37).iSelect = SFWEBPW
    tgFilterTypes(37).sContainAllowed = "Y"
    tgFilterTypes(37).sEqualAllowed = "Y"
    tgFilterTypes(37).sRangeAllowed = "N"
    tgFilterTypes(37).sNoEqualAllowed = "N"
    tgFilterTypes(37).sGreaterOrEqual = "N"
    tgFilterTypes(37).sCntrlType = "E"
    tgFilterTypes(37).iCountGroup = 24

    tgFilterTypes(38).sFieldName = "XDS Station ID"
    tgFilterTypes(38).iSelect = SFXDSID
    tgFilterTypes(38).sContainAllowed = "N"
    tgFilterTypes(38).sEqualAllowed = "Y"
    tgFilterTypes(38).sRangeAllowed = "N"
    tgFilterTypes(38).sNoEqualAllowed = "Y"
    tgFilterTypes(38).sGreaterOrEqual = "N"
    tgFilterTypes(38).sCntrlType = "E"
    tgFilterTypes(38).iCountGroup = 0
  
    tgFilterTypes(39).sFieldName = "Zip"
    tgFilterTypes(39).iSelect = SFZIP
    tgFilterTypes(39).sContainAllowed = "Y"
    tgFilterTypes(39).sEqualAllowed = "Y"
    tgFilterTypes(39).sRangeAllowed = "Y"
    tgFilterTypes(39).sNoEqualAllowed = "Y"
    tgFilterTypes(39).sGreaterOrEqual = "N"
    tgFilterTypes(39).sCntrlType = "E"
    tgFilterTypes(39).iCountGroup = 0


    tgFilterTypes(40).sFieldName = "Station with Multi-Casts"
    tgFilterTypes(40).iSelect = SFMULTICAST
    tgFilterTypes(40).sContainAllowed = "N"
    tgFilterTypes(40).sEqualAllowed = "Y"
    tgFilterTypes(40).sRangeAllowed = "N"
    tgFilterTypes(40).sNoEqualAllowed = "N"
    tgFilterTypes(40).sGreaterOrEqual = "N"
    tgFilterTypes(40).sCntrlType = "T"
    tgFilterTypes(40).iCountGroup = 25

    tgFilterTypes(41).sFieldName = "Station with Sisters"
    tgFilterTypes(41).iSelect = SFSISTER
    tgFilterTypes(41).sContainAllowed = "N"
    tgFilterTypes(41).sEqualAllowed = "Y"
    tgFilterTypes(41).sRangeAllowed = "N"
    tgFilterTypes(41).sNoEqualAllowed = "N"
    tgFilterTypes(41).sGreaterOrEqual = "N"
    tgFilterTypes(41).sCntrlType = "T"
    tgFilterTypes(41).iCountGroup = 26
  
    tgFilterTypes(42).sFieldName = "Watts"
    tgFilterTypes(42).iSelect = SFWATTS
    tgFilterTypes(42).sContainAllowed = "N"
    tgFilterTypes(42).sEqualAllowed = "Y"
    tgFilterTypes(42).sRangeAllowed = "Y"
    tgFilterTypes(42).sNoEqualAllowed = "N"
    tgFilterTypes(42).sGreaterOrEqual = "N"
    tgFilterTypes(42).sCntrlType = "E"
    tgFilterTypes(42).iCountGroup = 27

    tgFilterTypes(43).sFieldName = "DMA Rank"
    tgFilterTypes(43).iSelect = SFDMARANK
    tgFilterTypes(43).sContainAllowed = "N"
    tgFilterTypes(43).sEqualAllowed = "N"
    tgFilterTypes(43).sRangeAllowed = "Y"
    tgFilterTypes(43).sNoEqualAllowed = "N"
    tgFilterTypes(43).sGreaterOrEqual = "N"
    tgFilterTypes(43).sCntrlType = "E"
    tgFilterTypes(43).iCountGroup = 28

    tgFilterTypes(44).sFieldName = "MSA Rank"
    tgFilterTypes(44).iSelect = SFMSARANK
    tgFilterTypes(44).sContainAllowed = "N"
    tgFilterTypes(44).sEqualAllowed = "N"
    tgFilterTypes(44).sRangeAllowed = "Y"
    tgFilterTypes(44).sNoEqualAllowed = "N"
    tgFilterTypes(44).sGreaterOrEqual = "N"
    tgFilterTypes(44).sCntrlType = "E"
    tgFilterTypes(44).iCountGroup = 29
    '6048 Dan M
    tgFilterTypes(46).sFieldName = "Email Address"
    tgFilterTypes(46).iSelect = SFEMAILADDRESS
    tgFilterTypes(46).sContainAllowed = "Y"
    tgFilterTypes(46).sEqualAllowed = "Y"
    tgFilterTypes(46).sRangeAllowed = "N"
    tgFilterTypes(46).sNoEqualAllowed = "N"
    tgFilterTypes(46).sGreaterOrEqual = "N"
    tgFilterTypes(46).sCntrlType = "E"
    tgFilterTypes(46).iCountGroup = 0
    'Affidavits Due
    tgFilterTypes(47).sFieldName = "Affidavits Due"
    tgFilterTypes(47).iSelect = SFDUE
    tgFilterTypes(47).sContainAllowed = "N"
    tgFilterTypes(47).sEqualAllowed = "N"
    tgFilterTypes(47).sRangeAllowed = "N"
    tgFilterTypes(47).sNoEqualAllowed = "N"
    tgFilterTypes(47).sGreaterOrEqual = "Y"
    tgFilterTypes(47).sCntrlType = "E"
    tgFilterTypes(47).iCountGroup = 30
    '5/6/18: Vendor definitions
    tgFilterTypes(48).sFieldName = "Vendor- Log Delivery"
    tgFilterTypes(48).iSelect = SFLOGDELIVERY
    tgFilterTypes(48).sContainAllowed = "N"
    tgFilterTypes(48).sEqualAllowed = "Y"
    tgFilterTypes(48).sRangeAllowed = "N"
    tgFilterTypes(48).sNoEqualAllowed = "Y"
    tgFilterTypes(48).sGreaterOrEqual = "N"
    tgFilterTypes(48).sCntrlType = "L"
    tgFilterTypes(48).iCountGroup = 31
    tgFilterTypes(49).sFieldName = "Vendor- Audio Delivery"
    tgFilterTypes(49).iSelect = SFAUDIODELIVERY
    tgFilterTypes(49).sContainAllowed = "N"
    tgFilterTypes(49).sEqualAllowed = "Y"
    tgFilterTypes(49).sRangeAllowed = "N"
    tgFilterTypes(49).sNoEqualAllowed = "Y"
    tgFilterTypes(49).sGreaterOrEqual = "N"
    tgFilterTypes(49).sCntrlType = "L"
    tgFilterTypes(49).iCountGroup = 31
    'Service agreememts
    tgFilterTypes(50).sFieldName = "Service Agreements"
    tgFilterTypes(50).iSelect = SFSERVICEAGREEMENT
    tgFilterTypes(50).sContainAllowed = "N"
    tgFilterTypes(50).sEqualAllowed = "Y"
    tgFilterTypes(50).sRangeAllowed = "N"
    tgFilterTypes(50).sNoEqualAllowed = "N"
    tgFilterTypes(50).sGreaterOrEqual = "N"
    tgFilterTypes(50).sCntrlType = "T"
    tgFilterTypes(50).iCountGroup = 0
    
End Sub

Private Sub udcContactGrid_ContactFocus()
    If udcCommentGrid.Visible Then
        udcCommentGrid.Action 5 'Save
    End If
End Sub

Private Sub mPopAdvt()
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim iFound As Integer
    Dim iLoop As Integer
    On Error GoTo ErrHand
    
    lbcAdvertiser.Clear
    lbcContract.Clear
    ''If txtWeek.Text = "" Then
    ''    Exit Sub
    ''End If
    ''If Trim$(txtNoWeeks.Text) = "" Then
    ''    Exit Sub
    ''End If
    ''
    ''If gIsDate(txtWeek.Text) = False Then
    ''    Beep
    ''    gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
    ''    txtWeek.SetFocus
    ''Else
    ''    smFWkDate = Format(txtWeek.Text, "m/d/yy")
    ''End If
    
    ''dFWeek = CDate(smFWkDate)
    ''iNoWeeks = 7 * CInt(txtNoWeeks.Text) - 1
    ''dLWeek = DateAdd("d", iNoWeeks, dFWeek)
    ''sLWeek = CStr(dLWeek)
    ''smLWkDate = Format$(sLWeek, "m/d/yy")
    ''SQLQuery = "SELECT adf.adfName, adf.adfCode from adf, lst"
    ''SQLQuery = SQLQuery + " WHERE (adf.adfCode = lst.lstAdfCode"
    ''SQLQuery = SQLQuery + " AND lst.lstStartDate BETWEEN '" & smFWkDate & "' AND '" & smLWkDate & "')"
    ''SQLQuery = SQLQuery + " ORDER BY adf.adfName"
    'SQLQuery = "SELECT adf.adfName, adf.adfCode from ADF_Advertisers adf"
    SQLQuery = "SELECT adfName, adfCode"
    SQLQuery = SQLQuery & " FROM ADF_Advertisers"
    SQLQuery = SQLQuery + " ORDER BY adfName"

    

    'SQLQuery = "SELECT adf.adfName, adf.adfCode from adf"
    'SQLQuery = SQLQuery + " ORDER BY adf.adfName"
    
    Set rst = gSQLSelectCall(SQLQuery)
    'lbcAdvertiser.Clear
    While Not rst.EOF
        iFound = False
        'For iLoop = 0 To lbcAdvertiser.ListCount - 1 Step 1
        '    If lbcAdvertiser.ItemData(iLoop) = rst!adfCode Then
        '        iFound = True
        '        Exit For
        '    End If
        'Next iLoop
        If Not iFound Then
            lbcAdvertiser.AddItem Trim$(rst!adfName) '& ", " & rst(1).Value
            lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = rst!adfCode
        End If
        rst.MoveNext
    Wend
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopAdvt"
End Sub

Private Sub mPopContract()
    Dim slSDate As String
    Dim slEDate As String
    Dim ilAdfCode As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilVef As Integer
    
    On Error GoTo ErrHand
    lbcContract.Clear
    If smStationType = "" Then
        Exit Sub
    End If
    slSDate = Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX))
    If slSDate = "" Then
        'gMsgBox "Date must be specified.", vbOKOnly
        'txtWeek.SetFocus
        Exit Sub
    End If
    If slSDate <> "[All]" Then
        If gIsDate(slSDate) = False Then
            Beep
            'gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            'txtWeek.SetFocus
            Exit Sub
        End If
        slEDate = DateAdd("d", 6, slSDate)
    Else
        slSDate = smWeek54
        slEDate = smWeek1
    End If
    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    mMousePointer vbHourglass
    ilAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
    
    'SQLQuery = "SELECT DISTINCT lstCntrNo"
    'SQLQuery = SQLQuery & " FROM ADF_Advertisers, lst"
    'SQLQuery = SQLQuery + " WHERE (adfCode = lstAdfCode"
    'SQLQuery = SQLQuery + " AND lstCntrNo <> 0"
    'SQLQuery = SQLQuery + " AND lstStartDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND lstStartDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
    'SQLQuery = SQLQuery + " AND adfCode = " & ilAdfCode & ")"
    'SQLQuery = SQLQuery + " ORDER BY lstCntrNo"
    ReDim imPostBuyVef(0 To 0) As Integer
    For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
        SQLQuery = "SELECT DISTINCT lstCntrNo"
        SQLQuery = SQLQuery & " FROM lst"
        SQLQuery = SQLQuery + " WHERE ("
        SQLQuery = SQLQuery + " lstLogVefCode = " & tgVehicleInfo(ilVef).iCode
        SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery + " AND lstCntrNo <> 0"
        SQLQuery = SQLQuery + " AND lstAdfCode = " & ilAdfCode & ")"
        SQLQuery = SQLQuery + " ORDER BY lstCntrNo"
        
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            imPostBuyVef(UBound(imPostBuyVef)) = tgVehicleInfo(ilVef).iCode
            ReDim Preserve imPostBuyVef(0 To UBound(imPostBuyVef) + 1) As Integer
        End If
        Do While Not rst.EOF
            slStr = rst!lstCntrNo
            ilIndex = SendMessageByString(lbcContract.hwnd, LB_FINDSTRING, -1, slStr)
            If ilIndex < 0 Then
                lbcContract.AddItem rst!lstCntrNo  ', " & rst(1).Value & ""
            End If
            rst.MoveNext
        Loop
    Next ilVef
    If lbcContract.ListCount > 0 Then
        lbcContract.AddItem "[All]", 0
    Else
        lbcContract.AddItem "[None]", 0
    End If
    mMousePointer vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopContract"
End Sub

Private Sub mEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilLang                        slNameCode                *
'*  slCode                        ilCode                        ilRet                     *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (grdPostBuy.Row < grdPostBuy.FixedRows) Or (grdPostBuy.Row >= grdPostBuy.Rows) Or (grdPostBuy.Col < grdPostBuy.FixedCols) Or (grdPostBuy.Col > PSCONTRACTINDEX) Then
        Exit Sub
    End If
    lmEnableRow = grdPostBuy.Row
    lmEnableCol = grdPostBuy.Col
    imCtrlVisible = True
    'pbcArrow.Visible = False
    'pbcArrow.Move grdPostBuy.Left - pbcArrow.Width, grdPostBuy.Top + grdPostBuy.RowPos(grdPostBuy.Row) + (grdPostBuy.RowHeight(grdPostBuy.Row) - pbcArrow.Height) / 2
    'pbcArrow.Visible = True

    Select Case grdPostBuy.Col
        Case PSSTATIONTYPEINDEX
            pbcStationType.Move grdPostBuy.Left + grdPostBuy.ColPos(grdPostBuy.Col) + 30, grdPostBuy.Top + grdPostBuy.RowPos(grdPostBuy.Row) + 15, grdPostBuy.ColWidth(grdPostBuy.Col) - 30, grdPostBuy.RowHeight(grdPostBuy.Row) - 15
            pbcStationType.Visible = True
            pbcStationType.SetFocus
        Case PSADVERTISERINDEX
            mSetLbcGridControl lbcAdvertiser
            edcDropDown.MaxLength = 0
        Case PSWEEKOFINDEX
            mSetLbcGridControl lbcDate
            edcDropDown.MaxLength = 0
        Case PSCONTRACTINDEX
            mPopContract
            'mPopContractsAndGetVehicles
            mSetLbcGridControl lbcContract
            edcDropDown.MaxLength = 0
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
    Dim slStr As String

    If (lmEnableRow >= grdPostBuy.FixedRows) And (lmEnableRow < grdPostBuy.Rows) Then
        Select Case lmEnableCol
            Case PSSTATIONTYPEINDEX
                grdPostBuy.TextMatrix(lmEnableRow, lmEnableCol) = smStationType
            Case PSWEEKOFINDEX
                slStr = edcDropDown.Text
                grdPostBuy.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PSADVERTISERINDEX
                slStr = edcDropDown.Text
                grdPostBuy.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case PSCONTRACTINDEX
                slStr = edcDropDown.Text
                grdPostBuy.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    'pbcArrow.Visible = False
    edcDropDown.Visible = False
    cmcDropDown.Visible = False
    lbcDate.Visible = False
    lbcAdvertiser.Visible = False
    lbcContract.Visible = False
    pbcStationType.Visible = False
    mSetCommands
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdPostBuy.Row < grdPostBuy.FixedRows) Or (grdPostBuy.Row >= grdPostBuy.Rows) Or (grdPostBuy.Col < grdPostBuy.FixedCols) Or (grdPostBuy.Col > PSCONTRACTINDEX) Then
        Exit Sub
    End If
    imCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdPostBuy.Col - 1 Step 1
        llColPos = llColPos + grdPostBuy.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdPostBuy.ColWidth(grdPostBuy.Col)
    ilCol = grdPostBuy.Col
    Do While ilCol < grdPostBuy.Cols - 1
        If (Trim$(grdPostBuy.TextMatrix(grdPostBuy.Row - 1, grdPostBuy.Col)) <> "") And (Trim$(grdPostBuy.TextMatrix(grdPostBuy.Row - 1, grdPostBuy.Col)) = Trim$(grdPostBuy.TextMatrix(grdPostBuy.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdPostBuy.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdPostBuy.Col
        Case PSWEEKOFINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcDate.Visible = True
            edcDropDown.SetFocus
        Case PSADVERTISERINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcAdvertiser.Visible = True
        Case PSCONTRACTINDEX
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcContract.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub

Private Sub mSetLbcGridControl(lbcCtrl As ListBox)
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    
    If (grdPostBuy.Row < grdPostBuy.FixedRows) Or (grdPostBuy.Row >= grdPostBuy.Rows) Or (grdPostBuy.Col < grdPostBuy.FixedCols) Or (grdPostBuy.Col > PSCONTRACTINDEX) Then
        Exit Sub
    End If
    If lmEnableCol = PSWEEKOFINDEX Then
        edcDropDown.Move grdPostBuy.Left + grdPostBuy.ColPos(grdPostBuy.Col) + 30, grdPostBuy.Top + grdPostBuy.RowPos(grdPostBuy.Row) + 15, grdPostBuy.ColWidth(grdPostBuy.Col) - cmcDropDown.Width - 30, grdPostBuy.RowHeight(grdPostBuy.Row) - 15
    ElseIf lmEnableCol = PSADVERTISERINDEX Then
        edcDropDown.Move grdPostBuy.Left + grdPostBuy.ColPos(grdPostBuy.Col) + 30, grdPostBuy.Top + grdPostBuy.RowPos(grdPostBuy.Row) + 15, grdPostBuy.ColWidth(grdPostBuy.Col) - cmcDropDown.Width - 30, grdPostBuy.RowHeight(grdPostBuy.Row) - 15
    ElseIf lmEnableCol = PSCONTRACTINDEX Then
        edcDropDown.Move grdPostBuy.Left + grdPostBuy.ColPos(grdPostBuy.Col) + 30, grdPostBuy.Top + grdPostBuy.RowPos(grdPostBuy.Row) + 15, grdPostBuy.ColWidth(grdPostBuy.Col) - cmcDropDown.Width - 30, grdPostBuy.RowHeight(grdPostBuy.Row) - 15
    End If
    cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top, cmcDropDown.Width, edcDropDown.Height
    lbcCtrl.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height, edcDropDown.Width + cmcDropDown.Width
    lbcCtrl.ZOrder
    gSetListBoxHeight lbcCtrl, 6
    slStr = grdPostBuy.Text
    ilIndex = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRING, -1, slStr)
    If ilIndex >= 0 Then
        lbcCtrl.ListIndex = ilIndex
        edcDropDown.Text = lbcCtrl.List(lbcCtrl.ListIndex)
    Else
        lbcCtrl.ListIndex = -1
        edcDropDown.Text = ""
    End If
    If edcDropDown.Height > grdPostBuy.RowHeight(grdPostBuy.Row) - 15 Then
        edcDropDown.FontName = "Arial"
        edcDropDown.Height = grdPostBuy.RowHeight(grdPostBuy.Row) - 15
    End If
    edcDropDown.Visible = True
    cmcDropDown.Visible = True
    lbcCtrl.Visible = True
    edcDropDown.SetFocus
End Sub

Private Sub mPopDate()
    Dim llWeek1 As Long
    Dim llWeek54 As Long
    Dim llDate As Long
    lbcDate.Clear
    llWeek1 = gDateValue(smWeek1)
    llWeek54 = gDateValue(smWeek54)
    For llDate = llWeek54 To llWeek1 Step 7
        lbcDate.AddItem Format(llDate, sgShowDateForm)
    Next llDate
End Sub

Private Sub mDropdownChangeEvent(lbcCtrl As ListBox)
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer

    slStr = edcDropDown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        lbcCtrl.ListIndex = llRow
        edcDropDown.Text = lbcCtrl.List(lbcCtrl.ListIndex)
        edcDropDown.SelStart = ilLen
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub

Private Sub mSetPSCommands()
    cmcGen.Enabled = False
    If lbcDate.ListIndex < 0 Then
        Exit Sub
    End If
    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    If lbcContract.ListIndex < 0 Then
        Exit Sub
    End If
    cmcGen.Enabled = True
End Sub

Private Function mCreatePostBuyStationFilter() As Boolean
    Dim slSDate As String
    Dim slEDate As String
    Dim ilAdfCode As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim blFound As Boolean
    Dim ilVef As Integer
    Dim ilShtt As Integer
    
    On Error GoTo ErrHand
    mCreatePostBuyStationFilter = True
    ReDim tgFilterDef(0 To 0) As FILTERDEF
    ReDim tgNotFilterDef(0 To 0) As FILTERDEF
    ilIndex = -1
    For ilLoop = 0 To UBound(tgFilterTypes) - 1 Step 1
        If StrComp(Trim$(tgFilterTypes(ilLoop).sFieldName), "Call Letters", vbTextCompare) = 0 Then
            ilIndex = ilLoop
        End If
    Next ilLoop
    If ilIndex = -1 Then
        Exit Function
    End If
    slSDate = Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX))
    If slSDate <> "[All]" Then
        slEDate = DateAdd("d", 6, slSDate)
    Else
        slSDate = smWeek54
        slEDate = smWeek1
    End If
    ilAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
    'For ilVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
    For ilVef = 0 To UBound(imPostBuyVef) - 1 Step 1
        SQLQuery = "SELECT DISTINCT cpttVefCode, cpttShfCode"
        SQLQuery = SQLQuery & " FROM cptt WHERE "
        SQLQuery = SQLQuery & " (cpttvefCode = " & imPostBuyVef(ilVef)   'tgVehicleInfo(ilVef).iCode
        'If lacUserOption.Caption <> "All Stations" Then
        If smStationType <> "All Stations" Then
            SQLQuery = SQLQuery & " AND cpttStatus = 0 AND cpttPostingStatus < 2"
        End If
        SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slEDate, sgSQLDateForm) & "')"
        Set rst_Cptt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Cptt.EOF
            SQLQuery = "SELECT Count(lstCode)"
            SQLQuery = SQLQuery & " FROM lst"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery + " lstLogVefCode = " & rst_Cptt!cpttvefcode
            SQLQuery = SQLQuery + " AND lstLogDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
            If lbcContract.ListIndex > 0 Then
                SQLQuery = SQLQuery & " AND lstCntrNo = " & lbcContract.List(lbcContract.ListIndex)
            Else
                SQLQuery = SQLQuery + " AND lstCntrNo <> 0"
            End If
            SQLQuery = SQLQuery + " AND lstAdfCode = " & ilAdfCode & ")"
            Set rst_Lst = gSQLSelectCall(SQLQuery)
            If Not rst_Lst.EOF And (rst_Lst(0).Value <> 0) Then
                blFound = False
                For ilLoop = 0 To UBound(tgFilterDef) - 1 Step 1
                    If tgFilterDef(ilLoop).lFromValue = rst_Cptt!cpttshfcode Then
                        blFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not blFound Then
                    ilUpper = UBound(tgFilterDef)
                    ilShtt = gBinarySearchStationInfoByCode(rst_Cptt!cpttshfcode)
                    If ilShtt <> -1 Then
                        tgFilterDef(ilUpper).iCountGroup = tgFilterTypes(ilIndex).iCountGroup
                        tgFilterDef(ilUpper).sCntrlType = tgFilterTypes(ilIndex).sCntrlType
                        tgFilterDef(ilUpper).iSelect = tgFilterTypes(ilIndex).iSelect
                        tgFilterDef(ilUpper).iOperator = 1
                        tgFilterDef(ilUpper).lFromValue = rst_Cptt!cpttshfcode
                        tgFilterDef(ilUpper).sFromValue = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                        tgFilterDef(ilUpper).lToValue = 0
                        tgFilterDef(ilUpper).sToValue = ""
                        ReDim Preserve tgFilterDef(0 To ilUpper + 1) As FILTERDEF
                    End If
                End If
            End If
            rst_Cptt.MoveNext
        Loop
    Next ilVef
    If UBound(tgFilterDef) > 0 Then
        lacFilter.Caption = "On"
        mBuildFilter
        mPopStationsGrid
        udcCommentGrid.FilterStatus = True  'vbTrue
        mMousePointer vbDefault
    Else
        mMousePointer vbDefault
        MsgBox "No Stations Found"
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mCreatePostBuyStationFilter"
    mCreatePostBuyStationFilter = False
End Function

Private Sub mPopContractsAndGetVehicles()
    Dim slSDate As String
    Dim slEDate As String
    Dim ilAdfCode As Integer
    Dim llVef As Long
    Dim llVpf As Long
    Dim ilLoop As Integer
    Dim blFound As Boolean
    Dim ilTest As Integer
    Dim slVefType As String
    
    On Error GoTo ErrHand
    lbcContract.Clear
    ReDim imPostBuyVef(0 To 0) As Integer
    slSDate = Trim$(grdPostBuy.TextMatrix(grdPostBuy.FixedRows, PSWEEKOFINDEX))
    If slSDate = "" Then
        'gMsgBox "Date must be specified.", vbOKOnly
        'txtWeek.SetFocus
        Exit Sub
    End If
    If gIsDate(slSDate) = False Then
        Beep
        'gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        'txtWeek.SetFocus
        Exit Sub
    End If
    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    slEDate = DateAdd("d", 6, slSDate)
    ilAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
    SQLQuery = "SELECT DISTINCT chfCntrNo, chfCode"
    SQLQuery = SQLQuery + " FROM ADF_Advertisers, "
    SQLQuery = SQLQuery & "CHF_Contract_Header"
    SQLQuery = SQLQuery + " WHERE (adfCode = chfAdfCode"
    SQLQuery = SQLQuery + " AND adfCode = " & ilAdfCode
    SQLQuery = SQLQuery + " AND chfStartDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND chfEndDate >= '" & Format$(slSDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND chfDelete = 'N'"
    SQLQuery = SQLQuery + " AND (chfStatus = 'O'"
    SQLQuery = SQLQuery + " OR chfStatus = 'H'))"
    SQLQuery = SQLQuery + " ORDER BY chfCntrNo"
    
    Set rst_chf = gSQLSelectCall(SQLQuery)
    Do While Not rst_chf.EOF
        lbcContract.AddItem rst_chf!chfCntrNo  ', " & rst(1).Value & ""
        SQLQuery = "SELECT clfVefCode FROM clf_Contract_Line WHERE clfChfCode = " & rst_chf!chfCode
        Set rst_clf = gSQLSelectCall(SQLQuery)
        Do While Not rst_clf.EOF
            llVef = gBinarySearchVef(CLng(rst_clf!clfVefCode))
            If llVef <> -1 Then
                slVefType = tgVehicleInfo(llVef).sVehType
                If (slVefType = "C") Or (slVefType = "S") Or (slVefType = "G") Or (slVefType = "I") Or (slVefType = "L") Then
                    blFound = False
                    For ilTest = 0 To UBound(imPostBuyVef) - 1 Step 1
                        If imPostBuyVef(ilTest) = rst_clf!clfVefCode Then
                            blFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not blFound Then
                        imPostBuyVef(UBound(imPostBuyVef)) = rst_clf!clfVefCode
                        ReDim Preserve imPostBuyVef(0 + UBound(imPostBuyVef) + 1) As Integer
                        If slVefType = "S" Then
                            llVpf = gBinarySearchVpf(CLng(rst_clf!clfVefCode))
                            If llVpf <> -1 Then
                                For ilLoop = 0 To UBound(tgVpfOptions) Step 1
                                    llVef = gBinarySearchVef(CLng(tgVpfOptions(ilLoop).ivefKCode))
                                    If llVef <> -1 Then
                                        If (ilLoop <> llVpf) And (tgVpfOptions(llVpf).iSAGroupNo = tgVpfOptions(ilLoop).iSAGroupNo) And (tgVehicleInfo(llVef).sVehType = "A") Then
                                            blFound = False
                                            For ilTest = 0 To UBound(imPostBuyVef) - 1 Step 1
                                                If imPostBuyVef(ilTest) = tgVpfOptions(ilLoop).ivefKCode Then
                                                    blFound = True
                                                    Exit For
                                                End If
                                            Next ilTest
                                            If Not blFound Then
                                                imPostBuyVef(UBound(imPostBuyVef)) = tgVpfOptions(ilLoop).ivefKCode
                                                ReDim Preserve imPostBuyVef(0 + UBound(imPostBuyVef) + 1) As Integer
                                            End If
                                        End If
                                    End If
                                Next ilLoop
                            End If
                        End If
                    End If
                End If
            End If
            rst_clf.MoveNext
        Loop
        rst_chf.MoveNext
    Loop
    lbcContract.AddItem "[All]", 0

    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopContractsAndGetVehicles"
End Sub

Private Function mGetGameInfo(llGsfCode As Long) As String
    Dim slStr As String
    Dim ilTeam As Integer
    
    slStr = ""
    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfCode = " & llGsfCode & ")"
    Set rst_Gsf = gSQLSelectCall(SQLQuery)
    If Not rst_Gsf.EOF Then
        slStr = Trim$(Str$(rst_Gsf!gsfGameNo))
        ''Feed Source
        'If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
        '    If rst_gsf!gsfFeedSource = "V" Then
        '        slStr = "Feed: Visting"
        '    Else
        '        slStr = "Feed: Home"
        '    End If
        'End If
        ''Language
        'If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
        '    For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
        '        If tgLangInfo(ilLang).iCode = rst_gsf!gsfLangMnfCode Then
        '            If slStr = "" Then
        '                slStr = Trim$(tgLangInfo(ilLang).sName)
        '            Else
        '                slStr = slStr & " " & Trim$(tgLangInfo(ilLang).sName)
        '            End If
        '            Exit For
        '        End If
        '    Next ilLang
        'End If
        'Visiting Team
        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfVisitMnfCode Then
                If slStr = "" Then
                    slStr = Trim$(tgTeamInfo(ilTeam).sName)
                Else
                    slStr = slStr & " " & Trim$(tgTeamInfo(ilTeam).sName)
                End If
                Exit For
            End If
        Next ilTeam
        'Home Team
        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfHomeMnfCode Then
                If slStr = "" Then
                    slStr = Trim$(tgTeamInfo(ilTeam).sName)
                Else
                    slStr = slStr & " @ " & Trim$(tgTeamInfo(ilTeam).sName)
                End If
                Exit For
            End If
        Next ilTeam
        'Air Date
        slStr = slStr & " on " & Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
        'Start Time
        slStr = slStr & " at " & Format$(rst_Gsf!gsfAirTime, sgShowTimeWSecForm)
    End If
    mGetGameInfo = slStr
End Function

Private Function mGetMGAndMissedInfo(ilAst As Integer) As String
    Dim slDateTimeMsg As String
    Dim slMissedReason As String
    Dim ilMnfCode As Integer
    Dim slStr As String
    Dim llAdf As Long
    Dim slAdvtName As String
    
    On Error GoTo ErrHandle
    mGetMGAndMissedInfo = ""
    slDateTimeMsg = ""
    slMissedReason = ""
    ilMnfCode = 0
    smMGAirDate = ""
    smMGAirTime = ""
    If gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_MG) Or gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_REPLACEMENT) Then
        'SQLQuery = "SELECT altLinkToAstCode FROM alt WHERE altAstCode = " & tmAstInfo(ilAst).lCode
        'Set rst = gSQLSelectCall(SQLQuery)
        'If Not rst.EOF Then
        '    If rst!altLinkToAstCode > 0 Then
                'Get Missed
        '        SQLQuery = "SELECT astCode, astAirDate, astAirTime, astLsfCode FROM ast WHERE AstCode = " & rst!altLinkToAstCode
            If tmAstInfo(ilAst).lLkAstCode > 0 Then
                SQLQuery = "SELECT astCode, astDatCode, astAirDate, astAirTime, astLsfCode, astFeedDate, astFeedTime FROM ast WHERE AstCode = " & tmAstInfo(ilAst).lLkAstCode
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    smMissedFeedDate = Format(rst!astFeedDate, sgShowDateForm)
                    smMissedFeedTime = Format(rst!astFeedTime, sgShowTimeWOSecForm)
                    If gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_MG) Then
                        slDateTimeMsg = "Missed: " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    ElseIf gIsAstStatus(tmAstInfo(ilAst).iStatus, ASTEXTENDED_REPLACEMENT) Then
                        slAdvtName = ""
                        SQLQuery = "SELECT lstAdfCode FROM lst where lstcode = " & rst!astLsfCode
                        Set rst_Lst = gSQLSelectCall(SQLQuery)
                        If Not rst_Lst.EOF Then
                            llAdf = gBinarySearchAdf(CLng(rst_Lst!lstAdfCode))
                            If llAdf <> -1 Then
                                slAdvtName = " " & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                            End If
                        End If
                        slDateTimeMsg = "Replaced:" & slAdvtName & " " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    End If
                    ''Get Missed reason reference
                    'SQLQuery = "SELECT altMnfMissed FROM alt where altAstcode = " & rst!astCode
                    'Set rst = gSQLSelectCall(SQLQuery)
                    'If Not rst.EOF Then
                    '    ilMnfCode = rst!altMnfMissed
                    'End If
                End If
            End If
        'End If
    ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(ilAst).iStatus)).iPledged = 2 Then
        'SQLQuery = "SELECT altLinkToAstCode, altMnfMissed FROM alt WHERE altAstCode = " & tmAstInfo(ilAst).lCode
        'Set rst = gSQLSelectCall(SQLQuery)
        'If Not rst.EOF Then
        '    ilMnfCode = rst!altMnfMissed
        '    If rst!altLinkToAstCode > 0 Then
                'Get MG or Replacement
        '        SQLQuery = "SELECT astCode, astStatus, astAirDate, astAirTime, astLsfCode FROM ast WHERE AstCode = " & rst!altLinkToAstCode
            If tmAstInfo(ilAst).lLkAstCode > 0 Then
                SQLQuery = "SELECT astCode, astStatus, astAirDate, astAirTime, astLsfCode, astAdfCode FROM ast WHERE AstCode = " & tmAstInfo(ilAst).lLkAstCode
                Set rst = gSQLSelectCall(SQLQuery)
                If Not rst.EOF Then
                    If gIsAstStatus(rst!astStatus, ASTEXTENDED_MG) Then
                        'slDateTimeMsg = "MG: " & Format(rst!astAirDate, sgShowDateForm) & " " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                        smMGAirDate = Format(rst!astAirDate, sgShowDateForm)
                        smMGAirTime = Format(rst!astAirTime, sgShowTimeWOSecForm)
                    ElseIf gIsAstStatus(rst!astStatus, ASTEXTENDED_REPLACEMENT) Then
                        slAdvtName = ""
                        SQLQuery = "SELECT lstAdfCode FROM lst where lstcode = " & rst!astLsfCode
                        Set rst_Lst = gSQLSelectCall(SQLQuery)
                        If Not rst_Lst.EOF Then
                            llAdf = gBinarySearchAdf(CLng(rst_Lst!lstAdfCode))
                            If llAdf <> -1 Then
                                slAdvtName = " " & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                            End If
                        End If
                        slDateTimeMsg = "Replacement:" & slAdvtName & " " & Format(rst!astAirDate, sgShowDateForm) & " at " & Format(rst!astAirTime, sgShowTimeWOSecForm)
                    End If
                End If
            End If
        'End If
    Else
        mGetMGAndMissedInfo = ""
        Exit Function
    End If
    'Get missed reason
    ilMnfCode = tmAstInfo(ilAst).iMissedMnfCode
    If ilMnfCode > 0 Then
        SQLQuery = "select mnfName from MNF_Multi_Names where mnfCode = " & ilMnfCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            slMissedReason = rst!mnfName
        End If
    End If
    If slMissedReason = "" Then
        mGetMGAndMissedInfo = slDateTimeMsg
    Else
        If slDateTimeMsg = "" Then
            mGetMGAndMissedInfo = slMissedReason
        Else
            mGetMGAndMissedInfo = slDateTimeMsg & ", " & slMissedReason
        End If
    End If
    Exit Function
ErrHandle:
    gHandleError "AffErrorLog.txt", "CP Date/Time-mGetMGAndMissedInfo"
    'Exit Function
    Resume Next
End Function

Private Sub mBuildSingleFilter()
    Dim blSingle As Boolean
    Dim llFilterDefIndex As Integer
    Dim llNotFilterDefIndex As Integer
    Dim slSQLQuery As String
    Dim slFrom As String
    Dim slTo As String
    Dim slFieldName As String
    Dim slWhere As String
    
    On Error GoTo ErrHandle
    ReDim imFilterShttCode(0 To 0) As Integer
    blSingle = False
    If (UBound(tmFilterLink) = 1) And (lacFilter.Caption = "On") Then
        llFilterDefIndex = tmFilterLink(0).lFilterDefIndex
        If llFilterDefIndex >= 0 Then
            If tmFilterLink(0).lNextAnd = -1 Then
                If (tgFilterDef(llFilterDefIndex).iOperator = 0) Or (tgFilterDef(llFilterDefIndex).iOperator = 1) Or (tgFilterDef(llFilterDefIndex).iOperator = 4) Then
                    llNotFilterDefIndex = tmFilterLink(0).lNotFilterDefIndex
                    If llNotFilterDefIndex < 0 Then
                        blSingle = True
                    End If
                End If
            End If
        End If
    End If
    If Not blSingle Then
        Exit Sub
    End If
    slSQLQuery = ""
    slWhere = ""
    Select Case tgFilterDef(llFilterDefIndex).iSelect
        Case SFAREA
'            llValue = tgStationInfo(ilShtt).lAreaMntCode
'            ilMnt = gBinarySearchMnt(llValue, tgAreaInfo())
'            If ilMnt <> -1 Then
'                slStr = UCase$(Trim$(tgAreaInfo(ilMnt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            slSQLQuery = "SELECT Distinct shttCode FROM shtt LEFT OUTER JOIN mnt ON shttAreaMntCode = mntCode WHERE "
            slFieldName = "UCase(mntName)"
            slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
        Case SFCALLLETTERSCHGDATE
'            'Get date from clt
'            SQLQuery = "SELECT cltEndDate FROM clt"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " cltShfCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_clt = gSQLSelectCall(SQLQuery)
'            If Not rst_clt.EOF Then
'                Do While Not rst_clt.EOF
'                    mFilterCompareLong gDateValue(Format(rst_clt!cltEndDate, sgShowDateForm)), llFilterDefIndex, ilIncludeStation
'                    If ilIncludeStation Then
'                        Exit Do
'                    End If
'                    rst_clt.MoveNext
'                Loop
'            Else
'                ilIncludeStation = False
'            End If
        Case SFCALLLETTERS  'Station
'            If sgStationSearchCallSource <> "P" Then
'                If (tgFilterDef(llFilterDefIndex).iOperator = 1) Then
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters)) & ", " & Trim$(tgStationInfo(ilShtt).sMarket)
'                Else
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters))
'                End If
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                If (Not ilIncludeStation) And (tgFilterDef(llFilterDefIndex).iOperator = 1) Then
'                    slStr = Trim$(UCase(tgFilterDef(llFilterDefIndex).sFromValue))
'                    ilPos = InStr(1, slStr, ",", vbTextCompare)
'                    If ilPos > 0 Then
'                        slStr = Trim$(Left(slStr, ilPos - 1))
'                    End If
'                    SQLQuery = "SELECT cltCallLetters FROM clt"
'                    SQLQuery = SQLQuery + " WHERE ("
'                    SQLQuery = SQLQuery & " cltShfCode = " & tgStationInfo(ilShtt).iCode & " AND"
'                    SQLQuery = SQLQuery & " cltCallLetters = '" & slStr & "')"
'                    Set rst_clt = gSQLSelectCall(SQLQuery)
'                    If Not rst_clt.EOF Then
'                        ilIncludeStation = True
'                    End If
'                End If
'                If (Not ilIncludeStation) And (tgFilterDef(llFilterDefIndex).iOperator = 0) Then
'                    SQLQuery = "SELECT cltCallLetters FROM clt"
'                    SQLQuery = SQLQuery + " WHERE ("
'                    SQLQuery = SQLQuery & " cltShfCode = " & tgStationInfo(ilShtt).iCode & " AND"
'                    SQLQuery = SQLQuery & " cltCallLetters LIKE '%" & Trim$(UCase(tgFilterDef(llFilterDefIndex).sFromValue)) & "%')"
'                    Set rst_clt = gSQLSelectCall(SQLQuery)
'                    If Not rst_clt.EOF Then
'                        ilIncludeStation = True
'                    End If
'                End If
'            Else
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sCallLetters))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            End If
            slSQLQuery = "SELECT Distinct shttCode FROM shtt LEFT OUTER JOIN clt ON shttCode = cltShfCode WHERE "
            slWhere = "UCase(shttCallLetters) LIKE '%" & UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue)) & "%'"
            slWhere = slWhere & " OR UCase(cltCallLetters) LIKE '%" & UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue)) & "%'"
        Case SFCITYLIC
'            llValue = tgStationInfo(ilShtt).lCityLicMntCode
'            ilMnt = gBinarySearchMnt(llValue, tgCityInfo())
'            If ilMnt <> -1 Then
'                slStr = UCase$(Trim$(tgCityInfo(ilMnt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFCOMMERCIAL
'            slStr = tgStationInfo(ilShtt).sStationType
'            If slStr = "N" Then
'                slStr = "Non-Commercial"
'            Else
'                slStr = "Commercial"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFCOUNTYLIC
'            llValue = tgStationInfo(ilShtt).lCountyLicMntCode
'            ilMnt = gBinarySearchMnt(llValue, tgCountyInfo())
'            If ilMnt <> -1 Then
'                slStr = UCase$(Trim$(tgCountyInfo(ilMnt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFDAYLIGHT
'            ilValue = tgStationInfo(ilShtt).iAckDaylight
'            If ilValue = 1 Then
'                slStr = "Ignore Daylight Savings"
'            Else
'                slStr = "Honor Daylight Savings"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFDMA  'DMA
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMarket))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            slSQLQuery = "SELECT Distinct shttCode FROM shtt LEFT OUTER JOIN mkt ON shttMktCode = mktCode WHERE "
            slFieldName = "UCase(mktName)"
            slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
        Case SFDMARANK
'            llMkt = gBinarySearchMkt(CLng(tgStationInfo(ilShtt).iMktCode))
'            If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (Trim$(tgFilterDef(llFilterDefIndex).sToValue) <> "") Then
'                If llMkt <> -1 Then
'                    If tgMarketInfo(llMkt).iRank = 0 Then
'                        mFilterCompareInteger 999, llFilterDefIndex, ilIncludeStation
'                    Else
'                        mFilterCompareInteger tgMarketInfo(llMkt).iRank, llFilterDefIndex, ilIncludeStation
'                    End If
'                Else
'                    ilIncludeStation = False
'                End If
'            Else
'                If llMkt <> -1 Then
'                    If tgMarketInfo(llMkt).iRank <> 0 Then
'                        ilIncludeStation = False
'                    End If
'                Else
'                    ilIncludeStation = False
'                End If
'            End If
        Case SFFORMAT  'Format
'            llValue = tgStationInfo(ilShtt).iFormatCode
'            ilFmt = gBinarySearchFmt(CInt(llValue))
'            If ilFmt <> -1 Then
'                slStr = UCase$(Trim$(tgFormatInfo(ilFmt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFFREQ
'            mFilterCompareCurrency tgStationInfo(ilShtt).sFrequency, llFilterDefIndex, ilIncludeStation
        Case SFHISTSTARTDATE
'            mFilterCompareLong tgStationInfo(ilShtt).lHistStartDate, llFilterDefIndex, ilIncludeStation
        Case SFPERMID
'            mFilterCompareLong tgStationInfo(ilShtt).lPermStationID, llFilterDefIndex, ilIncludeStation
        Case SFMAILADDRESS
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMailAddress1))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            If Not ilIncludeStation Then
'                ilIncludeStation = True
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMailAddress2))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            End If
'            If Not ilIncludeStation Then
'                ilIncludeStation = True
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sMailState))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            End If
'            If Not ilIncludeStation Then
'                ilIncludeStation = True
'                llValue = tgStationInfo(ilShtt).lMailCityMntCode
'                ilMnt = gBinarySearchMnt(llValue, tgCityInfo())
'                If ilMnt <> -1 Then
'                    slStr = UCase$(Trim$(tgCityInfo(ilMnt).sName))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                Else
'                    ilIncludeStation = False
'                End If
'            End If
        Case SFMARKETREP
'            slStr = ""
'            ilValue = tgStationInfo(ilShtt).iMktRepUstCode
'            If ilValue <> 0 Then
'                For llLoop = 0 To UBound(tgMarketRepInfo) - 1 Step 1
'                    If tgMarketRepInfo(llLoop).iUstCode = ilValue Then
'                        slStr = UCase$(Trim$(tgMarketRepInfo(llLoop).sName))
'                        Exit For
'                    End If
'                Next llLoop
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFMONIKER
'            llValue = tgStationInfo(ilShtt).lMonikerMntCode
'            ilMnt = gBinarySearchMnt(llValue, tgMonikerInfo())
'            If ilMnt <> -1 Then
'                slStr = UCase$(Trim$(tgMonikerInfo(ilMnt).sName))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            Else
'                ilIncludeStation = False
'            End If
        Case SFMSA  'MSA
'            llValue = tgStationInfo(ilShtt).iMSAMktCode
'            llMSA = gBinarySearchMSAMkt(llValue)
'            If llMSA <> -1 Then
'                slStr = UCase$(Trim$(tgMSAMarketInfo(llMSA).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
            slSQLQuery = "SELECT Distinct shttCode FROM shtt LEFT OUTER JOIN met ON shttMetCode = metCode WHERE "
            slFieldName = "UCase(metName)"
            slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
        Case SFMSARANK
'            llMkt = gBinarySearchMSAMkt(CLng(tgStationInfo(ilShtt).iMSAMktCode))
'            If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (Trim$(tgFilterDef(llFilterDefIndex).sToValue) <> "") Then
'                If llMkt <> -1 Then
'                    If tgMSAMarketInfo(llMkt).iRank = 0 Then
'                        mFilterCompareInteger 999, llFilterDefIndex, ilIncludeStation
'                    Else
'                        mFilterCompareInteger tgMSAMarketInfo(llMkt).iRank, llFilterDefIndex, ilIncludeStation
'                    End If
'                Else
'                    ilIncludeStation = False
'                End If
'            Else
'                If llMkt <> -1 Then
'                    If tgMSAMarketInfo(llMkt).iRank <> 0 Then
'                        ilIncludeStation = False
'                    End If
'                Else
'                    ilIncludeStation = False
'                End If
'            End If
        Case SFONAIR
'            If tgStationInfo(ilShtt).sOnAir = "N" Then
'                slStr = "Off Air"
'            Else
'                slStr = "On Air"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFOPERATOR
'            llValue = tgStationInfo(ilShtt).lOperatorMntCode
'            ilMnt = gBinarySearchMnt(llValue, tgOperatorInfo())
'            If ilMnt <> -1 Then
'                slStr = UCase$(Trim$(tgOperatorInfo(ilMnt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFOWNER  'Owner
'            llValue = tgStationInfo(ilShtt).lOwnerCode
'            llOwner = mBinarySearchOwner(llValue)
'            If llOwner <> -1 Then
'                slStr = UCase$(Trim$(tgOwnerInfo(llOwner).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFP12PLUS
'            mFilterCompareLong tgStationInfo(ilShtt).lAudP12Plus, llFilterDefIndex, ilIncludeStation
        Case SFWATTS
'            mFilterCompareLong tgStationInfo(ilShtt).lWatts, llFilterDefIndex, ilIncludeStation
        Case SFPHONE
'            slStr = tgStationInfo(ilShtt).sPhone
'            If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
'                mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
'                If Not ilIncludeStation Then
'                    ilIncludeStation = True
'                    slStr = tgStationInfo(ilShtt).sFax
'                    mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
'                End If
'                If Not ilIncludeStation Then
'                    SQLQuery = "SELECT arttPhone, arttFax FROM artt"
'                    SQLQuery = SQLQuery + " WHERE ("
'                    SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
'                    Set rst_artt = gSQLSelectCall(SQLQuery)
'                    Do While Not rst_artt.EOF
'                        ilIncludeStation = True
'                        slStr = rst_artt!arttPhone
'                        mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
'                        If ilIncludeStation Then
'                            Exit Do
'                        End If
'                        ilIncludeStation = True
'                        slStr = rst_artt!arttFax
'                        mFilterComparePhone slStr, llFilterDefIndex, ilIncludeStation
'                        If ilIncludeStation Then
'                            Exit Do
'                        End If
'                        rst_artt.MoveNext
'                    Loop
'                End If
'            Else
'                If Trim$(slStr) <> "" Then
'                    ilIncludeStation = False
'                Else
'                    SQLQuery = "SELECT arttPhone FROM artt"
'                    SQLQuery = SQLQuery + " WHERE ("
'                    SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
'                    Set rst_artt = gSQLSelectCall(SQLQuery)
'                    Do While Not rst_artt.EOF
'                        slStr = rst_artt!arttPhone
'                        If Trim$(slStr) <> "" Then
'                            ilIncludeStation = False
'                            Exit Do
'                        End If
'                        rst_artt.MoveNext
'                    Loop
'                End If
'            End If
        Case SFPHYADDRESS
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyAddress1))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            If Not ilIncludeStation Then
'                ilIncludeStation = True
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyAddress2))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            End If
'            If Not ilIncludeStation Then
'                ilIncludeStation = True
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyState))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            End If
'            If Not ilIncludeStation Then
'                ilIncludeStation = True
'                llValue = tgStationInfo(ilShtt).lPhyCityMntCode
'                ilMnt = gBinarySearchMnt(llValue, tgCityInfo())
'                If ilMnt <> -1 Then
'                    slStr = UCase$(Trim$(tgCityInfo(ilMnt).sName))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                Else
'                    ilIncludeStation = False
'                End If
'            End If
        Case SFSERIAL
'            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
'                If (tgFilterDef(llFilterDefIndex).iOperator = 1) And (Trim(tgFilterDef(llFilterDefIndex).sFromValue) = "") Then
'                    'treat with AND operator
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo1))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    If ilIncludeStation Then
'                        slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo2))
'                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    End If
'                Else
'                    'Equal, treat with OR operator
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo1))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    If Not ilIncludeStation Then
'                        ilIncludeStation = True
'                        slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo2))
'                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    End If
'                End If
'            Else
'                'Not Equal, treat with AND operator
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo1))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                If ilIncludeStation Then
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sSerialNo2))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                End If
'            End If
        Case SFSERVICEREP
'            slStr = ""
'            ilValue = tgStationInfo(ilShtt).iServRepUstCode
'            If ilValue <> 0 Then
'                For llLoop = 0 To UBound(tgServiceRepInfo) - 1 Step 1
'                    If tgServiceRepInfo(llLoop).iUstCode = ilValue Then
'                        slStr = UCase$(Trim$(tgServiceRepInfo(llLoop).sName))
'                        Exit For
'                    End If
'                Next llLoop
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFSTATELIC
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sStateLic))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFEMAIL
'            blRecExist = False
'            slStr = "Not Checked"
'            SQLQuery = "SELECT arttWebEMail FROM artt"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_artt = gSQLSelectCall(SQLQuery)
'            Do While Not rst_artt.EOF
'                blRecExist = True
'                If rst_artt!arttWebEMail = "Y" Then
'                    slStr = "Checked"
'                    Exit Do
'                End If
'                rst_artt.MoveNext
'            Loop
'            If blRecExist Then
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            Else
'                ilIncludeStation = False
'            End If
        Case SFISCI
'            blRecExist = False
'            slStr = "Not Checked"
'            SQLQuery = "SELECT arttISCI2Contact FROM artt"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_artt = gSQLSelectCall(SQLQuery)
'            Do While Not rst_artt.EOF
'                blRecExist = True
'                If rst_artt!arttISCI2Contact = "1" Then
'                    slStr = "Checked"
'                    Exit Do
'                End If
'                rst_artt.MoveNext
'            Loop
'            If blRecExist Then
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            Else
'                ilIncludeStation = False
'            End If
        Case SFLABEL
'            blRecExist = False
'            slStr = "Not Checked"
'            SQLQuery = "SELECT arttAffContact FROM artt"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_artt = gSQLSelectCall(SQLQuery)
'            Do While Not rst_artt.EOF
'                blRecExist = True
'                If rst_artt!arttAffContact = "1" Then
'                    slStr = "Checked"
'                    Exit Do
'                End If
'                rst_artt.MoveNext
'            Loop
'            If blRecExist Then
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'            Else
'                ilIncludeStation = False
'            End If
        Case SFPERSONNEL
'            SQLQuery = "SELECT arttFirstName, arttLastName FROM artt"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " arttShttCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_artt = gSQLSelectCall(SQLQuery)
'            If Not rst_artt.EOF Then
'                Do While Not rst_artt.EOF
'                    If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
'                        ilIncludeStation = True
'                        slStr = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
'                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                        If ilIncludeStation Then
'                            Exit Do
'                        End If
'                    Else
'                        slStr = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
'                        If Trim$(slStr) <> "" Then
'                            ilIncludeStation = False
'                            Exit Do
'                        End If
'                    End If
'                    rst_artt.MoveNext
'                Loop
'            Else
'                If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) <> "") Or (tgFilterDef(llFilterDefIndex).iOperator <> 1) Then
'                    ilIncludeStation = False
'                End If
'            End If
        Case SFAGREEMENT
'            If tgStationInfo(ilShtt).sAgreementExist = "Y" Then
'                SQLQuery = "SELECT attVefCode FROM att"
'                SQLQuery = SQLQuery + " WHERE ("
'                SQLQuery = SQLQuery & " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' AND"
'                SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
'                Set rst_att = gSQLSelectCall(SQLQuery)
'                If Not rst_att.EOF Then
'                    slStr = "Active"
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    If Not ilIncludeStation Then
'                        ilIncludeStation = True
'                        slStr = "All"
'                    End If
'                Else
'                    slStr = "All"
'                End If
'            Else
'                slStr = "None"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFWEGENER
'            If tgStationInfo(ilShtt).sUsedForWegener = "Y" Then
'                slStr = "Yes"
'            Else
'                slStr = "No"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFXDS
'            If tgStationInfo(ilShtt).sUsedForXDigital = "Y" Then
'                slStr = "Yes"
'            Else
'                slStr = "No"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFZONE
'            ilValue = tgStationInfo(ilShtt).iTztCode
'            ilTzt = gBinarySearchTzt(ilValue)
'            If ilTzt <> -1 Then
'                slStr = UCase$(Trim$(tgTimeZoneInfo(ilTzt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFENTERPRISEID
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sEnterpriseID))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFVEHICLEACTIVE  'Vehicle
'            SQLQuery = "SELECT attVefCode FROM att"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' AND"
'            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_att = gSQLSelectCall(SQLQuery)
'            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
'                If Not rst_att.EOF Then
'                    Do While Not rst_att.EOF
'                        ilIncludeStation = True
'                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
'                        If ilVef <> -1 Then
'                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
'                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                        Else
'                            ilIncludeStation = False
'                        End If
'                        If ilIncludeStation Then
'                            Exit Do
'                        End If
'                        rst_att.MoveNext
'                    Loop
'                Else
'                    ilIncludeStation = False
'                End If
'            Else
'                If Not rst_att.EOF Then
'                    Do While Not rst_att.EOF
'                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
'                        If ilVef <> -1 Then
'                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
'                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                        Else
'                            ilIncludeStation = True
'                        End If
'                        If Not ilIncludeStation Then
'                            Exit Do
'                        End If
'                        rst_att.MoveNext
'                    Loop
'                Else
'                    ilIncludeStation = True
'                End If
'            End If
            slSQLQuery = "SELECT distinct attShfCode FROM att LEFT OUTER JOIN vef_Vehicles ON attVefCode = vefCode WHERE "
            slWhere = " attOffAir >= '" & Format(gNow(), sgSQLDateForm) & "' AND attDropDate >= '" & Format(gNow(), sgSQLDateForm) & "' "
            If tgFilterDef(llFilterDefIndex).iOperator = 0 Then
                slWhere = slWhere & " AND UCase(vefName) LIKE '%" & UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue)) & "%'"
            ElseIf tgFilterDef(llFilterDefIndex).iOperator = 1 Then
                slWhere = slWhere & " AND UCase(vefName) = '" & UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue)) & "'"
            Else
                slSQLQuery = ""
                slWhere = ""
            End If
        Case SFVEHICLEALL  'Vehicle
'            SQLQuery = "SELECT attVefCode FROM att"
'            SQLQuery = SQLQuery + " WHERE ("
'            SQLQuery = SQLQuery & " attShfCode = " & tgStationInfo(ilShtt).iCode & ")"
'            Set rst_att = gSQLSelectCall(SQLQuery)
'            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
'                If Not rst_att.EOF Then
'                    Do While Not rst_att.EOF
'                        ilIncludeStation = True
'                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
'                        If ilVef <> -1 Then
'                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
'                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                        Else
'                            ilIncludeStation = False
'                        End If
'                        If ilIncludeStation Then
'                            Exit Do
'                        End If
'                        rst_att.MoveNext
'                    Loop
'                Else
'                    ilIncludeStation = False
'                End If
'            Else
'                If Not rst_att.EOF Then
'                    Do While Not rst_att.EOF
'                        ilVef = gBinarySearchVef(CLng(rst_att!attvefCode))
'                        If ilVef <> -1 Then
'                            slStr = UCase$(Trim$(tgVehicleInfo(ilVef).sVehicle))
'                            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                        Else
'                            ilIncludeStation = True
'                        End If
'                        If Not ilIncludeStation Then
'                            Exit Do
'                        End If
'                        rst_att.MoveNext
'                    Loop
'                Else
'                    ilIncludeStation = True
'                End If
'            End If
            slSQLQuery = "SELECT Distinct attShfCode FROM att LEFT OUTER JOIN vef_Vehicles ON attVefCode = vefCode WHERE "
            slFieldName = "UCase(vefName)"
            slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
        Case SFWEBADDRESS
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sWebAddress))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFWEBPW
'            slStr = UCase$(Trim$(tgStationInfo(ilShtt).sWebPW))
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFXDSID
'            mFilterCompareLong tgStationInfo(ilShtt).lXDSStationID, llFilterDefIndex, ilIncludeStation
        Case SFZIP  'Zip
'            If tgFilterDef(llFilterDefIndex).iOperator <> 2 Then
'                If (Trim$(tgFilterDef(llFilterDefIndex).sFromValue) = "") And (tgFilterDef(llFilterDefIndex).iOperator = 1) Then
'                    ilIncludeStation = False
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
'                    If slStr = "" Then
'                        If (Trim$(tgStationInfo(ilShtt).sMailAddress1) <> "") Or (Trim$(tgStationInfo(ilShtt).sMailAddress2) <> "") Then
'                            ilIncludeStation = True
'                        End If
'                    End If
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyZip))
'                    If slStr = "" Then
'                        If (Trim$(tgStationInfo(ilShtt).sPhyAddress1) <> "") Or (Trim$(tgStationInfo(ilShtt).sPhyAddress2) <> "") Then
'                            ilIncludeStation = True
'                        End If
'                    End If
'                Else
'                    'Equal, treat with OR operator
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    If Not ilIncludeStation Then
'                        ilIncludeStation = True
'                        slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyZip))
'                        mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                    End If
'                End If
'            Else
'                'Not Equal, treat with AND operator
'                slStr = UCase$(Trim$(tgStationInfo(ilShtt).sZip))
'                mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                If ilIncludeStation Then
'                    slStr = UCase$(Trim$(tgStationInfo(ilShtt).sPhyZip))
'                    mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
'                End If
'            End If
        Case SFTERRITORY  'Territory
'            llValue = tgStationInfo(ilShtt).lMntCode
'            ilMnt = gBinarySearchMnt(llValue, tgTerritoryInfo())
'            If ilMnt <> -1 Then
'                slStr = UCase$(Trim$(tgTerritoryInfo(ilMnt).sName))
'            Else
'                slStr = ""
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFMULTICAST
'            If tgStationInfo(ilShtt).lMultiCastGroupID > 0 Then
'                slStr = "Yes"
'            Else
'                slStr = "No"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        Case SFSISTER
'            If tgStationInfo(ilShtt).lMarketClusterGroupID > 0 Then
'                slStr = "Yes"
'            Else
'                slStr = "No"
'            End If
'            mFilterCompareStr slStr, llFilterDefIndex, ilIncludeStation
        '6048
        Case SFEMAILADDRESS
            slFieldName = "UCase(arttEmail)"
            slFrom = UCase(Trim$(tgFilterDef(llFilterDefIndex).sFromValue))
            slSQLQuery = "select arttshttcode FROM artt WHERE  "
        Case SFDUE
        '5/6/18: Obtain stations with matching Vendor ID
        Case SFLOGDELIVERY
            slSQLQuery = "SELECT Distinct attShfCode FROM vat_Vendor_Agreement Left Outer Join att on vatAttCode = attCode WHERE "
            slFieldName = "UCase(vatWvtVendorID)"
            slFrom = Trim$(Str$(tgFilterDef(llFilterDefIndex).lFromValue))
        Case SFAUDIODELIVERY
            slSQLQuery = "SELECT Distinct attShfCode FROM vat_Vendor_Agreement Left Outer Join att on vatAttCode = attCode WHERE "
            slFieldName = "UCase(vatWvtVendorID)"
            slFrom = Trim$(Str$(tgFilterDef(llFilterDefIndex).lFromValue))
        Case SFSERVICEAGREEMENT
    End Select
    If slSQLQuery = "" Then
        Exit Sub
    End If
    If slWhere = "" Then
        'Operator: 0=Contains; 1=Equal; 2=Not Equal; 3=Range; 4=Greater or Equal
        Select Case tgFilterDef(llFilterDefIndex).iOperator
            Case 0  'Contains
                slSQLQuery = slSQLQuery & slFieldName & " LIKE '%" & slFrom & "%'"
            Case 1  'Equal
                slSQLQuery = slSQLQuery & slFieldName & " = '" & slFrom & "'"
            Case 2  'Not Equal
                slSQLQuery = ""
            Case 3  'Range
                slSQLQuery = ""
            Case 4  'GreaterOrEqual
                slSQLQuery = ""
        End Select
    Else
        slSQLQuery = slSQLQuery & slWhere
    End If
    If slSQLQuery = "" Then
        Exit Sub
    End If
    Set rst = gSQLSelectCall(slSQLQuery)
    If rst.EOF Then
        ReDim imFilterShttCode(0 To 1) As Integer
        imFilterShttCode(0) = -1
        Exit Sub
    End If
    Do While Not rst.EOF
        imFilterShttCode(UBound(imFilterShttCode)) = rst(0).Value
        ReDim Preserve imFilterShttCode(0 To UBound(imFilterShttCode) + 1) As Integer
        rst.MoveNext
    Loop
    Exit Sub
ErrHandle:
    gHandleError "AffErrorLog.txt", "Affiliate Management: mBuildSingleFilter"
    Resume Next

End Sub

Private Sub mPopUserOption()
    lbcUserOption(0).AddItem "Assigned"
    lbcUserOption(0).AddItem "All Affiliates"
    lbcUserOption(0).AddItem "Non-Affiliates"
    lbcUserOption(0).AddItem "All Stations"
    
    lbcUserOption(1).AddItem "Delinquent"
    lbcUserOption(1).AddItem "All Stations"
    
End Sub
Private Function mGetDueCountForStation(ilShttCode As Integer) As String
    Dim slCPTTSQLQuery As String
    Dim slAttSQLQuery As String
    Dim slEndWeek As String
    Dim ilMax As Integer
    
    
    mGetDueCountForStation = ""
    slEndWeek = DateAdd("d", -14, smWeek1)
    ilMax = 0
    slAttSQLQuery = "SELECT attCode FROM att"
    slAttSQLQuery = slAttSQLQuery + " WHERE ("
    slAttSQLQuery = slAttSQLQuery & " attShfCode = " & ilShttCode
    slAttSQLQuery = slAttSQLQuery & " And attOnAir >= '" & Format(smWeek54, sgSQLDateForm) & "'"
    slAttSQLQuery = slAttSQLQuery & ")"
    'Set rst_att = cnn.Execute(slAttSQLQuery)
    Set rst_att = gSQLSelectCall(slAttSQLQuery)
    If Not rst_att.EOF Then
        slCPTTSQLQuery = "Select Count(*) as Due from CPTT "
        slCPTTSQLQuery = slCPTTSQLQuery & " Where cpttPostingStatus <= 1 and "
        slCPTTSQLQuery = slCPTTSQLQuery & " cpttAtfCode = " & rst_att!attCode
        'slCPTTSQLQuery = slCPTTSQLQuery & " and cpttStartDate between '" & Format(smWeek54, sgSQLDateForm) & " ' and '" & Format(slEndWeek, sgSQLDateForm) & "'"
        slCPTTSQLQuery = slCPTTSQLQuery & " and cpttStartDate >= '" & Format(smWeek54, sgSQLDateForm) & " ' and cpttStartdate <= '" & Format(slEndWeek, sgSQLDateForm) & "'"
        slCPTTSQLQuery = slCPTTSQLQuery & " Order By Due Desc"
        'Set rst_Cptt = cnn.Execute(slCPTTSQLQuery)
        Set rst_Cptt = gSQLSelectCall(slCPTTSQLQuery)
        If Not rst_Cptt.EOF Then
            If rst_Cptt!Due > 0 Then
                ilMax = rst_Cptt!Due
            End If
        End If
    End If
    If ilMax > 0 Then
        mGetDueCountForStation = Trim(Str(ilMax))
    End If
End Function
Private Function mGetDueCountForAgreement(llAttCode As Long) As String
    Dim slSQLQuery As String
    Dim slEndWeek As String
    
    mGetDueCountForAgreement = ""
    slEndWeek = DateAdd("d", -14, smWeek1)
    slSQLQuery = "Select cpttShfCode, cpttVefCode, Count(*) as Due from CPTT "
    slSQLQuery = slSQLQuery & " Where cpttPostingStatus <= 1 and "
    slSQLQuery = slSQLQuery & " cpttAtfCode = " & llAttCode
    slSQLQuery = slSQLQuery & " and cpttStartDate >= '" & Format(smWeek54, sgSQLDateForm) & " ' and cpttStartdate <= '" & Format(slEndWeek, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " Group By cpttShfCode, cpttVefCode Order by cpttShfCode, Due Desc"
    'Set rst_Cptt = cnn.Execute(slSQLQuery)
    Set rst_Cptt = gSQLSelectCall(slSQLQuery)
    If Not rst_Cptt.EOF Then
        If rst_Cptt!Due > 0 Then
            mGetDueCountForAgreement = Trim$(Str(rst_Cptt!Due))
        End If
    End If
End Function

Private Sub mPopStationsGrid()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilShtt As Integer
    Dim llRet As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llCell As Long
    Dim llNext As Long
    Dim llCellColor As Long
    Dim ilIncludeStation As Integer
    Dim llFilterDefIndex As Long
    Dim llNotFilterDefIndex As Long
    Dim llYellowRow As Long
    Dim slMoniker As String
    Dim ilMnt As Integer
    Dim ilNoStations As Integer
    Dim slFilter As String
    Dim ilSingle As Integer
    Dim blMatch As Boolean
    Dim slWeek1 As String
    Dim slWeek54 As String
    Dim slSQLQuery As String
    Dim llFilter As Long
    
    On Error GoTo ErrHand:
    slWeek1 = Format$(smWeek1, sgSQLDateForm)
    slWeek54 = Format$(smWeek54, sgSQLDateForm)
    gPopStations
    ReDim igCommentShttCode(0 To 0) As Integer
    mBuildSingleFilter
    ilNoStations = 0
    grdStations.Rows = 2
    mClearGrid grdStations
    grdStations.Redraw = False
    grdStations.Row = 0
    For llCol = SCALLLETTERINDEX To SMCASTINDEX Step 1
        grdStations.Col = llCol
        grdStations.CellBackColor = LIGHTBLUE
    Next llCol
    grdStations.Col = SSELECTINDEX
    grdStations.CellBackColor = 16711808    'BROWN
    grdStations.CellForeColor = vbWhite
    grdStations.CellFontName = "Arial Narrow"
    
    '10/31/18: Replace how station grid is populated
    ReDim tmStationGridKey(0 To UBound(tgStationInfo)) As STATIONGRIDKEY
    ReDim smStationGridData(0 To (SSORTINDEX + 1) * UBound(tgStationInfo)) As String
    gGrid_FillWithRows grdStations
    llRow = grdStations.FixedRows
    For ilShtt = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        blMatch = False
        If UBound(imFilterShttCode) > LBound(imFilterShttCode) Then
            For ilSingle = 0 To UBound(imFilterShttCode) - 1 Step 1
                If tgStationInfo(ilShtt).iCode = imFilterShttCode(ilSingle) Then
                    blMatch = True
                    Exit For
                End If
            Next ilSingle
        Else
            blMatch = True
        End If
        If (tgStationInfo(ilShtt).iType = 0) And (blMatch) Then
            If (UBound(tmFilterLink) > 0) And (lacFilter.Caption = "On") Then
                For llCell = 0 To UBound(tmFilterLink) - 1 Step 1
                    ilIncludeStation = True
                    llFilterDefIndex = tmFilterLink(llCell).lFilterDefIndex
                    If llFilterDefIndex >= 0 Then
                        mTestFilter ilShtt, llFilterDefIndex, ilIncludeStation
                        If ilIncludeStation Then
                            llNext = tmFilterLink(llCell).lNextAnd
                            Do While llNext <> -1
                                llFilterDefIndex = tmAndFilterLink(llNext).lFilterDefIndex
                                If llFilterDefIndex >= 0 Then
                                    mTestFilter ilShtt, llFilterDefIndex, ilIncludeStation
                                    If Not ilIncludeStation Then
                                        Exit Do
                                    End If
                                End If
                                llNext = tmAndFilterLink(llNext).lNextAnd
                            Loop
                        End If
                    End If
                    'Test Not array
                    If ilIncludeStation Then
                        llNotFilterDefIndex = tmFilterLink(llCell).lNotFilterDefIndex
                        If llNotFilterDefIndex >= 0 Then
                            mTestFilter ilShtt, llNotFilterDefIndex, ilIncludeStation
                            If ilIncludeStation Then
                                llNext = tmFilterLink(llCell).lNextAnd
                                Do While llNext <> -1
                                    llNotFilterDefIndex = tmAndFilterLink(llNext).lNotFilterDefIndex
                                    If llNotFilterDefIndex >= 0 Then
                                        mTestFilter ilShtt, llNotFilterDefIndex, ilIncludeStation
                                        If Not ilIncludeStation Then
                                            Exit Do
                                        End If
                                    End If
                                    llNext = tmAndFilterLink(llNext).lNextAnd
                                Loop
                            End If
                        End If
                    End If
                    
                    If ilIncludeStation Then
                        Exit For
                    End If
                Next llCell
                
            Else
                ilIncludeStation = True
            End If
            If ilIncludeStation And (sgStationSearchCallSource <> "P") Then
                If lacUserOption.Caption = "Assigned" Then
                    If smUserType = "M" Then
                        If igUstCode <> tgStationInfo(ilShtt).iMktRepUstCode Then
                            ilIncludeStation = False
                        End If
                    ElseIf smUserType = "S" Then
                        If igUstCode <> tgStationInfo(ilShtt).iServRepUstCode Then
                            ilIncludeStation = False
                        End If
                    End If
                ElseIf lacUserOption.Caption = "All Affiliates" Then
                    If tgStationInfo(ilShtt).sAgreementExist <> "Y" Then
                        ilIncludeStation = False
                    End If
                ElseIf lacUserOption.Caption = "Non-Affiliates" Then
                    If tgStationInfo(ilShtt).sAgreementExist = "Y" Then
                        ilIncludeStation = False
                    End If
                ElseIf lacUserOption.Caption = "All Stations" Then
                End If
            End If
            '4/16/20: Moved Service Agreement filter here as it needs to be an AND operation. Taken out of mBuildFilter
            If ilIncludeStation And sgUsingServiceAgreement = "Y" Then
                For llFilter = 0 To UBound(tgFilterDef) - 1 Step 1
                    'Operator: 0=Contains; 1=Equal; 2=Not Equal; 3=Range; 4=Greater or Equal    '2=Greater Than; 3=Less Than; 4=Not Equal
                    'If ((tgFilterDef(ilFilter).iOperator = 0) And (tgFilterDef(ilFilter).iSelect <> SFCALLLETTERS)) Or (tgFilterDef(ilFilter).iOperator = 2) Then
                    If tgFilterDef(llFilter).iSelect = 51 Then
                        mTestFilter ilShtt, llFilter, ilIncludeStation
                    End If
                Next llFilter
            
            End If
            If ilIncludeStation Then
                ilNoStations = ilNoStations + 1
                igCommentShttCode(UBound(igCommentShttCode)) = tgStationInfo(ilShtt).iCode
                ReDim Preserve igCommentShttCode(0 To UBound(igCommentShttCode) + 1) As Integer
                
                'If llRow >= grdStations.Rows Then
                '    grdStations.AddItem ""
                'End If
                'grdStations.Row = llRow
                'grdStations.Col = SSELECTINDEX
                'grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
                'For llCol = SCALLLETTERINDEX To SPWINDEX Step 1
                '    grdStations.Col = llCol
                '    grdStations.CellBackColor = LIGHTYELLOW
                'Next llCol
                tmStationGridKey(llRow - grdStations.FixedRows).sKey = ""
                tmStationGridKey(llRow - grdStations.FixedRows).lRow = (SSORTINDEX + 1) * (llRow - grdStations.FixedRows)
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCALLLETTERINDEX) = Trim$(tgStationInfo(ilShtt).sCallLetters)
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SDUEINDEX) = gBinarySearchStationCount(tgStationInfo(ilShtt).iCode) 'mGetDueCountForStation(tgStationInfo(ilShtt).icode)
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SDMARANKINDEX) = ""
                llRet = gBinarySearchMkt(CLng(tgStationInfo(ilShtt).iMktCode))
                If llRet <> -1 Then
                    If tgMarketInfo(llRet).iRank <> 0 Then
                        smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SDMARANKINDEX) = tgMarketInfo(llRet).iRank
                    End If
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SDMAMARKETINDEX) = Trim$(tgStationInfo(ilShtt).sMarket)
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SAUDP12PLUSINDEX) = Format(tgStationInfo(ilShtt).lAudP12Plus, "##,###,###")
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SWATTSINDEX) = Format(tgStationInfo(ilShtt).lWatts, "##,###,###")
                llRet = gBinarySearchMSAMkt(CLng(tgStationInfo(ilShtt).iMSAMktCode))
                If llRet <> -1 Then
                    If tgMSAMarketInfo(llRet).iRank <> 0 Then
                        smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMSARANKINDEX) = tgMSAMarketInfo(llRet).iRank
                    End If
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMSAMARKETINDEX) = Trim$(tgMSAMarketInfo(llRet).sName)
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMSARANKINDEX) = ""
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMSAMARKETINDEX) = ""
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SSTATEINDEX) = Trim$(tgStationInfo(ilShtt).sPostalName)
                If tgStationInfo(ilShtt).iAckDaylight = 1 Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SZONEINDEX) = Trim$(tgStationInfo(ilShtt).sZone) & "*"
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SZONEINDEX) = Trim$(tgStationInfo(ilShtt).sZone)
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SZIPINDEX) = Trim$(tgStationInfo(ilShtt).sZip)
                ilRet = gBinarySearchFmt(CLng(tgStationInfo(ilShtt).iFormatCode))
                If ilRet <> -1 Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SFORMATINDEX) = Trim$(tgFormatInfo(ilRet).sName)
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SFORMATINDEX) = ""
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SOWNERINDEX) = ""
                llRet = mBinarySearchOwner(tgStationInfo(ilShtt).lOwnerCode)
                If llRet <> -1 Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SOWNERINDEX) = Trim$(tgOwnerInfo(llRet).sName)
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SOPERATORINDEX) = ""
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SOPERATORNAMEINDEX) = ""
                If tgStationInfo(ilShtt).lOperatorMntCode > 0 Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SOPERATORINDEX) = "*"
                    ilMnt = gBinarySearchMnt(tgStationInfo(ilShtt).lOperatorMntCode, tgOperatorInfo())
                    If ilMnt <> -1 Then
                        smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SOPERATORNAMEINDEX) = UCase$(Trim$(tgOperatorInfo(ilMnt).sName))
                    End If
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERINDEX) = ""
                If tgStationInfo(ilShtt).lMarketClusterGroupID <= 0 Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERINDEX) = ""
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERINDEX) = "*"
                    SQLQuery = "SELECT shttCallLetters FROM shtt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " shttClusterGroupID = " & tgStationInfo(ilShtt).lMarketClusterGroupID & ")"
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    Do While Not rst_Shtt.EOF
                        smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERNAMESINDEX) = smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERNAMESINDEX) & " " & Trim$(rst_Shtt!shttCallLetters)
                        rst_Shtt.MoveNext
                    Loop
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERNAMESINDEX) = Trim$(smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCLUSTERNAMESINDEX))
                End If
                
                If Trim$(tgStationInfo(ilShtt).sWebPW) <> "" Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SPASSWORDINDEX) = Trim$(tgStationInfo(ilShtt).sWebPW)
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SPWINDEX) = "*"
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SPASSWORDINDEX) = ""
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SPWINDEX) = ""
                End If
                
                If tgStationInfo(ilShtt).lMultiCastGroupID <= 0 Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMCASTINDEX) = ""
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMCASTINDEX) = "*"
                    SQLQuery = "SELECT shttCallLetters FROM shtt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " shttMultiCastGroupID = " & tgStationInfo(ilShtt).lMultiCastGroupID & ")"
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    Do While Not rst_Shtt.EOF
                        smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMCASTNAMESINDEX) = smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMCASTNAMESINDEX) & " " & Trim$(rst_Shtt!shttCallLetters)
                        rst_Shtt.MoveNext
                    Loop
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMCASTNAMESINDEX) = Trim$(smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SMCASTNAMESINDEX))
                End If
                'grdStations.Col = SWEBSITEINDEX
                If Trim$(tgStationInfo(ilShtt).sWebAddress) = "" Then
                '    grdStations.CellBackColor = GRAY    'vbWhite
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SWEBSITEINDEX) = ""
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SWEBADDRESSINDEX) = ""
                Else
                '    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SWEBSITEINDEX) = "W"
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SWEBADDRESSINDEX) = Trim$(tgStationInfo(ilShtt).sWebAddress)
                End If
                'grdStations.Col = SCALLLETTERINDEX
                'grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                
                'grdStations.Col = SAGRMNTINDEX
                'grdStations.CellAlignment = flexAlignCenterCenter
                If tgStationInfo(ilShtt).sAgreementExist = "Y" Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SAGRMNTINDEX) = "A"
                    If sgUsingServiceAgreement = "Y" Then
                        slSQLQuery = "Select Count(Case When attServiceAgreement = 'Y' Then 1 Else Null End) as ServiceCount, Count(Case When attServiceAgreement <> 'Y' Then 1 Else Null End ) As NonServiceCount from att "
                        slSQLQuery = slSQLQuery + " WHERE ("
                        If sgStationSearchCallSource = "P" Then
                            slSQLQuery = slSQLQuery & "(attOnAir <= '" & slWeek1 & "')"
                            slSQLQuery = slSQLQuery & " AND (attOffAir >= '" & slWeek54 & "')"
                            slSQLQuery = slSQLQuery & " AND (attDropDate >= '" & slWeek54 & "')"
                            slSQLQuery = slSQLQuery & " AND attServiceAgreement <> 'Y'"
                            slSQLQuery = slSQLQuery & " AND (attShfCode = " & tgStationInfo(ilShtt).iCode & "))"
                        Else
                            slSQLQuery = slSQLQuery & "(((attOnAir <= '" & slWeek1 & "')"
                            slSQLQuery = slSQLQuery & " AND (attOffAir >= '" & slWeek54 & "'))"
                            slSQLQuery = slSQLQuery & " OR ((attDropDate >= '" & slWeek54 & "')"
                            slSQLQuery = slSQLQuery & "AND (attOnAir <= '" & slWeek1 & "')))"
                            slSQLQuery = slSQLQuery & " AND (attShfCode = " & tgStationInfo(ilShtt).iCode & "))"
                        End If
                        
                        Set rst_att = gSQLSelectCall(slSQLQuery)
                        If Not rst_att.EOF Then
                            If rst_att!ServiceCount > 0 And rst_att!NonServiceCount > 0 Then
                                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SAGRMNTINDEX) = "A+SA"
                            ElseIf rst_att!ServiceCount > 0 Then
                                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SAGRMNTINDEX) = "SA"
                            End If
                        End If
                    End If
                '    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                '    grdStations.Col = SSELECTINDEX
                '    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
                '    grdStations.CellFontName = "Monotype Sorts"
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SSELECTINDEX) = "t"
                Else
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SAGRMNTINDEX) = ""
                '    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY    'vbWhite
                '    grdStations.Col = SSELECTINDEX
                '    grdStations.CellBackColor = GRAY    'vbWhite
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SSELECTINDEX) = ""
                End If
                
                'grdStations.Col = SCOMMENTINDEX
                'grdStations.CellAlignment = flexAlignCenterCenter
                If tgStationInfo(ilShtt).sCommentExist = "Y" Then
                    smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SCOMMENTINDEX) = "C"
                '    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
                Else
                '    grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY    'vbWhite
                End If
                            
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SSHTTCODEINDEX) = tgStationInfo(ilShtt).iCode
                slMoniker = ""
                If tgStationInfo(ilShtt).lMonikerMntCode > 0 Then
                    SQLQuery = "SELECT mntName FROM mnt"
                    SQLQuery = SQLQuery + " WHERE ("
                    SQLQuery = SQLQuery & " mntCode = " & tgStationInfo(ilShtt).lMonikerMntCode & ")"
                    Set rst_mnt = gSQLSelectCall(SQLQuery)
                    If Not rst_mnt.EOF Then
                        slMoniker = Trim$(rst_mnt!mntName)
                    End If
                End If
                smStationGridData((SSORTINDEX + 1) * (llRow - grdStations.FixedRows) + SFREQMONIKERINDEX) = Trim$(Trim$(tgStationInfo(ilShtt).sFrequency) & " " & slMoniker)
                llRow = llRow + 1
            End If
        End If
    Next ilShtt
    If llRow > grdStations.FixedRows Then
        lmStationMaxRow = llRow - 1
    Else
        lmStationMaxRow = llRow
    End If
    ReDim Preserve tmStationGridKey(0 To lmStationMaxRow) As STATIONGRIDKEY
    ReDim Preserve smStationGridData(0 To (SSORTINDEX + 1) * lmStationMaxRow) As String
    vbcStation.Min = 0
    vbcStation.Max = lmStationMaxRow - 1
    vbcStation.SmallChange = 1
    vbcStation.LargeChange = grdStations.Height / grdStations.RowHeight(0) - 1
    'grdStations.Rows = grdStations.Rows + ((cmcDone.Top - grdStations.Top) \ grdStations.RowHeight(1))
    grdStations.Rows = ((cmcDone.Top - grdStations.Top) \ grdStations.RowHeight(1))
    gGrid_IntegralHeight grdStations
    grdStations.Height = grdStations.Height + 30
    vbcStation.Height = grdStations.Height
    gGrid_FillWithRows grdStations
    For llYellowRow = llRow To grdStations.Rows - 1 Step 1
        grdStations.Row = llYellowRow
        For llCol = SSELECTINDEX To SCOMMENTINDEX Step 1
            grdStations.Col = llCol
            grdStations.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llYellowRow
    imLastStationSort = -1
    imLastStationColSorted = -1
    mStationSortCol SCALLLETTERINDEX
    grdStations.Row = 0
    grdStations.Col = SSHTTCODEINDEX
    lmRowSelected = -1
    If sgFilterName = "" Then
        If igFilterChgd Then
            slFilter = "Filter: Custom"
        Else
            slFilter = "Filter: None"
        End If
    ElseIf igFilterChgd Then
        slFilter = "Filter: " & sgFilterName & " modified"
    Else
        slFilter = "Filter: " & sgFilterName
    End If
    If sgStationSearchCallSource = "P" Then
        frmStationSearch.Caption = "Post-Buy Planning- " & Trim$(sgUserName) & ", " & ilNoStations & " Stations" & " " & slFilter
    Else
        frmStationSearch.Caption = "Management- " & Trim$(sgUserName) & ", " & ilNoStations & " Stations" & " " & slFilter
    End If
        
    mDisplayStationInfo

    grdStations.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmStationSearch-mPopStationsGrid"
End Sub
Private Sub mStationSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim llKey As Long
    
    On Error GoTo Error:
    For llKey = LBound(tmStationGridKey) To UBound(tmStationGridKey) - 1 Step 1
        llRow = tmStationGridKey(llKey).lRow
        slStr = Trim$(smStationGridData(llRow + SCALLLETTERINDEX))
        If slStr <> "" Then
            If (ilCol = SDMARANKINDEX) Or (ilCol = SMSARANKINDEX) Then
                slSort = UCase$(Trim$(smStationGridData(llRow + ilCol)))
                If slSort = "" Then
                    slSort = "999"
                End If
                Do While Len(slSort) < 3
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = SDUEINDEX Then
                slSort = UCase$(Trim$(smStationGridData(llRow + ilCol)))
                If slSort = "" Then
                    If (ilCol = imLastStationColSorted) Then
                        If imLastStationSort = flexSortStringNoCaseAscending Then
                            slSort = "000"
                        Else
                            slSort = "999"
                        End If
                    Else
                        If (ilCol <> imLastStationColSorted) Or (imLastStationSort = -1) Then
                            slSort = "000"
                        Else
                            slSort = "999"
                        End If
                    End If
                End If
                
                Do While Len(slSort) < 3
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(smStationGridData(llRow + ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slStr = smStationGridData(llRow + SSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastStationColSorted) Or ((ilCol = imLastStationColSorted) And (imLastStationSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 5
                    slRow = "0" & slRow
                Loop
                smStationGridData(llRow + SSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 5
                    slRow = "0" & slRow
                Loop
                smStationGridData(llRow + SSORTINDEX) = slSort & slStr & "|" & slRow
            End If
            tmStationGridKey(llKey).sKey = smStationGridData(llRow + SSORTINDEX)
        End If
    Next llKey
    If (ilCol = SDUEINDEX) And ((ilCol <> imLastStationColSorted) Or (imLastStationSort = -1)) Then
        imLastStationSort = flexSortStringNoCaseAscending
        imLastStationColSorted = ilCol
    End If
    'If ilCol = imLastStationColSorted Then
    '    imLastStationColSorted = SSORTINDEX
    'Else
    '    imLastStationColSorted = -1
    'End If
    'gGrid_SortByCol grdStations, SCALLLETTERINDEX, SSORTINDEX, imLastStationColSorted, imLastStationSort
    If UBound(tmStationGridKey) - 1 > 0 Then
        If imLastStationColSorted = ilCol Then
            If imLastStationSort = flexSortStringNoCaseAscending Then
                ArraySortTyp fnAV(tmStationGridKey(), 0), UBound(tmStationGridKey), 1, LenB(tmStationGridKey(0)), 0, LenB(tmStationGridKey(0).sKey), 0
                imLastStationSort = flexSortStringNoCaseDescending
            Else
                ArraySortTyp fnAV(tmStationGridKey(), 0), UBound(tmStationGridKey), 0, LenB(tmStationGridKey(0)), 0, LenB(tmStationGridKey(0).sKey), 0
                imLastStationSort = flexSortStringNoCaseAscending
            End If
        Else
            ArraySortTyp fnAV(tmStationGridKey(), 0), UBound(tmStationGridKey), 0, LenB(tmStationGridKey(0)), 0, LenB(tmStationGridKey(0).sKey), 0
            imLastStationSort = flexSortStringNoCaseAscending
        End If
    End If
    
    imLastStationColSorted = ilCol
    mSetCommands
    Exit Sub
Error:
    Resume Next
End Sub

Private Sub mDisplayStationInfo()
    Dim llRow As Long
    Dim llStationRow As Long
    Dim llKey As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim llSvCol As Long
    Dim llSvRow As Long
    
    On Error GoTo Error:
    llTopRow = vbcStation.Value
    llSvCol = grdStations.Col
    gGrid_Clear grdStations, True
    llRow = grdStations.FixedRows
    For llKey = llTopRow To UBound(tmStationGridKey) - 1 Step 1
        llStationRow = tmStationGridKey(llKey).lRow
        grdStations.Row = llRow
        grdStations.Col = SSELECTINDEX
        grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
        For llCol = SCALLLETTERINDEX To SPWINDEX Step 1
            grdStations.Col = llCol
            grdStations.CellBackColor = LIGHTYELLOW
        Next llCol
        grdStations.Col = SWEBSITEINDEX
        If Trim$(smStationGridData(llStationRow + SWEBADDRESSINDEX)) = "" Then
            grdStations.CellBackColor = GRAY    'vbWhite
        Else
            grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
        End If
        grdStations.Col = SCALLLETTERINDEX
        grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
        
        grdStations.Col = SAGRMNTINDEX
        grdStations.CellAlignment = flexAlignCenterCenter
        
        If smStationGridData(llStationRow + SAGRMNTINDEX) = "A" Or smStationGridData(llStationRow + SAGRMNTINDEX) = "SA" Or smStationGridData(llStationRow + SAGRMNTINDEX) = "A+SA" Then
            grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
            grdStations.Col = SSELECTINDEX
            grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'LIGHTBLUE
            grdStations.CellFontName = "Monotype Sorts"
        Else
            grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY    'vbWhite
            grdStations.Col = SSELECTINDEX
            grdStations.CellBackColor = GRAY    'vbWhite
            grdStations.CellFontName = "Arial Narrow"
        End If
        
        grdStations.Col = SCOMMENTINDEX
        grdStations.CellAlignment = flexAlignCenterCenter
        If smStationGridData(llStationRow + SCOMMENTINDEX) = "C" Then
            grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN
        Else
            grdStations.CellBackColor = LIGHTGREENCOLOR    'LIGHTGREEN   'GRAY    'vbWhite
        End If
        
        
        For llCol = SSELECTINDEX To SSORTINDEX Step 1
            grdStations.TextMatrix(llRow, llCol) = smStationGridData(llStationRow + llCol)
        Next llCol
        llRow = llRow + 1
        If llRow >= grdStations.Rows Then
            Exit For
        End If
        If Not grdStations.RowIsVisible(llRow) Then
            Exit For
        End If
    Next llKey
    If llRow > grdStations.FixedRows Then
        llSvRow = grdStations.Row
        For llRow = grdStations.FixedRows + 1 To grdStations.Rows - 1 Step 1
            If grdStations.TextMatrix(llRow, SCALLLETTERINDEX) = "" Then
                grdStations.Row = llRow
                For llCol = SSELECTINDEX To SCOMMENTINDEX Step 1
                    grdStations.Col = llCol
                    grdStations.CellBackColor = LIGHTYELLOW
                Next llCol
            End If
        Next llRow
        grdStations.Row = llSvRow
    End If
    grdStations.TopRow = grdStations.FixedRows  'llTopRow + grdStations.FixedRows
    grdStations.Col = llSvCol
    Exit Sub
Error:
    Resume Next
End Sub

Private Sub vbcStation_Change()
    mFillStationGrid
End Sub

Private Sub vbcStation_Scroll()
    mFillStationGrid
End Sub

Private Sub mFillStationGrid()
    mMousePointer vbHourglass
    grdStations.Redraw = False
    mDisplayStationInfo
    grdStations.Redraw = True
    mMousePointer vbDefault
End Sub

Private Sub mSetStationGridData(llRow As Long, slChar As String)
    Dim llShttCode As Long
    Dim llKey As Long
    Dim llStationRow As Long
    
    If grdStations.TextMatrix(llRow, SSHTTCODEINDEX) = "" Then
        Exit Sub
    End If
    llShttCode = grdStations.TextMatrix(llRow, SSHTTCODEINDEX)
    For llKey = LBound(tmStationGridKey) To UBound(tmStationGridKey) - 1 Step 1
        llStationRow = tmStationGridKey(llKey).lRow
        If llShttCode = Val(smStationGridData(llStationRow + SSHTTCODEINDEX)) Then
            smStationGridData(llStationRow + SSELECTINDEX) = slChar
            Exit Sub
        End If
    Next llKey
End Sub
Private Sub mFindAndDisplay(llRow As Long)
    Dim llShttCode As Long
    Dim llKey As Long
    Dim llStationRow As Long
    
    If grdStations.TextMatrix(llRow, SSHTTCODEINDEX) = "" Then
        Exit Sub
    End If
    llShttCode = grdStations.TextMatrix(llRow, SSHTTCODEINDEX)
    For llKey = LBound(tmStationGridKey) To UBound(tmStationGridKey) - 1 Step 1
        llStationRow = tmStationGridKey(llKey).lRow
        If llShttCode = Val(smStationGridData(llStationRow + SSHTTCODEINDEX)) Then
            vbcStation.Value = llKey
            mFillStationGrid
            Exit Sub
        End If
    Next llKey
End Sub

Private Sub mCompliantTypeChg()
    Dim blStationVisible As Boolean
    Dim llStationTopRow As Long
    Dim blAgreementVisible As Boolean
    Dim llAgreementTopRow As Long
    Dim blPostedInfoVisible As Boolean
    Dim llPostedInfoTopRow As Long
    Dim blSpotInfoVisible As Boolean
    Dim llSpotInfoTopRow As Long
    
    Dim slVehicleName As String
    Dim llAttCode As Long
    Dim llCpttCode As Long
    Dim llCptt As Long
    Dim llCpttRow As Long
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim llAgreementRow As Long
    
    mMousePointer vbHourglass
    blStationVisible = grdStations.Visible
    llStationTopRow = grdStations.TopRow
    blAgreementVisible = grdAgreementInfo.Visible
    llAgreementTopRow = grdAgreementInfo.TopRow
    blPostedInfoVisible = grdPostedInfo.Visible
    llPostedInfoTopRow = grdPostedInfo.TopRow
    blSpotInfoVisible = grdSpotInfo.Visible
    llSpotInfoTopRow = grdSpotInfo.TopRow
    imShttCode = Val(grdStations.TextMatrix(llStationTopRow, SSHTTCODEINDEX))
    slVehicleName = Trim$(grdAgreementInfo.TextMatrix(llAgreementTopRow, AVEHICLEINDEX))
    llAttCode = Val(Trim$(grdAgreementInfo.TextMatrix(llAgreementTopRow, AATTCODEINDEX)))
    llCpttCode = Val(Trim$(grdPostedInfo.TextMatrix(llPostedInfoTopRow, PCPTTINDEX)))
    If grdPostedInfo.TextMatrix(llPostedInfoTopRow, PSELECTINDEX) <> "" Then
        grdPostedInfo.TextMatrix(grdPostedInfo.TopRow, PSELECTINDEX) = "s"
    End If
    
    udcContactGrid.StationCode = imShttCode
    udcContactGrid.Action 3 'Populate
    udcCommentGrid.StationCode = imShttCode 'retain station so that correct comments show when Follow-up changed back to All Comments
    If rbcComments(0).Value Then
        udcCommentGrid.Action 3 'Populate
    End If
    If (udcCommentGrid.Visible) And (udcContactGrid.Visible = False) Then
        udcContactGrid.Visible = True
    End If
    If blAgreementVisible Then
        grdStations.TextMatrix(0, SSELECTINDEX) = "Home"    '"Stns"
        ilRet = mPopAgreementInfoGrid()
        If grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) <> "" Then
            grdStations.TextMatrix(grdStations.Row, SSELECTINDEX) = "s"
            mSetStationGridData grdStations.Row, "s"
        End If
        bmStationScrollAllowed = False
        lmStationTopRow = grdStations.TopRow
        If blPostedInfoVisible Then
            llAgreementRow = -1
            For ilVef = grdAgreementInfo.FixedRows To grdAgreementInfo.Rows - 1 Step 1
                If llAttCode = Val(grdAgreementInfo.TextMatrix(ilVef, AATTCODEINDEX)) Then
                    llAgreementRow = ilVef
                    Exit For
                End If
            Next ilVef
            If llAgreementRow <= 0 Then
                For ilVef = grdAgreementInfo.FixedRows To grdAgreementInfo.Rows - 1 Step 1
                    If StrComp(slVehicleName, grdAgreementInfo.TextMatrix(ilVef, AVEHICLEINDEX), vbTextCompare) = 0 Then
                        llAgreementRow = ilVef
                        Exit For
                    End If
                Next ilVef
            End If
            grdAgreementInfo.TextMatrix(llAgreementRow, ASELECTINDEX) = "s"
            bmAgreementScrollAllowed = True
            grdAgreementInfo.TopRow = llAgreementRow
            grdAgreementInfo.Row = llAgreementRow
            grdAgreementInfo.Col = ASELECTINDEX
            bmAgreementScrollAllowed = False
            lmAgreementTopRow = grdAgreementInfo.TopRow
            ilRet = mPopPostedInfoGrid()
            If grdSpotInfo.Visible Then
                llCpttRow = -1
                For llCptt = grdPostedInfo.FixedRows To grdPostedInfo.Rows - 1 Step 1
                    If llCpttCode = Val(grdPostedInfo.TextMatrix(llCptt, PCPTTINDEX)) Then
                        llCpttRow = llCptt
                        Exit For
                    End If
                Next llCptt
                If llCpttRow >= grdPostedInfo.FixedRows Then
                    grdPostedInfo.TextMatrix(llCpttRow, PSELECTINDEX) = "s"
                    bmPostedScrollAllowed = True
                    grdPostedInfo.TopRow = llCpttRow
                    grdPostedInfo.Row = llCpttRow
                    grdPostedInfo.Col = PSELECTINDEX
                    bmPostedScrollAllowed = False
                    lmPostedTopRow = grdPostedInfo.TopRow
                    ilRet = mPopSpotInfoGrid()
                End If
            End If
        End If
    End If
    grdStations.TopRow = llStationTopRow
    grdStations.Row = llStationTopRow
    If blAgreementVisible Then
        grdAgreementInfo.TopRow = llAgreementTopRow
        grdAgreementInfo.Row = llAgreementTopRow
        If grdPostedInfo.Visible Then
            grdPostedInfo.TopRow = llPostedInfoTopRow
            grdPostedInfo.Row = llPostedInfoTopRow
            If grdSpotInfo.Visible Then
                grdSpotInfo.TopRow = llSpotInfoTopRow
                grdSpotInfo.Row = llSpotInfoTopRow
            End If
        End If
    End If
    mSetGridPosition
    mMousePointer vbDefault
End Sub
