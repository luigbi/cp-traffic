VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MsComm32.ocx"
Begin VB.Form EngrSchdDef 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrSchdDef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.CommandButton cmcReload 
      Caption         =   "R&eload"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6975
      TabIndex        =   40
      Top             =   6855
      Width           =   1200
   End
   Begin VB.CommandButton cmcTask 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4920
      TabIndex        =   49
      Top             =   3630
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.PictureBox pbcHighlight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -60
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   6990
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11325
      Top             =   6255
   End
   Begin VB.ListBox lbcETE_Program 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":030A
      Left            =   405
      List            =   "EngrSchdDef.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3780
      Visible         =   0   'False
      Width           =   1410
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConflicts 
      Height          =   1035
      Left            =   225
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1826
      _Version        =   393216
      BackColor       =   12648447
      Rows            =   4
      Cols            =   12
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   12648447
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin V10EngineeringDev.CSI_TimeLength ltcEvent 
      Height          =   195
      Left            =   5370
      TabIndex        =   11
      Top             =   2070
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
      Text            =   "00:00.0"
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_UseHours    =   0   'False
      CSI_UseTenths   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin V10EngineeringDev.CSI_Calendar cccDate 
      Height          =   285
      Left            =   8685
      TabIndex        =   2
      Top             =   105
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   0   'False
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin VB.ListBox lbcCommercialSort 
      Height          =   255
      ItemData        =   "EngrSchdDef.frx":030E
      Left            =   4185
      List            =   "EngrSchdDef.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   46
      Top             =   180
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.ListBox lbcKey 
      BackColor       =   &H00C0FFFF&
      Height          =   2400
      ItemData        =   "EngrSchdDef.frx":0312
      Left            =   180
      List            =   "EngrSchdDef.frx":032E
      TabIndex        =   44
      Top             =   630
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Timer tmcCheck 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11520
      Top             =   6405
   End
   Begin VB.CommandButton cmcConflict 
      Caption         =   "Con&flict Check"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8325
      TabIndex        =   42
      Top             =   6855
      Width           =   1440
   End
   Begin MSCommLib.MSComm spcItemID 
      Left            =   11445
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmcMerge 
      Caption         =   "Commercial &Merge"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1935
      TabIndex        =   37
      Top             =   6855
      Width           =   1545
   End
   Begin VB.CommandButton cmcLoad 
      Caption         =   "&Load-Automation"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3615
      TabIndex        =   38
      Top             =   6855
      Width           =   1545
   End
   Begin VB.CommandButton cmcItemIDChk 
      Caption         =   "&Item ID Check"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10425
      TabIndex        =   41
      Top             =   7035
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmcTest 
      Caption         =   "&Test-Automation"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5295
      TabIndex        =   39
      Top             =   6855
      Width           =   1545
   End
   Begin VB.CommandButton cmcFilter 
      Caption         =   "&Filter"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8130
      TabIndex        =   36
      Top             =   6390
      Width           =   1200
   End
   Begin VB.ListBox lbcBDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0401
      Left            =   1845
      List            =   "EngrSchdDef.frx":0403
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCTE_1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0405
      Left            =   10080
      List            =   "EngrSchdDef.frx":0407
      Sorted          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4095
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcEDefine 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   2610
      ScaleHeight     =   165
      ScaleWidth      =   1035
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2685
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6555
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2745
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lbcFNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0409
      Left            =   10140
      List            =   "EngrSchdDef.frx":040B
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcMTE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":040D
      Left            =   8505
      List            =   "EngrSchdDef.frx":040F
      Sorted          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcANE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0411
      Left            =   3135
      List            =   "EngrSchdDef.frx":0413
      Sorted          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5190
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcSCE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0415
      Left            =   4695
      List            =   "EngrSchdDef.frx":0417
      Sorted          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5205
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcNNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0419
      Left            =   9030
      List            =   "EngrSchdDef.frx":041B
      Sorted          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCTE_2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":041D
      Left            =   8715
      List            =   "EngrSchdDef.frx":041F
      Sorted          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3225
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcASE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0421
      Left            =   2850
      List            =   "EngrSchdDef.frx":0423
      Sorted          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcRNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0425
      Left            =   1275
      List            =   "EngrSchdDef.frx":0427
      Sorted          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcETE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0429
      Left            =   405
      List            =   "EngrSchdDef.frx":042B
      Sorted          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2910
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcTTE_E 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":042D
      Left            =   6645
      List            =   "EngrSchdDef.frx":042F
      Sorted          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5115
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcTTE_S 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0431
      Left            =   6840
      List            =   "EngrSchdDef.frx":0433
      Sorted          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4155
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCCE_A 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0435
      Left            =   4950
      List            =   "EngrSchdDef.frx":0437
      Sorted          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2985
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCCE_B 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdDef.frx":0439
      Left            =   2850
      List            =   "EngrSchdDef.frx":043B
      Sorted          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3180
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   11340
      Top             =   6330
   End
   Begin VB.CommandButton cmcEDropDown 
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
      Left            =   3750
      Picture         =   "EngrSchdDef.frx":043D
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcEDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2805
      TabIndex        =   12
      Top             =   4530
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox edcEvent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6960
      TabIndex        =   10
      Top             =   3795
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmcReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   6690
      TabIndex        =   35
      Top             =   6390
      Width           =   1200
   End
   Begin VB.PictureBox pbcETab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   31
      Top             =   6675
      Width           =   60
   End
   Begin VB.PictureBox pbcESTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   540
      Width           =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   270
      Width           =   60
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   -15
      Picture         =   "EngrSchdDef.frx":0537
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save-Stay"
      Height          =   375
      Left            =   5265
      TabIndex        =   34
      Top             =   6390
      Width           =   1200
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11235
      Top             =   6825
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7290
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3825
      TabIndex        =   33
      Top             =   6390
      Width           =   1200
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2370
      TabIndex        =   32
      Top             =   6390
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLibEvents 
      Height          =   5655
      Left            =   165
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   630
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   4
      Cols            =   48
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      _Band(0).Cols   =   48
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.ListBox lbcSort 
      Height          =   255
      Left            =   10770
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6210
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmcShowEvents 
      Caption         =   "Show &Events"
      Height          =   300
      Left            =   10065
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Picture         =   "EngrSchdDef.frx":0841
      Top             =   315
      Width           =   480
   End
   Begin VB.Label lacHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   165
      TabIndex        =   43
      Top             =   6300
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1125
      Picture         =   "EngrSchdDef.frx":0B4B
      Top             =   6630
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Schedule"
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   2325
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   270
      Picture         =   "EngrSchdDef.frx":0E55
      Top             =   6630
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   10845
      Picture         =   "EngrSchdDef.frx":171F
      Top             =   6555
      Width           =   480
   End
   Begin VB.Label lacDate 
      Caption         =   "Schedule date"
      Height          =   270
      Left            =   7305
      TabIndex        =   1
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "EngrSchdDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrSchdDef - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private hmSEE As Integer
Private hmCME As Integer
Private hmSOE As Integer
Private hmCTE As Integer


Private imFieldChgd As Integer
Private bmIntegralSet As Boolean
Private smState As String
Private smYN As String
Private imInChg As Integer
Private imBSMode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer
Private imDoubleClickName As Integer
Private smDays() As String
Private smHours() As String
Private imOverlapCase As Integer    '1=Replace; 2=Terminate; 3=Change Start Date; 4=Split
Private lmCurrentDHE As Long        'DHE code that needs to be altered because of date overlap
Private smOverlapMsg As String
Private lmCharacterWidth As Long
Private imMaxColChars As Integer
Private smAirDate As String
Private hmMsg As Integer
Private hmMerge As Integer
Private bmMerged As Boolean
Private hmExport As Integer
Private smExportStr As String
Private smMsgFileName As String
Private smExportFileName As String
Private imSpotETECode As Integer
Private smSpotEventTypeName As String
Private lmCheckSHECode As Long
Private imAnyEvtChgs As Integer
Private imMaxCols As Integer
Private lmGridLibEventsHeight As Long
Private imMergeError As Integer '0=Merge not run from this screen; 1=Merge run but no errors; 2=Merge run but errors exist
Private bmInBranch As Boolean
Private bmInInsert As Boolean
Private bmInSave As Boolean
Private imEvtRet As Integer

Private bmPrinting As Boolean

Private imStartChgModeCompleted As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmFilterValues() As FILTERVALUES    'Same as tgFilterValues except the equals and Not Equals placed first
Private lmFilterStartTime As Long
Private lmFilterEndTime As Long
Private imFilterBus() As Integer
Private imFilterAudio() As Integer

Private lmChgStatusSEECode() As Long

Private tmSHE As SHE
Private smCurrDEEStamp
Private tmCurrSEE() As SEE
Private tmCTE As CTE

Private tmCCurrSEE() As SEE
Private tmPSHE As SHE
Private smPCurrSEEStamp As String
Private tmPCurrSEE() As SEE
Private tmPCTE As CTE
Private tmNSHE As SHE
Private smNCurrSEEStamp As String
Private tmNCurrSEE() As SEE
Private tmNCTE As CTE
Private tmSeeBracket() As SEEBRACKET
Private lmChgSEE() As Long

Private tmDee As DEE
Private tmDHE As DHE

Private tmSchdSort() As SCHDSORT

Private lmDeleteCodes() As Long


Private tmConflictList() As CONFLICTLIST
Private lmConflictRow As Long

Private tmConflictTest() As CONFLICTTEST

Private smCurrBSEStamp As String
Private tmCurrBSE() As BSE
'Private smBusGroups() As String
'Private smBuses() As String
Private smCurrDBEStamp As String
Private tmCurrDBE() As DBE
Private smT1Comment() As String
Private tmCurr1CTE_Name() As DEECTE
Private smT2Comment() As String
Private tmCurr2CTE_Name() As DEECTE
'Private smEBuses() As String

Private tmARE As ARE
Private smCurrLibDBEStamp As String
Private tmCurrLibDBE() As DBE
Private smCurrLibEBEStamp As String
Private tmCurrLibEBE() As EBE

Private tmDBE As DBE
Private tmEBE As EBE


Private fmUsedWidth As Single
Private fmUnusedWidth As Single
Private imUnusedCount As Integer


'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private imLastColSorted As Integer
Private imLastSort As Integer
Private lmEEnableRow As Long         'Current or last row focus was on
Private lmEEnableCol As Long         'Current or last column focus was on
Private imInsertState As Integer     'True is inserting a row
Private lmInsertRow As Long          'Row number that is being inserted
Private imDefaultProgIndex As Integer

Private lmHighlightRow As Long

Private Type CARTUNLOADTEST
    lGridRow As Long
    lEventID As Long
    lStartTime As Long
    lEndTime As Long
    AudioItemID As String * 30
End Type

Private tmCartUnloadTest() As CARTUNLOADTEST


Const HIGHLIGHTINDEX = 0
Const EVENTTYPEINDEX = 1
Const EVENTIDINDEX = 2
Const BUSNAMEINDEX = 3
Const BUSCTRLINDEX = 4
Const TIMEINDEX = 5
Const STARTTYPEINDEX = 6
Const FIXEDINDEX = 7
Const ENDTYPEINDEX = 8
Const DURATIONINDEX = 9
Const MATERIALINDEX = 10
Const AUDIONAMEINDEX = 11
Const AUDIOITEMIDINDEX = 12
Const AUDIOISCIINDEX = 13
Const AUDIOCTRLINDEX = 14
Const BACKUPNAMEINDEX = 15  '16
Const BACKUPCTRLINDEX = 16  '17
Const PROTNAMEINDEX = 17    '13
Const PROTITEMIDINDEX = 18  '14
Const PROTISCIINDEX = 19  '14
Const PROTCTRLINDEX = 20    '15
Const RELAY1INDEX = 21
Const RELAY2INDEX = 22
Const FOLLOWINDEX = 23
Const SILENCETIMEINDEX = 24
Const SILENCE1INDEX = 25
Const SILENCE2INDEX = 26
Const SILENCE3INDEX = 27
Const SILENCE4INDEX = 28
Const NETCUE1INDEX = 29
Const NETCUE2INDEX = 30
Const TITLE1INDEX = 31
Const TITLE2INDEX = 32
Const ABCFORMATINDEX = 33
Const ABCPGMCODEINDEX = 34
Const ABCXDSMODEINDEX = 35
Const ABCRECORDITEMINDEX = 36
Const PCODEINDEX = 37
Const SORTTIMEINDEX = 38
Const LIBNAMEINDEX = 39
Const TMCURRSEEINDEX = 40
Const SPOTAVAILTIMEINDEX = 41
Const AVAILDURATIONINDEX = 42
Const ERRORCONFLICTINDEX = 43
Const CHGSTATUSINDEX = 44
Const EVTCONFLICTINDEX = 45
Const DEECODEINDEX = 46
Const ERRORFIELDSORTINDEX = 47


Const CONFLICTNAMEINDEX = 0
Const CONFLICTSUBNAMEINDEX = 1
Const CONFLICTSTARTDATEINDEX = 2
Const CONFLICTENDDATEINDEX = 3
Const CONFLICTBUSESINDEX = 4
Const CONFLICTOFFSETINDEX = 5
Const CONFLICTHOURSINDEX = 6
Const CONFLICTDURATIONINDEX = 7
Const CONFLICTDAYSINDEX = 8
Const CONFLICTAUDIOINDEX = 9
Const CONFLICTBACKUPINDEX = 10
Const CONFLICTPROTINDEX = 11





Private Sub cccDate_CalendarChanged()
    Dim slStr As String
    Dim llStartDate As Long
    Dim llLastDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim slCCDate As Date
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim llNowDate As Long
    
    slNowDate = Format(gNow(), "mm/dd/yyyy")
    llNowDate = gDateValue(slNowDate)
    slStr = cccDate.GetCalendarDate
    llStartDate = gDateValue(cccDate.GetFirstDate)
    llLastDate = gDateValue(cccDate.GetLastDate)
    For llDate = llStartDate To llLastDate Step 1
        slDate = Format(llDate, "mm/dd/yyyy")
        slCCDate = slDate
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(slDate, "EngrSchedule-Get Schedule by Date", tmSHE)
        If Not ilRet Then
            cccDate.SetDateProperties slCCDate, CLng(vbBlack), True
        Else
            If llNowDate < llDate Then
                cccDate.SetDateProperties slCCDate, CLng(vbRed), True
            Else
                cccDate.SetDateProperties slCCDate, CLng(vbBlue), True
            End If
        End If
    Next llDate
End Sub

Private Sub cccDate_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub


Private Sub cmcCancel_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub



Private Sub cmcConflict_Click()
    Dim ilLibRet As Integer
    '9/13/11: ilEvtRet Place at Module level
    'Dim ilEvtRet As Integer
    Dim llLoop As Long
    Dim slCategory As String
    Dim ilETE As Integer
    Dim llAvailLength As Long
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    Dim ilRet As Integer
    Dim ilRemakeDay As Integer
    Dim llOldSHECode As Long
    Dim ilEvtRet As Integer
    
    If bmInSave Then
        Exit Sub
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    grdLibEvents.Redraw = False
    ilRemakeDay = False
    If (imFieldChgd) Or (UBound(tgFilterValues) > LBound(tgFilterValues)) Then
        ilRet = mSaveAndClearFilter(True, True)
        If Not ilRet Then
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            Exit Sub
        End If
    Else
        ilRemakeDay = True
        'mMoveSEECtrlsToRec
        'Remove filter
        'ReDim tmFilterValues(0 To 0) As FILTERVALUES
        'cbcApplyFilter.Value = vbUnchecked
    End If
    grdLibEvents.Redraw = False
    cmcTask.Caption = "Running Conflict Checker...."
    cmcTask.Visible = True
    ReDim tmConflictList(1 To 1) As CONFLICTLIST
    tmConflictList(UBound(tmConflictList)).iNextIndex = -1
    grdLibEvents.Visible = False
    If ilRemakeDay Then
        'mSetAvailTime
        'mMoveSEERecToCtrls
    End If
    grdLibEvents.Redraw = False
    imLastColSorted = -1
    mSortCol TIMEINDEX
    mInitConflictTest
    'ilLibRet = mCheckLibConflicts()
    'ilEvtRet = mSvCheckEventConflicts()
    ilLibRet = False
    ilEvtRet = mCheckEventConflicts()
    If ilEvtRet <> imEvtRet Then
        imFieldChgd = True
        imEvtRet = ilEvtRet
    End If
    lmConflictRow = -1
    mSortErrorsToTop
    cmcTask.Visible = False
    grdLibEvents.Redraw = True
    grdLibEvents.Visible = True
    imStartChgModeCompleted = True
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault

End Sub

Private Sub cmcConflict_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcEDropDown_Click()
    Select Case grdLibEvents.Col
        Case BUSNAMEINDEX
            lbcBDE.Visible = Not lbcBDE.Visible
        Case BUSCTRLINDEX
            lbcCCE_B.Visible = Not lbcCCE_B.Visible
        Case EVENTTYPEINDEX
            '2/9/12: Allow all events
            ''lbcETE.Visible = Not lbcETE.Visible
            'lbcETE_Program.Visible = Not lbcETE_Program.Visible
            lbcETE.Visible = Not lbcETE.Visible
        Case STARTTYPEINDEX
            lbcTTE_S.Visible = Not lbcTTE_S.Visible
        Case ENDTYPEINDEX
            lbcTTE_E.Visible = Not lbcTTE_E.Visible
        Case MATERIALINDEX
            lbcMTE.Visible = Not lbcMTE.Visible
        Case AUDIONAMEINDEX
            lbcASE.Visible = Not lbcASE.Visible
        Case AUDIOCTRLINDEX
            lbcCCE_A.Visible = Not lbcCCE_A.Visible
        Case BACKUPNAMEINDEX
            lbcANE.Visible = Not lbcANE.Visible
        Case BACKUPCTRLINDEX
            lbcCCE_A.Visible = Not lbcCCE_A.Visible
        Case PROTNAMEINDEX
            lbcANE.Visible = Not lbcANE.Visible
        Case PROTCTRLINDEX
            lbcCCE_A.Visible = Not lbcCCE_A.Visible
        Case RELAY1INDEX, RELAY2INDEX
            lbcRNE.Visible = Not lbcRNE.Visible
        Case FOLLOWINDEX
            lbcFNE.Visible = Not lbcFNE.Visible
        Case SILENCE1INDEX To SILENCE4INDEX
            lbcSCE.Visible = Not lbcSCE.Visible
        Case NETCUE1INDEX, NETCUE2INDEX
            lbcNNE.Visible = Not lbcNNE.Visible
        Case TITLE1INDEX
            lbcCTE_1.Visible = Not lbcCTE_1.Visible
        Case TITLE2INDEX
            lbcCTE_2.Visible = Not lbcCTE_2.Visible
    End Select
End Sub





Private Sub mSortCol(ilCol As Integer)
    Dim llEndRow As Long
    Dim llRow As Long
    Dim slStr As String
    Dim slBus As String
    Dim slTime As String
    Dim slType As String
    Dim ilLen As Integer
    Dim ilETE As Integer
    Dim slAvailTime As String
    Dim slEventCategory As String
    
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            If (ilCol = TIMEINDEX) Then
                slEventCategory = ""
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                        slEventCategory = tgCurrETE(ilETE).sCategory
                        If slEventCategory = "A" Then
                            slType = "C"    '"B"
                        ElseIf slEventCategory = "S" Then
                            slType = "B"    '"C"
                        Else
                            slType = "A"
                        End If
                        Exit For
                    End If
                Next ilETE
                'SPOTAVAILTIMEINDEX contains the offset time of all events
                'TIMEINDEX contains the Offset time for Programs and avails.  For spots it has the spot time
                slAvailTime = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
                slAvailTime = Trim$(Str$(gStrTimeInTenthToLong(slAvailTime, False)))
                Do While Len(slAvailTime) < 8
                    slAvailTime = "0" & slAvailTime
                Loop
                slBus = grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)
                ilLen = gSetMaxChars("BusName", 0)
                Do While Len(slBus) < ilLen
                    slBus = slBus & " "
                Loop
                slTime = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
                slTime = Trim$(Str$(gStrTimeInTenthToLong(slTime, False)))
                Do While Len(slTime) < 8
                    slTime = "0" & slTime
                Loop
                slStr = slAvailTime & slBus & slType & slTime & grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX)
            ElseIf (ilCol = DURATIONINDEX) Then
                slStr = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
                slStr = Trim$(Str$(gStrLengthInTenthToLong(slStr)))
                Do While Len(slStr) < 8
                    slStr = "0" & slStr
                Loop
            ElseIf (ilCol = SILENCETIMEINDEX) Then
                slStr = grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX)
                slStr = Trim$(Str$(gLengthToLong(slStr))) 'Trim$(Str$(gStrLengthInTenthToLong(slStr)))
                Do While Len(slStr) < 8
                    slStr = "0" & slStr
                Loop
            Else
                slStr = grdLibEvents.TextMatrix(llRow, ilCol)
            End If
            grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr & grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX)
            slStr = grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX)
            If Trim$(grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX)) = "" Then
                grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "1"
            End If
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX)) & slStr
            grdLibEvents.TextMatrix(llRow, SORTTIMEINDEX) = slStr
        End If
    Next llRow
    gGrid_SortByCol grdLibEvents, EVENTTYPEINDEX, SORTTIMEINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    Dim llRow As Long
    
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        cmcDone.Caption = "&Done"
        cmcSave.Enabled = False
        cmcReplace.Enabled = False
        cmcMerge.Enabled = False
        cmcTest.Enabled = False
        cmcLoad.Enabled = False
        cmcItemIDChk.Enabled = False
        cmcConflict.Enabled = False
        imcTrash.Enabled = False
        imcInsert.Enabled = False
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
            cmcDone.Caption = "Save-Go"
        Else
            cmcSave.Enabled = False
            cmcDone.Caption = "&Done"
        End If
    Else
        cmcSave.Enabled = False
        cmcDone.Caption = "&Done"
    End If
    cmcReplace.Enabled = False
    'cmcItemIDChk.Enabled = False
    cmcMerge.Enabled = False
    cmcTest.Enabled = False
    cmcLoad.Enabled = False
    cmcConflict.Enabled = False
    imcTrash.Enabled = False
    imcInsert.Enabled = True
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
            cmcReplace.Enabled = True
            'cmcItemIDChk.Enabled = True
            cmcMerge.Enabled = True
            If (tmSHE.sLoadedAutoStatus <> "L") Then
                cmcLoad.Enabled = True
            End If
            cmcTest.Enabled = True
            cmcConflict.Enabled = True
            imcTrash.Enabled = True
            Exit For
        End If
    Next llRow
    '7/27/11: hide button
    'cmcItemIDChk.Enabled = True
    cmcItemIDChk.Enabled = False
    If tmSHE.lCode = 0 Then
        cmcFilter.Enabled = False
        cmcMerge.Enabled = False
        cmcLoad.Enabled = False
    Else
        cmcFilter.Enabled = True
    End If
    
    If (tmSHE.lCode = 0) Or (tmSHE.sLoadedAutoStatus <> "L") Or (imFieldChgd) Then
        cmcReload.Enabled = False
    Else
        If mBusInFilter() Then
            cmcReload.Enabled = True
        Else
            cmcReload.Enabled = False
        End If
    End If
    
End Sub


Private Sub mEEnableBox()
    Dim ilCol As Integer
    Dim llColPos As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilFieldChgd As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilFound As Integer
    Dim ilETE As Integer
    Dim slBuses As String
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim slEventCategory As String
    
    If igLibCallType = 3 Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igJobStatus(SCHEDULEJOB) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdLibEvents.Row >= grdLibEvents.FixedRows) And (grdLibEvents.Row < grdLibEvents.Rows) And (grdLibEvents.Col >= 0) And (grdLibEvents.Col < grdLibEvents.Cols - 1) Then
        lmEEnableRow = grdLibEvents.Row
        mPaintRowColor grdLibEvents.Row
        ilCol = grdLibEvents.Col
        If grdLibEvents.Col >= TITLE1INDEX Then
            '8/26/11: Horizontal scroll bar required to move columns
'            grdLibEvents.ScrollBars = flexScrollBarBoth
            grdLibEvents.LeftCol = grdLibEvents.LeftCol + 6
            'This do event is required so that the column is moved now
            DoEvents
            '8/26/11: Hide Horizontal scroll bar
'            grdLibEvents.ScrollBars = flexScrollBarVertical
        End If
        If grdLibEvents.Col <= STARTTYPEINDEX Then
            '8/26/11: Horizontal scroll bar required to move columns
'            grdLibEvents.ScrollBars = flexScrollBarBoth
            grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
            'This do event is required so that the column is moved now
            DoEvents
            '8/26/11: Hide Horizontal scroll bar
'            grdLibEvents.ScrollBars = flexScrollBarVertical
        End If
        lmEEnableRow = grdLibEvents.Row
        grdLibEvents.Col = ilCol
        lmEEnableCol = grdLibEvents.Col
        imShowGridBox = True
        pbcArrow.Move grdLibEvents.Left - pbcArrow.Width - 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + (grdLibEvents.RowHeight(grdLibEvents.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        mShowConflictGrid
        'If (Val(grdLibEvents.TextMatrix(lmEEnableRow, PCODEINDEX)) = 0) And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, BUSNAMEINDEX)) <> "") Then
        If (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, EVENTTYPEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        lacHelp.Caption = ""
        lacHelp.Visible = True

        llColPos = 0
        For ilCol = 0 To grdLibEvents.Col - 1 Step 1
            If grdLibEvents.ColIsVisible(ilCol) Then
                llColPos = llColPos + grdLibEvents.ColWidth(ilCol)
            End If
        Next ilCol
        Select Case grdLibEvents.Col
            Case HIGHLIGHTINDEX
                pbcHighlight.Left = -400
                grdLibEvents.text = "»"
                pbcArrow.Visible = False
            Case BUSNAMEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("BusName", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("BusName", 6)
                imMaxColChars = gGetMaxChars("BusName")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcBDE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcBDE, CLng(grdLibEvents.Height / 2)
'                If lbcBDE.Top + lbcBDE.Height > cmcCancel.Top Then
'                    lbcBDE.Top = edcEDropdown.Top - lbcBDE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcBDE, slStr)
                If ilIndex >= 0 Then
                    lbcBDE.ListIndex = ilIndex
                    edcEDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcBDE.ListCount <= 0 Then
                        lbcBDE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcBDE.ListCount <= 1 Then
                            lbcBDE.ListIndex = 0
                            edcEDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                        Else
                            lbcBDE.ListIndex = 1
                            edcEDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Bus Name."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcBDE.Visible = True
'                edcEDropdown.SetFocus
            Case BUSCTRLINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("BusCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("BusCtrl", 6)
                imMaxColChars = gGetMaxChars("BusCtrl")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcCCE_B.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcCCE_B, CLng(grdLibEvents.Height / 2)
'                If lbcCCE_B.Top + lbcCCE_B.Height > cmcCancel.Top Then
'                    lbcCCE_B.Top = edcEDropdown.Top - lbcCCE_B.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcCCE_B, slStr)
                If ilIndex >= 0 Then
                    lbcCCE_B.ListIndex = ilIndex
                    edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcCCE_B.ListCount <= 0 Then
                        lbcCCE_B.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcCCE_B.ListCount <= 1 Then
                            lbcCCE_B.ListIndex = 0
                            edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
                        Else
                            lbcCCE_B.ListIndex = 1
                            edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Bus Control."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcCCE_B.Visible = True
'                edcEDropdown.SetFocus
            Case EVENTTYPEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("EventType", lmCharacterWidth, edcEDropdown.Width, Len(tgETE.sName))
                edcEDropdown.MaxLength = Len(tgETE.sName)
                imMaxColChars = edcEDropdown.MaxLength
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
'                If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
'                    lbcETE.Top = edcEDropdown.Top - lbcETE.Height
'                End If
                '2/9/12: Allow all events
                ''slStr = grdLibEvents.Text
                ''ilIndex = gListBoxFind(lbcETE, slStr)
                ''If ilIndex >= 0 Then
                ''    lbcETE.ListIndex = ilIndex
                ''    edcEDropdown.Text = lbcETE.List(lbcETE.ListIndex)
                ''Else
                ''    edcEDropdown.Text = ""
                ''    If lbcETE.ListCount <= 0 Then
                ''        lbcETE.ListIndex = -1
                ''        edcEDropdown.Text = ""
                ''    Else
                ''        lbcETE.ListIndex = 0
                ''        edcEDropdown.Text = lbcETE.List(lbcETE.ListIndex)
                ''    End If
                ''End If
                'slStr = grdLibEvents.text
                'ilIndex = gListBoxFind(lbcETE_Program, slStr)
                'If ilIndex >= 0 Then
                '    lbcETE_Program.ListIndex = ilIndex
                '    edcEDropdown.text = lbcETE_Program.List(lbcETE_Program.ListIndex)
                'Else
                '    edcEDropdown.text = ""
                '    If lbcETE_Program.ListCount <= 0 Then
                '        lbcETE_Program.ListIndex = -1
                '        edcEDropdown.text = ""
                '    Else
                '        If imDefaultProgIndex <> -1 Then
                '            lbcETE.ListIndex = imDefaultProgIndex
                '        Else
                '            lbcETE.ListIndex = 0
                '        End If
                '        edcEDropdown.text = lbcETE_Program.List(lbcETE_Program.ListIndex)
                '    End If
                'End If
                
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcETE, slStr)
                If ilIndex >= 0 Then
                    lbcETE.ListIndex = ilIndex
                    edcEDropdown.text = lbcETE.List(lbcETE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcETE.ListCount <= 0 Then
                        lbcETE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If imDefaultProgIndex <> -1 Then
                            lbcETE.ListIndex = imDefaultProgIndex
                        Else
                            lbcETE.ListIndex = 0
                        End If
                        edcEDropdown.text = lbcETE.List(lbcETE.ListIndex)
                    End If
                End If
                lacHelp.Caption = "Select Event Type.  The event type indicates which other fields are used and which are mandatory to be answered."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcETE.Visible = True
'                edcEDropdown.SetFocus
            Case TIMEINDEX
''                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
''                edcEvent.Width = gSetCtrlWidth("Time", lmCharacterWidth, edcEvent.Width, 0)
'                edcEvent.MaxLength = gSetMaxChars("Time", 0)
'                imMaxColChars = gGetMaxChars("Time")
'                edcEvent.Text = grdLibEvents.Text
''                lacHelp.Caption = "Enter Time offset of event from the Start Time defined in the Header area.  Time format is hh:mm:ss.t"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
                slStr = grdLibEvents.text
                ltcEvent.CSI_UseHours = True
                ltcEvent.CSI_UseTenths = True
                If Not gIsLengthTenths(slStr) Then
                    ltcEvent.text = ""
                Else
                    ltcEvent.text = ""
                    ltcEvent.text = slStr 'grdLibEvents.Text
                End If
                lacHelp.Caption = "Enter time of this event.  Format is hh:mm:ss.t"
            Case STARTTYPEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("StartType", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("StartType", 6)
                imMaxColChars = gGetMaxChars("StartType")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcTTE_S.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcTTE_S, CLng(grdLibEvents.Height / 2)
'                If lbcTTE_S.Top + lbcTTE_S.Height > cmcCancel.Top Then
'                    lbcTTE_S.Top = edcEDropdown.Top - lbcTTE_S.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcTTE_S, slStr)
                If ilIndex >= 0 Then
                    lbcTTE_S.ListIndex = ilIndex
                    edcEDropdown.text = lbcTTE_S.List(lbcTTE_S.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcTTE_S.ListCount <= 0 Then
                        lbcTTE_S.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcTTE_S.ListCount <= 1 Then
                            lbcTTE_S.ListIndex = 0
                            edcEDropdown.text = lbcTTE_S.List(lbcTTE_S.ListIndex)
                        Else
                            lbcTTE_S.ListIndex = 1
                            edcEDropdown.text = lbcTTE_S.List(lbcTTE_S.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Start Time Type parameter"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcTTE_S.Visible = True
'                edcEDropdown.SetFocus
            Case FIXEDINDEX
'                pbcYN.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
                smYN = grdLibEvents.text
                If (Trim$(smYN) = "") Or (smYN = "Missing") Then
                    smYN = "N"
                End If
                lacHelp.Caption = "Indicate if this is a fixed time event. Enter Y or N or Mouse click to cycle value"
'                pbcYN.Visible = True
'                pbcYN.SetFocus
            Case ENDTYPEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("EndType", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("EndType", 6)
                imMaxColChars = gGetMaxChars("EndType")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcTTE_E.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcTTE_E, CLng(grdLibEvents.Height / 2)
'                If lbcTTE_E.Top + lbcTTE_E.Height > cmcCancel.Top Then
'                    lbcTTE_E.Top = edcEDropdown.Top - lbcTTE_E.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcTTE_E, slStr)
                If ilIndex >= 0 Then
                    lbcTTE_E.ListIndex = ilIndex
                    edcEDropdown.text = lbcTTE_E.List(lbcTTE_E.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcTTE_E.ListCount <= 0 Then
                        lbcTTE_E.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcTTE_E.ListCount <= 1 Then
                            lbcTTE_E.ListIndex = 0
                            edcEDropdown.text = lbcTTE_E.List(lbcTTE_E.ListIndex)
                        Else
                            lbcTTE_E.ListIndex = 1
                            edcEDropdown.text = lbcTTE_E.List(lbcTTE_E.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select End Time Type parameter.  If selected, then the Duration must not be entered."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcTTE_E.Visible = True
'                edcEDropdown.SetFocus
            Case DURATIONINDEX
''                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
''                edcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, edcEvent.Width, 0)
'                edcEvent.MaxLength = gSetMaxChars("Duration", 0)
'                imMaxColChars = gGetMaxChars("Duration")
'                'edcEvent.Text = "00:" & grdLibEvents.Text
'                edcEvent.Text = grdLibEvents.Text
'                lacHelp.Caption = "Enter the length of this event.  If entered, then the End Time Type must not be entered. Format is mm:ss.t"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
                slEventCategory = ""
                slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, EVENTTYPEINDEX))
                If slStr <> "" Then
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                            slEventCategory = tgCurrETE(ilETE).sCategory
                        End If
                    Next ilETE
                End If
                slStr = grdLibEvents.text
                If slEventCategory <> "A" Then
                    ltcEvent.CSI_UseHours = True    'False
                Else
                    ltcEvent.CSI_UseHours = False
                End If
                ltcEvent.CSI_UseTenths = True
                If Not gIsLengthTenths(slStr) Then
                    ltcEvent.text = ""
                Else
                    ltcEvent.text = ""
                    ltcEvent.text = slStr 'grdLibEvents.Text
                End If
                lacHelp.Caption = "Enter the length of this event.  Format is hh:mm:ss.t"
            Case MATERIALINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("Material", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("Material", 6)
                imMaxColChars = gGetMaxChars("Material")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcMTE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcMTE, CLng(grdLibEvents.Height / 2)
'                If lbcMTE.Top + lbcMTE.Height > cmcCancel.Top Then
'                    lbcMTE.Top = edcEDropdown.Top - lbcMTE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcMTE, slStr)
                If ilIndex >= 0 Then
                    lbcMTE.ListIndex = ilIndex
                    edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcMTE.ListCount <= 0 Then
                        lbcMTE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcMTE.ListIndex = 0
                        edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
                        If lbcMTE.ListCount <= 1 Then
                            lbcMTE.ListIndex = 0
                            edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
                        Else
                            lbcMTE.ListIndex = 1
                            edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Material type parameter"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcMTE.Visible = True
'                edcEDropdown.SetFocus
            Case AUDIONAMEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("AudioName", lmCharacterWidth, edcEDropdown.Width, 0)
                edcEDropdown.MaxLength = gSetMaxChars("AudioName", 0)
                imMaxColChars = gGetMaxChars("AudioName")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcASE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcASE, CLng(grdLibEvents.Height / 2)
'                If lbcASE.Top + lbcASE.Height > cmcCancel.Top Then
'                    lbcASE.Top = edcEDropdown.Top - lbcASE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcASE, slStr)
                If ilIndex >= 0 Then
                    lbcASE.ListIndex = ilIndex
                    edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                Else
'                    edcEDropdown.Text = ""
'                    If lbcASE.ListCount <= 0 Then
'                        lbcASE.ListIndex = -1
'                        edcEDropdown.Text = ""
'                    Else
'                        lbcASE.ListIndex = 0
'                        edcEDropdown.Text = lbcASE.List(lbcASE.ListIndex)
'                    End If
                    lbcASE.ListIndex = 0
                    edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                    If lbcASE.ListCount <= 1 Then
                        lbcASE.ListIndex = 0
                        edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                    Else
                        lbcASE.ListIndex = 1
                        edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                    End If
                                                                                                                    
                End If
                lacHelp.Caption = "Select Primary Audio source. From this selection the default Backup and Protection will be set"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcASE.Visible = True
'                edcEDropdown.SetFocus
            Case AUDIOITEMIDINDEX
'                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEvent.Width = gSetCtrlWidth("AudioItemID", lmCharacterWidth, edcEvent.Width, 0)
                edcEvent.MaxLength = gSetMaxChars("AudioItemID", 0)
                imMaxColChars = gGetMaxChars("AudioItemID")
                edcEvent.text = grdLibEvents.text
                lacHelp.Caption = "Enter the Item ID that is to air for this event. Max" & Str$(tgNoCharAFE.iAudioItemID) & " characters"
'                edcEvent.Visible = True
'                edcEvent.SetFocus
            Case AUDIOISCIINDEX
                edcEvent.MaxLength = gSetMaxChars("AudioISCI", 0)
                imMaxColChars = gGetMaxChars("AudioISCI")
                edcEvent.text = grdLibEvents.text
                lacHelp.Caption = "Enter the ISCI that is to air for this event. Max" & Str$(tgNoCharAFE.iAudioISCI) & " characters"
            Case AUDIOCTRLINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("AudioCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("AudioCtrl", 6)
                imMaxColChars = gGetMaxChars("AudioCtrl")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
'                If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
'                    lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcCCE_A, slStr)
                If ilIndex >= 0 Then
                    lbcCCE_A.ListIndex = ilIndex
                    edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcCCE_A.ListCount <= 0 Then
                        lbcCCE_A.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcCCE_A.ListCount <= 1 Then
                            lbcCCE_A.ListIndex = 0
                            edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                        Else
                            lbcCCE_A.ListIndex = 1
                            edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Audio Control.  Default value set on Audio Screen."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcCCE_A.Visible = True
'                edcEDropdown.SetFocus
            Case BACKUPNAMEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("BkupName", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("BkupName", 6)
                imMaxColChars = gGetMaxChars("BkupName")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcANE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcANE, CLng(grdLibEvents.Height / 2)
'                If lbcANE.Top + lbcANE.Height > cmcCancel.Top Then
'                    lbcANE.Top = edcEDropdown.Top - lbcANE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcANE, slStr)
                If ilIndex >= 0 Then
                    lbcANE.ListIndex = ilIndex
                    edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcANE.ListCount <= 0 Then
                        lbcANE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcANE.ListIndex = 0
                        edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                        If lbcANE.ListCount <= 1 Then
                            lbcANE.ListIndex = 0
                            edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                        Else
                            lbcANE.ListIndex = 1
                            edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Backup Audio source."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcANE.Visible = True
'                edcEDropdown.SetFocus
            Case BACKUPCTRLINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("BkupCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("BkupCtrl", 6)
                imMaxColChars = gGetMaxChars("BkupCtrl")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
'                If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
'                    lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcCCE_A, slStr)
                If ilIndex >= 0 Then
                    lbcCCE_A.ListIndex = ilIndex
                    edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcCCE_A.ListCount <= 0 Then
                        lbcCCE_A.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcCCE_A.ListCount <= 1 Then
                            lbcCCE_A.ListIndex = 0
                            edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                        Else
                            lbcCCE_A.ListIndex = 1
                            edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Audio Control.  Default value set on Audio Screen."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcCCE_A.Visible = True
'                edcEDropdown.SetFocus
            Case PROTNAMEINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("ProtName", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("ProtName", 6)
                imMaxColChars = gGetMaxChars("ProtName")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcANE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcANE, CLng(grdLibEvents.Height / 2)
'                If lbcANE.Top + lbcANE.Height > cmcCancel.Top Then
'                    lbcANE.Top = edcEDropdown.Top - lbcANE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcANE, slStr)
                If ilIndex >= 0 Then
                    lbcANE.ListIndex = ilIndex
                    edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcANE.ListCount <= 0 Then
                        lbcANE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcANE.ListIndex = 0
                        edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                        If lbcANE.ListCount <= 1 Then
                            lbcANE.ListIndex = 0
                            edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                        Else
                            lbcANE.ListIndex = 1
                            edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Protection Audio source (Backup of the Backup Audio source)."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcANE.Visible = True
'                edcEDropdown.SetFocus
            Case PROTITEMIDINDEX
'                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEvent.Width = gSetCtrlWidth("ProtItemID", lmCharacterWidth, edcEvent.Width, 0)
                edcEvent.MaxLength = gSetMaxChars("ProtItemID", 0)
                imMaxColChars = gGetMaxChars("ProtItemID")
                edcEvent.text = grdLibEvents.text
                lacHelp.Caption = "Enter the Item ID that is to air for this event. Max" & Str$(tgNoCharAFE.iProtItemID) & " characters"
'                edcEvent.Visible = True
'                edcEvent.SetFocus
            Case PROTISCIINDEX
                edcEvent.MaxLength = gSetMaxChars("ProtISCI", 0)
                imMaxColChars = gGetMaxChars("ProtISCI")
                edcEvent.text = grdLibEvents.text
                lacHelp.Caption = "Enter the ISCI that is to air for this event. Max" & Str$(tgNoCharAFE.iProtISCI) & " characters"
            Case PROTCTRLINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("ProtCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("ProtCtrl", 6)
                imMaxColChars = gGetMaxChars("ProtCtrl")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
'                If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
'                    lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcCCE_A, slStr)
                If ilIndex >= 0 Then
                    lbcCCE_A.ListIndex = ilIndex
                    edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcCCE_A.ListCount <= 0 Then
                        lbcCCE_A.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        If lbcCCE_A.ListCount <= 1 Then
                            lbcCCE_A.ListIndex = 0
                            edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                        Else
                            lbcCCE_A.ListIndex = 1
                            edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Audio Control.  Default value set on Audio Screen."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcCCE_A.Visible = True
'                edcEDropdown.SetFocus
            Case RELAY1INDEX, RELAY2INDEX
                If grdLibEvents.Col = RELAY2INDEX Then
                    slStr = "Relay2"
                Else
                    slStr = "Relay1"
                End If
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars(slStr, 6)
                imMaxColChars = gGetMaxChars(slStr)
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcRNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcRNE, CLng(grdLibEvents.Height / 2)
'                If lbcRNE.Top + lbcRNE.Height > cmcCancel.Top Then
'                    lbcRNE.Top = edcEDropdown.Top - lbcRNE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcRNE, slStr)
                If ilIndex >= 0 Then
                    lbcRNE.ListIndex = ilIndex
                    edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcRNE.ListCount <= 0 Then
                        lbcRNE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcRNE.ListIndex = 0
                        edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
                        If lbcRNE.ListCount <= 1 Then
                            lbcRNE.ListIndex = 0
                            edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
                        Else
                            lbcRNE.ListIndex = 1
                            edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Relay parameter.  Relay 1 and 2 must be different"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcRNE.Visible = True
'                edcEDropdown.SetFocus
            Case FOLLOWINDEX
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("Follow", lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars("Follow", 6)
                imMaxColChars = gGetMaxChars("Follow")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcFNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcFNE, CLng(grdLibEvents.Height / 2)
'                If lbcFNE.Top + lbcFNE.Height > cmcCancel.Top Then
'                    lbcFNE.Top = edcEDropdown.Top - lbcFNE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcFNE, slStr)
                If ilIndex >= 0 Then
                    lbcFNE.ListIndex = ilIndex
                    edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcFNE.ListCount <= 0 Then
                        lbcFNE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcFNE.ListIndex = 0
                        edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
                        If lbcFNE.ListCount <= 1 Then
                            lbcFNE.ListIndex = 0
                            edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
                        Else
                            lbcFNE.ListIndex = 1
                            edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Follow parameter."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcFNE.Visible = True
'                edcEDropdown.SetFocus
            Case SILENCETIMEINDEX
''                edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
''                edcEvent.Width = gSetCtrlWidth("SilenceTime", lmCharacterWidth, edcEvent.Width, 0)
'                edcEvent.MaxLength = gSetMaxChars("SilenceTime", 0)
'                imMaxColChars = gGetMaxChars("SilenceTime")
'                edcEvent.Text = grdLibEvents.Text
'                lacHelp.Caption = "Enter the allowed silence time of this event. Format is mm:ss"
''                edcEvent.Visible = True
''                edcEvent.SetFocus
                slStr = grdLibEvents.text
                ltcEvent.CSI_UseHours = False
                ltcEvent.CSI_UseTenths = False
                If Not gIsLength(slStr) Then
                    ltcEvent.text = ""
                Else
                    ltcEvent.text = ""
                    ltcEvent.text = slStr 'grdLibEvents.Text
                End If
                lacHelp.Caption = "Enter the allowed silence time of this event.  Format is mm:ss"
            Case SILENCE1INDEX To SILENCE4INDEX
                If grdLibEvents.Col = SILENCE2INDEX Then
                    slStr = "Silence2"
                ElseIf grdLibEvents.Col = SILENCE3INDEX Then
                    slStr = "Silence3"
                ElseIf grdLibEvents.Col = SILENCE4INDEX Then
                    slStr = "Silence4"
                Else
                    slStr = "Silence1"
                End If
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars(slStr, 6)
                imMaxColChars = gGetMaxChars(slStr)
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcSCE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcSCE, CLng(grdLibEvents.Height / 2)
'                If lbcSCE.Top + lbcSCE.Height > cmcCancel.Top Then
'                    lbcSCE.Top = edcEDropdown.Top - lbcSCE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcSCE, slStr)
                If ilIndex >= 0 Then
                    lbcSCE.ListIndex = ilIndex
                    edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcSCE.ListCount <= 0 Then
                        lbcSCE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcSCE.ListIndex = 0
                        edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
                        If lbcSCE.ListCount <= 1 Then
                            lbcSCE.ListIndex = 0
                            edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
                        Else
                            lbcSCE.ListIndex = 1
                            edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
                        End If
                    End If
                End If
                lacHelp.Caption = "Select Silence parameter.  All must be different"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcSCE.Visible = True
'                edcEDropdown.SetFocus
            Case NETCUE1INDEX, NETCUE2INDEX
                If grdLibEvents.Col = NETCUE2INDEX Then
                    slStr = "Netcue2"
                Else
                    slStr = "Netcue1"
                End If
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
                edcEDropdown.MaxLength = gSetMaxChars(slStr, 6)
                imMaxColChars = gGetMaxChars(slStr)
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcNNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcNNE, CLng(grdLibEvents.Height / 2)
'                If lbcNNE.Top + lbcNNE.Height > cmcCancel.Top Then
'                    lbcNNE.Top = edcEDropdown.Top - lbcNNE.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcNNE, slStr)
                If ilIndex >= 0 Then
                    lbcNNE.ListIndex = ilIndex
                    edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
                Else
                    edcEDropdown.text = ""
                    If lbcNNE.ListCount <= 0 Then
                        lbcNNE.ListIndex = -1
                        edcEDropdown.text = ""
                    Else
                        lbcNNE.ListIndex = 0
                        edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
                        If lbcNNE.ListCount <= 1 Then
                            lbcNNE.ListIndex = 0
                            edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
                        Else
                            lbcNNE.ListIndex = 1
                            edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
                        End If
                    End If
                End If
                'lacHelp.Caption = "Select Netcue parameter.  Netque 1 and 2 must be different"
                lacHelp.Caption = "Select Netcue parameter."
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcNNE.Visible = True
'                edcEDropdown.SetFocus
            Case TITLE1INDEX
                mLoadCTE_1 Trim$(grdLibEvents.TextMatrix(lmEEnableRow, EVENTTYPEINDEX))
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("Title1", lmCharacterWidth, edcEDropdown.Width, 6)
'                edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
                edcEDropdown.MaxLength = gSetMaxChars("Title1", 6)
                imMaxColChars = gGetMaxChars("Title1")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcCTE_1.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcCTE_1, CLng(grdLibEvents.Height / 2)
'                If lbcCTE_1.Top + lbcCTE_1.Height > cmcCancel.Top Then
'                    lbcCTE_1.Top = edcEDropdown.Top - lbcCTE_1.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcCTE_1, slStr)
                If ilIndex >= 0 Then
                    lbcCTE_1.ListIndex = ilIndex
                    edcEDropdown.text = lbcCTE_1.List(lbcCTE_1.ListIndex)
                Else
                    edcEDropdown.text = slStr
                End If
                lacHelp.Caption = "Enter the First Title that is to air for this event. Max" & Str$(imMaxColChars) & " characters"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcCTE_1.Visible = True
'                edcEDropdown.SetFocus
            Case TITLE2INDEX
                mLoadCTE_2
'                edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'                edcEDropdown.Width = gSetCtrlWidth("Title2", lmCharacterWidth, edcEDropdown.Width, 6)
'                edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
                edcEDropdown.MaxLength = gSetMaxChars("Title2", 6)
                imMaxColChars = gGetMaxChars("Title2")
'                cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
'                lbcCTE_2.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
'                gSetListBoxHeight lbcCTE_2, CLng(grdLibEvents.Height / 2)
'                If lbcCTE_2.Top + lbcCTE_2.Height > cmcCancel.Top Then
'                    lbcCTE_2.Top = edcEDropdown.Top - lbcCTE_2.Height
'                End If
                slStr = grdLibEvents.text
                ilIndex = gListBoxFind(lbcCTE_2, slStr)
                If ilIndex >= 0 Then
                    lbcCTE_2.ListIndex = ilIndex
                    edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
                Else
                    '7/8/11: Make T2 work like T1
                    'edcEDropdown.text = ""
                    'If lbcCTE_2.ListCount <= 0 Then
                    '    lbcCTE_2.ListIndex = -1
                    '    edcEDropdown.text = ""
                    'Else
                    '    lbcCTE_2.ListIndex = 0
                    '    edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
                    '    If lbcCTE_2.ListCount <= 1 Then
                    '        lbcCTE_2.ListIndex = 0
                    '        edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
                    '    Else
                    '        lbcCTE_2.ListIndex = 1
                    '        edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
                    '    End If
                    'End If
                    edcEDropdown.text = slStr
                End If
                lacHelp.Caption = "Enter the Second Title that is to air for this event. Max" & Str$(imMaxColChars) & " characters"
'                edcEDropdown.Visible = True
'                cmcEDropDown.Visible = True
'                lbcCTE_2.Visible = True
'                edcEDropdown.SetFocus
            Case ABCFORMATINDEX
                edcEvent.MaxLength = gSetMaxChars("ABCFormat", 0)
                imMaxColChars = gGetMaxChars("ABCFormat")
                edcEvent.text = grdLibEvents.text
                If (Trim$(edcEvent.text) = "") And (Val(grdLibEvents.TextMatrix(lmEEnableRow, PCODEINDEX)) = 0) Then
                    edcEvent.text = ""
                    slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, EVENTTYPEINDEX))
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                            If tgCurrETE(ilETE).sCategory <> "P" Then
                                edcEvent.text = "0"
                            End If
                        End If
                    Next ilETE
                End If
                lacHelp.Caption = "Enter the ABC Format that is to air for this event. Max" & Str$(tgNoCharAFE.iABCFormat) & " characters"
            Case ABCPGMCODEINDEX
                edcEvent.MaxLength = gSetMaxChars("ABCPgmCode", 0)
                imMaxColChars = gGetMaxChars("ABCPgmCode")
                edcEvent.text = grdLibEvents.text
                lacHelp.Caption = "Enter the Program Code that is to air for this event. Max" & Str$(tgNoCharAFE.iABCPgmCode) & " characters"
            Case ABCXDSMODEINDEX
                edcEvent.MaxLength = gSetMaxChars("ABCXdsMode", 0)
                imMaxColChars = gGetMaxChars("ABCXdsMode")
                edcEvent.text = grdLibEvents.text
                If (Trim$(edcEvent.text) = "") And (Val(grdLibEvents.TextMatrix(lmEEnableRow, PCODEINDEX)) = 0) Then
                    edcEvent.text = ""
                    slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, EVENTTYPEINDEX))
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                            If tgCurrETE(ilETE).sCategory <> "P" Then
                                edcEvent.text = "*"
                            End If
                        End If
                    Next ilETE
                End If
                lacHelp.Caption = "Enter the XDS Mode that is to air for this event. Max" & Str$(tgNoCharAFE.iABCXDSMode) & " characters"
            Case ABCRECORDITEMINDEX
                edcEvent.MaxLength = gSetMaxChars("ABCRecordItem", 0)
                imMaxColChars = gGetMaxChars("ABCRecordItem")
                edcEvent.text = grdLibEvents.text
                lacHelp.Caption = "Enter the ABC Record that is to air for this event. Max" & Str$(tgNoCharAFE.iABCRecordItem) & " characters"
        End Select
        smESCValue = grdLibEvents.text
        mESetFocus
    End If
End Sub

Private Sub mESetShow()
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilSpot As Integer
    Dim slBus As String
    Dim llSpotTime As Long
    Dim llDurTime As Long
    Dim llRow As Long
    Dim slDur As String
    Dim llAvailTime As Long
    Dim llTime As Long
    Dim slOrigValue As String
    Dim ilRet As Integer
    Dim llSEE As Long
    Dim llRowSEE As Long
    
    If (lmEEnableRow >= grdLibEvents.FixedRows) And (lmEEnableRow < grdLibEvents.Rows) Then
        Select Case lmEEnableCol
            Case HIGHLIGHTINDEX
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = ""
            Case BUSNAMEINDEX
            Case BUSCTRLINDEX
            Case EVENTTYPEINDEX
                If imInsertState = True Then
                    ilSpot = False
                    slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                            If tgCurrETE(ilETE).sCategory = "S" Then
                                ilSpot = True
                            End If
                            Exit For
                        End If
                    Next ilETE
                    If Not ilSpot Then
                        imInsertState = False
                        lmInsertRow = -1
                    End If
                    mSetColExportColor lmEEnableRow
                End If
            Case TIMEINDEX
            Case STARTTYPEINDEX
            Case FIXEDINDEX
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = smYN
            Case ENDTYPEINDEX
            Case DURATIONINDEX
            Case MATERIALINDEX
            Case AUDIONAMEINDEX
            Case AUDIOITEMIDINDEX
            Case AUDIOISCIINDEX
            Case AUDIOCTRLINDEX
            Case BACKUPNAMEINDEX
            Case BACKUPCTRLINDEX
            Case PROTNAMEINDEX
            Case PROTITEMIDINDEX
            Case PROTISCIINDEX
            Case PROTCTRLINDEX
            Case RELAY1INDEX, RELAY2INDEX
            Case FOLLOWINDEX
            Case SILENCETIMEINDEX
            Case SILENCE1INDEX To SILENCE4INDEX
            Case NETCUE1INDEX, NETCUE2INDEX
            Case TITLE1INDEX
                slStr = UCase(Trim$(edcEDropdown.text))
                If (slStr <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)) <> "") Then
                    If UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))) <> slStr Then
                        ilRet = MsgBox("Change all occurrences of this Comment within this Schedule from the same Library", vbQuestion + vbYesNo + vbDefaultButton2, "Comment Changed")
                        If ilRet = vbYes Then
                            slOrigValue = UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)))
                            llSEE = Val(grdLibEvents.TextMatrix(lmEEnableRow, TMCURRSEEINDEX))
                            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                                llRowSEE = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
                                If tmCurrSEE(llSEE).lDheCode = tmCurrSEE(llRowSEE).lDheCode Then
                                    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                                        If UCase(Trim$(grdLibEvents.TextMatrix(llRow, lmEEnableCol))) = slOrigValue Then
                                            grdLibEvents.TextMatrix(llRow, lmEEnableCol) = Trim$(edcEDropdown.text)
                                        End If
                                    End If
                                End If
                            Next llRow
                        End If
                    End If
                End If
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = Trim$(edcEDropdown.text)
            Case TITLE2INDEX
                slStr = UCase(Trim$(edcEDropdown.text))
                If (slStr <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)) <> "") Then
                    If UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))) <> slStr Then
                        ilRet = MsgBox("Change all occurrences of this Comment within this Schedule from the same Library", vbQuestion + vbYesNo + vbDefaultButton2, "Comment Changed")
                        If ilRet = vbYes Then
                            slOrigValue = UCase(Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol)))
                            llSEE = Val(grdLibEvents.TextMatrix(lmEEnableRow, TMCURRSEEINDEX))
                            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                                llRowSEE = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
                                If tmCurrSEE(llSEE).lDheCode = tmCurrSEE(llRowSEE).lDheCode Then
                                    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                                        If UCase(Trim$(grdLibEvents.TextMatrix(llRow, lmEEnableCol))) = slOrigValue Then
                                            grdLibEvents.TextMatrix(llRow, lmEEnableCol) = Trim$(edcEDropdown.text)
                                        End If
                                    End If
                                End If
                            Next llRow
                        End If
                    End If
                End If
                grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol) = Trim$(edcEDropdown.text)
            Case ABCFORMATINDEX
            Case ABCPGMCODEINDEX
            Case ABCXDSMODEINDEX
            Case ABCRECORDITEMINDEX
        End Select
    End If
    lacHelp.Visible = False
    pbcEDefine.Visible = False
    edcEDropdown.Visible = False
    cmcEDropDown.Visible = False
    lbcBDE.Visible = False
    lbcCCE_B.Visible = False
    lbcETE_Program.Visible = False
    lbcETE.Visible = False
    edcEvent.Visible = False
    lbcTTE_S.Visible = False
    lbcTTE_E.Visible = False
    pbcYN.Visible = False
    lbcMTE.Visible = False
    lbcASE.Visible = False
    lbcCCE_A.Visible = False
    lbcANE.Visible = False
    lbcRNE.Visible = False
    lbcFNE.Visible = False
    lbcSCE.Visible = False
    lbcNNE.Visible = False
    lbcCTE_1.Visible = False
    lbcCTE_2.Visible = False
    ltcEvent.Visible = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    pbcHighlight.Visible = False
    'mHideConflictGrid
    imShowGridBox = False
    If imInsertState Then
        If (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, BUSNAMEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, EVENTTYPEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, TIMEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, DURATIONINDEX)) <> "") Then
            'Determine if avail is to be removed
            gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
            grdLibEvents.Redraw = False
            'use the moves so that the avail will be removed is required
            'If (UBound(tgFilterValues) <= LBound(tgFilterValues)) Then
            '    'Determine avail time
            '    slBus = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, BUSNAMEINDEX))
            '    slStr = (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, TIMEINDEX)))
            '    llSpotTime = gStrTimeInTenthToLong(slStr, False)
            '    slStr = (Trim$(grdLibEvents.TextMatrix(lmEEnableRow, DURATIONINDEX)))
            '    llDurTime = gStrTimeInTenthToLong(slStr, False)
            '    If Trim$(grdLibEvents.TextMatrix(lmEEnableRow, SPOTAVAILTIMEINDEX)) = "" Then
            '        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            '            If (Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX)) <> "") Then
            '                If StrComp(slBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0 Then
            '                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            '                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            '                        If StrComp(slBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0 Then
            '                            If tgCurrETE(ilETE).sCategory = "S" Then
            '                                slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
            '                                slDur = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
            '                                If (llSpotTime = gStrTimeInTenthToLong(slStr, False)) And (StrComp(slBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0) Then
            '                                    grdLibEvents.TextMatrix(lmEEnableRow, SPOTAVAILTIMEINDEX) = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
            '                                ElseIf (llSpotTime = gStrTimeInTenthToLong(slStr, False) + gStrTimeInTenthToLong(slDur, False)) And (StrComp(slBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0) Then
            '                                    grdLibEvents.TextMatrix(lmEEnableRow, SPOTAVAILTIMEINDEX) = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
            '                                End If
            '                            ElseIf tgCurrETE(ilETE).sCategory = "A" Then
            '                                slStr = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
            '                                If (llSpotTime = gStrTimeInTenthToLong(slStr, False)) And (StrComp(slBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0) Then
            '                                    grdLibEvents.TextMatrix(lmEEnableRow, SPOTAVAILTIMEINDEX) = slStr
            '                                End If
            '                            End If
            '                        End If
            '                        Exit For
            '                    Next ilETE
            '                    If Trim$(grdLibEvents.TextMatrix(lmEEnableRow, SPOTAVAILTIMEINDEX)) <> "" Then
            '                        Exit For
            '                    End If
            '                End If
            '            End If
            '        Next llRow
            '    End If
            '    'Adjust spot times
            '    slStr = grdLibEvents.TextMatrix(lmEEnableRow, SPOTAVAILTIMEINDEX)
            '    llAvailTime = gStrTimeInTenthToLong(slStr, False)
            '    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            '        If (Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX)) <> "") Then
            '            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            '            For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            '                If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
            '                    If tgCurrETE(ilETE).sCategory = "S" Then
            '                        slStr = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
            '                        If (llAvailTime = gStrTimeInTenthToLong(slStr, False)) And (StrComp(slBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0) Then
            '                            slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
            '                            llTime = gStrTimeInTenthToLong(slStr, False)
            '                            If llSpotTime < llTime Then
            '                                llTime = llTime + llDurTime
            '                                grdLibEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(llTime)
            '                            End If
            '                        End If
            '                    End If
            '                    Exit For
            '                End If
            '            Next ilETE
            '        End If
            '    Next llRow
            '    'Determine if avail should be removed
            'Else
'2/11/12: Not required
'                mMoveSEECtrlsToRec
'                mSetAvailTime
'                'Test actual record
'                grdLibEvents.Redraw = False
'                grdLibEvents.Visible = False
'                mMoveSEERecToCtrls
            'End If
            imInsertState = False
            lmInsertRow = -1
            grdLibEvents.Redraw = True
            grdLibEvents.Visible = True
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        End If
    End If
    llRow = lmEEnableRow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mPaintRowColor llRow
End Sub

Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llLength As Long
    Dim slDHEHours As String
    Dim slDEEHours As String
    Dim ilStartHour As Integer
    Dim ilEndHour As Integer
    'Dim slAllowedHours As String
    Dim ilHour As Integer
    Dim tlUsedEPE As EPE
    Dim tlManEPE As EPE
    Dim ilETE As Integer
    Dim ilEPE As Integer
    Dim llELength As Long
    Dim llDuration As Long
    Dim ilStartDay As Integer
    Dim ilEndDay As Integer
    Dim slAllowedDays As String
    Dim ilDay As Integer
    Dim slDEEDays As String
    Dim slDHEDays As String
    Dim llSilence As Long
    Dim llETime As Long
    Dim llLEndTime As Long
    Dim ilSHour As Integer
    Dim ilCol As Integer
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llErrorColor As Long
    
    If smAirDate = "" Then
        mCheckFields = True
        Exit Function
    End If
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    If ilTestState Then
        grdLibEvents.Redraw = False
        If ilTestState Then
            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "1"
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
                If slStr <> "" Then
                    grdLibEvents.Row = llRow
                    For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                        grdLibEvents.Col = ilCol
                        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                            grdLibEvents.CellForeColor = vbBlue
                        Else
                            grdLibEvents.CellForeColor = vbBlack
                        End If
                    Next ilCol
                End If
            Next llRow
        End If
        'Test if fields defined
        ilError = False
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, TIMEINDEX))
            If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sTime = "Y") And (tlManEPE.sTime = "Y") Then
                llErrorColor = vbRed
            Else
                If gIsTimeTenths(slStr) Then
                    If llAirDate = llNowDate Then
                        If llNowTime > gStrTimeInTenthToLong(slStr, False) Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
                            If (slStr <> "") And (Val(slStr) <> 0) Then
                                llErrorColor = BURGUNDY
                            Else
                                llErrorColor = vbRed
                            End If
                        Else
                            llErrorColor = vbRed
                        End If
                    Else
                        llErrorColor = vbRed
                    End If
                Else
                    llErrorColor = vbRed
                End If
            End If
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            If slStr = "" Then
                slStr = grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)
                If slStr <> "" Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = "Missing"
                    grdLibEvents.Row = llRow
                    grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
                    grdLibEvents.Col = EVENTTYPEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
            Else
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                        For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                            If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                                If tgCurrEPE(ilEPE).sType = "U" Then
                                    LSet tlUsedEPE = tgCurrEPE(ilEPE)
                                End If
                                If tgCurrEPE(ilEPE).sType = "M" Then
                                    LSet tlManEPE = tgCurrEPE(ilEPE)
                                End If
                            End If
                        Next ilEPE
                    End If
                Next ilETE
                slStr = grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sBus = "Y") And (tlManEPE.sBus = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = "Missing"
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = BUSNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TIMEINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sTime = "Y") And (tlManEPE.sTime = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, TIMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = TIMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                llETime = -1
                ilSHour = -1
                llELength = 0
                If (slStr <> "") And (tlUsedEPE.sTime = "Y") Then
                    If Not gIsTimeTenths(slStr) Then    'gIsLengthTenths(slStr) Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = TIMEINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    Else
                        llELength = gStrTimeInTenthToLong(slStr, False) 'gStrLengthInTenthToLong(slStr)
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sStartType = "Y") And (tlManEPE.sStartType = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = STARTTYPEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, FIXEDINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sFixedTime = "Y") And (tlManEPE.sFixedTime = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, FIXEDINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = FIXEDINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sEndType = "Y") And (tlManEPE.sEndType = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = ENDTYPEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, DURATIONINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sDuration = "Y") And (tlManEPE.sDuration = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = DURATIONINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                If slStr <> "" Then
                    If Not gIsLengthTenths(slStr) Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = DURATIONINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    Else
                        '3/29/06-Allow end time to overlap into next day
                        'llDuration = gStrLengthInTenthToLong(slStr)
                        'If llELength + llDuration <= 24 * CLng(36000) Then
                        '    'If (llLEndTime >= 0) And (ilSHour >= 0) Then
                        '    '    llETime = llELength + llDuration + (ilSHour - 1) * CLng(3600) * 10 - 1
                        '    '    If llETime > 10 * llLEndTime Then
                        '    '        ilError = True
                        '    '        grdLibEvents.Row = llRow
                        '    '        grdLibEvents.Col = DURATIONINDEX
                        '    '        grdLibEvents.CellForeColor = vbRed
                        '    '    End If
                        '    'End If
                        'Else
                        '    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        '    If llErrorColor = vbRed Then
                        '        ilError = True
                        '    End If
                        '    grdLibEvents.Row = llRow
                        '    grdLibEvents.Col = DURATIONINDEX
                        '    grdLibEvents.CellForeColor = llErrorColor
                        'End If
                    End If
                    '11/24/04- Allow end type and Duration to co-exist
                    'slStr = Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX))
                    'If (slStr <> "") Then
                    '    If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                    '        ilError = True
                    '        grdLibEvents.Row = llRow
                    '        grdLibEvents.Col = ENDTYPEINDEX
                    '        grdLibEvents.CellForeColor = vbRed
                    '    End If
                    'End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sEndType = "Y") Then
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, DURATIONINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sDuration = "Y") Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = DURATIONINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, MATERIALINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sMaterialType = "Y") And (tlManEPE.sMaterialType = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = MATERIALINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sAudioName = "Y") And (tlManEPE.sAudioName = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIONAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
'Moved to Conflict testing
'                    If slStr <> "" Then
'                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
'                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX))
'                            If slStr <> "" Then
'                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
'                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX)), slStr, vbTextCompare) = 0 Then
'                                        ilError = True
'                                        grdLibEvents.Row = llRow
'                                        grdLibEvents.Col = BACKUPNAMEINDEX
'                                        grdLibEvents.CellForeColor = vbRed
'                                    End If
'                                End If
'                            End If
'                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
'                            If slStr <> "" Then
'                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
'                                    'If (StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)), Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)), vbTextCompare) = 0) And (Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)) <> "") Then
'                                        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX)), slStr, vbTextCompare) = 0 Then
'                                            ilError = True
'                                            grdLibEvents.Row = llRow
'                                            grdLibEvents.Col = PROTNAMEINDEX
'                                            grdLibEvents.CellForeColor = vbRed
'                                        End If
'                                    'End If
'                                End If
'                            End If
'                        End If
'                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sAudioItemID = "Y") And (tlManEPE.sAudioItemID = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIOITEMIDINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sAudioISCI = "Y") And (tlManEPE.sAudioISCI = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIOISCIINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sAudioControl = "Y") And (tlManEPE.sAudioControl = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = AUDIOCTRLINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sBkupAudioName = "Y") And (tlManEPE.sBkupAudioName = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = BACKUPNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
'Moved to Conflict testing
'                    If slStr <> "" Then
'                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
'                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
'                            If slStr <> "" Then
'                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
'                                    'If (StrComp(Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)), Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)), vbTextCompare) = 0) And (Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)) <> "") Then
'                                        If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX)), slStr, vbTextCompare) = 0 Then
'                                            ilError = True
'                                            grdLibEvents.Row = llRow
'                                            grdLibEvents.Col = PROTNAMEINDEX
'                                            grdLibEvents.CellForeColor = vbRed
'                                        End If
'                                    'End If
'                                End If
'                            End If
'                        End If
'                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sBkupAudioControl = "Y") And (tlManEPE.sBkupAudioControl = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = BACKUPCTRLINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sProtAudioName = "Y") And (tlManEPE.sProtAudioName = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sProtAudioItemID = "Y") And (tlManEPE.sProtAudioItemID = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTITEMIDINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTISCIINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sProtAudioISCI = "Y") And (tlManEPE.sProtAudioISCI = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTISCIINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTISCIINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sProtAudioControl = "Y") And (tlManEPE.sProtAudioControl = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = PROTCTRLINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sRelay1 = "Y") And (tlManEPE.sRelay1 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = RELAY1INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = RELAY2INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sRelay2 = "Y") And (tlManEPE.sRelay2 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, RELAY2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = RELAY2INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, FOLLOWINDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sFollow = "Y") And (tlManEPE.sFollow = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = FOLLOWINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX))
                If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sSilenceTime = "Y") And (tlManEPE.sSilenceTime = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCETIMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                If slStr <> "" Then
                    If Not gIsLength(slStr) Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = SILENCETIMEINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    Else
                        '3/30/06-Allow silence to cross midnight
                        'llSilence = 10 * gLengthToLong(slStr) 'gStrLengthInTenthToLong(slStr)  'gLengthToLong(slStr)
                        'If llELength + llSilence < 24 * CLng(36000) Then
                        '    'If (llLEndTime >= 0) And (ilSHour >= 0) Then
                        '    '    llETime = llELength + llSilence + (ilSHour - 1) * CLng(3600) * 10 - 1
                        '    '    If llETime > 10 * llLEndTime Then
                        '    '        ilError = True
                        '    '        grdLibEvents.Row = llRow
                        '    '        grdLibEvents.Col = SILENCETIMEINDEX
                        '    '        grdLibEvents.CellForeColor = vbRed
                        '    '    End If
                        '    'End If
                        'Else
                        '    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        '    If llErrorColor = vbRed Then
                        '        ilError = True
                        '    End If
                        '    grdLibEvents.Row = llRow
                        '    grdLibEvents.Col = SILENCETIMEINDEX
                        '    grdLibEvents.CellForeColor = llErrorColor
                        'End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence1 = "Y") And (tlManEPE.sSilence1 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE1INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE2INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE3INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE4INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence2 = "Y") And (tlManEPE.sSilence2 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE2INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE3INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE4INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence3 = "Y") And (tlManEPE.sSilence3 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE3INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
                    If slStr <> "" Then
                        If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                            If slStr <> "" Then
                                If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX)), slStr, vbTextCompare) = 0 Then
                                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                                        If llErrorColor = vbRed Then
                                            ilError = True
                                        End If
                                        grdLibEvents.Row = llRow
                                        grdLibEvents.Col = SILENCE4INDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sSilence4 = "Y") And (tlManEPE.sSilence4 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = SILENCE4INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sStartNetcue = "Y") And (tlManEPE.sStartNetcue = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = NETCUE1INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                Else
                    '9/13/11:  Allow Netcue to be the same
                    'If slStr <> "" Then
                    '    If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                    '        slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX))
                    '        If slStr <> "" Then
                    '            If StrComp(slStr, "[None]", vbTextCompare) <> 0 Then
                    '                If StrComp(Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX)), slStr, vbTextCompare) = 0 Then
                    '                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    '                    If llErrorColor = vbRed Then
                    '                        ilError = True
                    '                    End If
                    '                    grdLibEvents.Row = llRow
                    '                    grdLibEvents.Col = NETCUE2INDEX
                    '                    grdLibEvents.CellForeColor = llErrorColor
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'End If
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sStopNetcue = "Y") And (tlManEPE.sStopNetcue = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = NETCUE2INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sTitle1 = "Y") And (tlManEPE.sTitle1 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = TITLE1INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
                If (((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) Or StrComp(slStr, "[None]", vbTextCompare) = 0) And (tlUsedEPE.sTitle2 = "Y") And (tlManEPE.sTitle2 = "Y") Then
                    grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    If llErrorColor = vbRed Then
                        ilError = True
                    End If
                    If slStr = "" Then
                        grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = "Missing"
                    End If
                    grdLibEvents.Row = llRow
                    grdLibEvents.Col = TITLE2INDEX
                    grdLibEvents.CellForeColor = llErrorColor
                End If
                If sgClientFields = "A" Then
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCFormat = "Y") And (tlManEPE.sABCFormat = "Y") Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCFORMATINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    End If
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCPgmCode = "Y") And (tlManEPE.sABCPgmCode = "Y") Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCPGMCODEINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    End If
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCXDSMode = "Y") And (tlManEPE.sABCXDSMode = "Y") Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCXDSMODEINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    End If
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX))
                    If ((slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0)) And (tlUsedEPE.sABCRecordItem = "Y") And (tlManEPE.sABCRecordItem = "Y") Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                        If llErrorColor = vbRed Then
                            ilError = True
                        End If
                        If slStr = "" Then
                            grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = "Missing"
                        End If
                        grdLibEvents.Row = llRow
                        grdLibEvents.Col = ABCRECORDITEMINDEX
                        grdLibEvents.CellForeColor = llErrorColor
                    End If
                End If
            End If
        Next llRow
    End If
    grdLibEvents.Redraw = True
    If ilError Then
        mCheckFields = False
        Exit Function
    Else
        mCheckFields = True
        Exit Function
    End If
End Function


Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    
    gGrid_AlignAllColsLeft grdLibEvents
    mGridColumnWidth
    'Set Titles
    'Set Titles
    For ilCol = BUSNAMEINDEX To BUSCTRLINDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Bus"
    Next ilCol
    For ilCol = AUDIONAMEINDEX To AUDIOCTRLINDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Audio"
    Next ilCol
    For ilCol = BACKUPNAMEINDEX To BACKUPCTRLINDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "B'kup"
    Next ilCol
    For ilCol = PROTNAMEINDEX To PROTCTRLINDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Protection"
    Next ilCol
    For ilCol = RELAY1INDEX To RELAY2INDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Relay"
    Next ilCol
    For ilCol = SILENCETIMEINDEX To SILENCE4INDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Silence"
    Next ilCol
    For ilCol = NETCUE1INDEX To NETCUE2INDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Netcue"
    Next ilCol
    For ilCol = TITLE1INDEX To TITLE2INDEX Step 1
        grdLibEvents.TextMatrix(0, ilCol) = "Title"
    Next ilCol
    grdLibEvents.TextMatrix(0, HIGHLIGHTINDEX) = "»"
    grdLibEvents.TextMatrix(1, BUSNAMEINDEX) = "Name"
    grdLibEvents.TextMatrix(1, BUSCTRLINDEX) = "C"
    grdLibEvents.TextMatrix(0, EVENTTYPEINDEX) = "Event"
    grdLibEvents.TextMatrix(1, EVENTTYPEINDEX) = "Type"
    grdLibEvents.TextMatrix(0, EVENTIDINDEX) = "Event"
    grdLibEvents.TextMatrix(1, EVENTIDINDEX) = "ID"
    grdLibEvents.TextMatrix(0, TIMEINDEX) = "Time"
    grdLibEvents.TextMatrix(1, TIMEINDEX) = ""
    grdLibEvents.TextMatrix(0, STARTTYPEINDEX) = "Start "
    grdLibEvents.TextMatrix(1, STARTTYPEINDEX) = "Type"
    grdLibEvents.TextMatrix(0, FIXEDINDEX) = "Fix"
    grdLibEvents.TextMatrix(0, ENDTYPEINDEX) = "End"
    grdLibEvents.TextMatrix(1, ENDTYPEINDEX) = "Type"
    grdLibEvents.TextMatrix(0, DURATIONINDEX) = "Duration"
    grdLibEvents.TextMatrix(0, MATERIALINDEX) = "Mat"
    grdLibEvents.TextMatrix(1, MATERIALINDEX) = "Type"
    grdLibEvents.TextMatrix(1, AUDIONAMEINDEX) = "Name"
    grdLibEvents.TextMatrix(1, AUDIOITEMIDINDEX) = "Item"
    grdLibEvents.TextMatrix(1, AUDIOISCIINDEX) = "ISCI"
    grdLibEvents.TextMatrix(1, AUDIOCTRLINDEX) = "C"
    grdLibEvents.TextMatrix(1, BACKUPNAMEINDEX) = "Name"
    grdLibEvents.TextMatrix(1, BACKUPCTRLINDEX) = "C"
    grdLibEvents.TextMatrix(1, PROTNAMEINDEX) = "Name"
    grdLibEvents.TextMatrix(1, PROTITEMIDINDEX) = "Item"
    grdLibEvents.TextMatrix(1, PROTISCIINDEX) = "ISCI"
    grdLibEvents.TextMatrix(1, PROTCTRLINDEX) = "C"
    grdLibEvents.TextMatrix(1, RELAY1INDEX) = "1"
    grdLibEvents.TextMatrix(1, RELAY2INDEX) = "2"
    grdLibEvents.TextMatrix(0, FOLLOWINDEX) = "Fol-"
    grdLibEvents.TextMatrix(1, FOLLOWINDEX) = "low"
    grdLibEvents.TextMatrix(1, SILENCETIMEINDEX) = "Time"
    grdLibEvents.TextMatrix(1, SILENCE1INDEX) = "1"
    grdLibEvents.TextMatrix(1, SILENCE2INDEX) = "2"
    grdLibEvents.TextMatrix(1, SILENCE3INDEX) = "3"
    grdLibEvents.TextMatrix(1, SILENCE4INDEX) = "4"
    grdLibEvents.TextMatrix(1, NETCUE1INDEX) = "Start"
    grdLibEvents.TextMatrix(1, NETCUE2INDEX) = "Stop"
    grdLibEvents.TextMatrix(1, TITLE1INDEX) = "1"
    grdLibEvents.TextMatrix(1, TITLE2INDEX) = "2"
    grdLibEvents.TextMatrix(0, ABCFORMATINDEX) = "For-"
    grdLibEvents.TextMatrix(1, ABCFORMATINDEX) = "mat"
    grdLibEvents.TextMatrix(0, ABCPGMCODEINDEX) = "Pgm"
    grdLibEvents.TextMatrix(1, ABCPGMCODEINDEX) = "Code"
    grdLibEvents.TextMatrix(0, ABCXDSMODEINDEX) = "XDS"
    grdLibEvents.TextMatrix(1, ABCXDSMODEINDEX) = "Mode"
    grdLibEvents.TextMatrix(0, ABCRECORDITEMINDEX) = "Rec'd"
    grdLibEvents.TextMatrix(1, ABCRECORDITEMINDEX) = "Item"
    
    grdLibEvents.Row = 1
    For ilCol = 0 To grdLibEvents.Cols - 1 Step 1
        grdLibEvents.Col = ilCol
        grdLibEvents.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdLibEvents.Row = 0
    grdLibEvents.MergeCells = flexMergeRestrictRows
    grdLibEvents.MergeRow(0) = True
    grdLibEvents.Row = 0
    grdLibEvents.Col = EVENTTYPEINDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = BUSNAMEINDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Row = 0
    grdLibEvents.Col = AUDIONAMEINDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = BACKUPNAMEINDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = PROTNAMEINDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = RELAY1INDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = SILENCETIMEINDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = NETCUE1INDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    grdLibEvents.Row = 0
    grdLibEvents.Col = TITLE1INDEX
    grdLibEvents.CellAlignment = flexAlignCenterCenter
    'grdLibEvents.Height = cmcCancel.Top - grdLibEvents.Top - 240    '4 * grdLibEvents.RowHeight(0) + 15
    'gGrid_IntegralHeight grdLibEvents
    'gGrid_Clear grdLibEvents, True
    
    gGrid_AlignAllColsLeft grdConflicts
    mGridColumnWidth
    'Set Titles
    grdConflicts.TextMatrix(0, CONFLICTNAMEINDEX) = "Name"
    grdConflicts.TextMatrix(0, CONFLICTSUBNAMEINDEX) = "Subname"
    grdConflicts.TextMatrix(0, CONFLICTSTARTDATEINDEX) = "Start"
    grdConflicts.TextMatrix(0, CONFLICTENDDATEINDEX) = "End"
    grdConflicts.TextMatrix(0, CONFLICTDAYSINDEX) = "Days"
    grdConflicts.TextMatrix(0, CONFLICTOFFSETINDEX) = "Event ID" '"Offset/Event ID"
    grdConflicts.TextMatrix(0, CONFLICTHOURSINDEX) = "Times"    '"Hours"
    grdConflicts.TextMatrix(0, CONFLICTDURATIONINDEX) = "Duration"
    grdConflicts.TextMatrix(0, CONFLICTBUSESINDEX) = "Buses"
    grdConflicts.TextMatrix(0, CONFLICTAUDIOINDEX) = "Audio"
    grdConflicts.TextMatrix(0, CONFLICTBACKUPINDEX) = "Backup"
    grdConflicts.TextMatrix(0, CONFLICTPROTINDEX) = "Protection"
    grdConflicts.Row = 1
    For ilCol = 0 To grdConflicts.Cols - 1 Step 1
        grdConflicts.Col = ilCol
        grdConflicts.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdConflicts.Height = 4 * grdConflicts.RowHeight(0) + 30    '15
    gGrid_IntegralHeight grdConflicts
    gGrid_Clear grdConflicts, True
    grdConflicts.Move grdLibEvents.Left, grdLibEvents.Top + grdLibEvents.Height - grdConflicts.Height + 60

End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    Dim ilPass As Integer
    
    
    grdLibEvents.ColWidth(PCODEINDEX) = 0
    grdLibEvents.ColWidth(SORTTIMEINDEX) = 0
    grdLibEvents.ColWidth(LIBNAMEINDEX) = 0
    grdLibEvents.ColWidth(TMCURRSEEINDEX) = 0
    grdLibEvents.ColWidth(SPOTAVAILTIMEINDEX) = 0
    grdLibEvents.ColWidth(AVAILDURATIONINDEX) = 0
    grdLibEvents.ColWidth(ERRORCONFLICTINDEX) = 0
    grdLibEvents.ColWidth(CHGSTATUSINDEX) = 0
    grdLibEvents.ColWidth(EVTCONFLICTINDEX) = 0
    grdLibEvents.ColWidth(DEECODEINDEX) = 0
    grdLibEvents.ColWidth(ERRORFIELDSORTINDEX) = 0
    imUnusedCount = 0
    fmUsedWidth = 0
    fmUnusedWidth = 0
    grdLibEvents.ColWidth(HIGHLIGHTINDEX) = (3 * pbcHighlight.TextWidth("»")) / 2
    For ilPass = 0 To 1 Step 1
'        grdLibEvents.ColWidth(EVENTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(EVENTTYPEINDEX), 23, "Y")
'        grdLibEvents.ColWidth(BUSNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSNAMEINDEX), 26, tgSchUsedSumEPE.sBus)
'        grdLibEvents.ColWidth(BUSCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSCTRLINDEX), 57, tgSchUsedSumEPE.sBusControl)
'        grdLibEvents.ColWidth(TIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TIMEINDEX), 17, tgSchUsedSumEPE.sTime)  '21
'        grdLibEvents.ColWidth(STARTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(STARTTYPEINDEX), 29, tgSchUsedSumEPE.sStartType)
'        grdLibEvents.ColWidth(FIXEDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FIXEDINDEX), 38, tgSchUsedSumEPE.sFixedTime)
'        grdLibEvents.ColWidth(ENDTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(ENDTYPEINDEX), 29, tgSchUsedSumEPE.sEndType)
'        grdLibEvents.ColWidth(DURATIONINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(DURATIONINDEX), 17, tgSchUsedSumEPE.sDuration)  '25
'        grdLibEvents.ColWidth(MATERIALINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(MATERIALINDEX), 29, tgSchUsedSumEPE.sMaterialType)
'        grdLibEvents.ColWidth(AUDIONAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIONAMEINDEX), 23, tgSchUsedSumEPE.sAudioName)
'        grdLibEvents.ColWidth(AUDIOITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOITEMIDINDEX), 24, tgSchUsedSumEPE.sAudioItemID)
'        grdLibEvents.ColWidth(AUDIOCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOCTRLINDEX), 58, tgSchUsedSumEPE.sAudioControl)
'        grdLibEvents.ColWidth(BACKUPNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPNAMEINDEX), 23, tgSchUsedSumEPE.sBkupAudioName)
'        grdLibEvents.ColWidth(BACKUPCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPCTRLINDEX), 58, tgSchUsedSumEPE.sBkupAudioControl)
'        grdLibEvents.ColWidth(PROTNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTNAMEINDEX), 23, tgSchUsedSumEPE.sProtAudioName)
'        grdLibEvents.ColWidth(PROTITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTITEMIDINDEX), 24, tgSchUsedSumEPE.sProtAudioItemID)
'        grdLibEvents.ColWidth(PROTCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTCTRLINDEX), 58, tgSchUsedSumEPE.sProtAudioControl)
'        grdLibEvents.ColWidth(RELAY1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY1INDEX), 30, tgSchUsedSumEPE.sRelay1)
'        grdLibEvents.ColWidth(RELAY2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY2INDEX), 30, tgSchUsedSumEPE.sRelay2)
'        grdLibEvents.ColWidth(FOLLOWINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FOLLOWINDEX), 35, tgSchUsedSumEPE.sFollow)
'        grdLibEvents.ColWidth(SILENCETIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCETIMEINDEX), 25, tgSchUsedSumEPE.sSilenceTime)
'        grdLibEvents.ColWidth(SILENCE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE1INDEX), 58, tgSchUsedSumEPE.sSilence1)
'        grdLibEvents.ColWidth(SILENCE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE2INDEX), 58, tgSchUsedSumEPE.sSilence2)
'        grdLibEvents.ColWidth(SILENCE3INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE3INDEX), 58, tgSchUsedSumEPE.sSilence3)
'        grdLibEvents.ColWidth(SILENCE4INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE4INDEX), 58, tgSchUsedSumEPE.sSilence4)
'        grdLibEvents.ColWidth(NETCUE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE1INDEX), 31, tgSchUsedSumEPE.sStartNetcue)
'        grdLibEvents.ColWidth(NETCUE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE2INDEX), 31, tgSchUsedSumEPE.sStopNetcue)
'        grdLibEvents.ColWidth(TITLE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE1INDEX), 53, tgSchUsedSumEPE.sTitle1)
'        grdLibEvents.ColWidth(TITLE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE2INDEX), 53, tgSchUsedSumEPE.sTitle2)
        grdLibEvents.ColWidth(EVENTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(EVENTTYPEINDEX), 65, "Y")
        grdLibEvents.ColWidth(EVENTIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(EVENTIDINDEX), 30, "Y")
        grdLibEvents.ColWidth(BUSNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSNAMEINDEX), 32, tgSchUsedSumEPE.sBus)
        grdLibEvents.ColWidth(BUSCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BUSCTRLINDEX), 65, tgSchUsedSumEPE.sBusControl)
        grdLibEvents.ColWidth(TIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TIMEINDEX), 24, tgSchUsedSumEPE.sTime)  '21
        grdLibEvents.ColWidth(STARTTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(STARTTYPEINDEX), 40, tgSchUsedSumEPE.sStartType)
        grdLibEvents.ColWidth(FIXEDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FIXEDINDEX), 50, tgSchUsedSumEPE.sFixedTime)
        grdLibEvents.ColWidth(ENDTYPEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(ENDTYPEINDEX), 40, tgSchUsedSumEPE.sEndType)
        grdLibEvents.ColWidth(DURATIONINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(DURATIONINDEX), 24, tgSchUsedSumEPE.sDuration)  '25
        grdLibEvents.ColWidth(MATERIALINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(MATERIALINDEX), 40, tgSchUsedSumEPE.sMaterialType)
        grdLibEvents.ColWidth(AUDIONAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIONAMEINDEX), 27, tgSchUsedSumEPE.sAudioName)
        grdLibEvents.ColWidth(AUDIOITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOITEMIDINDEX), 27, tgSchUsedSumEPE.sAudioItemID)
        grdLibEvents.ColWidth(AUDIOISCIINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOISCIINDEX), 27, tgSchUsedSumEPE.sAudioISCI)
        grdLibEvents.ColWidth(AUDIOCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(AUDIOCTRLINDEX), 65, tgSchUsedSumEPE.sAudioControl)
        grdLibEvents.ColWidth(BACKUPNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPNAMEINDEX), 27, tgSchUsedSumEPE.sBkupAudioName)
        grdLibEvents.ColWidth(BACKUPCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(BACKUPCTRLINDEX), 65, tgSchUsedSumEPE.sBkupAudioControl)
        grdLibEvents.ColWidth(PROTNAMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTNAMEINDEX), 27, tgSchUsedSumEPE.sProtAudioName)
        grdLibEvents.ColWidth(PROTITEMIDINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTITEMIDINDEX), 27, tgSchUsedSumEPE.sProtAudioItemID)
        grdLibEvents.ColWidth(PROTISCIINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTISCIINDEX), 27, tgSchUsedSumEPE.sProtAudioISCI)
        grdLibEvents.ColWidth(PROTCTRLINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(PROTCTRLINDEX), 65, tgSchUsedSumEPE.sProtAudioControl)
        grdLibEvents.ColWidth(RELAY1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY1INDEX), 32, tgSchUsedSumEPE.sRelay1)
        grdLibEvents.ColWidth(RELAY2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(RELAY2INDEX), 32, tgSchUsedSumEPE.sRelay2)
        grdLibEvents.ColWidth(FOLLOWINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(FOLLOWINDEX), 40, tgSchUsedSumEPE.sFollow)
        grdLibEvents.ColWidth(SILENCETIMEINDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCETIMEINDEX), 30, tgSchUsedSumEPE.sSilenceTime)
        grdLibEvents.ColWidth(SILENCE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE1INDEX), 65, tgSchUsedSumEPE.sSilence1)
        grdLibEvents.ColWidth(SILENCE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE2INDEX), 65, tgSchUsedSumEPE.sSilence2)
        grdLibEvents.ColWidth(SILENCE3INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE3INDEX), 65, tgSchUsedSumEPE.sSilence3)
        grdLibEvents.ColWidth(SILENCE4INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(SILENCE4INDEX), 65, tgSchUsedSumEPE.sSilence4)
        grdLibEvents.ColWidth(NETCUE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE1INDEX), 40, tgSchUsedSumEPE.sStartNetcue)
        grdLibEvents.ColWidth(NETCUE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(NETCUE2INDEX), 40, tgSchUsedSumEPE.sStopNetcue)
        'grdLibEvents.ColWidth(TITLE1INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE1INDEX), 53, tgSchUsedSumEPE.sTitle1)
        'grdLibEvents.ColWidth(TITLE2INDEX) = mComputeWidth(ilPass, grdLibEvents.ColWidth(TITLE2INDEX), 53, tgSchUsedSumEPE.sTitle2)
        If sgClientFields = "A" Then
            If tgSchUsedSumEPE.sABCFormat <> "Y" Then
                grdLibEvents.ColWidth(ABCFORMATINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCFORMATINDEX) = grdLibEvents.Width / 28
            End If
            If tgSchUsedSumEPE.sABCPgmCode <> "Y" Then
                grdLibEvents.ColWidth(ABCPGMCODEINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCPGMCODEINDEX) = grdLibEvents.Width / 28
            End If
            If tgSchUsedSumEPE.sABCXDSMode <> "Y" Then
                grdLibEvents.ColWidth(ABCXDSMODEINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCXDSMODEINDEX) = grdLibEvents.Width / 28
            End If
            If tgSchUsedSumEPE.sABCRecordItem <> "Y" Then
                grdLibEvents.ColWidth(ABCRECORDITEMINDEX) = 0
            Else
                grdLibEvents.ColWidth(ABCRECORDITEMINDEX) = grdLibEvents.Width / 28
            End If
        Else
            grdLibEvents.ColWidth(ABCFORMATINDEX) = 0
            grdLibEvents.ColWidth(ABCPGMCODEINDEX) = 0
            grdLibEvents.ColWidth(ABCXDSMODEINDEX) = 0
            grdLibEvents.ColWidth(ABCRECORDITEMINDEX) = 0
        End If
        If imUnusedCount = 0 Then
            Exit For
        End If
    Next ilPass
    
    grdLibEvents.ColWidth(TITLE1INDEX) = grdLibEvents.Width - GRIDSCROLLWIDTH
    'For ilCol = EVENTTYPEINDEX To TITLE2INDEX Step 1
    For ilCol = HIGHLIGHTINDEX To TITLE2INDEX Step 1
        If ilCol <> TITLE1INDEX And ilCol <> TITLE2INDEX Then
            If grdLibEvents.ColWidth(TITLE1INDEX) > grdLibEvents.ColWidth(ilCol) Then
                grdLibEvents.ColWidth(TITLE1INDEX) = grdLibEvents.ColWidth(TITLE1INDEX) - grdLibEvents.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    grdLibEvents.ColWidth(TITLE2INDEX) = grdLibEvents.ColWidth(TITLE1INDEX) / 3
    grdLibEvents.ColWidth(TITLE1INDEX) = grdLibEvents.ColWidth(TITLE1INDEX) - grdLibEvents.ColWidth(TITLE2INDEX)
    '8/26/11: Move here
    gGrid_IntegralHeight grdLibEvents

    grdConflicts.Width = EngrSchdDef.Width - lacScreen.Left - lacScreen.Width
    imUnusedCount = 0
    fmUsedWidth = 0
    fmUnusedWidth = 0
    For ilPass = 0 To 1 Step 1
        'grdConflicts.ColWidth(CONFLICTNAMEINDEX) = grdConflicts.Width / 10
        'grdConflicts.ColWidth(CONFLICTSUBNAMEINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTSUBNAMEINDEX), 10, tgSchUsedSumEPE.sBus)
        grdConflicts.ColWidth(CONFLICTSTARTDATEINDEX) = 0   'mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTSTARTDATEINDEX), 18, "Y")
        grdConflicts.ColWidth(CONFLICTENDDATEINDEX) = 0 'mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTENDDATEINDEX), 18, "Y")
        grdConflicts.ColWidth(CONFLICTDAYSINDEX) = 0    'mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTDAYSINDEX), 14, "Y")
        grdConflicts.ColWidth(CONFLICTOFFSETINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTOFFSETINDEX), 10, tgSchUsedSumEPE.sTime)
        grdConflicts.ColWidth(CONFLICTHOURSINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTHOURSINDEX), 12, "Y")
        grdConflicts.ColWidth(CONFLICTDURATIONINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTDURATIONINDEX), 12, tgSchUsedSumEPE.sDuration)
        grdConflicts.ColWidth(CONFLICTBUSESINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTBUSESINDEX), 12, tgSchUsedSumEPE.sBus)
        grdConflicts.ColWidth(CONFLICTAUDIOINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTAUDIOINDEX), 20, tgSchUsedSumEPE.sAudioName)
        grdConflicts.ColWidth(CONFLICTBACKUPINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTBACKUPINDEX), 20, tgSchUsedSumEPE.sBkupAudioName)
        grdConflicts.ColWidth(CONFLICTPROTINDEX) = mComputeWidth(ilPass, grdConflicts.ColWidth(CONFLICTPROTINDEX), 20, tgSchUsedSumEPE.sProtAudioName)
        If imUnusedCount = 0 Then
            Exit For
        End If
    Next ilPass
    grdConflicts.ColWidth(CONFLICTNAMEINDEX) = grdConflicts.Width - GRIDSCROLLWIDTH
    For ilCol = CONFLICTNAMEINDEX To CONFLICTPROTINDEX Step 1
        If (ilCol <> CONFLICTNAMEINDEX) And (ilCol <> CONFLICTSUBNAMEINDEX) Then
            If grdConflicts.ColWidth(CONFLICTNAMEINDEX) > grdConflicts.ColWidth(ilCol) Then
                grdConflicts.ColWidth(CONFLICTNAMEINDEX) = grdConflicts.ColWidth(CONFLICTNAMEINDEX) - grdConflicts.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
    grdConflicts.ColWidth(CONFLICTSUBNAMEINDEX) = grdConflicts.ColWidth(CONFLICTNAMEINDEX) / 3
    grdConflicts.ColWidth(CONFLICTNAMEINDEX) = grdConflicts.ColWidth(CONFLICTNAMEINDEX) - grdConflicts.ColWidth(CONFLICTSUBNAMEINDEX)

End Sub


Private Sub mClearControls()
    gGrid_Clear grdLibEvents, True
    grdLibEvents.BackColor = vbWhite
    'Can't be 0 to 0 because index stored into grid
    Dim lmDeleteCodes(0 To 0) As Long
End Sub

Private Sub mPopulate()
'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    Dim llRow As Long
'
'
'    ilRet = gGetRec_DHE_DayHeaderInfo(lgLibCallCode, "EngrSchdDef-mPopulation", tmDHE)
'    ilRet = gGetRecs_DEE_DayEvent(sgCurrDEEStamp, lgLibCallCode, "EngrSchdDef-mPopulate", tgCurrDEE())
'    If lgLibCallCode <= 0 Then
'        tmDHE.lCode = 0
'    End If
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
'    Dim ilIndex As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim ilNew As Integer
    Dim llOldSHECode As Long
    Dim llOldSEECode As Long
    Dim ilSEECompare As Integer
'    Dim slStartDate As String
'    Dim slEndDate As String
'    Dim llStartDate As Long
'    Dim llEndDate As Long
    Dim llNewAgedDHECode As Long
'    Dim ilNameOk As Integer
'    Dim tlDHE As DHE
'    Dim tlSvDHE As DHE
    Dim tlSHE As SHE
    Dim tlSEE As SEE
    Dim ilLibRet As Integer
    'ilEvtRet Placed at module level
    'Dim ilEvtRet As Integer
    Dim llLoop As Long
    Dim ilCount As Integer
    Dim llCTE As Long
    Dim llSEEOld As Long
    
    bmInSave = True
    'If (UBound(tgFilterValues) > LBound(tgFilterValues)) And (tmSHE.lCode = 0) Then
    '    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    '    'MsgBox "Times/Buses/Audio in Conflict with other Schedules or Libraries/Templates", vbCritical + vbOKOnly, "Schedule"
    '    MsgBox "Filter must be Removed before Saving", vbCritical, "Save Disallowed"
    '    mSave = False
    '    Exit Function
    'End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass

    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())

    'Sort moved here since Filter must be removed prior to saving.
    'Old logic first save changed records, then reloaded the grid
    
    ReDim tmCCurrSEE(0 To UBound(tgCurrSEE)) As SEE
    For llLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
        tmCCurrSEE(llLoop) = tgCurrSEE(llLoop)
    Next llLoop
    gAutoSortTime tmCCurrSEE
    
    ReDim tmPCurrSEE(0 To 0) As SEE
    ReDim tmNCurrSEE(0 To 0) As SEE
    ReDim tmSeeBracket(0 To 0) As SEEBRACKET
    ReDim lmChgSEE(0 To 0) As Long
    
    grdLibEvents.Redraw = False
    If tmSHE.lCode = 0 Then
        mMoveSEECtrlsToRec
        If (UBound(tgFilterValues) > LBound(tgFilterValues)) Then
            cmcTask.Caption = "Removing Filter...."
            cmcTask.Visible = True
            'Remove filter
            ReDim tgFilterValues(0 To 0) As FILTERVALUES
            mReorderFilter
            mMoveSEERecToCtrls
            cmcTask.Visible = False
        End If
        grdLibEvents.Redraw = False
        imLastColSorted = -1
        mSortCol TIMEINDEX
    Else
        'Conflict check requires that the rows be in time ascending order
        imLastColSorted = -1
        mSortCol TIMEINDEX
        mMoveSEECtrlsToRec
        ''Remove filter
        'ReDim tmFilterValues(0 To 0) As FILTERVALUES
        ''mMoveSEERecToCtrls
        grdLibEvents.Redraw = False
        ''imLastColSorted = -1
        ''mSortCol TIMEINDEX
    End If
    ReDim tmConflictList(1 To 1) As CONFLICTLIST
    tmConflictList(UBound(tmConflictList)).iNextIndex = -1
    mInitConflictTest
    If Not mCheckFields(True) Then
        'ilLibRet = mCheckLibConflicts()
        'ilEvtRet = mSvCheckEventConflicts()
        ilLibRet = False
        '9/13/11: Remove auto conflict testing within Save.  Retain the Conflict Button.
        '         ilEvtRet Placed at module level
        If tmSHE.lCode = 0 Then
            imEvtRet = mCheckEventConflicts()
        End If
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        If (Not ilLibRet) And (Not imEvtRet) Then
            MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Schedule"
        Else
            MsgBox "One or more required fields are missing or defined incorrectly and Times/Buses/Audio in Conflict", vbCritical + vbOKOnly, "Schedule"
        End If
        mSortErrorsToTop
        mSave = False
        Exit Function
    End If
    'ilLibRet = mCheckLibConflicts()
    'ilEvtRet = mSvCheckEventConflicts()
    ilLibRet = False
    '9/13/11: Remove auto conflict testing within Save.  Retain the Conflict Button.
    '         ilEvtRet placed at module level
    If tmSHE.lCode = 0 Then
        imEvtRet = mCheckEventConflicts()
    End If
    If ilLibRet Then
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        'MsgBox "Times/Buses/Audio in Conflict with other Schedules or Libraries/Templates", vbCritical + vbOKOnly, "Schedule"
        ilRet = MsgBox("Times/Buses/Audio in Conflict with other Schedules or Libraries/Templates, Continue with Save?", vbQuestion + vbYesNo, "Conflicts")
        If ilRet = vbNo Then
            mSortErrorsToTop
            mSave = False
            Exit Function
        End If
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    End If
    If (imEvtRet) And (tmSHE.lCode = 0) Then
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        'MsgBox "Times/Buses/Audio in Conflict within this Schedule", vbCritical + vbOKOnly, "Schedule"
        ilRet = MsgBox("Times/Buses/Audio in Conflict within this Schedule, Continue with Save?", vbQuestion + vbYesNo, "Conflicts")
        If ilRet = vbNo Then
            mSortErrorsToTop
            mSave = False
            Exit Function
        End If
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    End If
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If tmSHE.lCode <= 0 Then
        gCreateHeader smAirDate, tmSHE
        If imEvtRet Then
            tmSHE.sConflictExist = "Y"
        Else
            tmSHE.sConflictExist = "N"
        End If
        If imMergeError = 2 Then
            tmSHE.sSpotMergeStatus = "E"
        ElseIf imMergeError = 1 Then
            tmSHE.sSpotMergeStatus = "M"
        End If
        ilNew = True
        ilRet = gPutInsert_SHE_ScheduleHeader(0, tmSHE, "Schedule Definition-mSave: SHE")
    Else
        ilNew = False
        ilRet = gGetRec_SHE_ScheduleHeader(tmSHE.lCode, "EngrSchedule-Get Schedule to Check Load", tlSHE)
        If ilRet Then
            If tlSHE.iVersion > tmSHE.iVersion Then
                tmSHE.iVersion = tlSHE.iVersion + 1
            End If
            If tlSHE.iChgSeqNo > tmSHE.iChgSeqNo Then
                tmSHE.iChgSeqNo = tlSHE.iChgSeqNo
            End If
            If tlSHE.sLoadedAutoStatus = "L" Then
                tmSHE.sLoadedAutoStatus = "L"
                tmSHE.sLoadedAutoDate = tlSHE.sLoadedAutoDate
            End If
        Else
            tmSHE.iVersion = tmSHE.iVersion + 1
        End If
        If imEvtRet Then
            tmSHE.sConflictExist = "Y"
        Else
            tmSHE.sConflictExist = "N"
        End If
        If imMergeError = 2 Then
            tmSHE.sSpotMergeStatus = "E"
        ElseIf imMergeError = 1 Then
            tmSHE.sSpotMergeStatus = "M"
        End If
        If tmSHE.sLoadedAutoStatus = "L" Then
            If mBusInFilter() Then
                If Not mOkToGenUPD() Then
                    mSave = False
                    Exit Function
                End If
                'Get prev and next
                ilRet = gGetRec_SHE_ScheduleHeaderByDate(DateAdd("d", -1, smAirDate), "EngrSchedule-Get Schedule by Date", tmPSHE)
                If ilRet Then
                    ilRet = gGetRecs_SEE_ScheduleEventsAPIWithFilter(hmSEE, smPCurrSEEStamp, -1, tmPSHE.lCode, "EngrSchdDef-Get Events", tmPCurrSEE())
                    gAutoSortTime tmPCurrSEE
                End If
                ilRet = gGetRec_SHE_ScheduleHeaderByDate(DateAdd("d", 1, smAirDate), "EngrSchedule-Get Schedule by Date", tmNSHE)
                If ilRet Then
                    ilRet = gGetRecs_SEE_ScheduleEventsAPIWithFilter(hmSEE, smNCurrSEEStamp, -1, tmNSHE.lCode, "EngrSchdDef-Get Events", tmNCurrSEE())
                    gAutoSortTime tmNCurrSEE
                End If
            End If
        End If
        ilRet = gPutUpdate_SHE_ScheduleHeader(3, tmSHE, "Schedule Definition-mSave: Update SHE", llOldSHECode)
    End If
    For llLoop = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llLoop, EVENTTYPEINDEX)) <> "" Then
            If (grdLibEvents.TextMatrix(llLoop, CHGSTATUSINDEX) = "Y") Or (ilNew) Then
                llRow = Val(grdLibEvents.TextMatrix(llLoop, TMCURRSEEINDEX))
                llOldSEECode = tmCurrSEE(llRow).lCode
                
                If tmCurrSEE(llRow).lCode > 0 Then
                    ilSEECompare = mCompareSEE(llRow, llOldSEECode, Trim$(smT1Comment(llRow)), Trim$(smT2Comment(llRow)))
                    If ilSEECompare Then
                        ilSave = False
                    Else
                        ilSave = True
                    End If
                Else
                    ilSave = True
                End If
                If ilSave Then
                    tmCurrSEE(llRow).l1CteCode = 0
                    If (tmCurrSEE(llRow).iEteCode <> imSpotETECode) Then
                        If Trim$(smT1Comment(llRow)) <> "" Then
                            For llCTE = 0 To UBound(tmCurr1CTE_Name) - 1 Step 1
                                If (tmCurrSEE(llRow).lDheCode = tmCurr1CTE_Name(llCTE).lDheCode) Then
                                    If StrComp(UCase(Trim$(tmCurr1CTE_Name(llCTE).sComment)), UCase(Trim$(smT1Comment(llRow))), vbBinaryCompare) = 0 Then
                                        tmCurrSEE(llRow).l1CteCode = tmCurr1CTE_Name(llCTE).lCteCode
                                        Exit For
                                    End If
                                End If
                            Next llCTE
                            If tmCurrSEE(llRow).l1CteCode = 0 Then
                                gSetCTE smT1Comment(llRow), "T1", tmCTE
                                ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Schedule Definition-mSave: Insert CTE", hmCTE)
                                If ilRet Then
                                    tmCurrSEE(llRow).l1CteCode = tmCTE.lCode
                                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).sComment = tmCTE.sComment
                                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lCteCode = tmCTE.lCode
                                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lDeeCode = tmCurrSEE(llRow).lDeeCode
                                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lDheCode = tmCurrSEE(llRow).lDheCode
                                    ReDim Preserve tmCurr1CTE_Name(0 To UBound(tmCurr1CTE_Name) + 1) As DEECTE
                                End If
                            End If
                        End If
                    End If
                    '7/8/11: Make T2 work like T1
                    tmCurrSEE(llRow).l2CteCode = 0
                    'If (tmCurrSEE(llRow).iEteCode <> imSpotETECode) Then
                        If Trim$(smT2Comment(llRow)) <> "" Then
                            For llCTE = 0 To UBound(tmCurr2CTE_Name) - 1 Step 1
                                If (tmCurrSEE(llRow).lDheCode = tmCurr2CTE_Name(llCTE).lDheCode) Then
                                    If StrComp(UCase(Trim$(tmCurr2CTE_Name(llCTE).sComment)), UCase(Trim$(smT2Comment(llRow))), vbBinaryCompare) = 0 Then
                                        tmCurrSEE(llRow).l2CteCode = tmCurr2CTE_Name(llCTE).lCteCode
                                        Exit For
                                    End If
                                End If
                            Next llCTE
                            If tmCurrSEE(llRow).l2CteCode = 0 Then
                                gSetCTE smT2Comment(llRow), "T2", tmCTE
                                ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Schedule Definition-mSave: Insert CTE", hmCTE)
                                If ilRet Then
                                    tmCurrSEE(llRow).l2CteCode = tmCTE.lCode
                                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).sComment = tmCTE.sComment
                                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lCteCode = tmCTE.lCode
                                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lDeeCode = tmCurrSEE(llRow).lDeeCode
                                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lDheCode = tmCurrSEE(llRow).lDheCode
                                    ReDim Preserve tmCurr2CTE_Name(0 To UBound(tmCurr2CTE_Name) + 1) As DEECTE
                                End If
                            End If
                        End If
                    'End If
                    
                    tmCurrSEE(llRow).lCode = 0
                    tmCurrSEE(llRow).lSheCode = tmSHE.lCode
                    tmCurrSEE(llRow).sInsertFlag = "N"
                    tmCurrSEE(llRow).sUnused = ""
                    If llOldSEECode > 0 Then
                        tmCurrSEE(llRow).sAction = "C"
                        tmCurrSEE(llRow).sSentStatus = "N"
                        tmCurrSEE(llRow).sSentDate = Format$("12/31/2069", sgShowDateForm)
                    Else
                        tmCurrSEE(llRow).sAction = "N"
                        tmCurrSEE(llRow).sSentStatus = "N"
                        tmCurrSEE(llRow).sSentDate = Format$("12/31/2069", sgShowDateForm)
                    End If
                    ilRet = gPutInsert_SEE_ScheduleEvents(tmCurrSEE(llRow), "Schedule Definition-mSave: SEE", hmSEE, hmSOE)
                    If llOldSEECode > 0 Then
                        ilRet = gPutReplace_SEE_SHECode(llOldSEECode, llOldSHECode, "Schedule Replace-mSave: SEE")
                        ilRet = gUpdateAIE(1, tmSHE.iVersion, "SEE", llOldSEECode, tmCurrSEE(llRow).lCode, tmSHE.lOrigSheCode, "Schedule Definition- mSave: Insert SEE:AIE")
                        gSetUsedFlags tmCurrSEE(llRow), hmCTE
                        If tmCurrSEE(llRow).iEteCode <> imSpotETECode Then
                            ilRet = gPutDelete_CME_Conflict_Master("S", tmSHE.lCode, llOldSEECode, 0, "Schedule Definition-mSave: Delete SEE in CME", hmCME)
                            ilRet = gCreateCMEForSchd(tmSHE, tmCurrSEE(llRow), imSpotETECode, hmCME)
                        End If
                    Else
                        gSetUsedFlags tmCurrSEE(llRow), hmCTE
                        If tmCurrSEE(llRow).iEteCode <> imSpotETECode Then
                            ilRet = gCreateCMEForSchd(tmSHE, tmCurrSEE(llRow), imSpotETECode, hmCME)
                        End If
                    End If
                    If tmSHE.sLoadedAutoStatus = "L" Then
                        lmChgSEE(UBound(lmChgSEE)) = llRow
                        ReDim Preserve lmChgSEE(0 To UBound(lmChgSEE) + 1) As Long
                    End If
                End If
            Else
                If bmMerged Then
                    llRow = Val(grdLibEvents.TextMatrix(llLoop, TMCURRSEEINDEX))
                    '9/5/14: Fix subscript error
                    'If (tgCurrSEE(llRow).iEteCode = imSpotETECode) Then
                    If (tmCurrSEE(llRow).iEteCode = imSpotETECode) Then
                        llSEEOld = mBinarySearchOldSEE(tmCurrSEE(llRow).lCode)
                        If llSEEOld <> -1 Then
                            tgCurrSEE(llSEEOld).lCode = -tgCurrSEE(llSEEOld).lCode
                        End If
                    End If
                End If
            End If
        End If
    Next llLoop
    ilRet = gPutUpdate_SOE_SiteOption(1, tgSOE, "gPutInsert_See: Update EventID in SOE")
    For llRow = 0 To UBound(lmDeleteCodes) - 1 Step 1
        For ilLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
            If Abs(lmDeleteCodes(llRow)) = tgCurrSEE(ilLoop).lCode Then
                If lmDeleteCodes(llRow) >= 0 Then
                    mFindBrackets CLng(ilLoop)
                End If
                ilRet = gPutDelete_CME_Conflict_Master("S", tmSHE.lCode, tgCurrSEE(ilLoop).lCode, 0, "Schedule Definition- mSave: Delete SEE in CME", hmCME)
                '2/14/12: UPD file now created within the save
                'If tgCurrSEE(ilLoop).sSentStatus = "S" Then
                '    LSet tlSEE = tgCurrSEE(ilLoop)
                '    tlSEE.lCode = 0
                '    tlSEE.lSheCode = tmSHE.lCode
                '    tlSEE.sAction = "D"
                '    tlSEE.sSentStatus = "N"
                '    tlSEE.sSentDate = Format$("12/31/2069", sgShowDateForm)
                '    ilRet = gPutInsert_SEE_ScheduleEvents(tlSEE, "Schedule Definition-mSave: SEE", hmSEE, hmSOE)
                'Else
                '    'Delete record as it has not been sent
                '    ilRet = gPutDelete_SEE_Schedule_Events(lmDeleteCodes(llRow), "Schedule Definition-mSave: SEE")
                'End If
                ilRet = gPutDelete_SEE_Schedule_Events(Abs(lmDeleteCodes(llRow)), "Schedule Definition-mSave: SEE")
                Exit For
            End If
        Next ilLoop
    Next llRow
    '5/16/13: Handle case where spots removed
    'Remove extra spots if merged
    If bmMerged Then
        For llRow = 0 To UBound(tgCurrSEE) - 1 Step 1
            If tgCurrSEE(llRow).lCode > 0 Then
                If (tgCurrSEE(llRow).iEteCode = imSpotETECode) Then
                    ilRet = gPutDelete_CME_Conflict_Master("S", tmSHE.lCode, tgCurrSEE(llRow).lCode, 0, "Schedule Definition- mSave: Delete SEE in CME", hmCME)
                    ilRet = gPutDelete_SEE_Schedule_Events(tgCurrSEE(llRow).lCode, "Schedule Definition-mSave: SEE")
                End If
            End If
        Next llRow
        bmMerged = False
    End If

    For ilLoop = 0 To UBound(lgLibDheUsed) - 1 Step 1
        tmDHE.lCode = lgLibDheUsed(ilLoop)
        ilRet = gPutUpdate_DHE_DayHeaderInfo(2, tmDHE, "Schedule-mSave: Update DHE", llNewAgedDHECode)
    Next ilLoop
    If (tmSHE.sLoadedAutoStatus = "L") Then
        'ilRet = gGetRec_SHE_ScheduleHeader(tmSHE.lCode, "EngrSchedule-Get Schedule to obtain Change Sequence number", tlSHE)
        'tmSHE.sCreateLoad = "Y"
        'tmSHE.iChgSeqNo = tlSHE.iChgSeqNo
        'ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE", llOldSHECode)
        'lmCheckSHECode = tmSHE.lCode
        'tmcCheck.Enabled = True
        'Create upd file
        If Not mGenUPDFile() Then
        
        End If
    End If
    ReDim lgLibDheUsed(0 To 0) As Long
    imFieldChgd = False
    mSetCommands
    If igJobVisible Then
        ilCount = gGetCheckStatus()
        If ilCount > 0 Then
            EngrJob!lacTask(SCHEDULEJOB).ForeColor = vbRed
            EngrMain!imcTask(SCHEDULEJOB).ForeColor = vbRed
        Else
            EngrJob!lacTask(SCHEDULEJOB).ForeColor = vbButtonText
            EngrMain!imcTask(SCHEDULEJOB).ForeColor = vbButtonText
        End If
    End If
    mSave = True
    'If (imEvtRet) And (Not ilNew) Then
    '    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    '    'MsgBox "Times/Buses/Audio in Conflict within this Schedule", vbCritical + vbOKOnly, "Schedule"
    '    ilRet = MsgBox("Times/Buses/Audio in Conflict within this Schedule", vbQuestion + vbOKOnly, "Conflicts")
    '    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    'End If

End Function
Private Sub cmcCancel_Click()
    If bmInSave Then
        Exit Sub
    End If
    igReturnCallStatus = CALLCANCELLED
    Unload EngrSchdDef
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If bmInSave Then
        Exit Sub
    End If
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrSchdDef
        Exit Sub
    End If
    'If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        grdLibEvents.Redraw = False
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        ilRet = mSave()
        bmInSave = False
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        grdLibEvents.Redraw = True
        If Not ilRet Then
            Exit Sub
        End If
    'End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    Unload EngrSchdDef
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcLoad_Click()
    Dim llRow As Long
    Dim slStr As String
    Dim slExportFileName As String
    Dim ilLength As Integer
    Dim slDate As String
    Dim slEventCategory As String
    Dim slEventAutoCode As String
    Dim ilETE As Integer
    Dim llSEECode As Long
    Dim ilRet As Integer
    Dim ilSend As Integer
    Dim llTest As Long
    Dim llIndexRow As Long
    Dim llIndexTest As Long
    Dim ilSEEOld As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llAirDate As Long
    Dim llOldSHECode As Long
    Dim tlSHE As SHE
    
    If bmInSave Then
        Exit Sub
    End If
    If imFieldChgd Then
        If tmSHE.sLoadedAutoStatus = "L" Then
            MsgBox "Changes must be Saved prior to Generating Load for Automation.  Note: Day previously Loaded, then it will be reloaded as part of Save", vbInformation + vbOKOnly, "Schedule"
        Else
            MsgBox "Changes must be Saved prior to Generating Load for Automation.", vbInformation + vbOKOnly, "Schedule"
        End If
        Exit Sub
    End If
    If MsgBox("Generate Automation Export for Day", vbYesNo) = vbNo Then
        Exit Sub
    End If
    If tmSHE.sCreateLoad <> "Y" Then
        ilRet = gGetRec_SHE_ScheduleHeader(tmSHE.lCode, "EngrSchedule-Get Schedule to obtain Change Sequence number", tlSHE)
        tmSHE.sCreateLoad = "Y"
        tmSHE.iChgSeqNo = tlSHE.iChgSeqNo
        ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE", llOldSHECode)
    End If
    tmcCheck.Enabled = True
    lmCheckSHECode = tmSHE.lCode
    MsgBox "Command sent to Service program to load day.", vbInformation + vbOKOnly, "Schedule"
    'Add test that it is working, also add one in save
    Exit Sub
    
    
'    gSetMousePointer grdLibEvents, grdLibEvents, vbHourglass
'    grdLibEvents.Redraw = False
'    mMoveSEECtrlsToRec
'    'Remove filter
'    ReDim tmFilterValues(0 To 0) As FILTERVALUES
'    mMoveSEERecToCtrls
'    imLastColSorted = -1
'    mSortCol TIMEINDEX
'    If Not mOpenAutoExportFile(slExportFileName) Then
'        gSetMousePointer grdLibEvents, grdLibEvents, vbDefault
'        Exit Sub
'    End If
'    llAirDate = gDateValue(smAirDate)
'    slDateTime = gNow()
'    slNowDate = Format(slDateTime, "ddddd")
'    slNowTime = Format(slDateTime, "ttttt")
'    llNowDate = gDateValue(slNowDate)
'    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
'    ilLength = gExportStrLength()
'    If tgNoCharAFE.iDate = 8 Then
'        slDate = Format$(smAirDate, "yyyymmdd")
'    ElseIf tgNoCharAFE.iDate = 6 Then
'        slDate = Format$(smAirDate, "yymmdd")
'    End If
'
'    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
'        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
'        If slStr <> "" Then
'            'slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
'            'slEventCategory = ""
'            'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'            '    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
'            '        slEventCategory = tgCurrETE(ilETE).sCategory
'            '        slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
'            '        Exit For
'            '    End If
'            'Next ilETE
'            'If ((slEventCategory = "P") Or (slEventCategory = "S")) And (mExportRow(llRow)) Then
'            If mExportRow(llRow, slEventCategory, slEventAutoCode) Then
'                If tmSHE.sLoadedAutoStatus = "L" Then
'                    'Send unchanged that encompass changes and and send the change
'                    llIndexRow = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
'                    If tmCurrSEE(llIndexRow).sSentStatus = "S" Then
'                        'Look for next item on same bus, if it requires to be sent, then send this item
'                        For llTest = llRow + 1 To UBound(tmCurrSEE) - 1 Step 1
'                            llIndexTest = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
'                            If tmCurrSEE(llIndexRow).iBdeCode = tmCurrSEE(llIndexTest).iBdeCode Then
'                                If tmCurrSEE(llIndexTest).sSentStatus <> "S" Then
'                                    ilSend = True
'                                Else
'                                    ilSend = False
'                                End If
'                                Exit For
'                            End If
'                        Next llTest
'                        'Look for previous item on the same bus, if it required to be sent, then send this item
'                        If Not ilSend Then
'                            For llTest = llRow - 1 To LBound(tmCurrSEE) Step -1
'                                llIndexTest = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
'                                If tmCurrSEE(llIndexRow).iBdeCode = tmCurrSEE(llIndexTest).iBdeCode Then
'                                    If tmCurrSEE(llIndexTest).sSentStatus <> "S" Then
'                                        ilSend = True
'                                    Else
'                                        ilSend = False
'                                    End If
'                                    Exit For
'                                End If
'                            Next llTest
'                        End If
'                    Else
'                        ilSend = True
'                    End If
'                Else
'                    ilSend = True
'                End If
'            Else
'                ilSend = False
'            End If
'            'Check If today and enough time
'            If ilSend Then
'                'Check If today and enough time
'                llIndexRow = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
'                If llAirDate = llNowDate Then
'                    If llNowTime > tgCurrSEE(llIndexRow).lTime Then
'                        ilSend = False
'                    End If
'                End If
'            End If
'            If ilSend Then
'                smExportStr = String(ilLength, " ")
'                mMakeExportStr tgStartColAFE.iBus, tgNoCharAFE.iBus, llRow, BUSNAMEINDEX, True
'                mMakeExportStr tgStartColAFE.iBusControl, tgNoCharAFE.iBusControl, llRow, BUSCTRLINDEX, True
'                mMakeExportStr tgStartColAFE.iTime, tgNoCharAFE.iTime, llRow, TIMEINDEX, False
'                mMakeExportStr tgStartColAFE.iStartType, tgNoCharAFE.iStartType, llRow, STARTTYPEINDEX, False
'                mMakeExportStr tgStartColAFE.iEndType, tgNoCharAFE.iEndType, llRow, ENDTYPEINDEX, False
'                mMakeExportStr tgStartColAFE.iDuration, tgNoCharAFE.iDuration, llRow, DURATIONINDEX, False
'                mMakeExportStr tgStartColAFE.iMaterialType, tgNoCharAFE.iMaterialType, llRow, MATERIALINDEX, False
'                mMakeExportStr tgStartColAFE.iAudioName, tgNoCharAFE.iAudioName, llRow, AUDIONAMEINDEX, True
'                mMakeExportStr tgStartColAFE.iAudioItemID, tgNoCharAFE.iAudioItemID, llRow, AUDIOITEMIDINDEX, False
'                mMakeExportStr tgStartColAFE.iAudioControl, tgNoCharAFE.iAudioControl, llRow, AUDIOCTRLINDEX, True
'                mMakeExportStr tgStartColAFE.iBkupAudioName, tgNoCharAFE.iBkupAudioName, llRow, BACKUPNAMEINDEX, True
'                mMakeExportStr tgStartColAFE.iBkupAudioControl, tgNoCharAFE.iBkupAudioControl, llRow, BACKUPCTRLINDEX, True
'                mMakeExportStr tgStartColAFE.iProtAudioName, tgNoCharAFE.iProtAudioName, llRow, PROTNAMEINDEX, True
'                mMakeExportStr tgStartColAFE.iProtItemID, tgNoCharAFE.iProtItemID, llRow, PROTITEMIDINDEX, False
'                mMakeExportStr tgStartColAFE.iProtAudioControl, tgNoCharAFE.iProtAudioControl, llRow, PROTCTRLINDEX, True
'                mMakeExportStr tgStartColAFE.iRelay1, tgNoCharAFE.iRelay1, llRow, RELAY1INDEX, False
'                mMakeExportStr tgStartColAFE.iRelay2, tgNoCharAFE.iRelay2, llRow, RELAY2INDEX, False
'                mMakeExportStr tgStartColAFE.iFollow, tgNoCharAFE.iFollow, llRow, FOLLOWINDEX, False
'                mMakeExportStr tgStartColAFE.iSilenceTime, tgNoCharAFE.iSilenceTime, llRow, SILENCETIMEINDEX, False
'                mMakeExportStr tgStartColAFE.iSilence1, tgNoCharAFE.iSilence1, llRow, SILENCE1INDEX, False
'                mMakeExportStr tgStartColAFE.iSilence2, tgNoCharAFE.iSilence2, llRow, SILENCE2INDEX, False
'                mMakeExportStr tgStartColAFE.iSilence3, tgNoCharAFE.iSilence3, llRow, SILENCE3INDEX, False
'                mMakeExportStr tgStartColAFE.iSilence4, tgNoCharAFE.iSilence4, llRow, SILENCE4INDEX, False
'                mMakeExportStr tgStartColAFE.iStartNetcue, tgNoCharAFE.iStartNetcue, llRow, NETCUE1INDEX, False
'                mMakeExportStr tgStartColAFE.iStopNetcue, tgNoCharAFE.iStopNetcue, llRow, NETCUE2INDEX, False
'                If (slEventCategory = "P") Then
'                    mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, llRow, TITLE1INDEX, False
'                Else
'                    mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, llRow, TITLE1INDEX, False
'                End If
'                mMakeExportStr tgStartColAFE.iTitle2, tgNoCharAFE.iTitle2, llRow, TITLE2INDEX, False
'                'Event Type
'                'If mColOk(llRow, EVENTTYPEINDEX) Then
'                    If tgStartColAFE.iEventType > 0 Then
'                        slStr = slEventAutoCode
'                        Do While Len(slStr) < tgNoCharAFE.iEventType
'                            slStr = slStr & " "
'                        Loop
'                        Mid(smExportStr, tgStartColAFE.iEventType, tgNoCharAFE.iEventType) = slStr
'                    End If
'                'End If
'                'Fixed
'                If mExportCol(llRow, FIXEDINDEX) Then
'                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, FIXEDINDEX))
'                    If slStr = "Y" Then
'                        If tgStartColAFE.iFixedTime > 0 Then
'                            slStr = Trim$(tgAEE.sFixedTimeChar)
'                            Do While Len(slStr) < tgNoCharAFE.iFixedTime
'                                slStr = slStr & " "
'                            Loop
'                            Mid(smExportStr, tgStartColAFE.iFixedTime, tgNoCharAFE.iFixedTime) = slStr
'                        End If
'                    End If
'                End If
'                'Date
'                If tgStartColAFE.iDate > 0 Then
'                    slStr = slDate
'                    Do While Len(slStr) < tgNoCharAFE.iDate
'                        slStr = slStr & " "
'                    Loop
'                    Mid(smExportStr, tgStartColAFE.iDate, tgNoCharAFE.iDate) = slStr
'                End If
'                'End Time- If not exporting duration, then don't export end time
'                If mExportCol(llRow, DURATIONINDEX) Then
'                    If tgStartColAFE.iEndTime > 0 Then
'                        slStr = gLongToStrLengthInTenth(gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow, TIMEINDEX), False) + gStrLengthInTenthToLong(grdLibEvents.TextMatrix(llRow, DURATIONINDEX)), True)
'                        Do While Len(slStr) < tgNoCharAFE.iEndTime
'                            slStr = slStr & " "
'                        Loop
'                        Mid(smExportStr, tgStartColAFE.iEndTime, tgNoCharAFE.iEndTime) = slStr
'                    End If
'                End If
'                'Event ID
'                If tgStartColAFE.iEventID > 0 Then
'                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTIDINDEX))
'                    Do While Len(slStr) < tgNoCharAFE.iEventID
'                        slStr = "0" & slStr
'                    Loop
'                    Mid(smExportStr, tgStartColAFE.iEventID, tgNoCharAFE.iEventID) = slStr
'                End If
'                Print #hmExport, smExportStr
'                'Update SEE
'                llSEECode = Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
'                If llSEECode > 0 Then
'                    ilRet = gPutUpdate_SEE_SentFlag(llSEECode, "EngrSchdDef- Update SEE Sent Flag")
'                    llIndexRow = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
'                    For ilSEEOld = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
'                        If tmCurrSEE(llIndexRow).lCode = tgCurrSEE(ilSEEOld).lCode Then
'                            tgCurrSEE(ilSEEOld).sSentStatus = "S"
'                            tgCurrSEE(ilSEEOld).sSentDate = Format$(gNow(), sgShowDateForm)
'                            tmCurrSEE(llIndexRow).sSentStatus = tgCurrSEE(ilSEEOld).sSentStatus
'                            tmCurrSEE(llIndexRow).sSentDate = tgCurrSEE(ilSEEOld).sSentDate
'                            Exit For
'                        End If
'                    Next ilSEEOld
'                End If
'            End If
'        End If
'    Next llRow
'    Close hmExport
'    'Update SHE
'    ilRet = gPutUpdate_SHE_SentFlags(tmSHE.lCode, "EngrSchdDef- Update SHE Sent Flags")
'    If tmSHE.sLoadedAutoStatus <> "L" Then
'        tmSHE.sLoadedAutoStatus = "L"
'        tmSHE.iChgSeqNo = 0
'    Else
'        tmSHE.iChgSeqNo = tmSHE.iChgSeqNo + 1
'    End If
'    tmSHE.sLoadedAutoDate = Format$(gNow(), sgShowDateForm)
'    ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE")
'    gSetMousePointer grdLibEvents, grdLibEvents, vbDefault
'    MsgBox "Load Automation Export File: " & slExportFileName

End Sub

Private Sub cmcLoad_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcMerge_Click()
    Dim slMergeMsgFile As String
    Dim slMergeFileCP As String
    Dim slMergeFileCB As String
    Dim ilMergeRet As Integer
    Dim slMergeFileWOExt As String
    Dim ilPos As Integer
    Dim llStartSpot As Long
    Dim ilRemakeDay As Integer
    Dim ilRet As Integer
    
    If bmInSave Then
        Exit Sub
    End If
    If MsgBox("Merge Spots for Scheduled Date", vbYesNo) = vbNo Then
        Exit Sub
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    grdLibEvents.Redraw = False
    ilRemakeDay = False
    If (imFieldChgd) Or (UBound(tgFilterValues) > LBound(tgFilterValues)) Then
        ilRet = mSaveAndClearFilter(True, True)
        If Not ilRet Then
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            Exit Sub
        End If
    Else
        If Not mCheckFields(True) Then
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            MsgBox "One or more required fields are missing or defined incorrectly", vbInformation + vbOKOnly, "Schedule Merge"
            mSortErrorsToTop
            Exit Sub
        End If
        ilRemakeDay = True
    End If
    If Not mOpenMergeMsgFile(slMergeMsgFile) Then
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        Exit Sub
    End If
    If Not mOpenMergeFile(slMergeFileCP, slMergeFileCB) Then
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        Exit Sub
    End If
    If ilRemakeDay Then
        'mMoveSEECtrlsToRec
        imLastColSorted = -1
        mSortCol TIMEINDEX
    End If
    grdLibEvents.Redraw = False
    grdLibEvents.Visible = False
    'ilMergeRet = mMerge()
    llStartSpot = UBound(tmCurrSEE)
    ilMergeRet = gMerge(1, smAirDate, hmMerge, hmMsg, tmCurrSEE(), smT1Comment(), smT2Comment(), lbcCommercialSort, imMergeError)
    If imMergeError Then
        imMergeError = 2    'Merge done with errors
    Else
        imMergeError = 1    'Merge done without error
    End If
    Close hmMerge
    Close hmMsg
    If ilMergeRet Then
        '5/16/13: Handle case where spots removed
        bmMerged = True
        On Error Resume Next
        'Kill slMergeFileCP
        'If slMergeFileCB <> "" Then
        '    Kill slMergeFileCP
        'End If
        ilPos = InStr(1, slMergeFileCP, ".", vbTextCompare)
        If ilPos > 0 Then
            slMergeFileWOExt = Left$(slMergeFileCP, ilPos - 1)
        Else
            slMergeFileWOExt = slMergeFileCP
        End If
        Kill slMergeFileWOExt & ".old"
        On Error Resume Next
        Name slMergeFileCP As slMergeFileWOExt & ".old"
        If slMergeFileCB <> "" Then
            ilPos = InStr(1, slMergeFileCB, ".", vbTextCompare)
            If ilPos > 0 Then
                slMergeFileWOExt = Left$(slMergeFileCB, ilPos - 1)
            Else
                slMergeFileWOExt = slMergeFileCB
            End If
            Kill slMergeFileWOExt & ".old"
            On Error Resume Next
            Name slMergeFileCB As slMergeFileWOExt & ".old"
        End If
        ''ilRet = mCheckEventConflicts()
        '3/6/07- Remove automatic check.  Only done manually
        'gItemIDCheck spcItemID, tmCurrSEE()
    End If
    mPopARE
    'Remove filter
    mAdjustAvailTime llStartSpot
    mMoveSEERecToCtrls
    imLastColSorted = -1
    mSortCol TIMEINDEX
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    If (imMergeError = 2) Or (Not ilMergeRet) Then
        MsgBox "See Merge Results File: " & slMergeMsgFile & " for List of problems"
    End If
    grdLibEvents.Redraw = True
    grdLibEvents.Visible = True
    imFieldChgd = True
    mSetCommands
End Sub

Private Sub cmcMerge_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcReload_Click()
    Dim llLoop As Long
    Dim llRow As Long
    
    
    ReDim tmPCurrSEE(0 To 0) As SEE
    ReDim tmNCurrSEE(0 To 0) As SEE
    ReDim tmSeeBracket(0 To 0) As SEEBRACKET
    ReDim lmChgSEE(0 To 0) As Long
    
    If (tmSHE.lCode = 0) Or (tmSHE.sLoadedAutoStatus <> "L") Or (imFieldChgd) Then
        Exit Sub
    End If
    If Not mOkToGenUPD() Then
        Exit Sub
    End If
    
    For llLoop = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llLoop, EVENTTYPEINDEX)) <> "" Then
            llRow = Val(grdLibEvents.TextMatrix(llLoop, TMCURRSEEINDEX))
            lmChgSEE(UBound(lmChgSEE)) = llRow
            ReDim Preserve lmChgSEE(0 To UBound(lmChgSEE) + 1) As Long
        End If
    Next llLoop
    If Not mGenUPDFile() Then
    
    End If

End Sub

Private Sub cmcShowEvents_Click()
    Dim ilRet As Integer
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llLoop As Long
    Dim llRow As Long
    Dim ilETE As Integer
    Dim slCategory As String
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    Dim llAvailLength As Long
    
    If bmInSave Then
        Exit Sub
    End If
    smAirDate = cccDate.text    'edcDate.Text
    If Not gIsDate(smAirDate) Then
        Beep
        cccDate.SetFocus    'edcDate.SetFocus
        Exit Sub
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    'DoEvents
    imMergeError = 0    'Merge not done
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    ReDim lgLibDheUsed(0 To 0) As Long
    ReDim lmDeleteCodes(0 To 0) As Long
    ilRet = gGetRec_SHE_ScheduleHeaderByDate(smAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
    If Not ilRet Then
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        If llAirDate < llNowDate Then
            MsgBox "Schedule has not been created for specified date in past"
            Exit Sub
        End If
        If MsgBox("Day not created, Create from Libraries and Templates?", vbYesNo) = vbNo Then
            Exit Sub
        End If
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        DoEvents
        tmSHE.lCode = 0
        ilRet = gGetEventsFromLibraries(smAirDate)
        imEvtRet = False
    Else
        If igSchdCallType = 0 Then
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            MsgBox "Schedule date previously created. To change schedule return to Schedule selection screen and highlight date, then press Change button"
            Exit Sub
        End If
        ilRet = gGetRecs_SEE_ScheduleEventsAPIWithFilter(hmSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrSchdDef-Get Events", tgCurrSEE())
        If tmSHE.sConflictExist = "Y" Then
            imEvtRet = True
        Else
            imEvtRet = False
        End If
    End If
    mPopGrid
    '2/10/05-Remove conflict test as it is done when Libraries are save and when
    'this schedule is saved.
    'ilRet = mCheckEventConflicts()
    grdLibEvents.Redraw = True
    If tmSHE.lCode = 0 Then
        imFieldChgd = True
    Else
        imFieldChgd = False
    End If
    If (imSpotETECode > 0) And (llAirDate >= llNowDate) Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(SCHEDULEJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    Else
        'cmcDone.Enabled = False
        'imcInsert.Enabled = False
        'imcTrash.Enabled = False
    End If
    mSetCommands
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
End Sub

Private Sub cmcShowEvents_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcTest_Click()
    Dim llRow As Long
    Dim slStr As String
    Dim slExportFileName As String
    Dim ilLength As Integer
    Dim slDate As String
    Dim slEventCategory As String
    Dim slEventAutoCode As String
    Dim ilETE As Integer
    Dim llSort As Long
    Dim slBus As String
    Dim slTime As String
    ReDim tmSchdSort(0 To 0) As SCHDSORT
    
    If bmInSave Then
        Exit Sub
    End If
    If MsgBox("Generate a Test Export of Rows as Shown", vbYesNo) = vbNo Then
        Exit Sub
    End If
    If Not mOpenTestMsgFile(slExportFileName) Then
        Exit Sub
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    ilLength = gExportStrLength()
    If tgNoCharAFE.iDate = 8 Then
        slDate = Format$(smAirDate, "yyyymmdd")
    ElseIf tgNoCharAFE.iDate = 6 Then
        slDate = Format$(smAirDate, "yymmdd")
    End If
    
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            slEventCategory = ""
            For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                    slEventCategory = tgCurrETE(ilETE).sCategory
                    slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
                    Exit For
                End If
            Next ilETE
            If (slEventCategory = "P") Or (slEventCategory = "S") Then
                slBus = Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX))
                Do While Len(slBus) < 8
                    slBus = slBus & " "
                Loop
                slTime = Trim$(grdLibEvents.TextMatrix(llRow, TIMEINDEX))
                slTime = Trim$(Str$(gStrTimeInTenthToLong(slTime, False)))
                Do While Len(slTime) < 6
                    slTime = "0" & slTime
                Loop
                tmSchdSort(UBound(tmSchdSort)).sKey = slTime & "|" & slBus & "|" & slEventCategory
                tmSchdSort(UBound(tmSchdSort)).lRow = llRow
                ReDim Preserve tmSchdSort(0 To UBound(tmSchdSort) + 1) As SCHDSORT
            End If
        End If
    Next llRow
    
    If UBound(tmSchdSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmSchdSort(), 0), UBound(tmSchdSort), 0, LenB(tmSchdSort(0)), 0, LenB(tmSchdSort(0).sKey), 0
    End If
    
    'For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
    For llSort = 0 To UBound(tmSchdSort) - 1 Step 1
        llRow = tmSchdSort(llSort).lRow
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            slEventCategory = ""
            For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                    slEventCategory = tgCurrETE(ilETE).sCategory
                    slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
                    Exit For
                End If
            Next ilETE
            If (slEventCategory = "P") Or (slEventCategory = "S") Then
                smExportStr = String(ilLength, " ")
                mMakeExportStr tgStartColAFE.iBus, tgNoCharAFE.iBus, llRow, BUSNAMEINDEX, True
                mMakeExportStr tgStartColAFE.iBusControl, tgNoCharAFE.iBusControl, llRow, BUSCTRLINDEX, True
                mMakeExportStr tgStartColAFE.iTime, tgNoCharAFE.iTime, llRow, TIMEINDEX, False
                mMakeExportStr tgStartColAFE.iStartType, tgNoCharAFE.iStartType, llRow, STARTTYPEINDEX, False
                mMakeExportStr tgStartColAFE.iEndType, tgNoCharAFE.iEndType, llRow, ENDTYPEINDEX, False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, DURATIONINDEX))
                '2/22/13: Don't show duration if zero and Ent Type = MAN or EXT
                If (slStr <> "00:00:00.0") Or ((slStr = "00:00:00.0") And (Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)) <> "MAN") And (Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)) <> "EXT")) Then
                    mMakeExportStr tgStartColAFE.iDuration, tgNoCharAFE.iDuration, llRow, DURATIONINDEX, False
                End If
                mMakeExportStr tgStartColAFE.iMaterialType, tgNoCharAFE.iMaterialType, llRow, MATERIALINDEX, False
                mMakeExportStr tgStartColAFE.iAudioName, tgNoCharAFE.iAudioName, llRow, AUDIONAMEINDEX, True
                mMakeExportStr tgStartColAFE.iAudioItemID, tgNoCharAFE.iAudioItemID, llRow, AUDIOITEMIDINDEX, False
                mMakeExportStr tgStartColAFE.iAudioISCI, tgNoCharAFE.iAudioISCI, llRow, AUDIOISCIINDEX, False
                mMakeExportStr tgStartColAFE.iAudioControl, tgNoCharAFE.iAudioControl, llRow, AUDIOCTRLINDEX, True
                mMakeExportStr tgStartColAFE.iBkupAudioName, tgNoCharAFE.iBkupAudioName, llRow, BACKUPNAMEINDEX, True
                mMakeExportStr tgStartColAFE.iBkupAudioControl, tgNoCharAFE.iBkupAudioControl, llRow, BACKUPCTRLINDEX, True
                mMakeExportStr tgStartColAFE.iProtAudioName, tgNoCharAFE.iProtAudioName, llRow, PROTNAMEINDEX, True
                mMakeExportStr tgStartColAFE.iProtItemID, tgNoCharAFE.iProtItemID, llRow, PROTITEMIDINDEX, False
                mMakeExportStr tgStartColAFE.iProtISCI, tgNoCharAFE.iProtISCI, llRow, PROTISCIINDEX, False
                mMakeExportStr tgStartColAFE.iProtAudioControl, tgNoCharAFE.iProtAudioControl, llRow, PROTCTRLINDEX, True
                mMakeExportStr tgStartColAFE.iRelay1, tgNoCharAFE.iRelay1, llRow, RELAY1INDEX, False
                mMakeExportStr tgStartColAFE.iRelay2, tgNoCharAFE.iRelay2, llRow, RELAY2INDEX, False
                mMakeExportStr tgStartColAFE.iFollow, tgNoCharAFE.iFollow, llRow, FOLLOWINDEX, False
                mMakeExportStr tgStartColAFE.iSilenceTime, tgNoCharAFE.iSilenceTime, llRow, SILENCETIMEINDEX, False
                mMakeExportStr tgStartColAFE.iSilence1, tgNoCharAFE.iSilence1, llRow, SILENCE1INDEX, False
                mMakeExportStr tgStartColAFE.iSilence2, tgNoCharAFE.iSilence2, llRow, SILENCE2INDEX, False
                mMakeExportStr tgStartColAFE.iSilence3, tgNoCharAFE.iSilence3, llRow, SILENCE3INDEX, False
                mMakeExportStr tgStartColAFE.iSilence4, tgNoCharAFE.iSilence4, llRow, SILENCE4INDEX, False
                mMakeExportStr tgStartColAFE.iStartNetcue, tgNoCharAFE.iStartNetcue, llRow, NETCUE1INDEX, False
                mMakeExportStr tgStartColAFE.iStopNetcue, tgNoCharAFE.iStopNetcue, llRow, NETCUE2INDEX, False
                If (slEventCategory = "P") Then
                    mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, llRow, TITLE1INDEX, False
                Else
                    mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, llRow, TITLE1INDEX, False
                End If
                mMakeExportStr tgStartColAFE.iTitle2, tgNoCharAFE.iTitle2, llRow, TITLE2INDEX, False
                'Event Type
                'If mExportCol(llRow, EVENTTYPEINDEX) Then
                    If tgStartColAFE.iEventType > 0 Then
                        slStr = slEventAutoCode
                        Do While Len(slStr) < tgNoCharAFE.iEventType
                            slStr = slStr & " "
                        Loop
                        Mid(smExportStr, tgStartColAFE.iEventType, tgNoCharAFE.iEventType) = slStr
                    End If
                'End If
                'Fixed
                If mExportCol(llRow, FIXEDINDEX) Then
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, FIXEDINDEX))
                    If slStr = "Y" Then
                        If tgStartColAFE.iFixedTime > 0 Then
                            slStr = Trim$(tgAEE.sFixedTimeChar)
                            Do While Len(slStr) < tgNoCharAFE.iFixedTime
                                slStr = slStr & " "
                            Loop
                            Mid(smExportStr, tgStartColAFE.iFixedTime, tgNoCharAFE.iFixedTime) = slStr
                        End If
                    End If
                End If
                'Date
                If tgStartColAFE.iDate > 0 Then
                    slStr = slDate
                    Do While Len(slStr) < tgNoCharAFE.iDate
                        slStr = slStr & " "
                    Loop
                    Mid(smExportStr, tgStartColAFE.iDate, tgNoCharAFE.iDate) = slStr
                End If
                'End Time
                If mExportCol(llRow, DURATIONINDEX) Then
                    '2/22/13: Don't show Out time if duration is zero and End Type = MAN or EXT
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, DURATIONINDEX))
                    If (slStr <> "00:00:00.0") Or ((slStr = "00:00:00.0") And (Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)) <> "MAN") And (Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)) <> "EXT")) Then
                        If tgStartColAFE.iEndTime > 0 Then
                            slStr = gLongToStrLengthInTenth(gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow, TIMEINDEX), False) + gStrLengthInTenthToLong(grdLibEvents.TextMatrix(llRow, DURATIONINDEX)), True)
                            Do While Len(slStr) < tgNoCharAFE.iEndTime
                                slStr = slStr & " "
                            Loop
                            Mid(smExportStr, tgStartColAFE.iEndTime, tgNoCharAFE.iEndTime) = slStr
                        End If
                    End If
                End If
                'Event ID
                If tgStartColAFE.iEventID > 0 Then
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTIDINDEX))
                    Do While Len(slStr) < tgNoCharAFE.iEventID
                        slStr = "0" & slStr
                    Loop
                    Mid(smExportStr, tgStartColAFE.iEventID, tgNoCharAFE.iEventID) = slStr
                End If
                Print #hmMsg, smExportStr
            End If
        End If
    Next llSort
    'Next llRow
    Close hmMsg
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    MsgBox "Export File: " & slExportFileName
End Sub

Private Sub cmcTest_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcFilter_Click()
    Dim ilCol As Integer
    Dim ilFilter As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    
    If bmInSave Then
        Exit Sub
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    '6/14/06- remove check as done with save
    'If Not mCheckFields(True) Then
    '    gSetMousePointer grdLibEvents, grdLibEvents, vbDefault
    '    MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Schedule"
    '    mSortErrorsToTop
    '    Exit Sub
    'End If
    If (imFieldChgd) Then
        ilRet = mSaveAndClearFilter(True, False)
        If Not ilRet Then
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            Exit Sub
        End If
    End If
    'mMoveSEECtrlsToRec
    ''Moved to sch screen
    ''mCreateUsedArrays
    ''mInitFilterInfo
    igAnsFilter = 0
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    EngrSchdFilter.Show vbModal
    If igAnsFilter = CALLDONE Then 'Apply
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        mReorderFilter
        cmcShowEvents_Click
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
End Sub

Private Sub cmcFilter_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcItemIDChk_Click()
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slItemID As String
    Dim slTitle As String
    Dim llUpper As Long
    Dim llIndex As Long
    
    If bmInSave Then
        Exit Sub
    End If
    'If Not gIsDate(edcDate.Text) Then
    '    edcDate.SetFocus
    '    Exit Sub
    'End If
    igInitCallInfo = 0
    sgItemIDDate = cccDate.text  'edcDate.Text
    mBuildItemIDChk
    EngrItemIDChk.Show vbModal
    If igReturnCallStatus = CALLDONE Then 'Apply
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        grdLibEvents.Redraw = False
        'Check each item or pass item names to calling program
        For ilLoop = 0 To UBound(tgItemIDChk) - 1 Step 1
            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                    If StrComp("Spot", Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)), vbTextCompare) = 0 Then
                        slItemID = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX))
                        slTitle = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
                        If (slTitle <> "") And (StrComp(slTitle, "[None]", vbTextCompare) <> 0) Then
                            If StrComp(Trim$(tgItemIDChk(ilLoop).sItemID), slItemID, vbTextCompare) = 0 Then
                                If StrComp(Trim$(tgItemIDChk(ilLoop).sTitle), slTitle, vbTextCompare) = 0 Then
                                    If Trim$(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX)) = "" Then
                                        llUpper = UBound(tmCurrSEE)
                                        ReDim Preserve tmCurrSEE(0 To llUpper + 1) As SEE
                                        ReDim Preserve smT1Comment(0 To llUpper + 1) As String
                                        ReDim Preserve smT2Comment(0 To llUpper + 1) As String
                                        grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llUpper))
                                        gInitSEE tmCurrSEE(llUpper)
                                        smT1Comment(llUpper) = ""
                                        smT2Comment(llUpper) = ""
                                    End If
                                    llIndex = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
                                    If (tgItemIDChk(ilLoop).sAudioStatus = "F") And (tmCurrSEE(llIndex).sAudioItemIDChk <> "F") Then
                                        tmCurrSEE(llIndex).sAudioItemIDChk = "F"
                                        imFieldChgd = True
                                    ElseIf (tgItemIDChk(ilLoop).sAudioStatus = "O") And (tmCurrSEE(llIndex).sAudioItemIDChk <> "O") Then
                                        tmCurrSEE(llIndex).sAudioItemIDChk = "O"
                                        imFieldChgd = True
                                    End If
                                    grdLibEvents.Row = llRow
                                    grdLibEvents.Col = AUDIOITEMIDINDEX
                                    If (tmCurrSEE(llIndex).sAudioItemIDChk = "F") Then
                                        grdLibEvents.CellForeColor = vbRed
                                    ElseIf (tmCurrSEE(llIndex).sAudioItemIDChk = "O") Then
                                        grdLibEvents.CellForeColor = vbGreen
                                    Else
                                        grdLibEvents.CellForeColor = vbMagenta  'vbBlue
                                    End If
                                    If (tgItemIDChk(ilLoop).sProtStatus = "F") And (tmCurrSEE(llIndex).sProtItemIDChk <> "F") Then
                                        tmCurrSEE(llIndex).sProtItemIDChk = "F"
                                        imFieldChgd = True
                                    ElseIf (tgItemIDChk(ilLoop).sProtStatus = "O") And (tmCurrSEE(llIndex).sProtItemIDChk <> "O") Then
                                        tmCurrSEE(llIndex).sProtItemIDChk = "O"
                                        imFieldChgd = True
                                    End If
                                    grdLibEvents.Row = llRow
                                    grdLibEvents.Col = PROTITEMIDINDEX
                                    If (tmCurrSEE(llIndex).sProtItemIDChk = "F") Then
                                        grdLibEvents.CellForeColor = vbRed
                                    ElseIf (tmCurrSEE(llIndex).sProtItemIDChk = "O") Then
                                        grdLibEvents.CellForeColor = vbGreen
                                    Else
                                        grdLibEvents.CellForeColor = vbMagenta  'vbBlue
                                    End If
                                End If
                            End If
                        Else
                            If tmCurrSEE(llIndex).sAudioItemIDChk <> "N" Then
                                tmCurrSEE(llIndex).sAudioItemIDChk = "N"
                                imFieldChgd = True
                            End If
                            grdLibEvents.Row = llRow
                            grdLibEvents.Col = AUDIOITEMIDINDEX
                            If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                                grdLibEvents.CellForeColor = vbBlue
                            Else
                                grdLibEvents.CellForeColor = vbBlack
                            End If
                            If tmCurrSEE(llIndex).sProtItemIDChk <> "N" Then
                                tmCurrSEE(llIndex).sProtItemIDChk = "N"
                                imFieldChgd = True
                            End If
                            grdLibEvents.Row = llRow
                            grdLibEvents.Col = PROTITEMIDINDEX
                            If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                                grdLibEvents.CellForeColor = vbBlue
                            Else
                                grdLibEvents.CellForeColor = vbBlack
                            End If
                        End If
                    End If
                End If
            Next llRow
        Next ilLoop
        grdLibEvents.Redraw = True
    End If
    mSetCommands
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    
End Sub

Private Sub cmcItemIDChk_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcReplace_Click()
    Dim ilCol As Integer
    Dim ilFilter As Integer
    Dim ilIndex As Integer
    
    If bmInSave Then
        Exit Sub
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    grdLibEvents.Redraw = False
    If Not mCheckFields(True) Then
        grdLibEvents.Redraw = True
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Schedule"
        mSortErrorsToTop
        Exit Sub
    End If
    ReDim tgSchdReplaceValues(0 To 0) As SCHDREPLACEVALUES
    'mMoveSEECtrlsToRec
    ''mCreateUsedArrays
    ''mInitReplaceInfo
    igAnsReplace = 0
    igReplaceCallInfo = 0
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    EngrReplaceSchd.Show vbModal
    If igAnsReplace = CALLDONE Then 'Apply
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        grdLibEvents.Redraw = False
        'mReplaceValuesAvails
        mReplaceValues
    End If
    grdLibEvents.Redraw = True
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
End Sub

Private Sub cmcReplace_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    
    If bmInSave Then
        Exit Sub
    End If
    If imFieldChgd = True Then
        grdLibEvents.Redraw = False
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        ilRet = mSave()
        bmInSave = False
        If Not ilRet Then
            grdLibEvents.Redraw = True
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            Exit Sub
        End If
        grdLibEvents.Redraw = False
        mClearControls
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(smAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
        sgCurrSEEStamp = ""
        'ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrSchdDef-Get Events", tgCurrSEE())
        ilRet = gGetRecs_SEE_ScheduleEventsAPIWithFilter(hmSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrSchdDef-Get Events", tgCurrSEE())
        mPopGrid
        lmEEnableRow = -1
        lmEEnableCol = -1
        imFieldChgd = False
        mSetCommands
        grdLibEvents.Redraw = True
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub





Private Sub edcEDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    Dim ilANE As Integer
    Dim ilCCE As Integer
    Dim ilASE As Integer
    Dim ilANE2 As Integer
    
    slStr = edcEDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdLibEvents.Col
        Case BUSNAMEINDEX
            llRow = gListBoxFind(lbcBDE, slStr)
            If llRow >= 0 Then
                lbcBDE.ListIndex = llRow
                edcEDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case BUSCTRLINDEX
            llRow = gListBoxFind(lbcCCE_B, slStr)
            If llRow >= 0 Then
                lbcCCE_B.ListIndex = llRow
                edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case EVENTTYPEINDEX
            '2/9/12: Allow all events
            ''llRow = gListBoxFind(lbcETE, slStr)
            ''If llRow >= 0 Then
            ''    lbcETE.ListIndex = llRow
            ''    edcEDropdown.Text = lbcETE.List(lbcETE.ListIndex)
            ''    edcEDropdown.SelStart = ilLen
            ''    edcEDropdown.SelLength = Len(edcEDropdown.Text)
            ''End If
            'llRow = gListBoxFind(lbcETE_Program, slStr)
            'If llRow >= 0 Then
            '    lbcETE_Program.ListIndex = llRow
            '    edcEDropdown.text = lbcETE_Program.List(lbcETE_Program.ListIndex)
            '    edcEDropdown.SelStart = ilLen
            '    edcEDropdown.SelLength = Len(edcEDropdown.text)
            'End If
            llRow = gListBoxFind(lbcETE, slStr)
            If llRow >= 0 Then
                lbcETE.ListIndex = llRow
                edcEDropdown.text = lbcETE.List(lbcETE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
       Case STARTTYPEINDEX
            llRow = gListBoxFind(lbcTTE_S, slStr)
            If llRow >= 0 Then
                lbcTTE_S.ListIndex = llRow
                edcEDropdown.text = lbcTTE_S.List(lbcTTE_S.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case ENDTYPEINDEX
            llRow = gListBoxFind(lbcTTE_E, slStr)
            If llRow >= 0 Then
                lbcTTE_E.ListIndex = llRow
                edcEDropdown.text = lbcTTE_E.List(lbcTTE_E.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case MATERIALINDEX
            llRow = gListBoxFind(lbcMTE, slStr)
            If llRow >= 0 Then
                lbcMTE.ListIndex = llRow
                edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case AUDIONAMEINDEX
            llRow = gListBoxFind(lbcASE, slStr)
            If llRow >= 0 Then
                lbcASE.ListIndex = llRow
                edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case AUDIOCTRLINDEX
            llRow = gListBoxFind(lbcCCE_A, slStr)
            If llRow >= 0 Then
                lbcCCE_A.ListIndex = llRow
                edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case BACKUPNAMEINDEX
            llRow = gListBoxFind(lbcANE, slStr)
            If llRow >= 0 Then
                lbcANE.ListIndex = llRow
                edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case BACKUPCTRLINDEX
            llRow = gListBoxFind(lbcCCE_A, slStr)
            If llRow >= 0 Then
                lbcCCE_A.ListIndex = llRow
                edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case PROTNAMEINDEX
            llRow = gListBoxFind(lbcANE, slStr)
            If llRow >= 0 Then
                lbcANE.ListIndex = llRow
                edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case PROTCTRLINDEX
            llRow = gListBoxFind(lbcCCE_A, slStr)
            If llRow >= 0 Then
                lbcCCE_A.ListIndex = llRow
                edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case RELAY1INDEX, RELAY2INDEX
            llRow = gListBoxFind(lbcRNE, slStr)
            If llRow >= 0 Then
                lbcRNE.ListIndex = llRow
                edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case FOLLOWINDEX
            llRow = gListBoxFind(lbcFNE, slStr)
            If llRow >= 0 Then
                lbcFNE.ListIndex = llRow
                edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case SILENCE1INDEX To SILENCE4INDEX
            llRow = gListBoxFind(lbcSCE, slStr)
            If llRow >= 0 Then
                lbcSCE.ListIndex = llRow
                edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case NETCUE1INDEX, NETCUE2INDEX
            llRow = gListBoxFind(lbcNNE, slStr)
            If llRow >= 0 Then
                lbcNNE.ListIndex = llRow
                edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case TITLE1INDEX
            llRow = gListBoxFind(lbcCTE_1, slStr)
            If llRow >= 0 Then
                lbcCTE_1.ListIndex = llRow
                edcEDropdown.text = lbcCTE_1.List(lbcCTE_1.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
        Case TITLE2INDEX
            llRow = gListBoxFind(lbcCTE_2, slStr)
            If llRow >= 0 Then
                lbcCTE_2.ListIndex = llRow
                edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
                edcEDropdown.SelStart = ilLen
                edcEDropdown.SelLength = Len(edcEDropdown.text)
            End If
    End Select
    If (StrComp(grdLibEvents.text, edcEDropdown.text, vbTextCompare) <> 0) Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        Select Case grdLibEvents.Col
            Case AUDIONAMEINDEX
'                slStr = Trim$(edcEDropdown.Text)
'                For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
'                    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
'                        If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
'                            If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
'                                slStr = ""
'                                If tgCurrASE(ilASE).iPriCceCode > 0 Then
'                                    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
'                                        If tgCurrASE(ilASE).iPriCceCode = tgCurrAudioCCE(ilCCE).iCode Then
'                                            grdLibEvents.TextMatrix(grdLibEvents.Row, AUDIOCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
'                                            Exit For
'                                        End If
'                                    Next ilCCE
'                                Else
'                                    grdLibEvents.TextMatrix(grdLibEvents.Row, AUDIOCTRLINDEX) = ""
'                                End If
'                                If tgCurrASE(ilASE).iBkupAneCode > 0 Then
'                                    For ilANE2 = 0 To UBound(tgCurrANE) - 1 Step 1
'                                        If tgCurrASE(ilASE).iBkupAneCode = tgCurrANE(ilANE2).iCode Then
'                                            grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPNAMEINDEX) = Trim$(tgCurrANE(ilANE2).sName)
'                                            Exit For
'                                        End If
'                                    Next ilANE2
'                                Else
'                                    grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPNAMEINDEX) = ""
'                                End If
'                                If tgCurrASE(ilASE).iBkupCceCode > 0 Then
'                                    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
'                                        If tgCurrASE(ilASE).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
'                                            grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
'                                            Exit For
'                                        End If
'                                    Next ilCCE
'                                Else
'                                    grdLibEvents.TextMatrix(grdLibEvents.Row, BACKUPCTRLINDEX) = ""
'                                End If
'                                If tgCurrASE(ilASE).iProtAneCode > 0 Then
'                                    For ilANE2 = 0 To UBound(tgCurrANE) - 1 Step 1
'                                        If tgCurrASE(ilASE).iProtAneCode = tgCurrANE(ilANE2).iCode Then
'                                            grdLibEvents.TextMatrix(grdLibEvents.Row, PROTNAMEINDEX) = Trim$(tgCurrANE(ilANE2).sName)
'                                            Exit For
'                                        End If
'                                    Next ilANE2
'                                Else
'                                    grdLibEvents.TextMatrix(grdLibEvents.Row, PROTNAMEINDEX) = ""
'                                End If
'                                If tgCurrASE(ilASE).iProtCceCode > 0 Then
'                                    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
'                                        If tgCurrASE(ilASE).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
'                                            grdLibEvents.TextMatrix(grdLibEvents.Row, PROTCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
'                                            Exit For
'                                        End If
'                                    Next ilCCE
'                                Else
'                                    grdLibEvents.TextMatrix(grdLibEvents.Row, PROTCTRLINDEX) = ""
'                                End If
'                            End If
'                            Exit For
'                        End If
'                    Next ilANE
'                    If slStr = "" Then
'                        Exit For
'                    End If
'                Next ilASE
            Case ENDTYPEINDEX
                '11/24/04- Allow end type and Duration to co-exist
                'If lbcTTE_E.ListIndex > 1 Then
                '    grdLibEvents.TextMatrix(grdLibEvents.Row, DURATIONINDEX) = ""
                'End If
        End Select
        
        If (grdLibEvents.Col <> TITLE1INDEX) And (grdLibEvents.Col <> TITLE2INDEX) Then
            If StrComp(Trim$(edcEDropdown.text), "[None]", vbTextCompare) <> 0 Then
                grdLibEvents.text = edcEDropdown.text
            Else
                grdLibEvents.text = ""
            End If
        End If
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub edcEDropdown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub edcEDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcEDropdown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    If KeyAscii = 8 Then
        If edcEDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
    If (imMaxColChars < edcEDropdown.MaxLength) And (imMaxColChars > 0) And (KeyAscii <> 8) Then
        slStr = edcEDropdown.text
        slStr = Left$(slStr, edcEDropdown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcEDropdown.SelStart - edcEDropdown.SelLength)
        If (Len(slStr) > imMaxColChars) And (Left$(slStr, 1) <> "[") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcEDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdLibEvents.Col
            Case BUSNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcBDE, True
            Case BUSCTRLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE_B, True
            Case EVENTTYPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcETE, True
            Case STARTTYPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcTTE_S, True
            Case ENDTYPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcTTE_E, True
            Case MATERIALINDEX
                gProcessArrowKey Shift, KeyCode, lbcMTE, True
            Case AUDIONAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcASE, True
            Case AUDIOCTRLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE_A, True
            Case BACKUPNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcANE, True
            Case BACKUPCTRLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE_A, True
            Case PROTNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcANE, True
            Case PROTCTRLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE_A, True
            Case RELAY1INDEX, RELAY2INDEX
                gProcessArrowKey Shift, KeyCode, lbcRNE, True
            Case FOLLOWINDEX
                gProcessArrowKey Shift, KeyCode, lbcFNE, True
            Case SILENCE1INDEX To SILENCE4INDEX
                gProcessArrowKey Shift, KeyCode, lbcSCE, True
            Case NETCUE1INDEX, NETCUE2INDEX
                gProcessArrowKey Shift, KeyCode, lbcNNE, True
            Case TITLE1INDEX
                gProcessArrowKey Shift, KeyCode, lbcCTE_1, True
            Case TITLE2INDEX
                gProcessArrowKey Shift, KeyCode, lbcCTE_2, True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcEDropdown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilRet As Integer
    
    If imDoubleClickName Then
        ilRet = mEBranch()
        bmInBranch = False
    End If
End Sub

Private Sub edcEvent_Change()
    Dim slStr As String
    
    slStr = edcEvent.text
    Select Case grdLibEvents.Col
        Case TIMEINDEX
        Case AUDIOITEMIDINDEX
        Case AUDIOISCIINDEX
        Case PROTITEMIDINDEX
        Case PROTISCIINDEX
    End Select
    If grdLibEvents.text <> slStr Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        Select Case grdLibEvents.Col
            Case TIMEINDEX
            Case AUDIOITEMIDINDEX
            Case AUDIOISCIINDEX
            Case PROTITEMIDINDEX
            Case PROTISCIINDEX
        End Select
        grdLibEvents.text = slStr
        If grdLibEvents.CellForeColor <> vbMagenta Then
            If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                grdLibEvents.CellForeColor = vbBlue
            Else
                grdLibEvents.CellForeColor = vbBlack
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub edcEvent_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub edcEvent_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    '2/13/12: Disallow L in the Item ID.  The user should use the Delete instead
    If (grdLibEvents.Col = AUDIOITEMIDINDEX) Or (grdLibEvents.Col = PROTITEMIDINDEX) Then
        '               L                  l
        If (KeyAscii = 76) Or (KeyAscii = 108) Then
            '2/13/12: Disallow any L  to avoid L entered within the name then remove the previous characters
            'If (edcEvent.SelStart = 0) And (tmSHE.sLoadedAutoStatus = "L") Then
            If (tmSHE.sLoadedAutoStatus = "L") Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    cmcCancel.SetFocus
End Sub

Private Sub Form_Click()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    bmIntegralSet = False
    'Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    'Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    Me.Move Me.Left, Me.Top, 0.97 * Screen.Width, 0.82 * Screen.Height
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrSchdDef
    'gCenterFormModal EngrSchdDef
    gCenterForm EngrSchdDef
'    Unload EngrLib
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEEnableRow >= grdLibEvents.FixedRows) And (lmEEnableRow < grdLibEvents.Rows) Then
            If (lmEEnableCol >= grdLibEvents.FixedCols) And (lmEEnableCol < grdLibEvents.Cols) Then
                grdLibEvents.text = smESCValue
                mESetShow
                mEEnableBox
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
    igJobShowing(SCHEDULEJOB) = 2
End Sub

Private Sub Form_Resize()
    Dim llRow As Long
    Dim llHeight As Long
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdConflicts.Height = 4 * grdConflicts.RowHeight(0) + 15
    gGrid_IntegralHeight grdConflicts
    grdLibEvents.Height = cmcCancel.Top - lacHelp.Height - grdConflicts.Height - 240 '240    '4 * grdLibEvents.RowHeight(0) + 15
    '8/26/11: Moved
    'gGrid_IntegralHeight grdLibEvents
    gGrid_FillWithRows grdLibEvents
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        grdLibEvents.Row = llRow
        grdLibEvents.Col = EVENTIDINDEX
        grdLibEvents.CellBackColor = LIGHTYELLOW
        grdLibEvents.CellAlignment = flexAlignRightCenter
    Next llRow
    lacHelp.Height = 240
    lacHelp.Top = cmcCancel.Top - lacHelp.Height - 30
    imcTrash.Top = cmcCancel.Top
    imcPrint.Top = cmcCancel.Top
    imcInsert.Top = cmcCancel.Top
    lmCharacterWidth = CLng(pbcETab.TextWidth("n"))
    gSetListBoxHeight lbcKey, 2 * grdLibEvents.Height
    lbcKey.Height = lbcKey.Height + lbcKey.Height / 7
    'Adjust height so that the line under the scroll bar is not visible with IsRowVisible acll
    '8/26/11: Removed
    'grdLibEvents.Height = grdLibEvents.Height - 15
    lmGridLibEventsHeight = grdLibEvents.Height
    grdConflicts.Move grdLibEvents.Left + grdLibEvents.Width - grdConflicts.Width, 0    '60
    grdLibEvents.Top = grdConflicts.Top + grdConflicts.Height
    cmcTask.Top = (Me.Height - cmcTask.Height) / 2
    cmcTask.Left = (Me.Width - cmcTask.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmcCheck.Enabled = False
        
    gGetSchDates
    
    btrDestroy hmSEE
    btrDestroy hmCME
    btrDestroy hmSOE
    btrDestroy hmCTE
    
    Erase smHours
    Erase smDays
    Erase lmDeleteCodes
    Erase tmCurrBSE
    Erase tmCurrDBE
    Erase smT1Comment
    Erase tmCurr1CTE_Name
    Erase smT2Comment
    Erase tmCurr2CTE_Name
    Erase tmCurrSEE
    Erase tgSpotCurrSEE
    
    Erase tmCCurrSEE
    Erase tmPCurrSEE
    Erase tmNCurrSEE
    Erase tmSeeBracket
    Erase lmChgSEE
    
    Erase tmSchdSort
    
    Erase tmFilterValues
    Erase imFilterBus
    Erase imFilterAudio
    
    Erase lgLibDheUsed
    
    Erase tmCurrLibDBE
    Erase tmCurrLibEBE
    
    Erase tmConflictList
    Erase tmConflictTest
    Erase lmChgStatusSEECode
    
    Set EngrSchdDef = Nothing
    EngrSchd.Show vbModeless
End Sub





Private Sub mInit()
    Dim llRow As Long
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilETE As Integer
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim lmDeleteCodes(0 To 0) As Long
    ReDim tmConflictList(1 To 1) As CONFLICTLIST
    tmConflictList(UBound(tmConflictList)).iNextIndex = -1
    'ReDim tgFilterValues(0 To 0) As FILTERVALUES
    'ReDim tgFilterFields(0 To 0) As FIELDSELECTION
    ReDim tmFilterValues(0 To 0) As FILTERVALUES
    'ReDim tgSchdReplaceValues(0 To 0) As SCHDREPLACEVALUES
    'ReDim tgReplaceFields(0 To 0) As FIELDSELECTION
    ReDim tmCurrSEE(0 To 0) As SEE
    ReDim tmCurr1CTE_Name(0 To 0) As DEECTE
    ReDim tmCurr2CTE_Name(0 To 0) As DEECTE
    ReDim lmChgStatusSEECode(0 To 0) As Long
    '5/16/13: Handle case where spots removed
    bmMerged = False
    'Can't be 0 to 0 because of index in grid
'    cmcSearch.Top = 30
'    edcSearch.Top = cmcSearch.Top
    imDefaultProgIndex = -1
    bmInBranch = False
    bmInInsert = False
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCME = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCME, "", sgDBPath & "CME.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmSOE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSOE, "", sgDBPath & "SOE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    tmSHE.lCode = 0
    imStartChgModeCompleted = False
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEEnableRow = -1
    lmConflictRow = -1
    imFirstActivate = True
    lmInsertRow = -1
    imInsertState = False
    imInChg = True
    imMergeError = 0
    bmPrinting = False
    bmInSave = False
    lmFilterStartTime = 0
    lmFilterEndTime = 86399
    ReDim imFilterBus(0 To 0) As Integer
    ReDim imFilterAudio(0 To 0) As Integer
'Moved to start timer
'    mPopANE
'    mPopASE
'    mPopBDE
'    mPopCCE_Audio
'    mPopCCE_Bus
'    mPopCTE
'    mPopDNE
'    mPopDSE
'    mPopETE
'    mPopFNE
'    mPopMTE
'    mPopNNE
'    mPopRNE
'    mPopSCE
'    mPopTTE_EndType
'    mPopTTE_StartType
'    mPopARE
'    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
'    imSpotETECode = 0
'    smSpotEventTypeName = "Spot"
'    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'        If tgCurrETE(ilETE).sCategory = "S" Then
'            imSpotETECode = tgCurrETE(ilETE).iCode
'            smSpotEventTypeName = Trim$(tgCurrETE(ilETE).sName)
'            Exit For
'        End If
'    Next ilETE
'    If imSpotETECode <= 0 Then
'        gSetMousePointer grdLibEvents, grdLibEvents, vbDefault
'        MsgBox "Spot Event Type not defined", vbCritical + vbOKOnly, "Schedule"
'    End If
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If igSchdCallType = 0 Then
        lacDate.Visible = True
        cccDate.Visible = True
        cmcShowEvents.Visible = True
        lacScreen.Caption = "Schedule"
        slDate = gGetLatestSchdDate(True)
        If gDateValue(slDate) < gDateValue(smNowDate) Then
            slDate = smNowDate
        End If
        cccDate.text = DateAdd("d", 1, slDate)
    Else
        lacDate.Visible = False
        cccDate.Visible = False
        cmcShowEvents.Visible = False
        lacScreen.Caption = "Schedule: " & sgSchdDate
    End If
    imInChg = False
    imFieldChgd = False
    If imSpotETECode > 0 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igJobStatus(SCHEDULEJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    
    If sgClientFields = "A" Then
        '8/26/: Retained adding horizontal scroll bar
        grdLibEvents.ScrollBars = flexScrollBarBoth
        imMaxCols = ABCRECORDITEMINDEX
    Else
        imMaxCols = TITLE2INDEX
    End If
    
    gSetListBoxHeight lbcKey, grdLibEvents.Height
    mSetCommands
    'If igSchdCallType <> 0 Then
        tmcStart.Enabled = True
    'End If
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    'gMsg = ""
    'For Each gErrSQL In cnn.Errors  'rdoErrors
    '    If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
    '        gMsg = "A SQL error has occured in Relay Definition-Form Load: "
    '        MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
    '    End If
    'Next gErrSQL
    'If (Err.Number <> 0) And (gMsg = "") Then
    '    gMsg = "A general error has occured in Relay Definition-Form Load: "
    '    MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    'End If
    gHandleError "EngrErrors.Txt", "Schedule Definition-Form Load"
    Resume Next
End Sub

Private Sub grdLibEvents_Click()
    If grdLibEvents.Col >= grdLibEvents.Cols - 1 Then
        Exit Sub
    End If

End Sub

Private Sub grdLibEvents_EnterCell()
    mESetShow
End Sub

Private Sub grdLibEvents_GotFocus()
    If grdLibEvents.Col >= grdLibEvents.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdLibEvents_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdLibEvents.TopRow
    grdLibEvents.Redraw = False
End Sub

Private Sub grdLibEvents_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilFound As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    
    grdLibEvents.ToolTipText = ""
    If (y > grdLibEvents.RowHeight(0)) And (y < grdLibEvents.RowHeight(0) + grdLibEvents.RowHeight(1)) Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdLibEvents, x, y, llRow, llCol)
    If (ilFound) And (llCol = EVENTIDINDEX) Then
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, LIBNAMEINDEX))
        grdLibEvents.TextMatrix(llRow, LIBNAMEINDEX) = mGetLibName(slStr)
        grdLibEvents.ToolTipText = Trim$(grdLibEvents.TextMatrix(llRow, LIBNAMEINDEX))
    Else
        grdLibEvents.ToolTipText = Trim$(grdLibEvents.TextMatrix(llRow, llCol))
    End If
End Sub

Private Sub grdLibEvents_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    If bmInSave Then
        grdLibEvents.Redraw = True
        Exit Sub
    End If
    If (grdLibEvents.Row < grdLibEvents.FixedRows) Or (grdLibEvents.Row >= grdLibEvents.Rows) Then
        grdLibEvents.Redraw = True
        Exit Sub
    End If
    'Determine if in header
    If y < grdLibEvents.RowHeight(0) Then
        mSortCol grdLibEvents.Col
        Exit Sub
    End If
    If (y > grdLibEvents.RowHeight(0)) And (y < grdLibEvents.RowHeight(0) + grdLibEvents.RowHeight(1)) Then
        mSortCol grdLibEvents.Col
        Exit Sub
    End If
    'ilFound = gGrid_DetermineRowCol(grdLibEvents, x, y)
    'If Not ilFound Then
    '    grdLibEvents.Redraw = True
    '    pbcClickFocus.SetFocus
    '    Exit Sub
    'End If
    If grdLibEvents.Col >= grdLibEvents.Cols - 1 Then
        grdLibEvents.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdLibEvents.TopRow
    DoEvents
    llRow = grdLibEvents.Row
    If grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = "" Then
        grdLibEvents.Redraw = False
        Do
            llRow = llRow - 1
        Loop While (grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = "") And (llRow > grdLibEvents.FixedRows - 1)
        grdLibEvents.Row = llRow + 1
        grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
        grdLibEvents.Col = EVENTTYPEINDEX
        grdLibEvents.Redraw = True
    End If
    grdLibEvents.Redraw = True
    '8/26/11: Check that row is not behind scroll bar
    If grdLibEvents.RowPos(grdLibEvents.Row) + grdLibEvents.RowHeight(grdLibEvents.Row) + 60 >= grdLibEvents.Height Then
        imIgnoreScroll = True
        grdLibEvents.TopRow = grdLibEvents.TopRow + 1
    End If
    If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
        mEEnableBox
    Else
        Beep
        'pbcClickFocus.SetFocus
        lmEEnableRow = grdLibEvents.Row
        mShowConflictGrid
    End If
End Sub

Private Sub grdLibEvents_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdLibEvents.Redraw = False Then
        grdLibEvents.Redraw = True
        If lmTopRow < grdLibEvents.FixedRows Then
            grdLibEvents.TopRow = grdLibEvents.FixedRows
        Else
            grdLibEvents.TopRow = lmTopRow
        End If
        grdLibEvents.Refresh
        grdLibEvents.Redraw = False
    End If
    If (imShowGridBox) And (grdLibEvents.Row >= grdLibEvents.FixedRows) And (grdLibEvents.Col >= 0) And (grdLibEvents.Col < grdLibEvents.Cols - 1) Then
        If (grdLibEvents.RowIsVisible(grdLibEvents.Row)) And (grdLibEvents.ColIsVisible(grdLibEvents.Col)) Then
            pbcArrow.Move grdLibEvents.Left - pbcArrow.Width - 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + (grdLibEvents.RowHeight(grdLibEvents.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mShowConflictGrid
            mESetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            pbcEDefine.Visible = False
            edcEDropdown.Visible = False
            cmcEDropDown.Visible = False
            lbcBDE.Visible = False
            lbcBDE.Visible = False
            lbcCCE_B.Visible = False
            lbcETE_Program.Visible = False
            lbcETE.Visible = False
            edcEvent.Visible = False
            lbcTTE_S.Visible = False
            lbcTTE_E.Visible = False
            pbcYN.Visible = False
            lbcMTE.Visible = False
            lbcASE.Visible = False
            lbcCCE_A.Visible = False
            lbcANE.Visible = False
            lbcRNE.Visible = False
            lbcFNE.Visible = False
            lbcSCE.Visible = False
            lbcNNE.Visible = False
            lbcCTE_1.Visible = False
            lbcCTE_2.Visible = False
            '7/8/11: Make T2 work like T1
            'lbcCTE_2.Visible = False
            ltcEvent.Visible = False
            pbcArrow.Visible = False
            'mHideConflictGrid
        End If
    Else
        If Not grdConflicts.Visible Then
            pbcClickFocus.SetFocus
        End If
        pbcArrow.Visible = False
        'mHideConflictGrid
        imFromArrow = False
    End If

End Sub

Private Sub imcInsert_Click()
    If bmInSave Then
        Exit Sub
    End If
    bmInInsert = True
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
    mInsertRow
    bmInInsert = False
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lbcKey.Visible = False
End Sub

'
'               Snapshot of the Schedule events
'         Generate list of schedule events from the grid on the  screen
'
Private Sub imcPrint_Click()
    Dim ilRptDest As Integer            'disply, print, save as file
    Dim slRptName As String
    Dim slExportName As String
    Dim slRptType As String
    Dim llResult As Long
    Dim ilExportType As Integer
    Dim llGridRow As Long
    Dim slStr As String
    Dim llTime As Long
    Dim llAirDate As Long
    Dim slFilter As String              'filters selected by user
    Dim ilLoop As Integer
    Dim slOperator As String * 2        'operator for filter
    
    If bmInSave Then
        Exit Sub
    End If
    If bmPrinting Then
        Exit Sub
    End If
    bmPrinting = True
    igRptIndex = SCHED_RPT
    igRptSource = vbModal
    slRptName = "Sched.rpt"      'concatenate the crystal report name plus extension

    slExportName = ""               'no export for now
    
    Set rstSchedRpt = New Recordset
    gGenerateRstSchedule     'generate the ddfs for report
    
    rstSchedRpt.Open
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstSchedRpt, sgDBPath & "\SchedRpt.ttx", True)
    
    If smAirDate = "" Then
        bmPrinting = False
        MsgBox "Enter valid date"
        Exit Sub
    End If
    
    llAirDate = gDateValue(smAirDate)
    'loop thru the ItemID grid to print whats shown on the screen
    For llGridRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1
        slStr = Trim$(grdLibEvents.TextMatrix(llGridRow, EVENTTYPEINDEX))
        If slStr = "" Then
            Exit For
        Else
            rstSchedRpt.AddNew
            
            rstSchedRpt.Fields("StartDateSort") = llAirDate             'schedule date
            rstSchedRpt.Fields("EventType") = Left(grdLibEvents.TextMatrix(llGridRow, EVENTTYPEINDEX), 1)   'program, spot, avail
            rstSchedRpt.Fields("Event ID") = grdLibEvents.TextMatrix(llGridRow, EVENTIDINDEX)               'event ID
            rstSchedRpt.Fields("EvBusName") = grdLibEvents.TextMatrix(llGridRow, BUSNAMEINDEX)              'Bus name
            rstSchedRpt.Fields("EvBusCtl") = grdLibEvents.TextMatrix(llGridRow, BUSCTRLINDEX)                'Bus Control index
            rstSchedRpt.Fields("EvStarttime") = grdLibEvents.TextMatrix(llGridRow, TIMEINDEX)               'Event start time
            slStr = grdLibEvents.TextMatrix(llGridRow, TIMEINDEX)
            llTime = gStrTimeInTenthToLong(slStr, False)                'convert the start time of event to long for sorting
            rstSchedRpt.Fields("EvStartTimeSort") = llTime
            rstSchedRpt.Fields("EvStartType") = grdLibEvents.TextMatrix(llGridRow, STARTTYPEINDEX)          'start type
            rstSchedRpt.Fields("EvFix") = grdLibEvents.TextMatrix(llGridRow, FIXEDINDEX)                    'Fixed type
            rstSchedRpt.Fields("EvEndType") = grdLibEvents.TextMatrix(llGridRow, ENDTYPEINDEX)              'end type
            rstSchedRpt.Fields("EvDur") = grdLibEvents.TextMatrix(llGridRow, DURATIONINDEX)                 'duration
            rstSchedRpt.Fields("EvMatType") = grdLibEvents.TextMatrix(llGridRow, MATERIALINDEX)             'material type
            rstSchedRpt.Fields("EvAudName1") = grdLibEvents.TextMatrix(llGridRow, AUDIONAMEINDEX)           'primary audio name
            rstSchedRpt.Fields("EvItem1") = grdLibEvents.TextMatrix(llGridRow, AUDIOITEMIDINDEX)            'primary audio item id
            'rstSchedRpt.Fields("EvISCI1") = grdLibEvents.TextMatrix(llGridRow, AUDIOISCIINDEX)            'primary audio item id
            rstSchedRpt.Fields("EvCtl1") = grdLibEvents.TextMatrix(llGridRow, AUDIOCTRLINDEX)               'primary audio control
            rstSchedRpt.Fields("EvAudName2") = grdLibEvents.TextMatrix(llGridRow, BACKUPNAMEINDEX)          'backup audio name
            rstSchedRpt.Fields("EvCtl2") = grdLibEvents.TextMatrix(llGridRow, BACKUPCTRLINDEX)              'back control char
            rstSchedRpt.Fields("EvAudName3") = grdLibEvents.TextMatrix(llGridRow, PROTNAMEINDEX)            'protection audio name
            rstSchedRpt.Fields("EvItem3") = grdLibEvents.TextMatrix(llGridRow, PROTITEMIDINDEX)             'protection item id
            'rstSchedRpt.Fields("EvISCI3") = grdLibEvents.TextMatrix(llGridRow, PROTISCIINDEX)             'protection item id
            rstSchedRpt.Fields("EvCtl3") = grdLibEvents.TextMatrix(llGridRow, PROTCTRLINDEX)                'protection control
            rstSchedRpt.Fields("EvRelay1") = grdLibEvents.TextMatrix(llGridRow, RELAY1INDEX)                'relay 1 of 2
            rstSchedRpt.Fields("EvRelay2") = grdLibEvents.TextMatrix(llGridRow, RELAY2INDEX)                'relay 2 of 2
            rstSchedRpt.Fields("EvFollow") = grdLibEvents.TextMatrix(llGridRow, FOLLOWINDEX)                'follow name
            rstSchedRpt.Fields("EvSilenceTime") = grdLibEvents.TextMatrix(llGridRow, SILENCETIMEINDEX)      'silence time
            rstSchedRpt.Fields("EvSilence1") = grdLibEvents.TextMatrix(llGridRow, SILENCE1INDEX)            'silence name 1 of 4
            rstSchedRpt.Fields("EvSilence2") = grdLibEvents.TextMatrix(llGridRow, SILENCE2INDEX)            'silence name 2 of 4
            rstSchedRpt.Fields("EvSilence3") = grdLibEvents.TextMatrix(llGridRow, SILENCE3INDEX)            'silence name 3 of 4
            rstSchedRpt.Fields("EvSilence4") = grdLibEvents.TextMatrix(llGridRow, SILENCE4INDEX)            'silence name 4 of 4
            rstSchedRpt.Fields("EvNetCue1") = grdLibEvents.TextMatrix(llGridRow, NETCUE1INDEX)              'netcue name 1 of 2
            rstSchedRpt.Fields("EvNetCue2") = grdLibEvents.TextMatrix(llGridRow, NETCUE2INDEX)              'netcue name 2 of 2
            rstSchedRpt.Fields("EvTitle1") = grdLibEvents.TextMatrix(llGridRow, TITLE1INDEX)               'title 1 of 2
            rstSchedRpt.Fields("EvTitle2") = grdLibEvents.TextMatrix(llGridRow, TITLE2INDEX)               'title 2 of 2
            'rstSchedRpt.Fields("ABCFormat") = grdLibEvents.TextMatrix(llGridRow, ABCFORMATINDEX)
            'rstSchedRst.Fields("ABCPgmCode") = grdLibEvents.TextMatrix(llGridRow, ABCPGMCODEINDEX)
            'rstSchedRst.Fields("ABCXDSMode") = grdLibEvents.TextMatrix(llGridRow, ABCXDSMODEINDEX)
            'rstSchedRst.Fields("ABCRecordItem") = grdLibEvents.TextMatrix(llGridRow, ABCRECORDITEMINDEX)
            rstSchedRpt.Fields("EvABCCustomFields") = ""
            If sgClientFields = "A" Then            'abc client
                If Trim$(grdLibEvents.TextMatrix(llGridRow, ABCFORMATINDEX)) <> "" Or Trim$(grdLibEvents.TextMatrix(llGridRow, ABCPGMCODEINDEX)) <> "" Or Trim$(grdLibEvents.TextMatrix(llGridRow, ABCXDSMODEINDEX)) <> "" Or Trim$(grdLibEvents.TextMatrix(llGridRow, ABCRECORDITEMINDEX)) <> "" Then
                    rstSchedRpt.Fields("EvABCCustomFields") = "ABC: Format-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCFORMATINDEX)) & ", Pgm Code-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCPGMCODEINDEX)) & ", XDS Mode-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCXDSMODEINDEX)) & ", Record Item-" & Trim$(grdLibEvents.TextMatrix(llGridRow, ABCRECORDITEMINDEX))
                End If
            End If
        End If
    Next llGridRow
    
    slFilter = ""
    For ilLoop = 0 To UBound(tgFilterValues) - 1
        If slFilter <> "" Then               'not first time
            slFilter = slFilter & ", "
        End If
        If tgFilterValues(ilLoop).iOperator = 1 Then
            slOperator = "="
        ElseIf tgFilterValues(ilLoop).iOperator = 2 Then
            slOperator = "<>"
         ElseIf tgFilterValues(ilLoop).iOperator = 3 Then
            slOperator = ">"
        ElseIf tgFilterValues(ilLoop).iOperator = 4 Then
            slOperator = "<"
        ElseIf tgFilterValues(ilLoop).iOperator = 5 Then
            slOperator = ">="
        ElseIf tgFilterValues(ilLoop).iOperator = 6 Then
            slOperator = "<="
        End If
        slFilter = slFilter & Trim$(tgFilterValues(ilLoop).sFieldName) & " " & slOperator & " " & Trim$(tgFilterValues(ilLoop).sValue)
    
    Next ilLoop
    'pass the description of filters selected
    If slFilter = "" Then
        slFilter = "Filters: all events"
    Else
        slFilter = "Filters: " & slFilter
    End If
    
    sgCrystlFormula2 = "'" & Format$(llAirDate, "ddddd") & "'"         'air date for heading
    sgCrystlFormula3 = "'" & Trim$(slFilter) & "'"             'filter selected for heading

    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    igRptSource = vbModal
        
    EngrSchedRpt.Show vbModal
    
    'determine how the user responded, either cancel or produce output
    If igReturnCallStatus = CALLDONE Then           'produce the report flag
    
        slExportName = sgReturnCallName     'if exporting path and filename, this is filename; otherwise blank
        slRptType = ""
        'determine which version (condensed # of fields or all fields)
        If sgReturnOption = "ALL" Then
            slRptName = "SchedAll.rpt"
            EngrCrystal.gActiveCrystalReports igExportType, igRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstSchedRpt
        Else
            slRptName = "Sched.rpt"
            EngrCrystal.gActiveCrystalReports igExportType, igRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstSchedRpt
        End If
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    Set rstSchedRpt = Nothing
    cmcCancel.SetFocus
    bmPrinting = False
    Exit Sub
End Sub

Private Sub imcTrash_Click()
    If bmInSave Then
        Exit Sub
    End If
    mESetShow
    mHideConflictGrid
    mDeleteRow
End Sub




Private Sub lbcANE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcANE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcANE.Visible = False
End Sub

Private Sub lbcANE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcANE, y)
    If (llRow < lbcANE.ListCount) And (lbcANE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcANE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If ilCode = tgCurrANE(ilLoop).iCode Then
            ilLoop = gBinarySearchANE(ilCode, tgCurrANE())
            If ilLoop <> -1 Then
                lbcANE.ToolTipText = Trim$(tgCurrANE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub lbcASE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcASE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcASE.Visible = False
End Sub

Private Sub lbcASE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcASE, y)
    If (llRow < lbcASE.ListCount) And (lbcASE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcASE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
        '    If ilCode = tgCurrASE(ilLoop).iCode Then
            ilLoop = gBinarySearchASE(ilCode, tgCurrASE())
            If ilLoop <> -1 Then
                lbcASE.ToolTipText = Trim$(tgCurrASE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub


Private Sub lbcBDE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcBDE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcBDE.Visible = False
End Sub

Private Sub lbcBDE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcBDE, y)
    If (llRow < lbcBDE.ListCount) And (lbcBDE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcBDE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If ilCode = tgCurrBDE(ilLoop).iCode Then
            ilLoop = gBinarySearchBDE(ilCode, tgCurrBDE())
            If ilLoop <> -1 Then
                lbcBDE.ToolTipText = Trim$(tgCurrBDE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub lbcCCE_A_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcCCE_A_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcCCE_A.Visible = False
End Sub

Private Sub lbcCCE_A_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcCCE_A, y)
    If (llRow < lbcCCE_A.ListCount) And (lbcCCE_A.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcCCE_A.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If ilCode = tgCurrAudioCCE(ilLoop).iCode Then
                lbcCCE_A.ToolTipText = Trim$(tgCurrAudioCCE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcCCE_B_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcCCE_B_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcCCE_B.Visible = False
End Sub

Private Sub lbcCCE_B_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcCCE_B, y)
    If (llRow < lbcCCE_B.ListCount) And (lbcCCE_B.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcCCE_B.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            If ilCode = tgCurrBusCCE(ilLoop).iCode Then
                lbcCCE_B.ToolTipText = Trim$(tgCurrBusCCE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcCTE_1_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcCTE_1.List(lbcCTE_1.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcCTE_1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim llRow As Long
'    Dim llCode As Long
'    Dim ilLoop As Integer
    
'    llRow = gGetListBoxRow(lbcCTE_2, y)
'    If llRow <= lbcCTE_2.ListCount Then
'        llCode = lbcCTE_2.ItemData(llRow)
'        For ilLoop = 0 To UBound(tgCurrCTE) - 1 Step 1
'            If llCode = tgCurrCTE(ilLoop).lCode Then
'                lbcCTE_2.ToolTipText = Trim$(tgCurrCTE(ilLoop).sComment)
'                Exit For
'            End If
'        Next ilLoop
'    End If
End Sub

Private Sub lbcCTE_2_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If

End Sub

Private Sub lbcCTE_2_DblClick()
    '7/8/11: Make T2 work like T1
    'tmcClick.Enabled = False
    'Sleep 300
    'DoEvents
    'edcEDropdown.SetFocus
    'imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    'edcEDropdown_MouseUp 0, 0, 0, 0
    'lbcCTE_2.Visible = False
End Sub

Private Sub lbcCTE_2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '7/8/11: Make T2 work Like T1
    'Dim llRow As Long
    'Dim llCode As Long
    'Dim ilLoop As Integer
    
    'llRow = gGetListBoxRow(lbcCTE_2, y)
    'If (llRow < lbcCTE_2.ListCount) And (lbcCTE_2.ListCount > 0) And (llRow <> -1) Then
    '    llCode = lbcCTE_2.ItemData(llRow)
    '    For ilLoop = 0 To UBound(tgCurrCTE) - 1 Step 1
    '        If llCode = tgCurrCTE(ilLoop).lCode Then
    '            lbcCTE_2.ToolTipText = Trim$(tgCurrCTE(ilLoop).sComment)
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If
End Sub


Private Sub lbcETE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcETE.List(lbcETE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcETE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcETE.Visible = False
End Sub

Private Sub lbcETE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcETE, y)
    If (llRow < lbcETE.ListCount) And (lbcETE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcETE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
            If ilCode = tgCurrETE(ilLoop).iCode Then
                lbcETE.ToolTipText = Trim$(tgCurrETE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If

End Sub

Private Sub lbcETE_Program_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcETE_Program.List(lbcETE_Program.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcETE_Program_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcETE_Program.Visible = False
End Sub

Private Sub lbcETE_Program_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcETE_Program, y)
    If (llRow < lbcETE_Program.ListCount) And (lbcETE_Program.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcETE_Program.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
            If ilCode = tgCurrETE(ilLoop).iCode Then
                lbcETE_Program.ToolTipText = Trim$(tgCurrETE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcFNE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcFNE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcFNE.Visible = False
End Sub

Private Sub lbcFNE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcFNE, y)
    If (llRow < lbcFNE.ListCount) And (lbcFNE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcFNE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrFNE) - 1 Step 1
            If ilCode = tgCurrFNE(ilLoop).iCode Then
                lbcFNE.ToolTipText = Trim$(tgCurrFNE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcMTE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcMTE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcMTE.Visible = False
End Sub

Private Sub lbcMTE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcMTE, y)
    If (llRow < lbcMTE.ListCount) And (lbcMTE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcMTE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrMTE) - 1 Step 1
            If ilCode = tgCurrMTE(ilLoop).iCode Then
                lbcMTE.ToolTipText = Trim$(tgCurrMTE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcNNE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcNNE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcNNE.Visible = False
End Sub

Private Sub lbcNNE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcNNE, y)
    If (llRow < lbcNNE.ListCount) And (lbcNNE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcNNE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If ilCode = tgCurrNNE(ilLoop).iCode Then
            ilLoop = gBinarySearchNNE(ilCode, tgCurrNNE())
            If ilLoop <> -1 Then
                lbcNNE.ToolTipText = Trim$(tgCurrNNE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub lbcRNE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcRNE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcRNE.Visible = False
End Sub

Private Sub lbcRNE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcRNE, y)
    If (llRow < lbcRNE.ListCount) And (lbcRNE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcRNE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If ilCode = tgCurrRNE(ilLoop).iCode Then
            ilLoop = gBinarySearchRNE(ilCode, tgCurrRNE())
            If ilLoop <> -1 Then
                lbcRNE.ToolTipText = Trim$(tgCurrRNE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub lbcSCE_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcSCE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcSCE.Visible = False
End Sub

Private Sub lbcSCE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcSCE, y)
    If (llRow < lbcSCE.ListCount) And (lbcSCE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcSCE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
            If ilCode = tgCurrSCE(ilLoop).iCode Then
                lbcSCE.ToolTipText = Trim$(tgCurrSCE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcTTE_E_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcTTE_E.List(lbcTTE_E.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcTTE_E_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcTTE_E.Visible = False
End Sub

Private Sub lbcTTE_E_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcTTE_E, y)
    If (llRow < lbcTTE_E.ListCount) And (lbcTTE_E.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcTTE_E.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            If ilCode = tgCurrEndTTE(ilLoop).iCode Then
                lbcTTE_E.ToolTipText = Trim$(tgCurrEndTTE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcTTE_S_Click()
    tmcClick.Enabled = False
    edcEDropdown.text = lbcTTE_S.List(lbcTTE_S.ListIndex)
    If (edcEDropdown.Visible) And (edcEDropdown.Enabled) Then
        edcEDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcTTE_S_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcEDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcEDropdown_MouseUp 0, 0, 0, 0
    lbcTTE_S.Visible = False
End Sub

Private Sub lbcTTE_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcTTE_S, y)
    If (llRow < lbcTTE_S.ListCount) And (lbcTTE_S.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcTTE_S.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            If ilCode = tgCurrStartTTE(ilLoop).iCode Then
                lbcTTE_S.ToolTipText = Trim$(tgCurrStartTTE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub ltcEvent_OnChange()
    Dim slStr As String
    
    slStr = ltcEvent.text
    If grdLibEvents.text <> slStr Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        Select Case grdLibEvents.Col
            Case TIMEINDEX
            Case DURATIONINDEX
            Case SILENCETIMEINDEX
        End Select
        grdLibEvents.text = slStr
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcClickFocus_GotFocus()
    mESetShow
    lmEEnableRow = -1
    lmEEnableCol = -1
    mHideConflictGrid
End Sub




Private Sub pbcEDefine_Click()
'    Dim ilRet As Integer
'    ilRet = mEBranch()
'    pbcEDefine.SetFocus
End Sub

Private Sub pbcEDefine_Paint()
    pbcEDefine.CurrentX = 30
    pbcEDefine.CurrentY = 0
    pbcEDefine.Print "Multi-Select"
End Sub

Private Sub pbcESTab_GotFocus()
    Dim ilPrev As Integer
    
    If imStartChgModeCompleted = False Then
        Exit Sub
    End If
    If bmInBranch Then
        Exit Sub
    End If
    If bmInInsert Then
        Exit Sub
    End If
    If GetFocus() <> pbcESTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEEnableBox
        Exit Sub
    End If
    If pbcHighlight.Visible Or edcEvent.Visible Or edcEDropdown.Visible Or pbcYN.Visible Or pbcEDefine.Visible Or ltcEvent.Visible Then
'        If Not lbcBDE.Visible Then
'            If Not mEBranch() Then
'                mEEnableBox
'                bmInBranch=False
'                Exit Sub
'            End If
'        End If
        bmInBranch = False
        mESetShow
        Do
            ilPrev = False
            If grdLibEvents.Col = EVENTTYPEINDEX Then
                If grdLibEvents.Row > grdLibEvents.FixedRows Then
                    lmTopRow = -1
                    grdLibEvents.Row = grdLibEvents.Row - 1
                    If Not grdLibEvents.RowIsVisible(grdLibEvents.Row) Then
                        grdLibEvents.TopRow = grdLibEvents.TopRow - 1
                    End If
                    grdLibEvents.Col = imMaxCols    'TITLE2INDEX
                    If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
                        mEEnableBox
                    Else
                        ilPrev = True
                    End If
                Else
                    cmcCancel.SetFocus
                    Exit Do
                End If
            Else
                If grdLibEvents.Col <= HIGHLIGHTINDEX Then
                    cmcCancel.SetFocus
                    Exit Do
                End If
                If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
                    mEEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdLibEvents.TopRow = grdLibEvents.FixedRows
        grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
        grdLibEvents.Col = EVENTTYPEINDEX
        grdLibEvents.Row = grdLibEvents.FixedRows
        If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
            mEEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Sub pbcETab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEEnableRow As Long
    
    If bmInBranch Then
        Exit Sub
    End If
    If bmInInsert Then
        Exit Sub
    End If
    If GetFocus() <> pbcETab.hwnd Then
        Exit Sub
    End If
    If pbcHighlight.Visible Or edcEvent.Visible Or edcEDropdown.Visible Or pbcYN.Visible Or pbcEDefine.Visible Or ltcEvent.Visible Then
        '1/19/12: Reinstalled as a test
'        If Not lbcBDE.Visible Then
'            If Not mEBranch() Then
'                mEEnableBox
'                bmInBranch = False
'                Exit Sub
'            End If
'        End If
        If Not lbcBDE.Visible Then
            If Not mEBranch() Then
                mEEnableBox
                bmInBranch = False
                Exit Sub
            End If
        End If
        bmInBranch = False
        llEEnableRow = lmEEnableRow
        mESetShow
        Do
            ilNext = False
            If grdLibEvents.Col = imMaxCols Then
                llRow = grdLibEvents.Rows
                Do
                    llRow = llRow - 1
                Loop While grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = ""
                llRow = llRow + 1
                If (grdLibEvents.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdLibEvents.Row = grdLibEvents.Row + 1
                    If Not grdLibEvents.RowIsVisible(grdLibEvents.Row) Then
                        imIgnoreScroll = True
                        grdLibEvents.TopRow = grdLibEvents.TopRow + 1
                    End If
                    '8/26/11: Check that row is not behind scroll bar
                    If grdLibEvents.RowPos(grdLibEvents.Row) + grdLibEvents.RowHeight(grdLibEvents.Row) + 60 >= grdLibEvents.Height Then
                        imIgnoreScroll = True
                        grdLibEvents.TopRow = grdLibEvents.TopRow + 1
                    End If
                    grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
                    grdLibEvents.Col = EVENTTYPEINDEX
                    DoEvents
                    'grdLibEvents.TextMatrix(grdLibEvents.Row, CODEINDEX) = 0
                    If Trim$(grdLibEvents.TextMatrix(grdLibEvents.Row, EVENTTYPEINDEX)) <> "" Then
                        If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
                            mEEnableBox
                        Else
                            ilNext = True
                        End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdLibEvents.Left - pbcArrow.Width - 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + (grdLibEvents.RowHeight(grdLibEvents.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        mShowConflictGrid
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdLibEvents.TextMatrix(llEEnableRow, EVENTTYPEINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdLibEvents.Row + 1 >= grdLibEvents.Rows Then
                            grdLibEvents.AddItem ""
                        End If
                        grdLibEvents.Row = grdLibEvents.Row + 1
                        If Not grdLibEvents.RowIsVisible(grdLibEvents.Row) Then
                            imIgnoreScroll = True
                            grdLibEvents.TopRow = grdLibEvents.TopRow + 1
                        End If
                        '8/26/11: Check that row is not behind scroll bar
                        If grdLibEvents.RowPos(grdLibEvents.Row) + grdLibEvents.RowHeight(grdLibEvents.Row) + 60 >= grdLibEvents.Height Then
                            imIgnoreScroll = True
                            grdLibEvents.TopRow = grdLibEvents.TopRow + 1
                        End If
                        grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
                        grdLibEvents.Col = EVENTTYPEINDEX
                        DoEvents
                        grdLibEvents.TextMatrix(grdLibEvents.Row, PCODEINDEX) = 0
                        imFromArrow = True
                        pbcArrow.Move grdLibEvents.Left - pbcArrow.Width - 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + (grdLibEvents.RowHeight(grdLibEvents.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        mShowConflictGrid
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdLibEvents.Col = grdLibEvents.Col + 1
                If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
                    mEEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdLibEvents.TopRow = grdLibEvents.FixedRows
        grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
        grdLibEvents.Col = EVENTTYPEINDEX
        DoEvents
        grdLibEvents.Row = grdLibEvents.FixedRows
        If mColOk(grdLibEvents.Row, grdLibEvents.Col, True) Then
            mEEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub


Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    Dim llFirstInsertRow As Long
    Dim ilInsert As Integer
    Dim ilCol As Integer
    Dim llUpper As Long
    
    llTRow = grdLibEvents.TopRow
    llRow = grdLibEvents.Row
    'slMsg = "Insert above selected Row"
    'If MsgBox(slMsg, vbYesNo) = vbNo Then
    '    mInsertRow = False
    '    Exit Function
    'End If
    sgGenMsg = "Duplicate Contents from Select Event?"
    sgCMCTitle(0) = "Yes"
    sgCMCTitle(1) = "No"
    sgCMCTitle(2) = "Cancel"
    sgCMCTitle(3) = ""
    igDefCMC = 1
    igEditBox = 1
    sgMsgEditValue = "Insert How Many Events:"
    sgEditValue = "1"
    On Error Resume Next
    EngrGenMsg.Show vbModal
    If igAnsCMC = 2 Then
        mInsertRow = False
        Exit Function
    End If
    grdLibEvents.Redraw = False
    llFirstInsertRow = llRow + 1
    For ilInsert = 1 To Val(sgEditValue) Step 1
        llRow = grdLibEvents.Row + 1
        grdLibEvents.AddItem "", llRow '& vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
        If igAnsCMC = 0 Then
            For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                If ilCol <> EVENTIDINDEX Then
                    grdLibEvents.TextMatrix(llRow, ilCol) = grdLibEvents.TextMatrix(llFirstInsertRow - 1, ilCol)
                Else
                    grdLibEvents.TextMatrix(llRow, ilCol) = ""
                End If
            Next ilCol
        End If
        grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
        grdLibEvents.Row = llRow
        grdLibEvents.Col = EVENTIDINDEX
        grdLibEvents.CellBackColor = LIGHTYELLOW
        grdLibEvents.CellAlignment = flexAlignRightCenter
        llUpper = UBound(tmCurrSEE)
        ReDim Preserve tmCurrSEE(0 To llUpper + 1) As SEE
        ReDim Preserve smT1Comment(0 To llUpper + 1) As String
        ReDim Preserve smT2Comment(0 To llUpper + 1) As String
        grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llUpper))
        grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
        grdLibEvents.TextMatrix(llRow, DEECODEINDEX) = "0"
        grdLibEvents.TextMatrix(llRow, EVTCONFLICTINDEX) = "N"
        gInitSEE tmCurrSEE(llUpper)
        smT1Comment(llUpper) = ""
        smT2Comment(llUpper) = ""
        grdLibEvents.Redraw = False
    Next ilInsert
    
'    grdLibEvents.AddItem "", llRow '& vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
'    grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
'    grdLibEvents.Row = llRow
'    grdLibEvents.Col = EVENTIDINDEX
'    grdLibEvents.CellBackColor = LIGHTYELLOW
'    grdLibEvents.CellAlignment = flexAlignRightCenter
'    llUpper = UBound(tmCurrSEE)
'    ReDim Preserve tmCurrSEE(0 To llUpper + 1) As SEE
'    ReDim Preserve smT1Comment(0 To llUpper + 1) As String
'    ReDim Preserve smT2Comment(0 To llUpper + 1) As String
'    grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llUpper))
'    grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
'    grdLibEvents.TextMatrix(llRow, DEECODEINDEX) = "0"
'    grdLibEvents.TextMatrix(llRow, EVTCONFLICTINDEX) = "N"
'    gInitSEE tmCurrSEE(llUpper)
'    smT1Comment(llUpper) = ""
'    smT2Comment(llUpper) = ""
    grdLibEvents.Redraw = False
    grdLibEvents.TopRow = llTRow
    grdLibEvents.Redraw = True
    'DoEvents
    imInsertState = True
    'lmInsertRow = grdLibEvents.Row
    lmInsertRow = llFirstInsertRow
    grdLibEvents.Row = lmInsertRow
    grdLibEvents.LeftCol = HIGHLIGHTINDEX   'EVENTTYPEINDEX
    grdLibEvents.Col = EVENTTYPEINDEX
    mEEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    Dim llIndex As Long
    Dim llLoop As Long
    Dim llDelRow As Long
    Dim ilCol As Integer
    Dim slStr As String
    Dim ilSpotEvent As Integer
    Dim ilETE As Integer
    Dim llAvailTime As Long
    Dim llSpotDuration As Long
    Dim slAvailBus As String
    Dim llSpotTime As Long
    Dim llDelSpotTime As Long
    Dim slEventCategory As String
    Dim ilAvailAdj As Integer
    Dim ilBDE As Integer
    Dim llSEE As Long
    
    'Disallow deletion of Avails and Spot
    'Code related to removing spots retained in case we need to allow its deletion
    llRow = grdLibEvents.Row
    slEventCategory = ""
    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                slEventCategory = tgCurrETE(ilETE).sCategory
                '2/9/12: Allow spots to be removed
                'If (slEventCategory = "S") Or (slEventCategory = "A") Then
                If (slEventCategory = "A") Then
                    Beep
                    'MsgBox "Only Program events can be Removed"
                    MsgBox "Only Program and Spot events can be Removed"
                    mDeleteRow = False
                    Exit Function
                End If
            End If
        Next ilETE
    End If
    If tmSHE.lCode > 0 Then
        llSEE = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
        'If tmCurrSEE(llRow).lCode > 0 Then
        If tmCurrSEE(llSEE).lCode > 0 Then
            If Not mBusInFilter() Then
                Beep
                MsgBox "Events can only be deleted when Bus selected in Filter"
                mDeleteRow = False
                Exit Function
            End If
        End If
    End If
    llTRow = grdLibEvents.TopRow
    slMsg = "Delete selected Row"
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    grdLibEvents.Redraw = False
    ilSpotEvent = False
    'If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
    '    slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
    '    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
    '        If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
    '            If tgCurrETE(ilETE).sCategory = "S" Then
                If slEventCategory = "S" Then
                    ilSpotEvent = True
                    slStr = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
                    llAvailTime = gStrTimeInTenthToLong(slStr, False)
                    slAvailBus = Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX))
                    slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
                    llDelSpotTime = gStrLengthInTenthToLong(slStr)
                    slStr = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
                    llSpotDuration = gStrLengthInTenthToLong(slStr)
                End If
    '            Exit For
    '        End If
    '    Next ilETE
    'End If
    If (Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) <> 0) Then
        lmDeleteCodes(UBound(lmDeleteCodes)) = Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
        ReDim Preserve lmDeleteCodes(0 To UBound(lmDeleteCodes) + 1) As Long
    End If
    If Trim$(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX)) <> "" Then
        llIndex = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
        For llLoop = llIndex + 1 To UBound(tmCurrSEE) - 1 Step 1
            LSet tmCurrSEE(llLoop - 1) = tmCurrSEE(llLoop)
            smT1Comment(llLoop - 1) = smT1Comment(llLoop)
            smT2Comment(llLoop - 1) = smT2Comment(llLoop)
        Next llLoop
        ReDim Preserve tmCurrSEE(LBound(tmCurrSEE) To UBound(tmCurrSEE) - 1) As SEE
        ReDim Preserve smT1Comment(LBound(smT1Comment) To UBound(smT1Comment) - 1) As String
        ReDim Preserve smT2Comment(LBound(smT2Comment) To UBound(smT2Comment) - 1) As String
        For llDelRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            If (Trim$(grdLibEvents.TextMatrix(llDelRow, EVENTTYPEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(llDelRow, TMCURRSEEINDEX)) <> "") Then
                If Val(grdLibEvents.TextMatrix(llDelRow, TMCURRSEEINDEX)) > llIndex Then
                    grdLibEvents.TextMatrix(llDelRow, TMCURRSEEINDEX) = Val(grdLibEvents.TextMatrix(llDelRow, TMCURRSEEINDEX)) - 1
                End If
            End If
        Next llDelRow
    End If
    grdLibEvents.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Adjust time of any other spot and Avail Duration
    'If avail not showing, then tmCurrSEE would need to be adjusted
    If slEventCategory = "S" Then
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            If (Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "") And (Trim$(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX)) <> "") Then
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                        If tgCurrETE(ilETE).sCategory = "S" Then
                            slStr = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
                            If (llAvailTime = gStrTimeInTenthToLong(slStr, False)) And (StrComp(slAvailBus, Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX)), vbTextCompare) = 0) Then
                                slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
                                llSpotTime = gStrTimeInTenthToLong(slStr, False)
                                If llDelSpotTime < llSpotTime Then
                                    llSpotTime = llSpotTime - llSpotDuration
                                    grdLibEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(llSpotTime)
                                    grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                                End If
                            End If
                        End If
                        Exit For
                    End If
                Next ilETE
            End If
        Next llRow
    End If
    'Add row at bottom in case less rows showing then room in grid
    If ilSpotEvent Then
        mMoveSEECtrlsToRec
        ''Adjust Avail times (Duration and Running Time)
        'For ilAvailAdj = LBound(tmCurrSEE) To UBound(tmCurrSEE) - 1 Step 1
        '    slEventCategory = ""
        '    ilETE = gBinarySearchETE(tmCurrSEE(ilAvailAdj).iEteCode, tgCurrETE)
        '    If ilETE <> -1 Then
        '        slEventCategory = tgCurrETE(ilETE).sCategory
        '        If slEventCategory = "A" Then
        '            If (tmCurrSEE(ilAvailAdj).lTime = llAvailTime) Then
        '                'Test Bus
        '                'ilBDE = gBinarySearchBDE(tmCurrSEE(llLoop).iBdeCode, tgCurrBDE)
        '                ilBDE = gBinarySearchBDE(tmCurrSEE(ilAvailAdj).iBdeCode, tgCurrBDE)
        '                If ilBDE <> -1 Then
        '                    slStr = Trim$(tgCurrBDE(ilBDE).sName)
        '                    If StrComp(slStr, slAvailBus, vbTextCompare) = 0 Then
        '                        'tmCurrSEE(ilAvailAdj).lAvailLength = tmCurrSEE(ilAvailAdj).lAvailLength + llDelSpotTime
        '                        'tmCurrSEE(ilAvailAdj).lSpotTime = tmCurrSEE(ilAvailAdj).lSpotTime - llDelSpotTime
        '                        tmCurrSEE(ilAvailAdj).lAvailLength = tmCurrSEE(ilAvailAdj).lAvailLength + llSpotDuration
        '                        tmCurrSEE(ilAvailAdj).lSpotTime = tmCurrSEE(ilAvailAdj).lSpotTime + llSpotDuration  '- llDelSpotTime
        '                        Exit For
        '                    End If
        '                End If
        '            End If
        '        End If
        '    End If
        'Next ilAvailAdj
        mAdjustAvailTime 0
        grdLibEvents.Redraw = False
        grdLibEvents.Visible = False
        mMoveSEERecToCtrls
        grdLibEvents.Redraw = False
        If imLastColSorted >= 0 Then
            If imLastSort = flexSortStringNoCaseDescending Then
                imLastSort = flexSortStringNoCaseAscending
            Else
                imLastSort = flexSortStringNoCaseDescending
            End If
            ilCol = imLastColSorted
            mSortCol ilCol
        Else
            imLastSort = -1
            mSortCol TIMEINDEX
        End If
        grdLibEvents.TopRow = llTRow
        grdLibEvents.Redraw = True
        grdLibEvents.Visible = True
    Else
        grdLibEvents.AddItem ""
        grdLibEvents.Redraw = False
        grdLibEvents.TopRow = llTRow
    End If
    grdLibEvents.Redraw = True
    DoEvents
    'grdLibEvents.Col = CATEGORYINDEX
    'mEnableBox
    cmcCancel.SetFocus
    mSetCommands
    mDeleteRow = True
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
End Function



Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
'    If igInitCallInfo = 0 Then
'        If UBound(tgCurrDHE) > 0 Then
'            cmcCancel.SetFocus
'        End If
'        Exit Sub
'    End If
'    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
'        For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
'            slStr = Trim$(grdLib.TextMatrix(llRow, NAMEINDEX))
'            If (slStr <> "") Then
'                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
'                    grdLib.Row = llRow
'                    Do While Not grdLib.RowIsVisible(grdLib.Row)
'                        grdLib.TopRow = grdLib.TopRow + 1
'                    Loop
'                    grdLib.Col = NAMEINDEX
'                    mEnableBox
'                    Exit Sub
'                End If
'            End If
'        Next llRow
'    End If
'    If (Not ilCreateNew) Or (Not cmcDone.Enabled) Then
'        Exit Sub
'    End If
'    'Find first blank row
'    For llRow = grdLib.FixedRows To grdLib.Rows - 1 Step 1
'        slStr = Trim$(grdLib.TextMatrix(llRow, CATEGORYINDEX))
'        If (slStr = "") Then
'            grdLib.Row = llRow
'            Do While Not grdLib.RowIsVisible(grdLib.Row)
'                grdLib.TopRow = grdLib.TopRow + 1
'            Loop
'            grdLib.Col = CATEGORYINDEX
'            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
'                grdLib.TextMatrix(llRow, NAMEINDEX) = sgInitCallName
'            End If
'            mEnableBox
'            Exit Sub
'        End If
'    Next llRow
    
End Sub


Private Sub mMoveSEERecToCtrls()
    Dim llRow As Long
    Dim slStr As String
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim ilEBE As Integer
    Dim ilBDE As Integer
    Dim ilCCE As Integer
    Dim ilETE As Integer
    Dim ilTTE As Integer
    Dim ilMTE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilRNE As Integer
    Dim ilFNE As Integer
    Dim ilSCE As Integer
    Dim ilNNE As Integer
    Dim llCTE As Long
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slHours As String
    Dim llRet As Long
    Dim ilRowOk As Integer
    Dim slCategory As String
    Dim llAvailLength As Long
    Dim llTest As Long
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilCol As Integer
    Dim llChg As Long
    Dim llARE As Long
    Dim blBusInFilter As Boolean
    
    mClearControls
    If smAirDate = "" Then
        Exit Sub
    End If
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    blBusInFilter = mBusInFilter()
    llRow = grdLibEvents.FixedRows
    If UBound(tmCurrSEE) > grdLibEvents.FixedRows Then
        grdLibEvents.Rows = UBound(tmCurrSEE)
    End If
    For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
        slCategory = ""
        'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
        ilETE = gBinarySearchETE(tmCurrSEE(llLoop).iEteCode, tgCurrETE)
        If ilETE <> -1 Then
            slCategory = tgCurrETE(ilETE).sCategory
        End If
        '        Exit For
        '    End If
        'Next ilETE
        'If Avail, then check if ant time left
        ilRowOk = True
        If slCategory = "A" Then
            'Duration adjusted in pbcDateTab as the events for in time order
            'Avails and Spots can't be deleted
            'Avails can't be altered by the user
            'Spot duration can't be altered by the user (if later we need to allow the spot duration to be altered
            'then the avail duration can be adjusted
            llAvailLength = tmCurrSEE(llLoop).lAvailLength
            'For llTest = 0 To UBound(tmCurrSEE) - 1 Step 1
            '    If (tmCurrSEE(llLoop).iBdeCode = tmCurrSEE(llTest).iBdeCode) And (tmCurrSEE(llLoop).lTime = tmCurrSEE(llTest).lTime) And (tmCurrSEE(llTest).iEteCode = imSpotETECode) Then
            '        llAvailLength = llAvailLength - tmCurrSEE(llTest).lDuration
            '    End If
            'Next llTest
            If llAvailLength <= 0 Then
                ilRowOk = False
            End If
        End If
        'Add Criteria test here
        If ilRowOk Then
            ilRowOk = mCheckFilter(tmCurrSEE(llLoop), smT1Comment(llLoop))
        End If
        If ilRowOk Then
            If llRow + 1 > grdLibEvents.Rows Then
                grdLibEvents.AddItem ""
            End If
            grdLibEvents.Row = llRow
            'If avail, then set all columns to yellow
            '7/15/11: BecauseLibraries can only be modified after schedule days, allow Bus and Time change
            If slCategory = "A" Then
                ''For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                ''    grdLibEvents.Col = ilCol
                ''    grdLibEvents.CellBackColor = LIGHTYELLOW
                ''Next ilCol
                For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                    ''If (ilCol <> AUDIONAMEINDEX) And (ilCol <> AUDIOITEMIDINDEX) And (ilCol <> AUDIOISCIINDEX) And (ilCol <> AUDIOCTRLINDEX) _
                    ''    And (ilCol <> BACKUPNAMEINDEX) And (ilCol <> BACKUPCTRLINDEX) _
                    ''    And (ilCol <> PROTNAMEINDEX) And (ilCol <> PROTITEMIDINDEX) And (ilCol <> PROTISCIINDEX) And (ilCol <> PROTCTRLINDEX) _
                    ''    And (ilCol <> TIMEINDEX) And (ilCol <> DURATIONINDEX) _
                    ''    And (ilCol <> TITLE1INDEX) And (ilCol <> TITLE2INDEX) Then
                    'If (ilCol = EVENTTYPEINDEX) Or (ilCol = EVENTIDINDEX) Or (ilCol = BUSNAMEINDEX) Or (ilCol = BUSCTRLINDEX) Or (ilCol = TIMEINDEX) Then
                    If (ilCol = EVENTTYPEINDEX) Or (ilCol = EVENTIDINDEX) Then  'Or (ilCol = BUSNAMEINDEX) Or (ilCol = BUSCTRLINDEX) Or (ilCol = TIMEINDEX) Then
                            grdLibEvents.Col = ilCol
                            grdLibEvents.CellBackColor = LIGHTYELLOW
                    End If
                Next ilCol
            ElseIf tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                    'Allow all changes
                    'If (ilCol <> AUDIONAMEINDEX) And (ilCol <> AUDIOITEMIDINDEX) And (ilCol <> AUDIOISCIINDEX) And (ilCol <> AUDIOCTRLINDEX) _
                    '    And (ilCol <> BACKUPNAMEINDEX) And (ilCol <> BACKUPCTRLINDEX) _
                    '    And (ilCol <> PROTNAMEINDEX) And (ilCol <> PROTITEMIDINDEX) And (ilCol <> PROTISCIINDEX) And (ilCol <> PROTCTRLINDEX) _
                    '    And (ilCol <> TITLE1INDEX) And (ilCol <> TITLE2INDEX) Then
                    '        grdLibEvents.Col = ilCol
                    '        grdLibEvents.CellBackColor = LIGHTYELLOW
                    'End If
                    If (ilCol = EVENTTYPEINDEX) Or (ilCol = EVENTIDINDEX) Then  'Or (ilCol = BUSNAMEINDEX) Or (ilCol = BUSCTRLINDEX) Or (ilCol = TIMEINDEX) Then
                            grdLibEvents.Col = ilCol
                            grdLibEvents.CellBackColor = LIGHTYELLOW
                    End If
                Next ilCol
            End If
            grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llLoop))
            If tmCurrSEE(llLoop).lEventID > 0 Then
                grdLibEvents.TextMatrix(llRow, EVENTIDINDEX) = Trim$(Str$(tmCurrSEE(llLoop).lEventID))
            Else
                grdLibEvents.TextMatrix(llRow, EVENTIDINDEX) = ""
            End If
            
'            slStr = ""
'            If tmCurrSEE(llLoop).lDeeCode > 0 Then
'                ilRet = gGetRec_DEE_DayEvent(tmCurrSEE(llLoop).lDeeCode, "EngrSchdDef-mMoveSEERecToCtrls: DEE", tmDee)
'                ilRet = gGetRec_DHE_DayHeaderInfo(tmDee.lDheCode, "EngrSchdDef-mMoveSEERecToCtrls: DHE", tmDHE)
'
'                If tmDHE.sType <> "T" Then
'                    'For ilDNE = 0 To UBound(tgCurrLibDNE) - 1 Step 1
'                    '    If tmDHE.lDneCode = tgCurrLibDNE(ilDNE).lCode Then
'                    ilDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrLibDNE)
'                    If ilDNE <> -1 Then
'                        slStr = Trim$(tgCurrLibDNE(ilDNE).sName)
'                    End If
'                    '        Exit For
'                    '    End If
'                    'Next ilDNE
'                Else
'                    'For ilDNE = 0 To UBound(tgCurrTempDNE) - 1 Step 1
'                    '    If tmDHE.lDneCode = tgCurrTempDNE(ilDNE).lCode Then
'                    ilDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrTempDNE)
'                    If ilDNE <> -1 Then
'                        slStr = Trim$(tgCurrTempDNE(ilDNE).sName)
'                    End If
'                    '        Exit For
'                    '    End If
'                    'Next ilDNE
'                End If
'                'For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
'                '    If tmDHE.lDSECode = tgCurrDSE(ilDSE).lCode Then
'                ilDSE = gBinarySearchDSE(tmDHE.lDSECode, tgCurrDSE)
'                If ilDSE <> -1 Then
'                    slStr = slStr & "/" & Trim$(tgCurrDSE(ilDSE).sName)
'                End If
'                '        Exit For
'                '    End If
'                'Next ilDSE
'                If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
'                    slStr = slStr & "/" & gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
'                End If
'            End If
            slStr = ""
            If tmCurrSEE(llLoop).lDeeCode > 0 Then
                slStr = "DEE=" & tmCurrSEE(llLoop).lDeeCode
                If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                    slStr = slStr & "/" & gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                End If
            End If
            grdLibEvents.TextMatrix(llRow, LIBNAMEINDEX) = slStr
    
            
            slStr = ""
            'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iBdeCode = tgCurrBDE(ilBDE).iCode Then
            ilBDE = gBinarySearchBDE(tmCurrSEE(llLoop).iBdeCode, tgCurrBDE)
            If ilBDE <> -1 Then
                slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName)
            End If
            '        Exit For
            '    End If
            'Next ilBDE
            grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX) = slStr
            '11/7/11: Disallow Bus to be altered if Load created and Bus not part of Filter
            If tmSHE.sLoadedAutoStatus = "L" Then
                If Not blBusInFilter Then
                    grdLibEvents.Col = BUSNAMEINDEX
                    grdLibEvents.CellBackColor = LIGHTYELLOW
                End If
            End If
            grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX) = ""
            'For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iBusCceCode = tgCurrBusCCE(ilCCE).iCode Then
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iBusCceCode, tgCurrBusCCE)
            If ilCCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX) = Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_B.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrBusCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_B.AddItem Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
                    lbcCCE_B.ItemData(lbcCCE_B.NewIndex) = tgCurrBusCCE(ilCCE).iCode
                End If
            End If
            '       Exit For
            '    End If
            'Next ilCCE
            grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = ""
            'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
            ilETE = gBinarySearchETE(tmCurrSEE(llLoop).iEteCode, tgCurrETE)
            If ilETE <> -1 Then
                grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) = Trim$(tgCurrETE(ilETE).sName)
            End If
            '        Exit For
            '    End If
            'Next ilETE
            If tmCurrSEE(llLoop).iEteCode <> imSpotETECode Then
                If slCategory = "A" Then
                    If tmCurrSEE(llLoop).lSpotTime <> tmCurrSEE(llLoop).lTime Then
                        grdLibEvents.Col = TIMEINDEX
                        grdLibEvents.CellBackColor = LIGHTYELLOW
                    End If
                    grdLibEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lSpotTime)
                    grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                Else
                    grdLibEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                    grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                End If
            Else
                grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lTime)
                '2/11/12: Allow time to be altered
                'grdLibEvents.Col = TIMEINDEX
                'grdLibEvents.CellBackColor = LIGHTYELLOW
                grdLibEvents.TextMatrix(llRow, TIMEINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llLoop).lSpotTime)
            End If
            grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = ""
            'For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
            ilTTE = gBinarySearchTTE(tmCurrSEE(llLoop).iStartTteCode, tgCurrStartTTE)
            If ilTTE <> -1 Then
                grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX) = Trim$(tgCurrStartTTE(ilTTE).sName)
                llRet = SendMessageByString(lbcTTE_S.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrStartTTE(ilTTE).sName))
                If llRet < 0 Then
                    lbcTTE_S.AddItem Trim$(tgCurrStartTTE(ilTTE).sName)
                    lbcTTE_S.ItemData(lbcTTE_S.NewIndex) = tgCurrStartTTE(ilTTE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilTTE
            grdLibEvents.TextMatrix(llRow, FIXEDINDEX) = Trim$(tmCurrSEE(llLoop).sFixedTime)
            grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = ""
            'For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
            ilTTE = gBinarySearchTTE(tmCurrSEE(llLoop).iEndTteCode, tgCurrEndTTE)
            If ilTTE <> -1 Then
                grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX) = Trim$(tgCurrEndTTE(ilTTE).sName)
                llRet = SendMessageByString(lbcTTE_E.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrEndTTE(ilTTE).sName))
                If llRet < 0 Then
                    lbcTTE_E.AddItem Trim$(tgCurrEndTTE(ilTTE).sName)
                    lbcTTE_E.ItemData(lbcTTE_E.NewIndex) = tgCurrEndTTE(ilTTE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilTTE
            '11/24/04- Allow end type and Duration to co-exist
            'If (tmCurrSEE(llLoop).lDuration > 0) Or (Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX)) = "") Then
            If slCategory = "A" Then
                If (tmCurrSEE(llLoop).lDuration > 0) Then
                    grdLibEvents.TextMatrix(llRow, AVAILDURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llLoop).lDuration, True)
                Else
                    grdLibEvents.TextMatrix(llRow, AVAILDURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llLoop).lDuration, True)   '""
                End If
                '6/7/13: Change Open avails to include Hours
                'grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(llAvailLength, False)
                grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(llAvailLength, True)
            Else
                If (tmCurrSEE(llLoop).lDuration > 0) Then
                    grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llLoop).lDuration, True)
                Else
                    grdLibEvents.TextMatrix(llRow, DURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llLoop).lDuration, True)    '""
                End If
                grdLibEvents.TextMatrix(llRow, AVAILDURATIONINDEX) = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
            End If
            If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                '2/11/12: Allow duration to be altered
                'grdLibEvents.Col = DURATIONINDEX
                'grdLibEvents.CellBackColor = LIGHTYELLOW
            End If
            grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = ""
            'For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iMteCode = tgCurrMTE(ilMTE).iCode Then
            ilMTE = gBinarySearchMTE(tmCurrSEE(llLoop).iMteCode, tgCurrMTE)
            If ilMTE <> -1 Then
                grdLibEvents.TextMatrix(llRow, MATERIALINDEX) = Trim$(tgCurrMTE(ilMTE).sName)
                llRet = SendMessageByString(lbcMTE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrMTE(ilMTE).sName))
                If llRet < 0 Then
                    lbcMTE.AddItem Trim$(tgCurrMTE(ilMTE).sName)
                    lbcMTE.ItemData(lbcMTE.NewIndex) = tgCurrMTE(ilMTE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilMTE
            grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = ""
            'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASE).iCode Then
            ilASE = gBinarySearchASE(tmCurrSEE(llLoop).iAudioAseCode, tgCurrASE())
            If ilASE <> -1 Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                    llRet = SendMessageByString(lbcASE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrANE(ilANE).sName))
                    If llRet < 0 Then
                        lbcASE.AddItem Trim$(tgCurrANE(ilANE).sName)
                        lbcASE.ItemData(lbcASE.NewIndex) = tgCurrASE(ilASE).iCode
                    End If
                End If
                    '    End If
                    'Next ilANE
            End If
            '        Exit For
            '    End If
            'Next ilASE
            grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX) = Trim$(tmCurrSEE(llLoop).sAudioItemID)
            If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                grdLibEvents.Row = llRow
                grdLibEvents.Col = AUDIOITEMIDINDEX
                If tmCurrSEE(llLoop).sAudioItemIDChk = "F" Then
                    grdLibEvents.CellForeColor = vbRed
                ElseIf tmCurrSEE(llLoop).sAudioItemIDChk = "O" Then
                    grdLibEvents.CellForeColor = vbGreen
                Else
                    If tmCurrSEE(llLoop).lAreCode > 0 Then
                        grdLibEvents.CellForeColor = vbMagenta  'vbBlue
                    Else
                        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                            grdLibEvents.CellForeColor = vbBlue
                        Else
                            grdLibEvents.CellForeColor = vbBlack
                        End If
                    End If
                End If
            End If
            grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX) = Trim$(tmCurrSEE(llLoop).sAudioISCI)
            grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = ""
            'For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iAudioCceCode, tgCurrAudioCCE)
            If ilCCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrAudioCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilCCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilCCE
            'grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = ""
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iBkupAneCode, tgCurrANE())
            If ilANE <> -1 Then
                grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                llRet = SendMessageByString(lbcANE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrANE(ilANE).sName))
                If llRet < 0 Then
                    lbcANE.AddItem Trim$(tgCurrANE(ilANE).sName)
                    lbcANE.ItemData(lbcANE.NewIndex) = tgCurrANE(ilANE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilANE
            grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = ""
            'For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iBkupCceCode, tgCurrAudioCCE)
            If ilCCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrAudioCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilCCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilCCE
            grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = ""
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iProtAneCode, tgCurrANE())
            If ilANE <> -1 Then
                grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX) = Trim$(tgCurrANE(ilANE).sName)
                llRet = SendMessageByString(lbcANE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrANE(ilANE).sName))
                If llRet < 0 Then
                    lbcANE.AddItem Trim$(tgCurrANE(ilANE).sName)
                    lbcANE.ItemData(lbcANE.NewIndex) = tgCurrANE(ilANE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilANE
            grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX) = Trim$(tmCurrSEE(llLoop).sProtItemID)
            If tmCurrSEE(llLoop).iEteCode = imSpotETECode Then
                grdLibEvents.Row = llRow
                grdLibEvents.Col = PROTITEMIDINDEX
                If tmCurrSEE(llLoop).sProtItemIDChk = "F" Then
                    grdLibEvents.CellForeColor = vbRed
                ElseIf tmCurrSEE(llLoop).sProtItemIDChk = "O" Then
                    grdLibEvents.CellForeColor = vbGreen
                Else
                    If tmCurrSEE(llLoop).lAreCode > 0 Then
                        grdLibEvents.CellForeColor = vbMagenta  'vbBlue
                    Else
                        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                            grdLibEvents.CellForeColor = vbBlue
                        Else
                            grdLibEvents.CellForeColor = vbBlack
                        End If
                    End If
                End If
            End If
            grdLibEvents.TextMatrix(llRow, PROTISCIINDEX) = Trim$(tmCurrSEE(llLoop).sProtISCI)
            grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = ""
            'For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            ilCCE = gBinarySearchCCE(tmCurrSEE(llLoop).iProtCceCode, tgCurrAudioCCE)
            If ilCCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrAudioCCE(ilCCE).sAutoChar))
                If llRet < 0 Then
                    lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilCCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilCCE
            grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = ""
            'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
            '    If tmCurrSEE(llLoop).i1RneCode = tgCurrRNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i1RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, RELAY1INDEX) = Trim$(tgCurrRNE(ilRNE).sName)
                llRet = SendMessageByString(lbcRNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrRNE(ilRNE).sName))
                If llRet < 0 Then
                    lbcRNE.AddItem Trim$(tgCurrRNE(ilRNE).sName)
                    lbcRNE.ItemData(lbcRNE.NewIndex) = tgCurrRNE(ilRNE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilRNE
            grdLibEvents.TextMatrix(llRow, RELAY2INDEX) = ""
            'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
            '    If tmCurrSEE(llLoop).i2RneCode = tgCurrRNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i2RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, RELAY2INDEX) = Trim$(tgCurrRNE(ilRNE).sName)
                llRet = SendMessageByString(lbcRNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrRNE(ilRNE).sName))
                If llRet < 0 Then
                    lbcRNE.AddItem Trim$(tgCurrRNE(ilRNE).sName)
                    lbcRNE.ItemData(lbcRNE.NewIndex) = tgCurrRNE(ilRNE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilRNE
            grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = ""
            'For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iFneCode = tgCurrFNE(ilFNE).iCode Then
            ilFNE = gBinarySearchFNE(tmCurrSEE(llLoop).iFneCode, tgCurrFNE)
            If ilFNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, FOLLOWINDEX) = Trim$(tgCurrFNE(ilFNE).sName)
                llRet = SendMessageByString(lbcFNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrFNE(ilFNE).sName))
                If llRet < 0 Then
                    lbcFNE.AddItem Trim$(tgCurrFNE(ilFNE).sName)
                    lbcFNE.ItemData(lbcFNE.NewIndex) = tgCurrFNE(ilFNE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilFNE
            If tmCurrSEE(llLoop).lSilenceTime > 0 Then
                grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = gLongToLength(tmCurrSEE(llLoop).lSilenceTime, False)   'gLongToStrLengthInTenth(tmCurrSEE(llLoop).lSilenceTime, False)
            Else
                grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX) = ""
            End If
            grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = ""
            'For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).i1SceCode = tgCurrSCE(ilSCE).iCode Then
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i1SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, SILENCE1INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilSCE
            grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = ""
            'For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).i2SceCode = tgCurrSCE(ilSCE).iCode Then
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i2SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, SILENCE2INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilSCE
            grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = ""
            'For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).i3SceCode = tgCurrSCE(ilSCE).iCode Then
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i3SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, SILENCE3INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilSCE
            grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = ""
            'For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            '    If tmCurrSEE(llLoop).i4SceCode = tgCurrSCE(ilSCE).iCode Then
            ilSCE = gBinarySearchSCE(tmCurrSEE(llLoop).i4SceCode, tgCurrSCE)
            If ilSCE <> -1 Then
                grdLibEvents.TextMatrix(llRow, SILENCE4INDEX) = Trim$(tgCurrSCE(ilSCE).sAutoChar)
                llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrSCE(ilSCE).sAutoChar))
                If llRet < 0 Then
                    lbcSCE.AddItem Trim$(tgCurrSCE(ilSCE).sAutoChar)
                    lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilSCE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilSCE
            grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = ""
            'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iStartNneCode, tgCurrNNE)
            If ilNNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, NETCUE1INDEX) = Trim$(tgCurrNNE(ilNNE).sName)
                llRet = SendMessageByString(lbcNNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrNNE(ilNNE).sName))
                If llRet < 0 Then
                    lbcNNE.AddItem Trim$(tgCurrNNE(ilNNE).sName)
                    lbcNNE.ItemData(lbcNNE.NewIndex) = tgCurrNNE(ilNNE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilNNE
            grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = ""
            'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
            '    If tmCurrSEE(llLoop).iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iEndNneCode, tgCurrNNE)
            If ilNNE <> -1 Then
                grdLibEvents.TextMatrix(llRow, NETCUE2INDEX) = Trim$(tgCurrNNE(ilNNE).sName)
                llRet = SendMessageByString(lbcNNE.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrNNE(ilNNE).sName))
                If llRet < 0 Then
                    lbcNNE.AddItem Trim$(tgCurrNNE(ilNNE).sName)
                    lbcNNE.ItemData(lbcNNE.NewIndex) = tgCurrNNE(ilNNE).iCode
                End If
            End If
            '        Exit For
            '    End If
            'Next ilNNE
            If tmCurrSEE(llLoop).iEteCode <> imSpotETECode Then
                grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = smT1Comment(llLoop)
            Else
                'ilRet = gGetRec_ARE_AdvertiserRefer(tmCurrSEE(llLoop).lAreCode, "EngrSchdDef-mMoveSEERecToCtrls: Advertiser", tmARE)
                'slStr = Trim$(tmARE.sName)
                llARE = gBinarySearchARE(tmCurrSEE(llLoop).lAreCode, tgCurrARE())
                If llARE <> -1 Then
                    grdLibEvents.TextMatrix(llRow, TITLE1INDEX) = Trim$(tgCurrARE(llARE).sName)
                End If
            End If
            grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = ""
            '7/8/11: Make T2 work like T1
            ''For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
            ''    If tmCurrSEE(llLoop).l2CteCode = tgCurrCTE(ilCTE).lCode Then
            'llCTE = gBinarySearchCTE(tmCurrSEE(llLoop).l2CteCode, tgCurrCTE)
            'If llCTE <> -1 Then
            '    grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = Trim$(tgCurrCTE(llCTE).sName)
            '    llRet = SendMessageByString(lbcCTE_2.hwnd, LB_FINDSTRINGEXACT, -1, Trim$(tgCurrCTE(llCTE).sName))
            '    If llRet < 0 Then
            '        lbcCTE_2.AddItem Trim$(tgCurrCTE(llCTE).sName)
            '        lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = tgCurrCTE(llCTE).lCode
            '    End If
            'End If
            grdLibEvents.TextMatrix(llRow, TITLE2INDEX) = smT2Comment(llLoop)
            '        Exit For
            '    End If
            'Next ilCTE
            If sgClientFields = "A" Then
                grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = Trim$(tmCurrSEE(llLoop).sABCFormat)
                grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = Trim$(tmCurrSEE(llLoop).sABCPgmCode)
                grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = Trim$(tmCurrSEE(llLoop).sABCXDSMode)
                grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = Trim$(tmCurrSEE(llLoop).sABCRecordItem)
            Else
                grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX) = ""
                grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX) = ""
                grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX) = ""
                grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX) = ""
            End If
            grdLibEvents.TextMatrix(llRow, EVTCONFLICTINDEX) = tmCurrSEE(llLoop).sIgnoreConflicts
            grdLibEvents.TextMatrix(llRow, DEECODEINDEX) = tmCurrSEE(llLoop).lDeeCode
            grdLibEvents.TextMatrix(llRow, PCODEINDEX) = tmCurrSEE(llLoop).lCode
            If tmCurrSEE(llLoop).lCode <= 0 Then
                'Insert flag only set on for inserted rows or changed row of new scheduled.
                'Bypass testing of spots even if just Merged as Bus and audio with avails has been checked
                'Remove setting of change flag for spots
                'If (StrComp(Trim$(tmCurrSEE(llLoop).sInsertFlag), "Y", vbTextCompare) = 0) Or (tmCurrSEE(llLoop).iEteCode = imSpotETECode) Then
                If (StrComp(Trim$(tmCurrSEE(llLoop).sInsertFlag), "Y", vbTextCompare) = 0) Then
                    grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                Else
                    grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "N"
                End If
            Else
                grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "N"
                For llChg = 0 To UBound(lmChgStatusSEECode) - 1 Step 1
                    If lmChgStatusSEECode(llChg) = tmCurrSEE(llLoop).lCode Then
                        grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                        Exit For
                    End If
                Next llChg
            End If
            mSetColExportColor llRow
            llRow = llRow + 1
        End If
    Next llLoop
    ReDim lmChgStatusSEECode(0 To 0) As Long
    If (llRow > grdLibEvents.FixedRows) Then
        grdLibEvents.Rows = llRow
        If (llRow >= grdLibEvents.Rows) Then
            grdLibEvents.AddItem ""
        End If
        If llAirDate < llNowDate Then
            'Setting the background color is only setting rows not visible
            grdLibEvents.BackColor = LIGHTYELLOW
            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                grdLibEvents.Row = llRow
                grdLibEvents.Col = EVENTIDINDEX
                grdLibEvents.CellAlignment = flexAlignRightCenter
                For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                    grdLibEvents.Col = ilCol
                    grdLibEvents.CellBackColor = LIGHTYELLOW
                Next ilCol
            Next llRow
        Else
            For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                grdLibEvents.Row = llRow
                grdLibEvents.Col = EVENTIDINDEX
                grdLibEvents.CellBackColor = LIGHTYELLOW
                grdLibEvents.CellAlignment = flexAlignRightCenter
                If llAirDate = llNowDate Then
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, TIMEINDEX))
                    If slStr <> "" Then
                        If gIsTimeTenths(slStr) Then
                            If (llNowTime > gStrTimeInTenthToLong(slStr, False)) Then
                                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
                                If (slStr = "") Or (Val(slStr) = 0) Then
    '                                If tmSHE.lCode <= 0 Then
    '                                    For ilCol = EVENTTYPEINDEX To TITLE2INDEX Step 1
    '                                        grdLibEvents.Col = ilCol
    '                                        grdLibEvents.CellBackColor = LIGHTYELLOW
    '                                    Next ilCol
    '                                Else
    '                                    grdLibEvents.Col = TIMEINDEX
    '                                    grdLibEvents.CellForeColor = vbRed
    '                                End If
                                Else
                                    For ilCol = EVENTTYPEINDEX To imMaxCols Step 1
                                        grdLibEvents.Col = ilCol
                                        grdLibEvents.CellBackColor = LIGHTYELLOW
                                    Next ilCol
                                End If
                            End If
                        End If
                    End If
                End If
            Next llRow
        End If
    End If
    '8/26/11:  Moved Integral here in addition to ColumnWidth
    If Not bmIntegralSet Then
        bmIntegralSet = True
        gGrid_IntegralHeight grdLibEvents
        gGrid_FillWithRows grdLibEvents
        grdLibEvents.Height = grdLibEvents.Height '+ 30
    End If
End Sub

Private Sub mMoveSEECtrlsToRec()
    Dim llIndex As Long
    Dim llRow As Long
    Dim ilEBE As Integer
    Dim ilBDE As Integer
    Dim ilCCE As Integer
    Dim ilETE As Integer
    Dim ilTTE As Integer
    Dim ilMTE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilRNE As Integer
    Dim ilFNE As Integer
    Dim ilSCE As Integer
    Dim ilNNE As Integer
    Dim llCTE As Long
    Dim slStr As String
    Dim ilDays As Integer
    Dim ilHours As Integer
    Dim ilSet As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilPos As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim llSEEOld As Long
    Dim llUpper As Long
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llRet As Long
    Dim slEventCategory As String
    
    ReDim lmChgStatusSEECode(0 To 0)
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
            If Trim$(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX)) = "" Then
                llUpper = UBound(tmCurrSEE)
                ReDim Preserve tmCurrSEE(0 To llUpper + 1) As SEE
                ReDim Preserve smT1Comment(0 To llUpper + 1) As String
                ReDim Preserve smT2Comment(0 To llUpper + 1) As String
                grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX) = Trim$(Str$(llUpper))
                gInitSEE tmCurrSEE(llUpper)
                smT1Comment(llUpper) = ""
                smT2Comment(llUpper) = ""
            End If
            llIndex = Val(grdLibEvents.TextMatrix(llRow, TMCURRSEEINDEX))
            'Get Category- if avails bypass updating of image
            slEventCategory = ""
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            If slStr <> "" Then
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                        slEventCategory = tgCurrETE(ilETE).sCategory
                    End If
                Next ilETE
            End If
            If (grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y") Or (tmSHE.lCode <= 0) Then
            '7/15/11: Allow Bus, Bus Ctrl and Time for avails to be altered
            'If slEventCategory <> "A" Then
                'Set Later- Bus selected
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BUSNAMEINDEX))
                ilBDE = gBinarySearchBDE(tmCurrSEE(llIndex).iBdeCode, tgCurrBDE)
                If ilBDE <> -1 Then
                    If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iBdeCode = 0
                    llRet = SendMessageByString(lbcBDE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iBdeCode = lbcBDE.ItemData(llRet)
                    Else
                        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slStr, vbTextCompare) = 0 Then
                            ilBDE = gBinarySearchName(slStr, tgCurrBDE_Name())
                            If ilBDE <> -1 Then
                                tmCurrSEE(llIndex).iBdeCode = tgCurrBDE_Name(ilBDE).iCode   'tgCurrBDE(ilBDE).iCode
                        '        Exit For
                            End If
                        'Next ilBDE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BUSCTRLINDEX))
                ilCCE = gBinarySearchCCE(tmCurrSEE(llIndex).iBusCceCode, tgCurrBusCCE)
                If ilCCE <> -1 Then
                    If StrComp(Trim$(tgCurrBusCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iBusCceCode = 0
                    llRet = SendMessageByString(lbcCCE_B.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iBusCceCode = lbcCCE_B.ItemData(llRet)
                    Else
                        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrBusCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iBusCceCode = tgCurrBusCCE(ilCCE).iCode
                                Exit For
                            End If
                        Next ilCCE
                    End If
                End If
            'If slEventCategory <> "A" Then
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
                ilETE = gBinarySearchETE(tmCurrSEE(llIndex).iEteCode, tgCurrETE)
                If ilETE <> -1 Then
                    If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iEteCode = 0
                    llRet = SendMessageByString(lbcETE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iEteCode = lbcETE.ItemData(llRet)
                    Else
                        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iEteCode = tgCurrETE(ilETE).iCode
                                Exit For
                            End If
                        Next ilETE
                    End If
                End If
            'End If
                slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
                If (tmCurrSEE(llIndex).iEteCode = imSpotETECode) Then
                    tmCurrSEE(llIndex).lSpotTime = gStrTimeInTenthToLong(slStr, False)
                    If (grdLibEvents.TextMatrix(llRow, DEECODEINDEX) <> "") And (grdLibEvents.TextMatrix(llRow, DEECODEINDEX) <> "0") Then
                        slStr = grdLibEvents.TextMatrix(llRow, SPOTAVAILTIMEINDEX)
                        tmCurrSEE(llIndex).lTime = gStrTimeInTenthToLong(slStr, False)
                    Else
                        tmCurrSEE(llIndex).lTime = tmCurrSEE(llIndex).lSpotTime
                    End If
                Else
                    If slEventCategory <> "A" Then
                        tmCurrSEE(llIndex).lTime = gStrTimeInTenthToLong(slStr, False)
                        tmCurrSEE(llIndex).lSpotTime = -1
                    Else
                        grdLibEvents.Col = TIMEINDEX
                        If grdLibEvents.CellBackColor <> LIGHTYELLOW Then
                            tmCurrSEE(llIndex).lSpotTime = gStrTimeInTenthToLong(slStr, False)
                            tmCurrSEE(llIndex).lTime = gStrTimeInTenthToLong(slStr, False)
                        End If
                    End If
                End If
            '7/15/11
            'End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, STARTTYPEINDEX))
                ilTTE = gBinarySearchTTE(tmCurrSEE(llIndex).iStartTteCode, tgCurrStartTTE)
                If ilTTE <> -1 Then
                    If StrComp(Trim$(tgCurrStartTTE(ilTTE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iStartTteCode = 0
                    llRet = SendMessageByString(lbcTTE_S.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iStartTteCode = lbcTTE_S.ItemData(llRet)
                    Else
                        For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
                            If StrComp(Trim$(tgCurrStartTTE(ilTTE).sName), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iStartTteCode = tgCurrStartTTE(ilTTE).iCode
                                Exit For
                            End If
                        Next ilTTE
                    End If
                End If
                tmCurrSEE(llIndex).sFixedTime = grdLibEvents.TextMatrix(llRow, FIXEDINDEX)
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, ENDTYPEINDEX))
                ilTTE = gBinarySearchTTE(tmCurrSEE(llIndex).iEndTteCode, tgCurrEndTTE)
                If ilTTE <> -1 Then
                    If StrComp(Trim$(tgCurrEndTTE(ilTTE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iEndTteCode = 0
                    llRet = SendMessageByString(lbcTTE_E.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iEndTteCode = lbcTTE_E.ItemData(llRet)
                    Else
                        For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
                            If StrComp(Trim$(tgCurrEndTTE(ilTTE).sName), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iEndTteCode = tgCurrEndTTE(ilTTE).iCode
                                Exit For
                            End If
                        Next ilTTE
                    End If
                End If
                slStr = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
                tmCurrSEE(llIndex).lDuration = gStrLengthInTenthToLong(slStr)
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, MATERIALINDEX))
                ilMTE = gBinarySearchMTE(tmCurrSEE(llIndex).iMteCode, tgCurrMTE)
                If ilMTE <> -1 Then
                    If StrComp(Trim$(tgCurrMTE(ilMTE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iMteCode = 0
                    llRet = SendMessageByString(lbcMTE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iMteCode = lbcMTE.ItemData(llRet)
                    Else
                        For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
                            If StrComp(Trim$(tgCurrMTE(ilMTE).sName), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iMteCode = tgCurrMTE(ilMTE).iCode
                                Exit For
                            End If
                        Next ilMTE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIONAMEINDEX))
                ilASE = gBinarySearchASE(tmCurrSEE(llIndex).iAudioAseCode, tgCurrASE())
                If ilASE <> -1 Then
                    ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                            ilFound = True
                        End If
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iAudioAseCode = 0
                    llRet = SendMessageByString(lbcASE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iAudioAseCode = lbcASE.ItemData(llRet)
                        ilASE = gBinarySearchASE(tmCurrSEE(llIndex).iAudioAseCode, tgCurrASE())
                        If ilASE <> -1 Then
                            ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                            If ilANE <> -1 Then
                                If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                                    ilFound = True
                                End If
                            End If
                        End If
                    End If
                    If Not ilFound Then
                        tmCurrSEE(llIndex).iAudioAseCode = 0
                        For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                            '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                                ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                                If ilANE <> -1 Then
                                    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                                        tmCurrSEE(llIndex).iAudioAseCode = tgCurrASE(ilASE).iCode
                                    End If
                            '        Exit For
                                End If
                            'Next ilANE
                            If tmCurrSEE(llIndex).iAudioAseCode <> 0 Then
                                Exit For
                            End If
                        Next ilASE
                    End If
                End If
                tmCurrSEE(llIndex).sAudioItemID = grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX)
                tmCurrSEE(llIndex).sAudioISCI = grdLibEvents.TextMatrix(llRow, AUDIOISCIINDEX)
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOCTRLINDEX))
                ilCCE = gBinarySearchCCE(tmCurrSEE(llIndex).iAudioCceCode, tgCurrAudioCCE)
                If ilCCE <> -1 Then
                    If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iAudioCceCode = 0
                    llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iAudioCceCode = lbcCCE_A.ItemData(llRet)
                    Else
                        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode
                                Exit For
                            End If
                        Next ilCCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPNAMEINDEX))
                ilANE = gBinarySearchANE(tmCurrSEE(llIndex).iBkupAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iBkupAneCode = 0
                    llRet = SendMessageByString(lbcANE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iBkupAneCode = lbcANE.ItemData(llRet)
                    Else
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                            ilANE = gBinarySearchName(slStr, tgCurrANE_Name())
                            If ilANE <> -1 Then
                                tmCurrSEE(llIndex).iBkupAneCode = tgCurrANE_Name(ilANE).iCode   'tgCurrANE(ilANE).iCode
                        '        Exit For
                            End If
                        'Next ilANE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, BACKUPCTRLINDEX))
                ilCCE = gBinarySearchCCE(tmCurrSEE(llIndex).iBkupCceCode, tgCurrAudioCCE)
                If ilCCE <> -1 Then
                    If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iBkupCceCode = 0
                    llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iBkupCceCode = lbcCCE_A.ItemData(llRet)
                    Else
                        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode
                                Exit For
                            End If
                        Next ilCCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTNAMEINDEX))
                ilANE = gBinarySearchANE(tmCurrSEE(llIndex).iProtAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iProtAneCode = 0
                    llRet = SendMessageByString(lbcANE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iProtAneCode = lbcANE.ItemData(llRet)
                    Else
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                            ilANE = gBinarySearchName(slStr, tgCurrANE_Name())
                            If ilANE <> -1 Then
                                tmCurrSEE(llIndex).iProtAneCode = tgCurrANE_Name(ilANE).iCode   'tgCurrANE(ilANE).iCode
                        '        Exit For
                            End If
                        'Next ilANE
                    End If
                End If
                tmCurrSEE(llIndex).sProtItemID = grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX)
                tmCurrSEE(llIndex).sProtISCI = grdLibEvents.TextMatrix(llRow, PROTISCIINDEX)
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, PROTCTRLINDEX))
                ilCCE = gBinarySearchCCE(tmCurrSEE(llIndex).iProtCceCode, tgCurrAudioCCE)
                If ilCCE <> -1 Then
                    If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iProtCceCode = 0
                    llRet = SendMessageByString(lbcCCE_A.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iProtCceCode = lbcCCE_A.ItemData(llRet)
                    Else
                        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode
                                Exit For
                            End If
                        Next ilCCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY1INDEX))
                ilRNE = gBinarySearchRNE(tmCurrSEE(llIndex).i1RneCode, tgCurrRNE)
                If ilRNE <> -1 Then
                    If StrComp(Trim$(tgCurrRNE(ilRNE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).i1RneCode = 0
                    llRet = SendMessageByString(lbcRNE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).i1RneCode = lbcRNE.ItemData(llRet)
                    Else
                        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrRNE(ilRNE).sName), slStr, vbTextCompare) = 0 Then
                            ilRNE = gBinarySearchName(slStr, tgCurrRNE_Name())
                            If ilRNE <> -1 Then
                                tmCurrSEE(llIndex).i1RneCode = tgCurrRNE_Name(ilRNE).iCode  'tgCurrRNE(ilRNE).iCode
                        '        Exit For
                            End If
                        'Next ilRNE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, RELAY2INDEX))
                ilRNE = gBinarySearchRNE(tmCurrSEE(llIndex).i2RneCode, tgCurrRNE)
                If ilRNE <> -1 Then
                    If StrComp(Trim$(tgCurrRNE(ilRNE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).i2RneCode = 0
                    llRet = SendMessageByString(lbcRNE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).i2RneCode = lbcRNE.ItemData(llRet)
                    Else
                        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrRNE(ilRNE).sName), slStr, vbTextCompare) = 0 Then
                            ilRNE = gBinarySearchName(slStr, tgCurrRNE_Name())
                            If ilRNE <> -1 Then
                                tmCurrSEE(llIndex).i2RneCode = tgCurrRNE_Name(ilRNE).iCode  'tgCurrRNE(ilRNE).iCode
                        '        Exit For
                            End If
                        'Next ilRNE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, FOLLOWINDEX))
                ilFNE = gBinarySearchFNE(tmCurrSEE(llIndex).iFneCode, tgCurrFNE)
                If ilFNE <> -1 Then
                    If StrComp(Trim$(tgCurrFNE(ilFNE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iFneCode = 0
                    llRet = SendMessageByString(lbcFNE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iFneCode = lbcFNE.ItemData(llRet)
                    Else
                        For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
                            If StrComp(Trim$(tgCurrFNE(ilFNE).sName), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).iFneCode = tgCurrFNE(ilFNE).iCode
                                Exit For
                            End If
                        Next ilFNE
                    End If
                End If
                slStr = grdLibEvents.TextMatrix(llRow, SILENCETIMEINDEX)
                tmCurrSEE(llIndex).lSilenceTime = gLengthToLong(slStr)  'gStrLengthInTenthToLong(slStr)
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE1INDEX))
                ilSCE = gBinarySearchSCE(tmCurrSEE(llIndex).i1SceCode, tgCurrSCE)
                If ilSCE <> -1 Then
                    If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).i1SceCode = 0
                    llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).i1SceCode = lbcSCE.ItemData(llRet)
                    Else
                        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).i1SceCode = tgCurrSCE(ilSCE).iCode
                                Exit For
                            End If
                        Next ilSCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE2INDEX))
                ilSCE = gBinarySearchSCE(tmCurrSEE(llIndex).i2SceCode, tgCurrSCE)
                If ilSCE <> -1 Then
                    If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).i2SceCode = 0
                    llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).i2SceCode = lbcSCE.ItemData(llRet)
                    Else
                        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).i2SceCode = tgCurrSCE(ilSCE).iCode
                                Exit For
                            End If
                        Next ilSCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE3INDEX))
                ilSCE = gBinarySearchSCE(tmCurrSEE(llIndex).i3SceCode, tgCurrSCE)
                If ilSCE <> -1 Then
                    If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).i3SceCode = 0
                    llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).i3SceCode = lbcSCE.ItemData(llRet)
                    Else
                        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).i3SceCode = tgCurrSCE(ilSCE).iCode
                                Exit For
                            End If
                        Next ilSCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, SILENCE4INDEX))
                ilSCE = gBinarySearchSCE(tmCurrSEE(llIndex).i4SceCode, tgCurrSCE)
                If ilSCE <> -1 Then
                    If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).i4SceCode = 0
                    llRet = SendMessageByString(lbcSCE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).i4SceCode = lbcSCE.ItemData(llRet)
                    Else
                        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
                            If StrComp(Trim$(tgCurrSCE(ilSCE).sAutoChar), slStr, vbTextCompare) = 0 Then
                                tmCurrSEE(llIndex).i4SceCode = tgCurrSCE(ilSCE).iCode
                                Exit For
                            End If
                        Next ilSCE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE1INDEX))
                ilNNE = gBinarySearchNNE(tmCurrSEE(llIndex).iStartNneCode, tgCurrNNE)
                If ilNNE <> -1 Then
                    If StrComp(Trim$(tgCurrNNE(ilNNE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iStartNneCode = 0
                    llRet = SendMessageByString(lbcNNE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iStartNneCode = lbcNNE.ItemData(llRet)
                    Else
                        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrNNE(ilNNE).sName), slStr, vbTextCompare) = 0 Then
                            ilNNE = gBinarySearchName(slStr, tgCurrNNE_Name())
                            If ilNNE <> -1 Then
                                tmCurrSEE(llIndex).iStartNneCode = tgCurrNNE_Name(ilNNE).iCode  'tgCurrNNE(ilNNE).iCode
                        '        Exit For
                            End If
                        'Next ilNNE
                    End If
                End If
                ilFound = False
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, NETCUE2INDEX))
                ilNNE = gBinarySearchNNE(tmCurrSEE(llIndex).iEndNneCode, tgCurrNNE)
                If ilNNE <> -1 Then
                    If StrComp(Trim$(tgCurrNNE(ilNNE).sName), slStr, vbTextCompare) = 0 Then
                        ilFound = True
                    End If
                End If
                If Not ilFound Then
                    tmCurrSEE(llIndex).iEndNneCode = 0
                    llRet = SendMessageByString(lbcNNE.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                    If llRet > 0 Then
                        tmCurrSEE(llIndex).iEndNneCode = lbcNNE.ItemData(llRet)
                    Else
                        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
                        '    If StrComp(Trim$(tgCurrNNE(ilNNE).sName), slStr, vbTextCompare) = 0 Then
                            ilNNE = gBinarySearchName(slStr, tgCurrNNE_Name())
                            If ilNNE <> -1 Then
                                tmCurrSEE(llIndex).iEndNneCode = tgCurrNNE_Name(ilNNE).iCode  'tgCurrNNE(ilNNE).iCode
                        '        Exit For
                            End If
                        'Next ilNNE
                    End If
                End If
                'Set later
                If (tmCurrSEE(llIndex).iEteCode = imSpotETECode) Then
                    smT1Comment(llIndex) = ""
                    tmARE.lCode = 0
                    tmARE.sName = grdLibEvents.TextMatrix(llRow, TITLE1INDEX)
                    tmARE.sUnusued = ""
                    If (Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX)) <> "") And (StrComp(Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX)), "[None]", vbTextCompare) <> 0) Then
                        ilRet = gPutInsert_ARE_AdvertiserRefer(tmARE, "EngrSchdDef-Merge Insert Advertiser Name")
                        tmCurrSEE(llIndex).lAreCode = tmARE.lCode
                    Else
                        tmCurrSEE(llIndex).lAreCode = 0
                    End If
                Else
                    tmCurrSEE(llIndex).lAreCode = 0
                    smT1Comment(llIndex) = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
                End If
                '7/8/11: Make T2 work like T1
                'ilFound = False
                'slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
                'llCTE = gBinarySearchCTE(tmCurrSEE(llIndex).l2CteCode, tgCurrCTE)
                'If llCTE <> -1 Then
                '    If StrComp(Trim$(tgCurrCTE(llCTE).sName), slStr, vbTextCompare) = 0 Then
                '        ilFound = True
                '    End If
                'End If
                'If Not ilFound Then
                '    tmCurrSEE(llIndex).l2CteCode = 0
                '    llRet = SendMessageByString(lbcCTE_2.hwnd, LB_FINDSTRINGEXACT, -1, slStr)
                '    If llRet > 0 Then
                '        tmCurrSEE(llIndex).l2CteCode = lbcCTE_2.ItemData(llRet)
                '    Else
                '        'For llCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
                '        '    If StrComp(Trim$(tgCurrCTE(llCTE).sName), slStr, vbTextCompare) = 0 Then
                '            llCTE = gBinarySearchCTEName(slStr, tgCurr2CTE_Name())
                '            If llCTE <> -1 Then
                '                tmCurrSEE(llIndex).l2CteCode = tgCurr2CTE_Name(llCTE).lCode
                '        '        Exit For
                '            End If
                '        'Next llCTE
                '    End If
                'End If
                smT2Comment(llIndex) = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
                If sgClientFields = "A" Then
                    tmCurrSEE(llIndex).sABCFormat = Trim$(grdLibEvents.TextMatrix(llRow, ABCFORMATINDEX))
                    tmCurrSEE(llIndex).sABCPgmCode = Trim$(grdLibEvents.TextMatrix(llRow, ABCPGMCODEINDEX))
                    tmCurrSEE(llIndex).sABCXDSMode = Trim$(grdLibEvents.TextMatrix(llRow, ABCXDSMODEINDEX))
                    tmCurrSEE(llIndex).sABCRecordItem = Trim$(grdLibEvents.TextMatrix(llRow, ABCRECORDITEMINDEX))
                Else
                    tmCurrSEE(llIndex).sABCFormat = ""
                    tmCurrSEE(llIndex).sABCPgmCode = ""
                    tmCurrSEE(llIndex).sABCXDSMode = ""
                    tmCurrSEE(llIndex).sABCRecordItem = ""
                End If
                tmCurrSEE(llIndex).sUnused = ""
                If Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) = "" Then
                    grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
                    grdLibEvents.TextMatrix(llRow, EVTCONFLICTINDEX) = "N"
                End If
            'End If
            End If
            
            tmCurrSEE(llIndex).sIgnoreConflicts = grdLibEvents.TextMatrix(llRow, EVTCONFLICTINDEX)
            tmCurrSEE(llIndex).lCode = Val(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
            If tmCurrSEE(llIndex).lCode > 0 Then
                If grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y" Then
                    lmChgStatusSEECode(UBound(lmChgStatusSEECode)) = tmCurrSEE(llIndex).lCode
                    ReDim Preserve lmChgStatusSEECode(0 To UBound(lmChgStatusSEECode) + 1) As Long
                End If
                'For llSEEOld = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
                '    If tmCurrSEE(llIndex).lCode = tgCurrSEE(llSEEOld).lCode Then
                llSEEOld = mBinarySearchOldSEE(tmCurrSEE(llIndex).lCode)
                If llSEEOld <> -1 Then
                        If slEventCategory <> "A" Then
                            tmCurrSEE(llIndex).sSentStatus = tgCurrSEE(llSEEOld).sSentStatus
                            tmCurrSEE(llIndex).sSentDate = tgCurrSEE(llSEEOld).sSentDate
                            tmCurrSEE(llIndex).sIgnoreConflicts = tgCurrSEE(llSEEOld).sIgnoreConflicts
                        Else
                            ''LSet tmCurrSEE(llIndex) = tgCurrSEE(llSEEOld)
                            'tmCurrSEE(llIndex).iBdeCode = tgCurrSEE(llSEEOld).iBdeCode
                            'tmCurrSEE(llIndex).iBusCceCode = tgCurrSEE(llSEEOld).iBusCceCode
                            'tmCurrSEE(llIndex).iEteCode = tgCurrSEE(llSEEOld).iEteCode
                            'tmCurrSEE(llIndex).lTime = tgCurrSEE(llSEEOld).lTime
                            'tmCurrSEE(llIndex).lSpotTime = tgCurrSEE(llSEEOld).lSpotTime
                            tmCurrSEE(llIndex).sSentStatus = tgCurrSEE(llSEEOld).sSentStatus
                            tmCurrSEE(llIndex).sSentDate = tgCurrSEE(llSEEOld).sSentDate
                            tmCurrSEE(llIndex).sIgnoreConflicts = tgCurrSEE(llSEEOld).sIgnoreConflicts
                        End If
                        'Exit For
                End If
                '    End If
                'Next llSEEOld
            Else
                tmCurrSEE(llIndex).sInsertFlag = "N"
                'Using Unused as a temporary flag field to know row was added by Insert or altered after creating a schedule
                If grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y" Then
                    tmCurrSEE(llIndex).sInsertFlag = "Y"
                End If
            End If
        End If
    Next llRow
    
End Sub




Private Function mCompareSEE(llSEENew As Long, llCode As Long, slT1Comment As String, slT2Comment As String) As Integer
    'Dim llSEENew As Long
    Dim llSEEOld As Long
    Dim ilEBE As Integer
    Dim slStr As String
    Dim ilBDE As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    If llCode > 0 Then
        'For llSEENew = LBound(tmCurrSEE) To UBound(tmCurrSEE) - 1 Step 1
        '    If llCode = tmCurrSEE(llSEENew).lCode Then
                'For llSEEOld = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
                '    If llCode = tgCurrSEE(llSEEOld).lCode Then
                llSEEOld = mBinarySearchOldSEE(llCode)
                If llSEEOld <> -1 Then
                        'Compare fields
                        'Buses
                        
                        If tmCurrSEE(llSEENew).iBdeCode <> tgCurrSEE(llSEEOld).iBdeCode Then
                            mFindBrackets llSEEOld
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iBusCceCode <> tgCurrSEE(llSEEOld).iBusCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iEteCode <> tgCurrSEE(llSEEOld).iEteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).lTime <> tgCurrSEE(llSEEOld).lTime Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If (tmCurrSEE(llSEENew).iEteCode = imSpotETECode) Then
                            If tmCurrSEE(llSEENew).lSpotTime <> tgCurrSEE(llSEEOld).lSpotTime Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        If tmCurrSEE(llSEENew).iStartTteCode <> tgCurrSEE(llSEEOld).iStartTteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).sFixedTime <> tgCurrSEE(llSEEOld).sFixedTime Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iEndTteCode <> tgCurrSEE(llSEEOld).iEndTteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).lDuration <> tgCurrSEE(llSEEOld).lDuration Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iMteCode <> tgCurrSEE(llSEEOld).iMteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iAudioAseCode <> tgCurrSEE(llSEEOld).iAudioAseCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrSEE(llSEENew).sAudioItemID, tgCurrSEE(llSEEOld).sAudioItemID, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrSEE(llSEENew).sAudioISCI, tgCurrSEE(llSEEOld).sAudioISCI, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If (tmCurrSEE(llSEENew).iEteCode = imSpotETECode) Then
                            If tmCurrSEE(llSEENew).sAudioItemIDChk <> tgCurrSEE(llSEEOld).sAudioItemIDChk Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        If tmCurrSEE(llSEENew).iAudioCceCode <> tgCurrSEE(llSEEOld).iAudioCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iBkupAneCode <> tgCurrSEE(llSEEOld).iBkupAneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iBkupCceCode <> tgCurrSEE(llSEEOld).iBkupCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iProtAneCode <> tgCurrSEE(llSEEOld).iProtAneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrSEE(llSEENew).sProtItemID, tgCurrSEE(llSEEOld).sProtItemID, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tmCurrSEE(llSEENew).sProtISCI, tgCurrSEE(llSEEOld).sProtISCI, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If (tmCurrSEE(llSEENew).iEteCode = imSpotETECode) Then
                            If tmCurrSEE(llSEENew).sProtItemIDChk <> tgCurrSEE(llSEEOld).sProtItemIDChk Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        If tmCurrSEE(llSEENew).iProtCceCode <> tgCurrSEE(llSEEOld).iProtCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).i1RneCode <> tgCurrSEE(llSEEOld).i1RneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).i2RneCode <> tgCurrSEE(llSEEOld).i2RneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iFneCode <> tgCurrSEE(llSEEOld).iFneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).lSilenceTime <> tgCurrSEE(llSEEOld).lSilenceTime Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).i1SceCode <> tgCurrSEE(llSEEOld).i1SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).i2SceCode <> tgCurrSEE(llSEEOld).i2SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).i3SceCode <> tgCurrSEE(llSEEOld).i3SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).i4SceCode <> tgCurrSEE(llSEEOld).i4SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iStartNneCode <> tgCurrSEE(llSEEOld).iStartNneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tmCurrSEE(llSEENew).iEndNneCode <> tgCurrSEE(llSEEOld).iEndNneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        'Comment
                        ilRet = gGetRec_CTE_CommtsTitle(tgCurrSEE(llSEEOld).l1CteCode, "EngrSchdDef- mCompaerSEE for CTE", tmCTE)
                        If ilRet Then
                            If StrComp(Trim$(slT1Comment), Trim$(tmCTE.sComment), vbTextCompare) <> 0 Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        Else
                            If Trim$(slT1Comment) <> "" Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        '7/8/11: Make T2 work like T1
                        'If tmCurrSEE(llSEENew).l2CteCode <> tgCurrSEE(llSEEOld).l2CteCode Then
                        '    mCompareSEE = False
                        '    Exit Function
                        'End If
                        ilRet = gGetRec_CTE_CommtsTitle(tgCurrSEE(llSEEOld).l2CteCode, "EngrSchdDef- mCompaerSEE for CTE", tmCTE)
                        If ilRet Then
                            If StrComp(Trim$(slT2Comment), Trim$(tmCTE.sComment), vbTextCompare) <> 0 Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        Else
                            If Trim$(slT2Comment) <> "" Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        If sgClientFields = "A" Then
                            If StrComp(tmCurrSEE(llSEENew).sABCFormat, tgCurrSEE(llSEEOld).sABCFormat, vbTextCompare) <> 0 Then
                                mCompareSEE = False
                                Exit Function
                            End If
                            If StrComp(tmCurrSEE(llSEENew).sABCPgmCode, tgCurrSEE(llSEEOld).sABCPgmCode, vbTextCompare) <> 0 Then
                                mCompareSEE = False
                                Exit Function
                            End If
                            If StrComp(tmCurrSEE(llSEENew).sABCXDSMode, tgCurrSEE(llSEEOld).sABCXDSMode, vbTextCompare) <> 0 Then
                                mCompareSEE = False
                                Exit Function
                            End If
                            If StrComp(tmCurrSEE(llSEENew).sABCRecordItem, tgCurrSEE(llSEEOld).sABCRecordItem, vbTextCompare) <> 0 Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        '5/16/13: Handle case where spots removed
                        tgCurrSEE(llSEEOld).lCode = -tgCurrSEE(llSEEOld).lCode
                        
                        mCompareSEE = True
                        Exit Function
                End If
                    'End If
                'Next llSEEOld
                mCompareSEE = False 'True
                Exit Function
        '    End If
        'Next llSEENew
    Else
        mCompareSEE = False
    End If
    
    
    
End Function


Private Function mEBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    
    bmInBranch = True
    mEBranch = True
    If (lmEEnableRow >= grdLibEvents.FixedRows) And (lmEEnableRow < grdLibEvents.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, lmEEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEEnableCol
                Case BUSNAMEINDEX
                    llRow = gListBoxFind(lbcBDE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrBus.Show vbModal
                        sgCurrBDEStamp = ""
                        mPopBDE
                        lbcBDE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcBDE, CLng(grdLibEvents.Height / 2)
                        If lbcBDE.Top + lbcBDE.Height > cmcCancel.Top Then
                            lbcBDE.Top = edcEDropdown.Top - lbcBDE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcBDE, slStr)
                            If llRow > 0 Then
                                lbcBDE.ListIndex = llRow
                                edcEDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case BUSCTRLINDEX
                    'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcCCE_B, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 2
                        sgInitCallName = slStr
                        EngrControlChar.Show vbModal
                        sgCurrBusCCEStamp = ""
                        mPopCCE_Bus
                        lbcCCE_B.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcCCE_B, CLng(grdLibEvents.Height / 2)
                        If lbcCCE_B.Top + lbcCCE_B.Height > cmcCancel.Top Then
                            lbcCCE_B.Top = edcEDropdown.Top - lbcCCE_B.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcCCE_B, slStr)
                            If llRow > 0 Then
                                lbcCCE_B.ListIndex = llRow
                                edcEDropdown.text = lbcCCE_B.List(lbcCCE_B.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case EVENTTYPEINDEX
                    '2/9/12: Allow all event types
                    ''llRow = gListBoxFind(lbcETE, slStr)
                    'llRow = gListBoxFind(lbcETE_Program, slStr)
                    llRow = gListBoxFind(lbcETE, slStr)
                    If (llRow < 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrEventType.Show vbModal
                        sgCurrETEStamp = ""
                        mPopETE
                        ''2/9/12: Allow all event types
                        ''lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        ''gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
                        ''If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
                        ''    lbcETE.Top = edcEDropdown.Top - lbcETE.Height
                        ''End If
                        'lbcETE_Program.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        'gSetListBoxHeight lbcETE_Program, CLng(grdLibEvents.Height / 2)
                        'If lbcETE_Program.Top + lbcETE_Program.Height > cmcCancel.Top Then
                        '    lbcETE_Program.Top = edcEDropdown.Top - lbcETE_Program.Height
                        'End If
                        lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
                        If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
                            lbcETE.Top = edcEDropdown.Top - lbcETE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            ''llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            '2/9/12: Allow all event types
                            ''llRow = gListBoxFind(lbcETE, slStr)
                            ''If llRow > 0 Then
                            ''    lbcETE.ListIndex = llRow
                            ''    edcEDropdown.Text = lbcETE.List(lbcETE.ListIndex)
                            'llRow = gListBoxFind(lbcETE_Program, slStr)
                            'If llRow > 0 Then
                            '    lbcETE_Program.ListIndex = llRow
                            '    edcEDropdown.text = lbcETE_Program.List(lbcETE_Program.ListIndex)
                            llRow = gListBoxFind(lbcETE, slStr)
                            If llRow > 0 Then
                                lbcETE.ListIndex = llRow
                                edcEDropdown.text = lbcETE.List(lbcETE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case STARTTYPEINDEX
                    llRow = gListBoxFind(lbcTTE_S, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrTimeType.Show vbModal
                        sgCurrStartTTEStamp = ""
                        mPopTTE_StartType
                        lbcTTE_S.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcTTE_S, CLng(grdLibEvents.Height / 2)
                        If lbcTTE_S.Top + lbcTTE_S.Height > cmcCancel.Top Then
                            lbcTTE_S.Top = edcEDropdown.Top - lbcTTE_S.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcTTE_S, slStr)
                            If llRow > 0 Then
                                lbcTTE_S.ListIndex = llRow
                                edcEDropdown.text = lbcTTE_S.List(lbcTTE_S.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case ENDTYPEINDEX
                    llRow = gListBoxFind(lbcTTE_E, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 2
                        sgInitCallName = slStr
                        EngrTimeType.Show vbModal
                        sgCurrEndTTEStamp = ""
                        mPopTTE_EndType
                        lbcTTE_E.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcTTE_E, CLng(grdLibEvents.Height / 2)
                        If lbcTTE_E.Top + lbcTTE_E.Height > cmcCancel.Top Then
                            lbcTTE_E.Top = edcEDropdown.Top - lbcTTE_E.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcTTE_E, slStr)
                            If llRow > 0 Then
                                lbcTTE_E.ListIndex = llRow
                                edcEDropdown.text = lbcTTE_E.List(lbcTTE_E.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case MATERIALINDEX
                    llRow = gListBoxFind(lbcMTE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrMatType.Show vbModal
                        sgCurrMTEStamp = ""
                        mPopMTE
                        lbcMTE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcMTE, CLng(grdLibEvents.Height / 2)
                        If lbcMTE.Top + lbcMTE.Height > cmcCancel.Top Then
                            lbcMTE.Top = edcEDropdown.Top - lbcMTE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcMTE, slStr)
                            If llRow > 0 Then
                                lbcMTE.ListIndex = llRow
                                edcEDropdown.text = lbcMTE.List(lbcMTE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case AUDIONAMEINDEX
                    llRow = gListBoxFind(lbcASE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrAudio.Show vbModal
                        sgCurrASEStamp = ""
                        mPopASE
                        lbcASE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcASE, CLng(grdLibEvents.Height / 2)
                        If lbcASE.Top + lbcASE.Height > cmcCancel.Top Then
                            lbcASE.Top = edcEDropdown.Top - lbcASE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcASE, slStr)
                            If llRow > 0 Then
                                lbcASE.ListIndex = llRow
                                edcEDropdown.text = lbcASE.List(lbcASE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case AUDIOCTRLINDEX
                    llRow = gListBoxFind(lbcCCE_A, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrControlChar.Show vbModal
                        sgCurrAudioCCEStamp = ""
                        mPopCCE_Audio
                        lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
                        If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
                            lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcCCE_A, slStr)
                            If llRow > 0 Then
                                lbcCCE_A.ListIndex = llRow
                                edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case BACKUPNAMEINDEX, PROTNAMEINDEX
                    llRow = gListBoxFind(lbcANE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrAudioName.Show vbModal
                        sgCurrANEStamp = ""
                        mPopANE
                        lbcANE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcANE, CLng(grdLibEvents.Height / 2)
                        If lbcANE.Top + lbcANE.Height > cmcCancel.Top Then
                            lbcANE.Top = edcEDropdown.Top - lbcANE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcANE, slStr)
                            If llRow > 0 Then
                                lbcANE.ListIndex = llRow
                                edcEDropdown.text = lbcANE.List(lbcANE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case BACKUPCTRLINDEX, PROTCTRLINDEX
                    llRow = gListBoxFind(lbcCCE_A, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrControlChar.Show vbModal
                        sgCurrAudioCCEStamp = ""
                        mPopCCE_Audio
                        lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
                        If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
                            lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcCCE_A, slStr)
                            If llRow > 0 Then
                                lbcCCE_A.ListIndex = llRow
                                edcEDropdown.text = lbcCCE_A.List(lbcCCE_A.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case RELAY1INDEX, RELAY2INDEX
                    llRow = gListBoxFind(lbcRNE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrRelay.Show vbModal
                        sgCurrRNEStamp = ""
                        mPopRNE
                        lbcRNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcRNE, CLng(grdLibEvents.Height / 2)
                        If lbcRNE.Top + lbcRNE.Height > cmcCancel.Top Then
                            lbcRNE.Top = edcEDropdown.Top - lbcRNE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcRNE, slStr)
                            If llRow > 0 Then
                                lbcRNE.ListIndex = llRow
                                edcEDropdown.text = lbcRNE.List(lbcRNE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case FOLLOWINDEX
                    llRow = gListBoxFind(lbcFNE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrFollow.Show vbModal
                        sgCurrFNEStamp = ""
                        mPopFNE
                        lbcFNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcFNE, CLng(grdLibEvents.Height / 2)
                        If lbcFNE.Top + lbcFNE.Height > cmcCancel.Top Then
                            lbcFNE.Top = edcEDropdown.Top - lbcFNE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcFNE, slStr)
                            If llRow > 0 Then
                                lbcFNE.ListIndex = llRow
                                edcEDropdown.text = lbcFNE.List(lbcFNE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case SILENCE1INDEX To SILENCE4INDEX
                    llRow = gListBoxFind(lbcSCE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrSilence.Show vbModal
                        sgCurrSCEStamp = ""
                        mPopSCE
                        lbcSCE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcSCE, CLng(grdLibEvents.Height / 2)
                        If lbcSCE.Top + lbcSCE.Height > cmcCancel.Top Then
                            lbcSCE.Top = edcEDropdown.Top - lbcSCE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcSCE, slStr)
                            If llRow > 0 Then
                                lbcSCE.ListIndex = llRow
                                edcEDropdown.text = lbcSCE.List(lbcSCE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                Case NETCUE1INDEX, NETCUE2INDEX
                    llRow = gListBoxFind(lbcNNE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrNetcue.Show vbModal
                        sgCurrNNEStamp = ""
                        mPopNNE
                        lbcNNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                        gSetListBoxHeight lbcNNE, CLng(grdLibEvents.Height / 2)
                        If lbcNNE.Top + lbcNNE.Height > cmcCancel.Top Then
                            lbcNNE.Top = edcEDropdown.Top - lbcNNE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcNNE, slStr)
                            If llRow > 0 Then
                                lbcNNE.ListIndex = llRow
                                edcEDropdown.text = lbcNNE.List(lbcNNE.ListIndex)
                                edcEDropdown.SelStart = 0
                                edcEDropdown.SelLength = Len(edcEDropdown.text)
                            Else
                                mEBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mEBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mEBranch = False
                        End If
                    End If
                '7/8/11: Make T2 work like T1
                'Case TITLE2INDEX
                '    llRow = gListBoxFind(lbcCTE_2, slStr)
                '    If (llRow <= 0) Or (imDoubleClickName) Then
                '        igInitCallInfo = 2
                '        sgInitCallName = slStr
                '        EngrComment.Show vbModal
                '        sgCurrCTEStamp = ""
                '        mPopCTE
                '        lbcCTE_2.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
                '        gSetListBoxHeight lbcCTE_2, CLng(grdLibEvents.Height / 2)
                '        If lbcCTE_2.Top + lbcCTE_2.Height > cmcCancel.Top Then
                '            lbcCTE_2.Top = edcEDropdown.Top - lbcCTE_2.Height
                '        End If
                '        If igReturnCallStatus = CALLDONE Then
                '            slStr = sgReturnCallName
                '            'llRow = SendMessageByString(lbcDSE.hwnd, LB_FINDSTRING, -1, slStr)
                '            llRow = gListBoxFind(lbcCTE_2, slStr)
                '            If llRow > 0 Then
                '                lbcCTE_2.ListIndex = llRow
                '                edcEDropdown.text = lbcCTE_2.List(lbcCTE_2.ListIndex)
                '                edcEDropdown.SelStart = 0
                '                edcEDropdown.SelLength = Len(edcEDropdown.text)
                '            Else
                '                mEBranch = False
                '            End If
                '        ElseIf igReturnCallStatus = CALLCANCELLED Then
                '            mEBranch = False
                '        ElseIf igReturnCallStatus = CALLTERMINATED Then
                '            mEBranch = False
                '        End If
                '    End If
            End Select
        End If
    End If
    imDoubleClickName = False
End Function


Private Sub mPopDNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrLibDNEStamp, "EngrSchdDef-mPopulate Library Names", tgCurrLibDNE())
    ilRet = gGetTypeOfRecs_DNE_DayName("C", "T", sgCurrTempDNEStamp, "EngrSchdDef-mPopulate Library Names", tgCurrTempDNE())
End Sub

Private Sub mPopDSE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrSchdDef-mPopDSE Day Subname", tgCurrDSE())
End Sub


Private Sub mPopBDE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrSchdDef-mPopBDE Bus Definition", tgCurrBDE())
    lbcBDE.Clear
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        If tgCurrBDE(ilLoop).sState = "A" Then
            lbcBDE.AddItem Trim$(tgCurrBDE(ilLoop).sName)
            lbcBDE.ItemData(lbcBDE.NewIndex) = tgCurrBDE(ilLoop).iCode
        End If
    Next ilLoop
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
        lbcBDE.AddItem "[New]", 0
        lbcBDE.ItemData(lbcBDE.NewIndex) = 0
    Else
        lbcBDE.AddItem "[View]", 0
        lbcBDE.ItemData(lbcBDE.NewIndex) = 0
    End If
End Sub

Private Sub spcItemID_OnComm()
    gErrorMsgPort spcItemID
End Sub

Private Sub tmcCheck_Timer()
    Dim tlSHE As SHE
    Dim ilRet As Integer
    
    tmcCheck.Enabled = False
    ilRet = gGetRec_SHE_ScheduleHeader(lmCheckSHECode, "EngrSchedule-Get Schedule to Check Load", tlSHE)
    If Not ilRet Then
        If tlSHE.sCreateLoad = "Y" Then
            MsgBox "Command sent to Service program has not been retrieved, please that the EngrService is running", vbCritical + vbOKOnly, "Schedule"
        End If
    End If
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If edcEvent.Visible Or pbcEDefine.Visible Or edcEDropdown.Visible Or pbcYN.Visible Then
        Select Case grdLibEvents.Col
            Case BUSNAMEINDEX
                lbcBDE.Visible = False
            Case BUSCTRLINDEX
                lbcCCE_B.Visible = False
            Case EVENTTYPEINDEX
                '2/9/12: Allow all event types
                ''lbcETE.Visible = False
                'lbcETE_Program.Visible = False
                lbcETE.Visible = False
            Case STARTTYPEINDEX
                lbcTTE_S.Visible = False
            Case ENDTYPEINDEX
                lbcTTE_E.Visible = False
            Case MATERIALINDEX
                lbcMTE.Visible = False
            Case AUDIONAMEINDEX
                lbcASE.Visible = False
            Case AUDIOCTRLINDEX
                lbcCCE_A.Visible = False
            Case BACKUPNAMEINDEX
                lbcANE.Visible = False
            Case BACKUPCTRLINDEX
                lbcCCE_A.Visible = False
            Case PROTNAMEINDEX
                lbcANE.Visible = False
            Case PROTCTRLINDEX
                lbcCCE_A.Visible = False
            Case RELAY1INDEX, RELAY2INDEX
                lbcRNE.Visible = False
            Case FOLLOWINDEX
                lbcFNE.Visible = False
            Case SILENCE1INDEX To SILENCE4INDEX
                lbcSCE.Visible = False
            Case NETCUE1INDEX, NETCUE2INDEX
                lbcNNE.Visible = False
            Case TITLE1INDEX
                lbcCTE_1.Visible = False
            Case TITLE2INDEX
                lbcCTE_2.Visible = False
        End Select
    End If
End Sub

Private Sub mPopCCE_Audio()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrSchdDef-mPopCCE_Audio Control Character", tgCurrAudioCCE())
    lbcCCE_A.Clear
    For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tgCurrAudioCCE(ilLoop).sState = "A" Then
            lbcCCE_A.AddItem Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
            lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = tgCurrAudioCCE(ilLoop).iCode
        End If
    Next ilLoop
    lbcCCE_A.AddItem "[None]", 0
    lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcCCE_A.AddItem "[New]", 0
        lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = 0
    Else
        lbcCCE_A.AddItem "[View]", 0
        lbcCCE_A.ItemData(lbcCCE_A.NewIndex) = 0
    End If
End Sub

Private Sub mPopCCE_Bus()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrSchdDef-mPopCCE_Bus Control Character", tgCurrBusCCE())
    lbcCCE_B.Clear
    For ilLoop = 0 To UBound(tgCurrBusCCE) - 1 Step 1
        If tgCurrBusCCE(ilLoop).sState = "A" Then
            lbcCCE_B.AddItem Trim$(tgCurrBusCCE(ilLoop).sAutoChar)
            lbcCCE_B.ItemData(lbcCCE_B.NewIndex) = tgCurrBusCCE(ilLoop).iCode
        End If
    Next ilLoop
    lbcCCE_B.AddItem "[None]", 0
    lbcCCE_B.ItemData(lbcCCE_B.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
        lbcCCE_B.AddItem "[New]", 0
        lbcCCE_B.ItemData(lbcCCE_B.NewIndex) = 0
    Else
        lbcCCE_B.AddItem "[View]", 0
        lbcCCE_B.ItemData(lbcCCE_B.NewIndex) = 0
    End If
End Sub

Private Sub mPopTTE_StartType()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrSchdDef-mPopTTE_StartType Start Type", tgCurrStartTTE())
    lbcTTE_S.Clear
    For ilLoop = 0 To UBound(tgCurrStartTTE) - 1 Step 1
        If tgCurrStartTTE(ilLoop).sState = "A" Then
            lbcTTE_S.AddItem Trim$(tgCurrStartTTE(ilLoop).sName)
            lbcTTE_S.ItemData(lbcTTE_S.NewIndex) = tgCurrStartTTE(ilLoop).iCode
        End If
    Next ilLoop
    lbcTTE_S.AddItem "[None]", 0
    lbcTTE_S.ItemData(lbcTTE_S.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TIMETYPELIST) = 2) Then
        lbcTTE_S.AddItem "[New]", 0
        lbcTTE_S.ItemData(lbcTTE_S.NewIndex) = 0
    Else
        lbcTTE_S.AddItem "[View]", 0
        lbcTTE_S.ItemData(lbcTTE_S.NewIndex) = 0
    End If
End Sub

Private Sub mPopTTE_EndType()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrSchdDef-mPopTTE_EndType End Type", tgCurrEndTTE())
    lbcTTE_E.Clear
    For ilLoop = 0 To UBound(tgCurrEndTTE) - 1 Step 1
        If tgCurrEndTTE(ilLoop).sState = "A" Then
            lbcTTE_E.AddItem Trim$(tgCurrEndTTE(ilLoop).sName)
            lbcTTE_E.ItemData(lbcTTE_E.NewIndex) = tgCurrEndTTE(ilLoop).iCode
        End If
    Next ilLoop
    lbcTTE_E.AddItem "[None]", 0
    lbcTTE_E.ItemData(lbcTTE_E.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TIMETYPELIST) = 2) Then
        lbcTTE_E.AddItem "[New]", 0
        lbcTTE_E.ItemData(lbcTTE_E.NewIndex) = 0
    Else
        lbcTTE_E.AddItem "[View]", 0
        lbcTTE_E.ItemData(lbcTTE_E.NewIndex) = 0
    End If
End Sub

Private Sub mPopASE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilANE As Integer

    mPopANE
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrSchdDef-mPopASE Audio Source", tgCurrASE())
    lbcASE.Clear
    For ilLoop = 0 To UBound(tgCurrASE) - 1 Step 1
        If tgCurrASE(ilLoop).sState = "A" Then
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tgCurrASE(ilLoop).iPriAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tgCurrASE(ilLoop).iPriAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    lbcASE.AddItem Trim$(tgCurrANE(ilANE).sName)
                    lbcASE.ItemData(lbcASE.NewIndex) = tgCurrASE(ilLoop).iCode
                End If
            'Next ilANE
        End If
    Next ilLoop
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcASE.AddItem "[New]", 0
        lbcASE.ItemData(lbcASE.NewIndex) = 0
    Else
        lbcASE.AddItem "[View]", 0
        lbcASE.ItemData(lbcASE.NewIndex) = 0
    End If
End Sub

Private Sub mPopSCE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrSchdDef-mPopSCE Silence Character", tgCurrSCE())
    lbcSCE.Clear
    For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tgCurrSCE(ilLoop).sState = "A" Then
            lbcSCE.AddItem Trim$(tgCurrSCE(ilLoop).sAutoChar)
            lbcSCE.ItemData(lbcSCE.NewIndex) = tgCurrSCE(ilLoop).iCode
        End If
    Next ilLoop
    lbcSCE.AddItem "[None]", 0
    lbcSCE.ItemData(lbcSCE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(SILENCELIST) = 2) Then
        lbcSCE.AddItem "[New]", 0
        lbcSCE.ItemData(lbcSCE.NewIndex) = 0
    Else
        lbcSCE.AddItem "[View]", 0
        lbcSCE.ItemData(lbcSCE.NewIndex) = 0
    End If
End Sub

Private Sub mPopNNE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrSchdDef-mPopNNE Netcue", tgCurrNNE())
    lbcNNE.Clear
    For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
        If tgCurrNNE(ilLoop).sState = "A" Then
            lbcNNE.AddItem Trim$(tgCurrNNE(ilLoop).sName)
            lbcNNE.ItemData(lbcNNE.NewIndex) = tgCurrNNE(ilLoop).iCode
        End If
    Next ilLoop
    lbcNNE.AddItem "[None]", 0
    lbcNNE.ItemData(lbcNNE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(NETCUELIST) = 2) Then
        lbcNNE.AddItem "[New]", 0
        lbcNNE.ItemData(lbcNNE.NewIndex) = 0
    Else
        lbcNNE.AddItem "[View]", 0
        lbcNNE.ItemData(lbcNNE.NewIndex) = 0
    End If
End Sub

Private Sub mPopCTE()
    '7/8/11: Make T2 work like T1
    'Dim ilRet As Integer
    'Dim ilLoop As Integer

    'ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrSchdDef-mPopCTE Title 2", tgCurrCTE())
    'lbcCTE_2.Clear
    'For ilLoop = 0 To UBound(tgCurrCTE) - 1 Step 1
    '    If tgCurrCTE(ilLoop).sState = "A" Then
    '        lbcCTE_2.AddItem Trim$(tgCurrCTE(ilLoop).sName)
    '        lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = tgCurrCTE(ilLoop).lCode
    '    End If
    'Next ilLoop
    'lbcCTE_2.AddItem "[None]", 0
    'lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = 0
    'If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(COMMENTLIST) = 2) Then
    '    lbcCTE_2.AddItem "[New]", 0
    '    lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = 0
    'Else
    '    lbcCTE_2.AddItem "[View]", 0
    '    lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = 0
    'End If
End Sub

Private Sub mPopANE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrSchdDef-mPopANE Audio Audio Names", tgCurrANE())
    lbcANE.Clear
    For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
        If tgCurrANE(ilLoop).sState = "A" Then
            lbcANE.AddItem Trim$(tgCurrANE(ilLoop).sName)
            lbcANE.ItemData(lbcANE.NewIndex) = tgCurrANE(ilLoop).iCode
        End If
    Next ilLoop
    lbcANE.AddItem "[None]", 0
    lbcANE.ItemData(lbcANE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcANE.AddItem "[New]", 0
        lbcANE.ItemData(lbcANE.NewIndex) = 0
    Else
        lbcANE.AddItem "[View]", 0
        lbcANE.ItemData(lbcANE.NewIndex) = 0
    End If
End Sub

Private Sub mPopARE()

    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetRecs_ARE_AdvertiserRefer(sgCurrAREStamp, "EngrSchdDef-mPopARE Advertiser Names", tgCurrARE())
End Sub

Private Sub mPopETE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    imDefaultProgIndex = -1
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrSchdDef-mPopETE Event Properties", tgCurrEPE())
    lbcETE.Clear
    lbcETE_Program.Clear
    For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilLoop).sState = "A" Then
            lbcETE.AddItem Trim$(tgCurrETE(ilLoop).sName)
            lbcETE.ItemData(lbcETE.NewIndex) = tgCurrETE(ilLoop).iCode
            If tgCurrETE(ilLoop).sCategory = "P" Then
                lbcETE_Program.AddItem Trim$(tgCurrETE(ilLoop).sName)
                lbcETE_Program.ItemData(lbcETE_Program.NewIndex) = tgCurrETE(ilLoop).iCode
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcETE.ListCount - 1 Step 1
        If Trim$(lbcETE.List(ilLoop)) = "Program" Then
            imDefaultProgIndex = ilLoop
            Exit For
        End If
    Next ilLoop
'    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(EVENTTYPELIST) = 2) Then
'        lbcETE.AddItem "[New]", 0
'        lbcETE.ItemData(lbcETE.NewIndex) = 0
'    Else
'        lbcETE.AddItem "[View]", 0
'        lbcETE.ItemData(lbcETE.NewIndex) = 0
'    End If
End Sub

Private Sub mPopMTE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrSchdDef-mPopMTE Material Type", tgCurrMTE())
    lbcMTE.Clear
    For ilLoop = 0 To UBound(tgCurrMTE) - 1 Step 1
        If tgCurrMTE(ilLoop).sState = "A" Then
            lbcMTE.AddItem Trim$(tgCurrMTE(ilLoop).sName)
            lbcMTE.ItemData(lbcMTE.NewIndex) = tgCurrMTE(ilLoop).iCode
        End If
    Next ilLoop
    lbcMTE.AddItem "[None]", 0
    lbcMTE.ItemData(lbcMTE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(MATERIALTYPELIST) = 2) Then
        lbcMTE.AddItem "[New]", 0
        lbcMTE.ItemData(lbcMTE.NewIndex) = 0
    Else
        lbcMTE.AddItem "[View]", 0
        lbcMTE.ItemData(lbcMTE.NewIndex) = 0
    End If
End Sub

Private Sub mPopRNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrSchdDef-mPopRNE Relay", tgCurrRNE())
    lbcRNE.Clear
    For ilLoop = 0 To UBound(tgCurrRNE) - 1 Step 1
        If tgCurrRNE(ilLoop).sState = "A" Then
            lbcRNE.AddItem Trim$(tgCurrRNE(ilLoop).sName)
            lbcRNE.ItemData(lbcRNE.NewIndex) = tgCurrRNE(ilLoop).iCode
        End If
    Next ilLoop
    lbcRNE.AddItem "[None]", 0
    lbcRNE.ItemData(lbcRNE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(RELAYLIST) = 2) Then
        lbcRNE.AddItem "[New]", 0
        lbcRNE.ItemData(lbcRNE.NewIndex) = 0
    Else
        lbcRNE.AddItem "[View]", 0
        lbcRNE.ItemData(lbcRNE.NewIndex) = 0
    End If
End Sub

Private Sub mPopFNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrSchdDef-mPopFNE Follow", tgCurrFNE())
    lbcFNE.Clear
    For ilLoop = 0 To UBound(tgCurrFNE) - 1 Step 1
        If tgCurrFNE(ilLoop).sState = "A" Then
            lbcFNE.AddItem Trim$(tgCurrFNE(ilLoop).sName)
            lbcFNE.ItemData(lbcFNE.NewIndex) = tgCurrFNE(ilLoop).iCode
        End If
    Next ilLoop
    lbcFNE.AddItem "[None]", 0
    lbcFNE.ItemData(lbcFNE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(FOLLOWLIST) = 2) Then
        lbcFNE.AddItem "[New]", 0
        lbcFNE.ItemData(lbcFNE.NewIndex) = 0
    Else
        lbcFNE.AddItem "[View]", 0
        lbcFNE.ItemData(lbcFNE.NewIndex) = 0
    End If
End Sub

Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If smYN <> "Y" Then
            imFieldChgd = True
            grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        End If
        smYN = "Y"
        pbcYN_Paint
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If smYN <> "N" Then
            imFieldChgd = True
            grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        End If
        smYN = "N"
        pbcYN_Paint
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If smYN = "Y" Then
            imFieldChgd = True
            grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
            smYN = "N"
            pbcYN_Paint
            If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                grdLibEvents.CellForeColor = vbBlue
            Else
                grdLibEvents.CellForeColor = vbBlack
            End If
        ElseIf smYN = "N" Then
            imFieldChgd = True
            grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
            smYN = "Y"
            pbcYN_Paint
            If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                grdLibEvents.CellForeColor = vbBlue
            Else
                grdLibEvents.CellForeColor = vbBlack
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smYN = "Y" Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        smYN = "N"
        pbcYN_Paint
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    ElseIf smYN = "N" Then
        imFieldChgd = True
        grdLibEvents.TextMatrix(grdLibEvents.Row, CHGSTATUSINDEX) = "Y"
        smYN = "Y"
        pbcYN_Paint
        If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
            grdLibEvents.CellForeColor = vbBlue
        Else
            grdLibEvents.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = 30  'fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    pbcYN.Print smYN
End Sub


Private Sub mESetFocus()
    Dim ilCol As Integer
    Dim llColPos As Integer
    Dim slStr As String
    
    llColPos = 0
    For ilCol = 0 To grdLibEvents.Col - 1 Step 1
        If grdLibEvents.ColIsVisible(ilCol) Then
            llColPos = llColPos + grdLibEvents.ColWidth(ilCol)
        End If
    Next ilCol
    '8/26/11: Check that row is not behind scroll bar
    If grdLibEvents.RowPos(grdLibEvents.Row) + grdLibEvents.RowHeight(grdLibEvents.Row) + 60 >= grdLibEvents.Height Then
        imIgnoreScroll = True
        grdLibEvents.TopRow = grdLibEvents.TopRow + 1
    End If
    Select Case grdLibEvents.Col
        Case HIGHLIGHTINDEX
            pbcHighlight.Visible = True
            pbcHighlight.SetFocus
        Case BUSNAMEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("BusName", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcBDE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcBDE, CLng(grdLibEvents.Height / 2)
            If lbcBDE.Top + lbcBDE.Height > cmcCancel.Top Then
                lbcBDE.Top = edcEDropdown.Top - lbcBDE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            edcEDropdown.SetFocus
        Case BUSCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("BusCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcCCE_B.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcCCE_B, CLng(grdLibEvents.Height / 2)
            If lbcCCE_B.Top + lbcCCE_B.Height > cmcCancel.Top Then
                lbcCCE_B.Top = edcEDropdown.Top - lbcCCE_B.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcCCE_B.Visible = True
            edcEDropdown.SetFocus
        Case EVENTTYPEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("EventType", lmCharacterWidth, edcEDropdown.Width, Len(tgETE.sName)) / 2
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            '2/9/12: Allow all event
            ''lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            ''gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
            ''If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
            ''    lbcETE.Top = edcEDropdown.Top - lbcETE.Height
            ''End If
            'lbcETE_Program.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            'gSetListBoxHeight lbcETE_Program, CLng(grdLibEvents.Height / 2)
            'If lbcETE_Program.Top + lbcETE_Program.Height > cmcCancel.Top Then
            '    lbcETE_Program.Top = edcEDropdown.Top - lbcETE_Program.Height
            'End If
            lbcETE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcETE, CLng(grdLibEvents.Height / 2)
            If lbcETE.Top + lbcETE.Height > cmcCancel.Top Then
                lbcETE.Top = edcEDropdown.Top - lbcETE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            '2/9/12: Allow all events
            ''lbcETE.Visible = True
            'lbcETE_Program.Visible = True
            lbcETE.Visible = True
            edcEDropdown.SetFocus
        Case TIMEINDEX
'            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'            edcEvent.Width = gSetCtrlWidth("Time", lmCharacterWidth, edcEvent.Width, 0)
'            edcEvent.Visible = True
'            edcEvent.SetFocus
            ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            ltcEvent.Width = gSetCtrlWidth("Time", lmCharacterWidth, ltcEvent.Width, 0)
            ltcEvent.Visible = True
            ltcEvent.SetFocus
        Case STARTTYPEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("StartType", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcTTE_S.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcTTE_S, CLng(grdLibEvents.Height / 2)
            If lbcTTE_S.Top + lbcTTE_S.Height > cmcCancel.Top Then
                lbcTTE_S.Top = edcEDropdown.Top - lbcTTE_S.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcTTE_S.Visible = True
            edcEDropdown.SetFocus
        Case FIXEDINDEX
            pbcYN.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case ENDTYPEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("EndType", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcTTE_E.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcTTE_E, CLng(grdLibEvents.Height / 2)
            If lbcTTE_E.Top + lbcTTE_E.Height > cmcCancel.Top Then
                lbcTTE_E.Top = edcEDropdown.Top - lbcTTE_E.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcTTE_E.Visible = True
            edcEDropdown.SetFocus
        Case DURATIONINDEX
'            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'            edcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, edcEvent.Width, 0)
'            edcEvent.Visible = True
'            edcEvent.SetFocus
            ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            ltcEvent.Width = gSetCtrlWidth("Duration", lmCharacterWidth, ltcEvent.Width, 0)
            ltcEvent.Visible = True
            ltcEvent.SetFocus
        Case MATERIALINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("Material", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcMTE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcMTE, CLng(grdLibEvents.Height / 2)
            If lbcMTE.Top + lbcMTE.Height > cmcCancel.Top Then
                lbcMTE.Top = edcEDropdown.Top - lbcMTE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcMTE.Visible = True
            edcEDropdown.SetFocus
        Case AUDIONAMEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("AudioName", lmCharacterWidth, edcEDropdown.Width, 0)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcASE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcASE, CLng(grdLibEvents.Height / 2)
            If lbcASE.Top + lbcASE.Height > cmcCancel.Top Then
                lbcASE.Top = edcEDropdown.Top - lbcASE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcASE.Visible = True
            edcEDropdown.SetFocus
        Case AUDIOITEMIDINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("AudioItemID", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case AUDIOISCIINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("AudioISCI", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case AUDIOCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("AudioCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
            If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
                lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcCCE_A.Visible = True
            edcEDropdown.SetFocus
        Case BACKUPNAMEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("BkupName", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcANE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcANE, CLng(grdLibEvents.Height / 2)
            If lbcANE.Top + lbcANE.Height > cmcCancel.Top Then
                lbcANE.Top = edcEDropdown.Top - lbcANE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcANE.Visible = True
            edcEDropdown.SetFocus
        Case BACKUPCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("BkupCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
            If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
                lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcCCE_A.Visible = True
            edcEDropdown.SetFocus
        Case PROTNAMEINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("ProtName", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcANE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcANE, CLng(grdLibEvents.Height / 2)
            If lbcANE.Top + lbcANE.Height > cmcCancel.Top Then
                lbcANE.Top = edcEDropdown.Top - lbcANE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcANE.Visible = True
            edcEDropdown.SetFocus
        Case PROTITEMIDINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ProtItemID", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case PROTISCIINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ProtISCI", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case PROTCTRLINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("ProtCtrl", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcCCE_A.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcCCE_A, CLng(grdLibEvents.Height / 2)
            If lbcCCE_A.Top + lbcCCE_A.Height > cmcCancel.Top Then
                lbcCCE_A.Top = edcEDropdown.Top - lbcCCE_A.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcCCE_A.Visible = True
            edcEDropdown.SetFocus
        Case RELAY1INDEX, RELAY2INDEX
            If grdLibEvents.Col = RELAY2INDEX Then
                slStr = "Relay2"
            Else
                slStr = "Relay1"
            End If
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcRNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcRNE, CLng(grdLibEvents.Height / 2)
            If lbcRNE.Top + lbcRNE.Height > cmcCancel.Top Then
                lbcRNE.Top = edcEDropdown.Top - lbcRNE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcRNE.Visible = True
            edcEDropdown.SetFocus
        Case FOLLOWINDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("Follow", lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcFNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcFNE, CLng(grdLibEvents.Height / 2)
            If lbcFNE.Top + lbcFNE.Height > cmcCancel.Top Then
                lbcFNE.Top = edcEDropdown.Top - lbcFNE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcFNE.Visible = True
            edcEDropdown.SetFocus
        Case SILENCETIMEINDEX
'            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
'            edcEvent.Width = gSetCtrlWidth("SilenceTime", lmCharacterWidth, edcEvent.Width, 0)
'            edcEvent.Visible = True
'            edcEvent.SetFocus
            ltcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            ltcEvent.Width = gSetCtrlWidth("SILENCETIME", lmCharacterWidth, ltcEvent.Width, 0)
            ltcEvent.Visible = True
            ltcEvent.SetFocus
        Case SILENCE1INDEX To SILENCE4INDEX
            If grdLibEvents.Col = SILENCE2INDEX Then
                slStr = "Silence2"
            ElseIf grdLibEvents.Col = SILENCE3INDEX Then
                slStr = "Silence3"
            ElseIf grdLibEvents.Col = SILENCE4INDEX Then
                slStr = "Silence4"
            Else
                slStr = "Silence1"
            End If
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcSCE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcSCE, CLng(grdLibEvents.Height / 2)
            If lbcSCE.Top + lbcSCE.Height > cmcCancel.Top Then
                lbcSCE.Top = edcEDropdown.Top - lbcSCE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcSCE.Visible = True
            edcEDropdown.SetFocus
        Case NETCUE1INDEX, NETCUE2INDEX
            If grdLibEvents.Col = NETCUE2INDEX Then
                slStr = "Netcue2"
            Else
                slStr = "Netcue1"
            End If
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth(slStr, lmCharacterWidth, edcEDropdown.Width, 6)
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcNNE.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcNNE, CLng(grdLibEvents.Height / 2)
            If lbcNNE.Top + lbcNNE.Height > cmcCancel.Top Then
                lbcNNE.Top = edcEDropdown.Top - lbcNNE.Height
            End If
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcNNE.Visible = True
            edcEDropdown.SetFocus
        Case TITLE1INDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("Title1", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcCTE_1.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcCTE_1, CLng(grdLibEvents.Height / 2)
            If lbcCTE_1.Top + lbcCTE_1.Height > cmcCancel.Top Then
                lbcCTE_1.Top = edcEDropdown.Top - lbcCTE_1.Height
            End If
            '9/26/11: Reset edit box with to be width of title
            edcEDropdown.Width = grdLibEvents.ColWidth(grdLibEvents.Col) - cmcEDropDown.Width
            edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcCTE_1.Visible = True
            edcEDropdown.SetFocus
        Case TITLE2INDEX
            edcEDropdown.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEDropdown.Width = gSetCtrlWidth("Title2", lmCharacterWidth, edcEDropdown.Width, 6)
            edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
            cmcEDropDown.Move edcEDropdown.Left + edcEDropdown.Width, edcEDropdown.Top, cmcEDropDown.Width, edcEDropdown.Height
            lbcCTE_2.Move edcEDropdown.Left, edcEDropdown.Top + edcEDropdown.Height, edcEDropdown.Width + cmcEDropDown.Width
            gSetListBoxHeight lbcCTE_2, CLng(grdLibEvents.Height / 2)
            If lbcCTE_2.Top + lbcCTE_2.Height > cmcCancel.Top Then
                lbcCTE_2.Top = edcEDropdown.Top - lbcCTE_2.Height
            End If
            '9/26/11: Reset edit box with to be width of title
            edcEDropdown.Width = grdLibEvents.ColWidth(TITLE1INDEX) - cmcEDropDown.Width
            edcEDropdown.Left = grdLibEvents.Left + llColPos + grdLibEvents.ColWidth(grdLibEvents.Col) - edcEDropdown.Width - cmcEDropDown.Width
            edcEDropdown.Visible = True
            cmcEDropDown.Visible = True
            lbcCTE_2.Visible = True
            edcEDropdown.SetFocus
        Case ABCFORMATINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ABCFormat", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.MaxLength = gSetMaxChars("ABCFormat", 0)
            imMaxColChars = gGetMaxChars("ABCFormat")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case ABCPGMCODEINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ABCPgmCode", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.Left = grdLibEvents.Left + llColPos + 30 + grdLibEvents.ColWidth(ABCXDSMODEINDEX) - edcEvent.Width
            edcEvent.MaxLength = gSetMaxChars("ABCPgmCode", 0)
            imMaxColChars = gGetMaxChars("ABCPgmCode")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case ABCXDSMODEINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ABCXdsMode", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.MaxLength = gSetMaxChars("ABCXdsMode", 0)
            imMaxColChars = gGetMaxChars("ABCXdsMode")
            edcEvent.Visible = True
            edcEvent.SetFocus
        Case ABCRECORDITEMINDEX
            edcEvent.Move grdLibEvents.Left + llColPos + 30, grdLibEvents.Top + grdLibEvents.RowPos(grdLibEvents.Row) + 15, grdLibEvents.ColWidth(grdLibEvents.Col) - 30, grdLibEvents.RowHeight(grdLibEvents.Row) - 15
            edcEvent.Width = gSetCtrlWidth("ABCRecordItem", lmCharacterWidth, edcEvent.Width, 0)
            edcEvent.MaxLength = gSetMaxChars("ABCRecordItem", 0)
            imMaxColChars = gGetMaxChars("ABCRecordItem")
            edcEvent.Visible = True
            edcEvent.SetFocus
    End Select

End Sub








Private Function mComputeWidth(ilPass As Integer, CtrlWidth As Single, ilAdjValue As Integer, slUsedFlag As String) As Single
    If ilPass = 0 Then
        CtrlWidth = grdLibEvents.Width / ilAdjValue
        If slUsedFlag <> "Y" Then
            imUnusedCount = imUnusedCount + 1
            fmUnusedWidth = fmUnusedWidth + CtrlWidth
            CtrlWidth = 0
        Else
            fmUsedWidth = fmUsedWidth + CtrlWidth
        End If
    Else
        CtrlWidth = CtrlWidth + ((fmUnusedWidth * CtrlWidth) / fmUsedWidth)
    End If
    mComputeWidth = CtrlWidth
End Function

Private Function mColOk(llRow As Long, llCol As Long, ilCheckYellow As Integer) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llAirDate As Long
    
    mColOk = True
    If grdLibEvents.ColWidth(llCol) <= 0 Then
        mColOk = False
        Exit Function
    End If
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    If llAirDate = llNowDate Then
        slStr = grdLibEvents.TextMatrix(llRow, TIMEINDEX)
        If slStr <> "" Then
            If gIsTimeTenths(slStr) Then
                'slStr = Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX))
                If (slStr <> "") And (Val(slStr) <> 0) Then
                    If llNowTime > gStrTimeInTenthToLong(slStr, False) Then
                        mColOk = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    grdLibEvents.Row = llRow
    grdLibEvents.Col = llCol
    If (grdLibEvents.CellBackColor = LIGHTYELLOW) And (ilCheckYellow) Then
        mColOk = False
        Exit Function
    End If
    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "U" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            Select Case llCol
                                Case BUSNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                        mColOk = False
                                    End If
                                Case BUSCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                        mColOk = False
                                    End If
                                Case EVENTTYPEINDEX
                                Case EVENTIDINDEX
                                    mColOk = False
                                Case TIMEINDEX
                                    If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                        mColOk = False
                                    End If
                                Case STARTTYPEINDEX
                                    If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                        mColOk = False
                                    End If
                                Case FIXEDINDEX
                                    If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                        mColOk = False
                                    End If
                                Case ENDTYPEINDEX
                                    If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                        mColOk = False
                                    End If
                                Case DURATIONINDEX
                                    If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                        mColOk = False
                                    End If
                                Case MATERIALINDEX
                                    If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                        mColOk = False
                                    End If
                                Case AUDIONAMEINDEX
                                    If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                        mColOk = False
                                    End If
                                Case AUDIOITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                        mColOk = False
                                    End If
                                Case AUDIOISCIINDEX
                                    If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                        mColOk = False
                                    End If
                                Case AUDIOCTRLINDEX
                                    If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                        mColOk = False
                                    End If
                                Case BACKUPNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                        mColOk = False
                                    End If
                                Case BACKUPCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                        mColOk = False
                                    End If
                                Case PROTNAMEINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                        mColOk = False
                                    End If
                                Case PROTITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                        mColOk = False
                                    End If
                                Case PROTISCIINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                        mColOk = False
                                    End If
                                Case PROTCTRLINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                        mColOk = False
                                    End If
                                Case RELAY1INDEX
                                    If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case RELAY2INDEX
                                    If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case FOLLOWINDEX
                                    If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                        mColOk = False
                                    End If
                                Case SILENCETIMEINDEX
                                    If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                        mColOk = False
                                    End If
                                Case SILENCE1INDEX
                                    If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case SILENCE2INDEX
                                    If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case SILENCE3INDEX
                                    If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case SILENCE4INDEX
                                    If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case NETCUE1INDEX
                                    If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                        mColOk = False
                                    End If
                                Case NETCUE2INDEX
                                    If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                        mColOk = False
                                    End If
                                Case TITLE1INDEX
                                    If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case TITLE2INDEX
                                    If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                        mColOk = False
                                    End If
                                Case ABCFORMATINDEX
                                    If tgCurrEPE(ilEPE).sABCFormat <> "Y" Then
                                        mColOk = False
                                    End If
                                Case ABCPGMCODEINDEX
                                    If tgCurrEPE(ilEPE).sABCPgmCode <> "Y" Then
                                        mColOk = False
                                    End If
                                Case ABCXDSMODEINDEX
                                    If tgCurrEPE(ilEPE).sABCXDSMode <> "Y" Then
                                        mColOk = False
                                    End If
                                Case ABCRECORDITEMINDEX
                                    If tgCurrEPE(ilEPE).sABCRecordItem <> "Y" Then
                                        mColOk = False
                                    End If
                            End Select
                            Exit For
                        End If
                    End If
                Next ilEPE
                Exit For
            End If
        Next ilETE
    End If
End Function







Private Function mSvCheckEventConflicts() As Integer
    Dim llRow1 As Long
    Dim llRow2 As Long
    Dim llRow3 As Long
    Dim llRowT1 As Long
    Dim llRowT2 As Long
    Dim llRowT3 As Long
    Dim ilHour1 As Integer
    Dim ilHour2 As Integer
    Dim slHours1 As String
    Dim slHours2 As String
    Dim ilDay1 As Integer
    Dim ilDay2 As Integer
    Dim slDays1 As String
    Dim slDays2 As String
    Dim slStr As String
    Dim slEvtType1 As String
    Dim slEvtType2 As String
    Dim ilBus1 As Integer
    Dim ilBus2 As Integer
    Dim slBuses1 As String
    Dim slBuses2 As String
    ReDim slBus1(0 To 0) As String
    ReDim slBus2(0 To 0) As String
    Dim llTime1 As Long
    Dim llTime2 As Long
    Dim llDur1 As Long
    Dim llDur2 As Long
    Dim llStartTime1 As Long
    Dim llEndTime1 As Long
    Dim llStartTime2 As Long
    Dim llEndTime2 As Long
    Dim slPriAudio1 As String
    Dim slProtAudio1 As String
    Dim slBkupAudio1 As String
    Dim slPriAudio2 As String
    Dim slProtAudio2 As String
    Dim slBkupAudio2 As String
    Dim slPriItemID1 As String
    Dim slPriItemID2 As String
    Dim slProtItemID1 As String
    Dim slProtItemID2 As String
    Dim slBkupItemID1 As String
    Dim slBkupItemID2 As String
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llErrorColor As Long
    Dim ilSpotsExist As Integer
    Dim slEventCategory As String
    Dim ilETE As Integer
    Dim ilETE3 As Integer
    Dim ilError As Integer
    Dim ilStartConflictIndex As Integer
    Dim ilConflictIndex As Integer
    Dim llRow1TmIndex As Long
    Dim llRow2TmIndex As Long
    Dim ilTest As Integer
    Dim llPostTime As Long
    Dim llPreTime As Long
    Dim ilATE As Integer
    Dim llUpper As Long
    Dim llEventStartTime As Long
    Dim llEventEndTime As Long
    Dim slIgnoreConflicts As String
    Dim slRow1Chg As String
    Dim slRow2Chg As String
    
    mSvCheckEventConflicts = False
    'If Not imAnyEvtChgs Then
    '    Exit Function
    'End If
    
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    llPostTime = 0
    llPreTime = 0
    For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
        If tgCurrATE(ilATE).lPreBufferTime > llPreTime Then
            llPreTime = tgCurrATE(ilATE).lPreBufferTime
        End If
        If tgCurrATE(ilATE).lPostBufferTime > llPostTime Then
            llPostTime = tgCurrATE(ilATE).lPostBufferTime
        End If
    Next ilATE
    ReDim tmConflictTest(1 To 1) As CONFLICTTEST
    llUpper = 1
    For llRow1 = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If (Trim$(grdLibEvents.TextMatrix(llRow1, EVENTTYPEINDEX)) <> "") Then
            slStr = grdLibEvents.TextMatrix(llRow1, TIMEINDEX)
            llEventStartTime = gStrTimeInTenthToLong(slStr, False)
            slStr = grdLibEvents.TextMatrix(llRow1, DURATIONINDEX)
            llEventEndTime = llEventStartTime + gStrLengthInTenthToLong(slStr)  ' - 1
            If llEventEndTime < llEventStartTime Then
                llEventEndTime = llEventStartTime
            End If
            slDays1 = Trim$(Str$(llAirDate))
            slIgnoreConflicts = Trim$(grdLibEvents.TextMatrix(llRow1, EVTCONFLICTINDEX))
            If (slIgnoreConflicts = "A") Or (slIgnoreConflicts = "I") Then
                slPriAudio1 = ""
                slProtAudio1 = ""
                slBkupAudio1 = ""
            Else
                slPriAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, AUDIONAMEINDEX))
                slProtAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, PROTNAMEINDEX))
                slBkupAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, BACKUPNAMEINDEX))
            End If
            mCreateBusRecs llRow1, "B", slIgnoreConflicts, llEventStartTime, llEventEndTime, slDays1, tmConflictTest()
            mCreateAudioRecs llRow1, "1", slPriAudio1, llEventStartTime, llEventEndTime, slDays1, tmConflictTest()
            mCreateAudioRecs llRow1, "2", slProtAudio1, llEventStartTime, llEventEndTime, slDays1, tmConflictTest()
            mCreateAudioRecs llRow1, "3", slBkupAudio1, llEventStartTime, llEventEndTime, slDays1, tmConflictTest()
        End If
    Next llRow1
    'For llRow1 = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
    For llRowT1 = 1 To UBound(tmConflictTest) - 1 Step 1
        llRow1 = tmConflictTest(llRowT1).lRow
        slStr = Trim$(grdLibEvents.TextMatrix(llRow1, EVENTTYPEINDEX))
        slEvtType1 = slStr
        slRow1Chg = grdLibEvents.TextMatrix(llRow1, CHGSTATUSINDEX)
        slRow1Chg = "Y"
        If (slStr <> "") And (slRow1Chg = "Y") Then
            slEventCategory = ""
            ilSpotsExist = False
            For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                    slEventCategory = tgCurrETE(ilETE).sCategory
                    If slEventCategory = "A" Then
                        'Avail are after spots
                        'For llRow3 = llRow1 + 1 To grdLibEvents.Rows - 1 Step 1
                        For llRow3 = llRow1 - 1 To grdLibEvents.FixedRows Step -1
                            slStr = Trim$(grdLibEvents.TextMatrix(llRow3, EVENTTYPEINDEX))
                            If slStr <> "" Then
                                If StrComp(grdLibEvents.TextMatrix(llRow1, BUSNAMEINDEX), grdLibEvents.TextMatrix(llRow3, BUSNAMEINDEX), vbTextCompare) = 0 Then
                                    For ilETE3 = 0 To UBound(tgCurrETE) - 1 Step 1
                                        If StrComp(Trim$(tgCurrETE(ilETE3).sName), slStr, vbTextCompare) = 0 Then
                                            If tgCurrETE(ilETE3).sCategory = "S" Then
                                                If gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow1, SPOTAVAILTIMEINDEX), False) = gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow3, SPOTAVAILTIMEINDEX), False) Then
                                                    ilSpotsExist = True
                                                End If
                                            End If
                                            Exit For
                                        End If
                                    Next ilETE3
                                    If ilSpotsExist = True Then
                                        Exit For
                                    End If
                                End If
                                If gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow1, TIMEINDEX), False) + 3000 < gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow3, TIMEINDEX), False) Then
                                    Exit For
                                End If
                            End If
                        Next llRow3
                    End If
                    Exit For
                End If
            Next ilETE
            If (slEventCategory = "P") Or ((slEventCategory = "A") And (ilSpotsExist = False)) Or (slEventCategory = "S") Then
                'slStr = grdLibEvents.TextMatrix(llRow1, TIMEINDEX)
                'llTime1 = gStrTimeInTenthToLong(slStr, False)
                'slStr = Trim$(grdLibEvents.TextMatrix(llRow1, DURATIONINDEX))
                'llDur1 = gStrLengthInTenthToLong(slStr)
                'llStartTime1 = llTime1
                'llEndTime1 = llStartTime1 + llDur1 - 1
                'If llEndTime1 < llStartTime1 Then
                '    llEndTime1 = llStartTime1
                'End If
                llStartTime1 = tmConflictTest(llRowT1).lEventStartTime
                llEndTime1 = tmConflictTest(llRowT1).lEventEndTime
                slIgnoreConflicts = Trim$(grdLibEvents.TextMatrix(llRow1, EVTCONFLICTINDEX))
                If (slIgnoreConflicts = "A") Or (slIgnoreConflicts = "I") Then
                    slPriAudio1 = ""
                    slProtAudio1 = ""
                    slBkupAudio1 = ""
                Else
                    slPriAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, AUDIONAMEINDEX))
                    slProtAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, PROTNAMEINDEX))
                    slBkupAudio1 = Trim$(grdLibEvents.TextMatrix(llRow1, BACKUPNAMEINDEX))
                End If
                slPriItemID1 = Trim$(grdLibEvents.TextMatrix(llRow1, AUDIOITEMIDINDEX))
                slProtItemID1 = Trim$(grdLibEvents.TextMatrix(llRow1, PROTITEMIDINDEX))
                slBkupItemID1 = Trim$(grdLibEvents.TextMatrix(llRow1, AUDIOITEMIDINDEX))
                slBuses1 = Trim$(grdLibEvents.TextMatrix(llRow1, BUSNAMEINDEX))
                llErrorColor = vbRed
                If llAirDate = llNowDate Then
                    If llNowTime > llTime1 Then
                        slStr = Trim$(grdLibEvents.TextMatrix(llRow1, PCODEINDEX))
                        If (slStr <> "") And (Val(slStr) <> 0) Then
                            llErrorColor = BURGUNDY
                        End If
                    End If
                ElseIf llAirDate < llNowDate Then
                    llErrorColor = BURGUNDY
                End If
                If (slPriAudio1 <> "") And (slPriAudio1 <> "[None]") And (slPriAudio1 = slProtAudio1) Then
                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                    grdLibEvents.Row = llRow1
                    grdLibEvents.Col = AUDIONAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    grdLibEvents.Col = PROTNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    If llErrorColor = vbRed Then
                        mSvCheckEventConflicts = True
                    End If
                End If
                If (slPriAudio1 <> "") And (slPriAudio1 <> "[None]") And (slPriAudio1 = slBkupAudio1) Then
                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                    grdLibEvents.Row = llRow1
                    grdLibEvents.Col = AUDIONAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    grdLibEvents.Col = BACKUPNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    If llErrorColor = vbRed Then
                        mSvCheckEventConflicts = True
                    End If
                End If
                If (slBkupAudio1 <> "") And (slBkupAudio1 <> "[None]") And (slBkupAudio1 = slProtAudio1) Then
                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                    grdLibEvents.Row = llRow1
                    grdLibEvents.Col = BACKUPNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    grdLibEvents.Col = PROTNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    If llErrorColor = vbRed Then
                        mSvCheckEventConflicts = True
                    End If
                End If
                ''For llRow2 = llRow1 + 1 To grdLibEvents.Rows - 1 Step 1
                'For llRowT2 = llRowT1 + 1 To UBound(tmConflictTest) - 1 Step 1
                For llRowT2 = 1 To UBound(tmConflictTest) - 1 Step 1
                    llRow2 = tmConflictTest(llRowT2).lRow
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow2, EVENTTYPEINDEX))
                    slEvtType2 = slStr
                    ilTest = False
                    If (slStr <> "") And (llRow1 <> llRow2) Then
                        slIgnoreConflicts = Trim$(grdLibEvents.TextMatrix(llRow2, EVTCONFLICTINDEX))
                        If (slIgnoreConflicts = "A") Or (slIgnoreConflicts = "I") Then
                            slPriAudio2 = ""
                            slProtAudio2 = ""
                            slBkupAudio2 = ""
                        Else
                            slPriAudio2 = Trim$(grdLibEvents.TextMatrix(llRow2, AUDIONAMEINDEX))
                            slProtAudio2 = Trim$(grdLibEvents.TextMatrix(llRow2, PROTNAMEINDEX))
                            slBkupAudio2 = Trim$(grdLibEvents.TextMatrix(llRow2, BACKUPNAMEINDEX))
                        End If
                        slPriItemID2 = Trim$(grdLibEvents.TextMatrix(llRow2, AUDIOITEMIDINDEX))
                        slProtItemID2 = Trim$(grdLibEvents.TextMatrix(llRow2, PROTITEMIDINDEX))
                        slBkupItemID2 = Trim$(grdLibEvents.TextMatrix(llRow2, AUDIOITEMIDINDEX))
                        slBuses2 = Trim$(grdLibEvents.TextMatrix(llRow2, BUSNAMEINDEX))
                        llStartTime2 = tmConflictTest(llRowT2).lEventStartTime
                        llEndTime2 = tmConflictTest(llRowT2).lEventEndTime
                        If (tmConflictTest(llRowT1).sType = tmConflictTest(llRowT2).sType) Then
                            If llEndTime1 + 600 < llStartTime2 Then
                                Exit For
                            End If
                        End If
                        If (llEndTime2 >= llStartTime1) And (llStartTime2 <= llEndTime1) Then
                            If (StrComp(slBuses1, slBuses2, vbTextCompare) = 0) Then
                                ilTest = True
                            Else
                                If (slPriAudio1 <> "") And ((StrComp(slPriAudio1, slPriAudio2, vbTextCompare) = 0) Or (StrComp(slPriAudio1, slProtAudio2, vbTextCompare) = 0) Or (StrComp(slPriAudio1, slBkupAudio2, vbTextCompare) = 0)) Then
                                    ilTest = True
                                Else
                                    If (slProtAudio1 <> "") And ((StrComp(slProtAudio1, slPriAudio2, vbTextCompare) = 0) Or (StrComp(slProtAudio1, slProtAudio2, vbTextCompare) = 0) Or (StrComp(slProtAudio1, slBkupAudio2, vbTextCompare) = 0)) Then
                                        ilTest = True
                                    Else
                                        If (slBkupAudio1 <> "") And ((StrComp(slBkupAudio1, slPriAudio2, vbTextCompare) = 0) Or (StrComp(slBkupAudio1, slProtAudio2, vbTextCompare) = 0) Or (StrComp(slBkupAudio1, slBkupAudio2, vbTextCompare) = 0)) Then
                                            ilTest = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If ilTest Then
                        slEventCategory = ""
                        ilSpotsExist = False
                        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                                slEventCategory = tgCurrETE(ilETE).sCategory
                                If slEventCategory = "A" Then
                                    'Avail are after spots
                                    'For llRow3 = llRow2 + 1 To grdLibEvents.Rows - 1 Step 1
                                    For llRow3 = llRow2 - 1 To grdLibEvents.FixedRows Step -1
                                        slStr = Trim$(grdLibEvents.TextMatrix(llRow3, EVENTTYPEINDEX))
                                        If slStr <> "" Then
                                            If StrComp(grdLibEvents.TextMatrix(llRow2, BUSNAMEINDEX), grdLibEvents.TextMatrix(llRow3, BUSNAMEINDEX), vbTextCompare) = 0 Then
                                                For ilETE3 = 0 To UBound(tgCurrETE) - 1 Step 1
                                                    If StrComp(Trim$(tgCurrETE(ilETE3).sName), slStr, vbTextCompare) = 0 Then
                                                        If tgCurrETE(ilETE3).sCategory = "S" Then
                                                            If gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow2, SPOTAVAILTIMEINDEX), False) = gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow3, SPOTAVAILTIMEINDEX), False) Then
                                                                ilSpotsExist = True
                                                            End If
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilETE3
                                                If ilSpotsExist = True Then
                                                    Exit For
                                                End If
                                            End If
                                            If gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow2, TIMEINDEX), False) + 3000 < gStrTimeInTenthToLong(grdLibEvents.TextMatrix(llRow3, TIMEINDEX), False) Then
                                                Exit For
                                            End If
                                        End If
                                    Next llRow3
                                End If
                                Exit For
                            End If
                        Next ilETE
                        If (slEventCategory = "P") Or ((slEventCategory = "A") And (ilSpotsExist = False)) Or (slEventCategory = "S") Then
                            llStartTime2 = tmConflictTest(llRowT2).lEventStartTime
                            llEndTime2 = tmConflictTest(llRowT2).lEventEndTime
'                            slPriAudio2 = Trim$(grdLibEvents.TextMatrix(llRow2, AUDIONAMEINDEX))
'                            slProtAudio2 = Trim$(grdLibEvents.TextMatrix(llRow2, PROTNAMEINDEX))
'                            slBkupAudio2 = Trim$(grdLibEvents.TextMatrix(llRow2, BACKUPNAMEINDEX))
'                            slPriItemID2 = Trim$(grdLibEvents.TextMatrix(llRow2, AUDIOITEMIDINDEX))
'                            slProtItemID2 = Trim$(grdLibEvents.TextMatrix(llRow2, PROTITEMIDINDEX))
'                            slBkupItemID2 = Trim$(grdLibEvents.TextMatrix(llRow2, AUDIOITEMIDINDEX))
'                            slBuses2 = Trim$(grdLibEvents.TextMatrix(llRow2, BUSNAMEINDEX))
                            ilError = False
                            ilConflictIndex = UBound(tmConflictList)
                            tmConflictList(ilConflictIndex).sType = "E"
                            tmConflictList(ilConflictIndex).sStartDate = ""
                            tmConflictList(ilConflictIndex).sEndDate = ""
                            tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow2, TMCURRSEEINDEX))
                            tmConflictList(ilConflictIndex).iNextIndex = -1
                            ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX))
                            llErrorColor = vbRed
                            If llAirDate = llNowDate Then
                                If llNowTime > llTime2 Then
                                    slStr = Trim$(grdLibEvents.TextMatrix(llRow2, PCODEINDEX))
                                    If (slStr <> "") And (Val(slStr) <> 0) Then
                                        llErrorColor = BURGUNDY
                                    End If
                                End If
                            ElseIf llAirDate < llNowDate Then
                                llErrorColor = BURGUNDY
                            End If
                            If (tmConflictTest(llRowT1).sType = "B") And (tmConflictTest(llRowT2).sType = "B") Then
                                If StrComp(slBuses1, slBuses2, vbTextCompare) = 0 Then
                                    If (llEndTime2 > llStartTime1) And (llStartTime2 < llEndTime1) Or (llStartTime1 = llStartTime2) Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                        grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                        grdLibEvents.Row = llRow1   'llRow2
                                        grdLibEvents.Col = BUSNAMEINDEX 'TIMEINDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                        grdLibEvents.Row = llRow2
                                        grdLibEvents.Col = BUSNAMEINDEX 'DURATIONINDEX
                                        grdLibEvents.CellForeColor = llErrorColor
                                        If llErrorColor = vbRed Then
                                            mSvCheckEventConflicts = True
                                        End If
                                        If Not ilError Then
                                            grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                            ilConflictIndex = UBound(tmConflictList)
                                            tmConflictList(ilConflictIndex).sType = "E"
                                            tmConflictList(ilConflictIndex).sStartDate = ""
                                            tmConflictList(ilConflictIndex).sEndDate = ""
                                            tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                            tmConflictList(ilConflictIndex).iNextIndex = -1
                                            ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                            grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        End If
                                        ilError = True
                                    End If
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "1") And (tmConflictTest(llRowT2).sType = "1") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio1, slPriAudio2, slPriItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1   'llRow2
                                    grdLibEvents.Col = AUDIONAMEINDEX   'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = AUDIONAMEINDEX   'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "1") And (tmConflictTest(llRowT2).sType = "2") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio1, slProtAudio2, slPriItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = AUDIONAMEINDEX   'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = PROTNAMEINDEX    'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "1") And (tmConflictTest(llRowT2).sType = "3") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio1, slBkupAudio2, slPriItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = AUDIONAMEINDEX   'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = BACKUPNAMEINDEX  'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "2") And (tmConflictTest(llRowT2).sType = "1") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio1, slPriAudio2, slProtItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = PROTNAMEINDEX    'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = AUDIONAMEINDEX    'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "2") And (tmConflictTest(llRowT2).sType = "2") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio1, slProtAudio2, slProtItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = PROTNAMEINDEX    'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = PROTNAMEINDEX    'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "2") And (tmConflictTest(llRowT2).sType = "3") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio1, slBkupAudio2, slProtItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = PROTNAMEINDEX    'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = BACKUPNAMEINDEX  'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "3") And (tmConflictTest(llRowT2).sType = "1") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio1, slPriAudio2, slBkupItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = BACKUPNAMEINDEX  'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = AUDIONAMEINDEX    'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "3") And (tmConflictTest(llRowT2).sType = "2") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio1, slProtAudio2, slBkupItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = BACKUPNAMEINDEX  'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = PROTNAMEINDEX    'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                            If (tmConflictTest(llRowT1).sType = "3") And (tmConflictTest(llRowT2).sType = "3") Then
                                If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio1, slBkupAudio2, slBkupItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, False, slBuses1, slBuses2) Then
                                    grdLibEvents.TextMatrix(llRow1, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.TextMatrix(llRow2, ERRORFIELDSORTINDEX) = "0"
                                    grdLibEvents.Row = llRow1
                                    grdLibEvents.Col = BACKUPNAMEINDEX  'TIMEINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    grdLibEvents.Row = llRow2
                                    grdLibEvents.Col = BACKUPNAMEINDEX  'DURATIONINDEX
                                    grdLibEvents.CellForeColor = llErrorColor
                                    If llErrorColor = vbRed Then
                                        mSvCheckEventConflicts = True
                                    End If
                                    If Not ilError Then
                                        grdLibEvents.TextMatrix(llRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                        ilConflictIndex = UBound(tmConflictList)
                                        tmConflictList(ilConflictIndex).sType = "E"
                                        tmConflictList(ilConflictIndex).sStartDate = ""
                                        tmConflictList(ilConflictIndex).sEndDate = ""
                                        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
                                        tmConflictList(ilConflictIndex).iNextIndex = -1
                                        ilStartConflictIndex = Val(grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX))
                                        grdLibEvents.TextMatrix(llRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tmConflictList())))
                                    End If
                                    ilError = True
                                End If
                            End If
                        End If
                    End If
                Next llRowT2
            End If
        End If
    Next llRowT1
    Erase tmConflictTest
End Function

Private Function mCheckEventConflicts() As Integer
    Dim llRow1 As Long
    Dim llRow2 As Long
    Dim ilHour1 As Integer
    Dim ilHour2 As Integer
    Dim slHours1 As String
    Dim slHours2 As String
    Dim ilDay1 As Integer
    Dim ilDay2 As Integer
    Dim slDays1 As String
    Dim slDays2 As String
    Dim slStr As String
    Dim ilBus1 As Integer
    Dim ilBus2 As Integer
    Dim llTime1 As Long
    Dim llTime2 As Long
    Dim llDur1 As Long
    Dim llDur2 As Long
    Dim llStartTime1 As Long
    Dim llEndTime1 As Long
    Dim llStartTime2 As Long
    Dim llEndTime2 As Long
    Dim ilPriAudio1 As Integer
    Dim ilProtAudio1 As Integer
    Dim ilBkupAudio1 As Integer
    Dim ilPriAudio2 As Integer
    Dim ilProtAudio2 As Integer
    Dim ilBkupAudio2 As Integer
    Dim slPriItemID1 As String
    Dim slPriItemID2 As String
    Dim slProtItemID1 As String
    Dim slProtItemID2 As String
    Dim slBkupItemID1 As String
    Dim slBkupItemID2 As String
    Dim ilASE As Integer
    Dim ilETE As Integer
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llErrorColor As Long
    Dim slEventCategory1 As String
    Dim slEventCategory2 As String
    Dim ilError As Integer
    Dim ilConflictIndex As Integer
    Dim slMsgFileName As String
    Dim ilRet As Integer
    Dim llLoop1 As Long
    Dim llLoop2 As Long
    Dim llPostTime As Long
    Dim llPreTime As Long
    Dim ilATE As Integer
    Dim ilCheckBus As Integer
    Dim ilCheckAudio As Integer
    Dim llEventID1 As Long
    Dim llEventID2 As Long
    Dim llGridRow1 As Long
    Dim llGridRow2 As Long
    Dim llStart As Long
    Dim llEnd As Long
    Dim ilStep As Integer
    Dim ilSpotsExist As Integer
    Dim ilBDE As Integer
    Dim llSEEIndex As Integer
    
    mCheckEventConflicts = False
    'Conflict test not required when schedule created as each time library/template is added a conflict test is done
    'Removed conflict checking for schedule
    'Remove Conflict checking after merge
    'These test can be removed because library and templates are checked for conflicts.
    'If avail will not be defined to use the same audio on different buses
    ilRet = mOpenConflictMsgFile(smAirDate, slMsgFileName)
    If Not ilRet Then
        Exit Function
    End If
    Print #hmMsg, "** Conflict Test: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, "For: " & smAirDate
    
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    
    'Find max Pre and Post adjustment time to help minimize compare time
    llPostTime = 0
    llPreTime = 0
    For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
        If tgCurrATE(ilATE).lPreBufferTime > llPreTime Then
            llPreTime = tgCurrATE(ilATE).lPreBufferTime
        End If
        If tgCurrATE(ilATE).lPostBufferTime > llPostTime Then
            llPostTime = tgCurrATE(ilATE).lPostBufferTime
        End If
    Next ilATE

    lbcSort.Clear
    'For llRow1 = LBound(tmCurrSEE) To UBound(tmCurrSEE) - 1 Step 1
    '    llTime1 = tmCurrSEE(llRow1).lTime
    '    slStr = Trim$(Str$(llTime1))
    '    Do While Len(slStr) < 8
    '        slStr = "0" & slStr
    '    Loop
    '    lbcSort.AddItem slStr
    '    lbcSort.ItemData(lbcSort.NewIndex) = llRow1
    'Next llRow1
    
    For llRow1 = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow1, EVENTTYPEINDEX)) <> "" Then
            slStr = Trim$(Str$(llRow1))
            Do While Len(slStr) < 8
                slStr = "0" & slStr
            Loop
            lbcSort.AddItem slStr
            lbcSort.ItemData(lbcSort.NewIndex) = Val(grdLibEvents.TextMatrix(llRow1, TMCURRSEEINDEX))
        End If
    Next llRow1
'    For llRow1 = LBound(tmCurrSEE) To UBound(tmCurrSEE) - 1 Step 1
    ilSpotsExist = False
    For llLoop1 = 0 To lbcSort.ListCount - 1 Step 1
        llGridRow1 = Val(lbcSort.List(llLoop1))
        llRow1 = lbcSort.ItemData(llLoop1)
        slEventCategory1 = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tmCurrSEE(llRow1).iEteCode = tgCurrETE(ilETE).iCode Then
                If tgCurrETE(ilETE).sCategory = "S" Then
                    ilSpotsExist = True
                End If
                Exit For
            End If
        Next ilETE
        If ilSpotsExist Then
            Exit For
        End If
    Next llLoop1
    For llLoop1 = 0 To lbcSort.ListCount - 1 Step 1
        llGridRow1 = Val(lbcSort.List(llLoop1))
        llRow1 = lbcSort.ItemData(llLoop1)
        slEventCategory1 = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tmCurrSEE(llRow1).iEteCode = tgCurrETE(ilETE).iCode Then
                slEventCategory1 = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        If (slEventCategory1 = "P") Or ((slEventCategory1 = "A") And (Not ilSpotsExist)) Or ((slEventCategory1 = "S") And (ilSpotsExist)) Then
            If (tmCurrSEE(llRow1).sAction <> "D") And (tmCurrSEE(llRow1).sAction <> "R") Then
                llEventID1 = tmCurrSEE(llRow1).lEventID
                If (slEventCategory1 = "S") And (ilSpotsExist) Then
                    llTime1 = tmCurrSEE(llRow1).lSpotTime
                Else
                    llTime1 = tmCurrSEE(llRow1).lTime
                End If
                llDur1 = tmCurrSEE(llRow1).lDuration
                llStartTime1 = llTime1
                llEndTime1 = llStartTime1 + llDur1
                If llEndTime1 < llStartTime1 Then
                    llEndTime1 = llStartTime1
                End If
                If llEndTime1 > 864000 Then
                    llEndTime1 = 864000
                End If
                llErrorColor = vbRed
                If llAirDate = llNowDate Then
                    If llNowTime > llTime1 Then
                        slStr = Trim$(grdLibEvents.TextMatrix(llGridRow1, PCODEINDEX))
                        If (slStr <> "") And (Val(slStr) <> 0) Then
                            llErrorColor = BURGUNDY
                        End If
                    End If
                ElseIf llAirDate < llNowDate Then
                    llErrorColor = BURGUNDY
                End If
                ilPriAudio1 = -1
                For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                    If tmCurrSEE(llRow1).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                        ilPriAudio1 = tgCurrASE(ilASE).iPriAneCode
                        Exit For
                    End If
                Next ilASE
                ilProtAudio1 = tmCurrSEE(llRow1).iProtAneCode
                ilBkupAudio1 = tmCurrSEE(llRow1).iBkupAneCode
                slPriItemID1 = Trim$(tmCurrSEE(llRow1).sAudioItemID)
                slProtItemID1 = Trim$(tmCurrSEE(llRow1).sProtItemID)
                slBkupItemID1 = Trim$(tmCurrSEE(llRow1).sAudioItemID)
                ilBus1 = tmCurrSEE(llRow1).iBdeCode
                If (ilPriAudio1 > 0) And (ilPriAudio1 = ilProtAudio1) Then
                    grdLibEvents.TextMatrix(llGridRow1, ERRORFIELDSORTINDEX) = "0"
                    grdLibEvents.Row = llGridRow1
                    grdLibEvents.Col = AUDIONAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    grdLibEvents.Col = PROTNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    If llErrorColor = vbRed Then
                        mCheckEventConflicts = True
                    End If
                    mPrintEventMsg "Primary and Protection defined with same Audio Name", ilPriAudio1, llEventID1, ilBus1, llTime1
                End If
                If (ilPriAudio1 > 0) And (ilPriAudio1 = ilBkupAudio1) Then
                    grdLibEvents.TextMatrix(llGridRow1, ERRORFIELDSORTINDEX) = "0"
                    grdLibEvents.Row = llGridRow1
                    grdLibEvents.Col = AUDIONAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    grdLibEvents.Col = BACKUPNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    If llErrorColor = vbRed Then
                        mCheckEventConflicts = True
                    End If
                    mPrintEventMsg "Primary and Backup defined with same Audio Name", ilPriAudio1, llEventID1, ilBus1, llTime1
                End If
                If (ilBkupAudio1 > 0) And (ilBkupAudio1 = ilProtAudio1) Then
                    grdLibEvents.TextMatrix(llGridRow1, ERRORFIELDSORTINDEX) = "0"
                    grdLibEvents.Row = llGridRow1
                    grdLibEvents.Col = BACKUPNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    grdLibEvents.Col = PROTNAMEINDEX
                    grdLibEvents.CellForeColor = llErrorColor
                    If llErrorColor = vbRed Then
                        mCheckEventConflicts = True
                    End If
                    mPrintEventMsg "Backup and Protection defined with same Audio Name", ilBkupAudio1, llEventID1, ilBus1, llTime1
                End If
    '            For llRow2 = llRow1 + 1 To UBound(tmCurrSEE) - 1 Step 1
                'For llLoop2 = 0 To lbcSort.ListCount - 1 Step 1
                llStart = llLoop1 + 1
                llEnd = lbcSort.ListCount - 1
                ilStep = 1
                For llLoop2 = llStart To llEnd Step ilStep
                    llGridRow2 = Val(lbcSort.List(llLoop2))
                    llRow2 = lbcSort.ItemData(llLoop2)
                    slEventCategory2 = ""
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If tmCurrSEE(llRow2).iEteCode = tgCurrETE(ilETE).iCode Then
                            slEventCategory2 = tgCurrETE(ilETE).sCategory
                            Exit For
                        End If
                    Next ilETE
                    llEventID2 = tmCurrSEE(llRow2).lEventID
                    If (slEventCategory2 = "S") And (ilSpotsExist) Then
                        llTime2 = tmCurrSEE(llRow2).lSpotTime
                    Else
                        llTime2 = tmCurrSEE(llRow2).lTime
                    End If
                    llDur2 = tmCurrSEE(llRow2).lDuration
                    llStartTime2 = llTime2
                    llEndTime2 = llStartTime2 + llDur2
                    If llEndTime2 < llStartTime2 Then
                        llEndTime2 = llStartTime2
                    End If
                    If llEndTime2 > 864000 Then
                        llEndTime2 = 864000
                    End If
                    'Compare can stop once the start time of the llLoop2 item is beyond compare time
                    If llEndTime1 + llPostTime + llPreTime + 3000 < llStartTime2 Then
                        Exit For
                    End If
                    If (slEventCategory2 = "P") Or ((slEventCategory2 = "A") And (Not ilSpotsExist)) Or ((slEventCategory2 = "S") And (ilSpotsExist)) Then
                        If (tmCurrSEE(llRow2).sAction <> "D") And (tmCurrSEE(llRow2).sAction <> "R") Then
                            ilPriAudio2 = -1
                            For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                If tmCurrSEE(llRow2).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                    ilPriAudio2 = tgCurrASE(ilASE).iPriAneCode
                                    Exit For
                                End If
                            Next ilASE
                            ilProtAudio2 = tmCurrSEE(llRow2).iProtAneCode
                            ilBkupAudio2 = tmCurrSEE(llRow2).iBkupAneCode
                            slPriItemID2 = Trim$(tmCurrSEE(llRow2).sAudioItemID)
                            slProtItemID2 = Trim$(tmCurrSEE(llRow2).sProtItemID)
                            slBkupItemID2 = Trim$(tmCurrSEE(llRow2).sAudioItemID)
                            ilError = False
                            ilConflictIndex = UBound(tmConflictList)
                            tmConflictList(ilConflictIndex).sType = "E"
                            tmConflictList(ilConflictIndex).sStartDate = ""
                            tmConflictList(ilConflictIndex).sEndDate = ""
                            tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llGridRow2, TMCURRSEEINDEX))
                            tmConflictList(ilConflictIndex).iNextIndex = -1
                            llErrorColor = vbRed
                            If llAirDate = llNowDate Then
                                If llNowTime > llTime2 Then
                                    slStr = Trim$(grdLibEvents.TextMatrix(llRow2, PCODEINDEX))
                                    If (slStr <> "") And (Val(slStr) <> 0) Then
                                        llErrorColor = BURGUNDY
                                    End If
                                End If
                            ElseIf llAirDate < llNowDate Then
                                llErrorColor = BURGUNDY
                            End If
                            ilBus2 = tmCurrSEE(llRow2).iBdeCode
                            ilCheckBus = True
                            If (tmCurrSEE(llRow1).lDheCode <> tmCurrSEE(llRow2).lDheCode) Then
                                If ((tmCurrSEE(llRow1).sIgnoreConflicts = "B") Or (tmCurrSEE(llRow1).sIgnoreConflicts = "I")) Then
                                    ilCheckBus = False
                                End If
                                If ((tmCurrSEE(llRow2).sIgnoreConflicts = "B") Or (tmCurrSEE(llRow2).sIgnoreConflicts = "I")) Then
                                    ilCheckBus = False
                                End If
                            End If
                            If tgSOE.sMatchBNotT = "N" Then
                                ilCheckBus = False
                            End If
                            If (ilBus1 = ilBus2) And (ilCheckBus) Then
                                If (llEndTime2 > llStartTime1) And (llStartTime2 < llEndTime1) Or (llStartTime1 = llStartTime2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, BUSNAMEINDEX, llGridRow2, BUSNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintBusConflictMsg "Bus Conflict", ilBus1, llEventID1, llEventID2, llTime1, llTime2
                                End If
                            End If
                            ilCheckAudio = True
                            If (tmCurrSEE(llRow1).lDheCode <> tmCurrSEE(llRow2).lDheCode) Then
                                If ((tmCurrSEE(llRow1).sIgnoreConflicts = "A") Or (tmCurrSEE(llRow1).sIgnoreConflicts = "I")) Then
                                    ilCheckAudio = False
                                End If
                                If ((tmCurrSEE(llRow2).sIgnoreConflicts = "A") Or (tmCurrSEE(llRow2).sIgnoreConflicts = "I")) Then
                                    ilCheckAudio = False
                                End If
                            End If
                            If ilCheckAudio Then
                                If mAudioConflicts(ilPriAudio1, ilPriAudio2, slPriItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, AUDIONAMEINDEX, llGridRow2, AUDIONAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Primary and Primary Audio Conflict", ilPriAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilPriAudio1, ilProtAudio2, slPriItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, AUDIONAMEINDEX, llGridRow2, PROTNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Primary and Protection Audio Conflict", ilPriAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilPriAudio1, ilBkupAudio2, slPriItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, AUDIONAMEINDEX, llGridRow2, BACKUPNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Primary and Backup Audio Conflict", ilPriAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilProtAudio1, ilPriAudio2, slProtItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, PROTNAMEINDEX, llGridRow2, AUDIONAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Protection and Primary Audio Conflict", ilProtAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilProtAudio1, ilProtAudio2, slProtItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, PROTNAMEINDEX, llGridRow2, PROTNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Protection and Protection Audio Conflict", ilProtAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilProtAudio1, ilBkupAudio2, slProtItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, PROTNAMEINDEX, llGridRow2, BACKUPNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Protection and Backup Audio Conflict", ilProtAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilBkupAudio1, ilPriAudio2, slBkupItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, BACKUPNAMEINDEX, llGridRow2, AUDIONAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Backup and Primary Audio Conflict", ilBkupAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilBkupAudio1, ilProtAudio2, slBkupItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, BACKUPNAMEINDEX, llGridRow2, PROTNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Backup and Protection Audio Conflict", ilBkupAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                                If mAudioConflicts(ilBkupAudio1, ilBkupAudio2, slBkupItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mGridErrorRows llErrorColor, llGridRow1, BACKUPNAMEINDEX, llGridRow2, BACKUPNAMEINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintAudioConflictMsg "Backup and Backup Audio Conflict", ilBkupAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                End If
                            End If
                        End If
                    End If
                Next llLoop2
    '            Next llRow2
            End If
        End If
'    Next llRow1
    Next llLoop1
    'Check that same bus/cart are not within the lgCartUnloadTime
    If (ilSpotsExist) And (lgCartUnloadTime >= 0) Then
        ReDim tmCartUnloadTest(0 To UBound(tgCurrBDE)) As CARTUNLOADTEST
        For llRow1 = 0 To UBound(tmCartUnloadTest) Step 1
            tmCartUnloadTest(llRow1).lEventID = -1
        Next llRow1
        'Reload with bus codes
        For llLoop1 = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            llSEEIndex = Val(grdLibEvents.TextMatrix(llLoop1, TMCURRSEEINDEX))
            If Trim$(grdLibEvents.TextMatrix(llLoop1, EVENTTYPEINDEX)) <> "" Then
                slEventCategory1 = ""
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If tmCurrSEE(llSEEIndex).iEteCode = tgCurrETE(ilETE).iCode Then
                        slEventCategory1 = tgCurrETE(ilETE).sCategory
                        Exit For
                    End If
                Next ilETE
                If slEventCategory1 = "S" Then
                    If (tmCurrSEE(llSEEIndex).sAction <> "D") And (tmCurrSEE(llSEEIndex).sAction <> "R") And (Left(tmCurrSEE(llSEEIndex).sAudioItemID, 1) <> "L") Then
                        ilError = False
                        ilConflictIndex = UBound(tmConflictList)
                        tmConflictList(ilConflictIndex).sType = "E"
                        tmConflictList(ilConflictIndex).sStartDate = ""
                        tmConflictList(ilConflictIndex).sEndDate = ""
                        '9/5/14: Fix what is displayed within the conflict table.
                        tmConflictList(ilConflictIndex).lIndex = llSEEIndex  'llLoop1
                        tmConflictList(ilConflictIndex).iNextIndex = -1
                        llErrorColor = vbRed
                        If llAirDate = llNowDate Then
                            If llNowTime > llTime2 Then
                                slStr = Trim$(grdLibEvents.TextMatrix(llLoop1, PCODEINDEX))
                                If (slStr <> "") And (Val(slStr) <> 0) Then
                                    llErrorColor = BURGUNDY
                                End If
                            End If
                        ElseIf llAirDate < llNowDate Then
                            llErrorColor = BURGUNDY
                        End If
                        llEventID1 = tmCurrSEE(llSEEIndex).lEventID
                        llTime1 = tmCurrSEE(llSEEIndex).lSpotTime
                        llDur1 = tmCurrSEE(llSEEIndex).lDuration
                        llStartTime1 = llTime1
                        llEndTime1 = llStartTime1 + llDur1
                        If llEndTime1 < llStartTime1 Then
                            llEndTime1 = llStartTime1
                        End If
                        If llEndTime1 > 864000 Then
                            llEndTime1 = 864000
                        End If
                        ilBDE = gBinarySearchBDE(tmCurrSEE(llSEEIndex).iBdeCode, tgCurrBDE())
                        If ilBDE <> -1 Then
                            If tmCartUnloadTest(ilBDE).lEventID >= 0 Then
                                If (tmCartUnloadTest(ilBDE).lEndTime + 10 * lgCartUnloadTime >= llStartTime1) And (Trim$(tmCartUnloadTest(ilBDE).AudioItemID) = Trim$(tmCurrSEE(llSEEIndex).sAudioItemID)) Then
                                    mGridErrorRows llErrorColor, tmCartUnloadTest(ilBDE).lGridRow, AUDIOITEMIDINDEX, llLoop1, AUDIOITEMIDINDEX, ilError
                                    If llErrorColor = vbRed Then
                                        mCheckEventConflicts = True
                                    End If
                                    mPrintCartConflictMsg "Cart Conflict", tmCurrSEE(llSEEIndex).iBdeCode, tmCurrSEE(llSEEIndex).sAudioItemID, tmCartUnloadTest(ilBDE).lEventID, llEventID1, tmCartUnloadTest(ilBDE).lStartTime, llStartTime1
                                End If
                            End If
                            tmCartUnloadTest(ilBDE).lGridRow = llLoop1
                            tmCartUnloadTest(ilBDE).lEventID = llEventID1
                            tmCartUnloadTest(ilBDE).lStartTime = llStartTime1
                            tmCartUnloadTest(ilBDE).lEndTime = llEndTime1
                            tmCartUnloadTest(ilBDE).AudioItemID = tmCurrSEE(llSEEIndex).sAudioItemID
                        End If
                    End If
                End If
            End If
        Next llLoop1
    End If
    Close hmMsg

End Function

Private Function mAudioConflicts(ilAudio1 As Integer, ilAudio2 As Integer, slItemID1 As String, slItemID2 As String, llStartTime1 As Long, llEndTime1 As Long, llStartTime2 As Long, llEndTime2 As Long, ilBus1 As Integer, ilBus2 As Integer) As Integer
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim llAdjStartTime1 As Long
    Dim llAdjEndTime1 As Long
    Dim llAdjStartTime2 As Long
    Dim llAdjEndTime2 As Long
    
    mAudioConflicts = False
    If ilAudio1 <= 0 Then
        Exit Function
    End If
    If ilAudio1 = ilAudio2 Then
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            If tgCurrANE(ilANE).iCode = ilAudio1 Then
                If tgCurrANE(ilANE).sCheckConflicts <> "N" Then
                    If (llStartTime1 <> llStartTime2) Or (llEndTime1 <> llEndTime2) Then
                        If tgSOE.sMatchANotT <> "N" Then
                            For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                                If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                    llPreTime = tgCurrATE(ilATE).lPreBufferTime
                                    llPostTime = tgCurrATE(ilATE).lPostBufferTime
                                    llAdjStartTime1 = llStartTime1 - llPreTime
                                    If llAdjStartTime1 < 0 Then
                                        llAdjStartTime1 = 0
                                    End If
                                    llAdjEndTime1 = llEndTime1 + llPostTime
                                    If llAdjEndTime1 > 864000 Then
                                        llAdjEndTime1 = 864000
                                    End If
                                    llAdjStartTime2 = llStartTime2 - llPreTime
                                    If llAdjStartTime2 < 0 Then
                                        llAdjStartTime2 = 0
                                    End If
                                    llAdjEndTime2 = llEndTime2 + llPostTime
                                    If llAdjEndTime2 > 864000 Then
                                        llAdjEndTime2 = 864000
                                    End If
                                    If (llAdjEndTime2 > llAdjStartTime1) And (llAdjStartTime2 < llAdjEndTime1) Then
                                        mAudioConflicts = True
                                        Exit Function
                                    End If
                                End If
                            Next ilATE
                        End If
                    Else
                        If ilBus1 <> ilBus2 Then
                            If tgSOE.sMatchATNotB <> "N" Then
                                mAudioConflicts = True
                                Exit Function
                            End If
                        Else
                            If tgSOE.sMatchATBNotI <> "N" Then
                                If (Trim$(slItemID1) <> "") And (Trim$(slItemID2) <> "") Then
                                    If StrComp(slItemID1, slItemID2, vbTextCompare) <> 0 Then
                                        mAudioConflicts = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilANE
    End If
    
End Function

Private Function mOpenConflictMsgFile(slAirDate As String, slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String
    Dim slNowDate As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String

    On Error GoTo mOpenConflictMsgFileErr:
    ilRet = 0
    slAirYear = Year(slAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(slAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(slAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If
    
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgMsgDirectory & "ConflictFromSchedule_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenConflictMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        mOpenConflictMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMsgFileName = slToFile
    mOpenConflictMsgFile = True
    Exit Function
mOpenConflictMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Sub mPrintAudioConflictMsg(slMsg As String, ilAudio1 As Integer, llEventID1 As Long, llEventID2 As Long, ilBus1 As Integer, ilBus2 As Integer, llTime1 As Long, llTime2 As Long)
    Dim slTime1 As String
    Dim slTime2 As String
    Dim slBus1 As String
    Dim slBus2 As String
    Dim ilBDE As Integer
    Dim slAudio As String
    Dim ilANE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slTime2 = gLongToStrLengthInTenth(llTime2, True)
    slAudio = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tgCurrANE(ilANE).iCode = ilAudio1 Then
            slAudio = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    slBus2 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus2 = tgCurrBDE(ilBDE).iCode Then
            slBus2 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID1 > 0 Then
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slAudio & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & "(Bus " & slBus1 & ")" & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2 & "(Bus " & slBus2 & ")"
        Else
            Print #hmMsg, slMsg & " on " & slAudio & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & "(Bus " & slBus1 & ")" & " and at " & slTime2 & "(Bus " & slBus2 & ")"
        End If
    Else
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slAudio & " for events at " & slTime1 & "(Bus " & slBus1 & ")" & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2 & "(Bus " & slBus2 & ")"
        Else
            Print #hmMsg, slMsg & " on " & slAudio & " for events at " & slTime1 & "(Bus " & slBus1 & ")" & " and " & slTime2 & "(Bus " & slBus2 & ")"
        End If
    End If
End Sub
Private Sub mPrintBusConflictMsg(slMsg As String, ilBus1 As Integer, llEventID1 As Long, llEventID2 As Long, llTime1 As Long, llTime2 As Long)
    Dim slTime1 As String
    Dim slTime2 As String
    Dim slBus1 As String
    Dim slBus2 As String
    Dim ilBDE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slTime2 = gLongToStrLengthInTenth(llTime2, True)
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID1 > 0 Then
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slBus1 & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2
        Else
            Print #hmMsg, slMsg & " on " & slBus1 & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & " and at " & slTime2
        End If
    Else
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slBus1 & " for events at " & slTime1 & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2
        Else
            Print #hmMsg, slMsg & " on " & slBus1 & " for events at " & slTime1 & " and " & slTime2
        End If
    End If
End Sub

Private Sub mPrintCartConflictMsg(slMsg As String, ilBus1 As Integer, slAudioItemID As String, llEventID1 As Long, llEventID2 As Long, llTime1 As Long, llTime2 As Long)
    Dim slTime1 As String
    Dim slTime2 As String
    Dim slBus1 As String
    Dim ilBDE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slTime2 = gLongToStrLengthInTenth(llTime2, True)
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID1 > 0 Then
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slBus1 & " Audio Item ID " & Trim$(slAudioItemID) & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2
        Else
            Print #hmMsg, slMsg & " on " & slBus1 & " Audio Item ID " & Trim$(slAudioItemID) & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & " and at " & slTime2
        End If
    Else
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slBus1 & " Audio Item ID " & Trim$(slAudioItemID) & " for events at " & slTime1 & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2
        Else
            Print #hmMsg, slMsg & " on " & slBus1 & " Audio Item ID " & Trim$(slAudioItemID) & " for events at " & slTime1 & " and " & slTime2
        End If
    End If
End Sub

Private Sub mPrintEventMsg(slMsg As String, ilAudio1 As Integer, llEventID As Long, ilBus1 As Integer, llTime1 As Long)
    Dim slTime1 As String
    Dim slBus1 As String
    Dim ilBDE As Integer
    Dim slAudio As String
    Dim ilANE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slAudio = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tgCurrANE(ilANE).iCode = ilAudio1 Then
            slAudio = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID > 0 Then
        Print #hmMsg, slMsg & " on " & slAudio & " for event ID " & Trim$(Str$(llEventID)) & " at " & slTime1 & "(Bus " & slBus1 & ")"
    Else
        Print #hmMsg, slMsg & " on " & slAudio & " for events at " & slTime1 & "(Bus " & slBus1 & ")"
    End If
End Sub

Private Sub mLoadCTE_1(slEventType As String)
    Dim llRow As Long
    Dim slStr As String
    
    lbcCTE_1.Clear
    If StrComp(slEventType, smSpotEventTypeName, vbTextCompare) <> 0 Then
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
            If (slStr <> "") And (StrComp(slStr, smSpotEventTypeName, vbTextCompare) <> 0) Then
                slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
                If slStr <> "" Then
                    If gListBoxFind(lbcCTE_1, slStr, True) < 0 Then
                        lbcCTE_1.AddItem slStr
                    End If
                End If
            End If
        Next llRow
        lbcCTE_1.AddItem "[None]", 0
        lbcCTE_1.ItemData(lbcCTE_1.NewIndex) = 0
    Else
        For llRow = 0 To UBound(tgCurrARE) - 1 Step 1
            If (Trim$(tgCurrARE(llRow).sName) <> "") And (StrComp(Trim$(tgCurrARE(llRow).sName), "[None]", vbTextCompare) <> 0) Then
                lbcCTE_1.AddItem Trim$(tgCurrARE(llRow).sName)
                lbcCTE_1.ItemData(lbcCTE_1.NewIndex) = tgCurrARE(llRow).lCode
            End If
        Next llRow
        lbcCTE_1.AddItem "[None]", 0
        lbcCTE_1.ItemData(lbcCTE_1.NewIndex) = 0
    End If
End Sub


Private Sub mLoadCTE_2()
    Dim llRow As Long
    Dim slStr As String
    
    lbcCTE_2.Clear
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If (slStr <> "") Then
            slStr = Trim$(grdLibEvents.TextMatrix(llRow, TITLE2INDEX))
            If slStr <> "" Then
                If gListBoxFind(lbcCTE_2, slStr, True) < 0 Then
                    lbcCTE_2.AddItem slStr
                End If
            End If
        End If
    Next llRow
    lbcCTE_2.AddItem "[None]", 0
    lbcCTE_2.ItemData(lbcCTE_2.NewIndex) = 0
End Sub





Private Function mOpenAutoExportFile(slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilPosE As Integer
    Dim slName As String
    Dim slPath As String
    Dim slDateTime As String
    Dim slChar As String
    Dim slSeqNo As String

    On Error GoTo mOpenAutoExportFileErr:
    slNowDate = Format$(gNow(), sgShowDateForm)
    slName = ""
    slPath = ""
    For ilLoop = 0 To UBound(tgCurrAPE) - 1 Step 1
        If tgCurrAPE(ilLoop).sType = "CE" Then
            If ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                If tmSHE.sLoadedAutoStatus = "L" Then
                    slName = Trim$(tgCurrAPE(ilLoop).sChgFileName) & "." & Trim$(tgCurrAPE(ilLoop).sChgFileExt)
                Else
                    slName = Trim$(tgCurrAPE(ilLoop).sNewFileName) & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
                End If
                'Check for date
                ilPos = InStr(1, slName, "Date", vbTextCompare)
                If ilPos > 0 Then
                    '2/6/12: Switch to Air date
                    'If Trim$(tgCurrAPE(ilLoop).sDateFormat) <> "" Then
                    '    slDate = Format$(slNowDate, Trim$(tgCurrAPE(ilLoop).sDateFormat))
                    'Else
                    '    slDate = Format$(slNowDate, "yymmdd")
                    'End If
                    If Trim$(tgCurrAPE(ilLoop).sDateFormat) <> "" Then
                        slDate = Format$(smAirDate, Trim$(tgCurrAPE(ilLoop).sDateFormat))
                    Else
                        slDate = Format$(smAirDate, "yymmdd")
                    End If
                    slName = Left$(slName, ilPos - 1) & slDate & Mid(slName, ilPos + 4)
                End If
                'Check for Time
                ilPos = InStr(1, slName, "Time", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilLoop).sTimeFormat) <> "" Then
                        slTime = Format$(slNowDate, Trim$(tgCurrAPE(ilLoop).sTimeFormat))
                    Else
                        slTime = Format$(slNowDate, "hhmmss")
                    End If
                    slName = Left$(slName, ilPos - 1) & slTime & Mid(slName, ilPos + 4)
                End If
                'Check for Sequence number
                If tmSHE.sLoadedAutoStatus = "L" Then
                    ilPos = InStr(1, slName, "S", vbTextCompare)
                    If ilPos > 0 Then
                        ilPosE = ilPos + 1
                        Do While ilPosE <= Len(slName)
                            slChar = Mid$(slName, ilPosE, 1)
                            If StrComp(slChar, "S", vbTextCompare) <> 0 Then
                                Exit Do
                            End If
                            ilPosE = ilPosE + 1
                        Loop
                    End If
                    slSeqNo = Trim$(Str$(tmSHE.iChgSeqNo + 1))
                    Do While Len(slSeqNo) < ilPosE - ilPos
                        slSeqNo = "0" & slSeqNo
                    Loop
                    Mid$(slName, ilPos, ilPosE - ilPos) = slSeqNo
                End If
            End If
            'slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            If (Not igTestSystem) And ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            ElseIf (igTestSystem) And (tgCurrAPE(ilLoop).sSubType = "T") Then
                slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            End If
            If slPath <> "" Then
                If right(slPath, 1) <> "\" Then
                    slPath = slPath & "\"
                End If
            End If
            'Exit For
        End If
    Next ilLoop
    If slName = "" Then
        MsgBox "Load File Name missing for Client from Automation Equipment Definition", vbOKOnly
        mOpenAutoExportFile = False
        Exit Function
    End If
    If slPath = "" Then
        MsgBox "Load Path missing for Client from Automation Equipment Definition", vbOKOnly
        mOpenAutoExportFile = False
        Exit Function
    End If
    
    ilRet = 0
    slToFile = slPath & slName
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    
    '3/31/13: Create temp file name
    ilRet = 0
    sgLoadFileName = slToFile
    ilPos = InStr(1, slToFile, ".", vbBinaryCompare)
    If ilPos > 0 Then
        Mid(slToFile, ilPos, 1) = "_"
        slToFile = slToFile & ".txt"
    Else
        mOpenAutoExportFile = False
        Exit Function
    End If
    sgTmpLoadFileName = slToFile
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    
    ilRet = 0
    On Error GoTo mOpenAutoExportFileErr:
    hmExport = FreeFile
    Open slToFile For Output As hmExport
    If ilRet <> 0 Then
        Close hmExport
        hmExport = -1
        MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        mOpenAutoExportFile = False
        Exit Function
    End If
    On Error GoTo 0
'    Print #hmExport, "** Test : " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    Print #hmExport, ""
    slMsgFileName = slToFile
    mOpenAutoExportFile = True
    Exit Function
mOpenAutoExportFileErr:
    ilRet = 1
    Resume Next
End Function


Private Function mOpenMergeFile(slMergeFileCP As String, slMergeFileCB As String) As Integer
    Dim slToFile As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slName As String
    Dim slPathCP As String
    Dim slPathCB As String
    Dim slPathCPTest As String
    Dim slPathCBTest As String
    Dim slDateTime As String

    On Error GoTo mOpenMergeFileErr:
    slName = ""
    slPathCP = ""
    slPathCB = ""
    slPathCPTest = ""
    slPathCBTest = ""
    
    slName = Trim$(tgSOE.sMergeFileFormat) & "." & Trim$(tgSOE.sMergeFileExt)
    ilPos = InStr(1, slName, "Date", vbTextCompare)
    If ilPos > 0 Then
        If Trim$(tgSOE.sMergeDateFormat) <> "" Then
            slDate = Format$(smAirDate, Trim$(tgSOE.sMergeDateFormat))
        Else
            slNowDate = Format$(gNow(), "ddddd")
            slDate = Format$(slNowDate, "yymmdd")
        End If
        slName = Left$(slName, ilPos - 1) & slDate & Mid(slName, ilPos + 4)
    End If
    For ilLoop = 0 To UBound(tgCurrSPE) - 1 Step 1
        If Not igTestSystem Then
            If (tgCurrSPE(ilLoop).sType = "CP") And ((tgCurrSPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrSPE(ilLoop).sSubType) = "")) Then
                slPathCP = Trim$(tgCurrSPE(ilLoop).sPath)
                If slPathCP <> "" Then
                    If right(slPathCP, 1) <> "\" Then
                        slPathCP = slPathCP & "\"
                    End If
                End If
                'Exit For
            End If
            If (tgCurrSPE(ilLoop).sType = "CP") And (tgCurrSPE(ilLoop).sSubType = "T") Then
                slPathCPTest = Trim$(tgCurrSPE(ilLoop).sPath)
                If slPathCPTest <> "" Then
                    If right(slPathCPTest, 1) <> "\" Then
                        slPathCPTest = slPathCPTest & "\"
                    End If
                End If
                'Exit For
            End If
        ElseIf igTestSystem Then
            If (tgCurrSPE(ilLoop).sType = "CP") And (tgCurrSPE(ilLoop).sSubType = "T") Then
                slPathCP = Trim$(tgCurrSPE(ilLoop).sPath)
                If slPathCP <> "" Then
                    If right(slPathCP, 1) <> "\" Then
                        slPathCP = slPathCP & "\"
                    End If
                End If
                'Exit For
            End If
        End If
    Next ilLoop
'    For ilLoop = 0 To UBound(tgCurrSPE) - 1 Step 1
'        If tgCurrSPE(ilLoop).sType = "CB" Then
'            slPathCB = Trim$(tgCurrSPE(ilLoop).sPath)
'            If slPathCB <> "" Then
'                If Right(slPathCB, 1) <> "\" Then
'                    slPathCB = slPathCB & "\"
'                End If
'            End If
'            Exit For
'        End If
'    Next ilLoop
    If slName = "" Then
        MsgBox "Merge File Name missing for Client from Site Option", vbOKOnly
        mOpenMergeFile = False
        Exit Function
    End If
    If slPathCP = "" Then
        MsgBox "Merge Path missing for Client from Site Option", vbOKOnly
        mOpenMergeFile = False
        Exit Function
    End If
    
    ilRet = 0
    slToFile = slPathCP & slName
    slDateTime = FileDateTime(slToFile)
    If ilRet <> 0 Then
        If slPathCB <> "" Then
            ilRet = 0
            slToFile = slPathCB & slName
            slDateTime = FileDateTime(slToFile)
            If ilRet <> 0 Then
                MsgBox "Merge File missing from " & slPathCP & slName & " and from " & slToFile, vbOKOnly
                mOpenMergeFile = False
                Exit Function
            End If
        Else
            MsgBox "Merge File missing from " & slPathCP & slName, vbOKOnly
            mOpenMergeFile = False
            Exit Function
        End If
    End If
    ilRet = 0
    On Error GoTo mOpenMergeFileErr:
    hmMerge = FreeFile
    Open slToFile For Input Access Read As hmMerge
    If ilRet <> 0 Then
        Close hmMerge
        hmMerge = -1
        MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        mOpenMergeFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMergeFileCP = slPathCP & slName
    If slPathCB <> "" Then
        slMergeFileCB = slPathCB & slName
    Else
        slMergeFileCB = ""
    End If
    If (Not igTestSystem) And (tgSOE.sMergeStopFlagTst = "N") And (slPathCPTest <> "") Then
        On Error GoTo mOpenMergeFileErr:
        ilRet = 0
        slToFile = slPathCPTest & slName
        slDateTime = FileDateTime(slToFile)
        If ilRet = 0 Then
            Kill slPathCPTest & slName
        End If
        FileCopy slPathCP & slName, slPathCPTest & slName
    End If
    mOpenMergeFile = True
    Exit Function
mOpenMergeFileErr:
    ilRet = 1
    Resume Next
End Function
Private Function mOpenMergeMsgFile(slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String

    On Error GoTo mOpenMergeMsgFileErr:
    ilRet = 0
    
    slAirYear = Year(smAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(smAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(smAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If

    
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgMsgDirectory & "MergeSpots_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenMergeMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        mOpenMergeMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
'    Print #hmMsg, "** Test : " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    Print #hmMsg, ""
    slMsgFileName = slToFile
    mOpenMergeMsgFile = True
    Exit Function
mOpenMergeMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Function mOpenTestMsgFile(slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String

    On Error GoTo mOpenTestMsgFileErr:
    ilRet = 0
    slAirYear = Year(smAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(smAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(smAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If

    
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgExportDirectory & "TestAuto_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenTestMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        mOpenTestMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
'    Print #hmMsg, "** Test : " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    Print #hmMsg, ""
    slMsgFileName = slToFile
    mOpenTestMsgFile = True
    Exit Function
mOpenTestMsgFileErr:
    ilRet = 1
    Resume Next
End Function



Private Sub mMakeExportStr(ilStartCol As Integer, ilNoChar As Integer, llRow As Long, llCol As Long, ilUCase As Integer)
    Dim slStr As String
    If (ilStartCol > 0) And (mExportCol(llRow, llCol)) Then
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, llCol))
        Do While Len(slStr) < ilNoChar
            slStr = slStr & " "
        Loop
        If ilUCase Then
            slStr = UCase$(slStr)
        End If
        Mid(smExportStr, ilStartCol, ilNoChar) = slStr
    End If
End Sub

Private Sub mInitFilterInfo()
    Dim ilUpper As Integer
    ReDim tgFilterFields(0 To 0) As FIELDSELECTION
    
    ilUpper = 0
    If (UBound(tgUsedATE) > 0) And ((tgSchUsedSumEPE.sAudioName <> "N") Or (tgSchUsedSumEPE.sProtAudioName <> "N") Or (tgSchUsedSumEPE.sBkupAudioName <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Audio Types"
        tgFilterFields(ilUpper).iFieldType = 5
        If Len(tgATE.sName) >= 6 Then
            tgFilterFields(ilUpper).iMaxNoChar = Len(tgATE.sName)
        Else
            tgFilterFields(ilUpper).iMaxNoChar = 6
        End If
        tgFilterFields(ilUpper).sListFile = "ATE"
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedANE) > 0) And ((tgSchUsedSumEPE.sAudioName <> "N") Or (tgSchUsedSumEPE.sProtAudioName <> "N") Or (tgSchUsedSumEPE.sBkupAudioName <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Audio Name"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioName", 6)
        tgFilterFields(ilUpper).sListFile = "ANE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sAudioName
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedBDE) > 0) And (tgSchUsedSumEPE.sBus <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Bus"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("BusName", 6)
        tgFilterFields(ilUpper).sListFile = "BDE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sBus
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If UBound(tgUsedETE) > 0 Then
        tgFilterFields(ilUpper).sFieldName = "Event Types"
        tgFilterFields(ilUpper).iFieldType = 5
        If Len(tgETE.sName) >= 6 Then
            tgFilterFields(ilUpper).iMaxNoChar = Len(tgETE.sName)
        Else
            tgFilterFields(ilUpper).iMaxNoChar = 6
        End If
        tgFilterFields(ilUpper).sListFile = "ETE"
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedFNE) > 0) And (tgSchUsedSumEPE.sFollow <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Follow"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Follow", 6)
        tgFilterFields(ilUpper).sListFile = "FNE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sFollow
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedMTE) > 0) And (tgSchUsedSumEPE.sMaterialType <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Material"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Material", 6)
        tgFilterFields(ilUpper).sListFile = "MTE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sMaterialType
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedNNE) > 0) And ((tgSchUsedSumEPE.sStartNetcue <> "N") Or (tgSchUsedSumEPE.sStopNetcue <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Netcue"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Netcue1", 6)
        tgFilterFields(ilUpper).sListFile = "NNE"
        If (tgSchManSumEPE.sStartNetcue = "Y") Or (tgSchManSumEPE.sStopNetcue = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedRNE) > 0) And ((tgSchUsedSumEPE.sRelay1 <> "N") Or (tgSchUsedSumEPE.sRelay2 <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Relay"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Relay1", 6)
        tgFilterFields(ilUpper).sListFile = "RNE"
        If (tgSchManSumEPE.sRelay1 = "Y") Or (tgSchManSumEPE.sRelay2 = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedStartTTE) > 0) And (tgSchUsedSumEPE.sStartType <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Start Type"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("StartType", 6)
        tgFilterFields(ilUpper).sListFile = "TTES"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sStartType
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedEndTTE) > 0) And (tgSchUsedSumEPE.sEndType <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "End Type"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("EndType", 6)
        tgFilterFields(ilUpper).sListFile = "TTEE"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sEndType
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedAudioCCE) > 0) And ((tgSchUsedSumEPE.sAudioControl <> "N") Or (tgSchUsedSumEPE.sProtAudioControl <> "N") Or (tgSchUsedSumEPE.sBkupAudioControl <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Audio Control"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioCtrl", 6)
        tgFilterFields(ilUpper).sListFile = "CCEA"
        If (tgSchManSumEPE.sAudioControl = "Y") Or (tgSchManSumEPE.sProtAudioControl = "Y") Or (tgSchManSumEPE.sBkupAudioControl = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedBusCCE) > 0) And (tgSchUsedSumEPE.sBusControl <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Bus Control"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("BusCtrl", 6)
        tgFilterFields(ilUpper).sListFile = "CCEB"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sBusControl
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    '7/8/11: Make T2 work like T1
    'If (UBound(tgUsedT2CTE) > 0) And (tgSchUsedSumEPE.sTitle2 <> "N") Then
    '    tgFilterFields(ilUpper).sFieldName = "Title 2"
    '    tgFilterFields(ilUpper).iFieldType = 5
    '    tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
    '    tgFilterFields(ilUpper).sListFile = "CTE2"
    '    tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle2
    '    ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
    '    ilUpper = ilUpper + 1
    'End If
    If (UBound(tgT2MatchList) > 0) And (tgSchUsedSumEPE.sTitle2 <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Title 2"
        tgFilterFields(ilUpper).iFieldType = 9
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
        tgFilterFields(ilUpper).sListFile = "CTE2"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle2
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    
    If (UBound(tgT1MatchList) > 0) And (tgSchUsedSumEPE.sTitle1 <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "Title 1"
        tgFilterFields(ilUpper).iFieldType = 9
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Title1", 6)
        tgFilterFields(ilUpper).sListFile = "CTE1"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle1
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedSCE) > 0) And ((tgSchUsedSumEPE.sSilence1 <> "N") Or (tgSchUsedSumEPE.sSilence2 <> "N") Or (tgSchUsedSumEPE.sSilence3 <> "N") Or (tgSchUsedSumEPE.sSilence4 <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Silence Control"
        tgFilterFields(ilUpper).iFieldType = 5
        tgFilterFields(ilUpper).iMaxNoChar = gSetMaxChars("Silence1", 6)
        tgFilterFields(ilUpper).sListFile = "SCE"
        If (tgSchManSumEPE.sSilence1 = "Y") Or (tgSchManSumEPE.sSilence2 = "Y") Or (tgSchManSumEPE.sSilence3 = "Y") Or (tgSchManSumEPE.sSilence4 = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sFixedTime <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Fixed Time"
        tgFilterFields(ilUpper).iFieldType = 9
        tgFilterFields(ilUpper).iMaxNoChar = 1
        tgFilterFields(ilUpper).sListFile = "FTYN"
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sFixedTime
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sTime <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Time"
        tgFilterFields(ilUpper).iFieldType = 6
        tgFilterFields(ilUpper).iMaxNoChar = 10
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sTime
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sDuration <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Duration"
        tgFilterFields(ilUpper).iFieldType = 8
        tgFilterFields(ilUpper).iMaxNoChar = 10
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sDuration
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (UBound(tgUsedANE) > 0) And ((tgSchUsedSumEPE.sAudioItemID <> "N") Or (tgSchUsedSumEPE.sProtAudioItemID <> "N")) Then
        tgFilterFields(ilUpper).sFieldName = "Item ID"
        tgFilterFields(ilUpper).iFieldType = 2
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioItemID = "Y") Or (tgSchManSumEPE.sProtAudioItemID = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sAudioISCI <> "N") Or (tgSchUsedSumEPE.sProtAudioISCI <> "N") Then
        tgFilterFields(ilUpper).sFieldName = "ISCI"
        tgFilterFields(ilUpper).iFieldType = 2
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioISCI = "Y") Or (tgSchManSumEPE.sProtAudioISCI = "Y") Then
            tgFilterFields(ilUpper).sMandatory = "Y"
        Else
            tgFilterFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sSilenceTime <> "N" Then
        tgFilterFields(ilUpper).sFieldName = "Silence Time"
        tgFilterFields(ilUpper).iFieldType = 8
        tgFilterFields(ilUpper).iMaxNoChar = 5
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = tgSchManSumEPE.sSilenceTime
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
    End If
    If tmSHE.lCode <> 0 Then
        tgFilterFields(ilUpper).sFieldName = "Event ID"
        tgFilterFields(ilUpper).iFieldType = 1
        tgFilterFields(ilUpper).iMaxNoChar = 0
        tgFilterFields(ilUpper).sListFile = ""
        tgFilterFields(ilUpper).sMandatory = "Y"
        ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (sgClientFields = "A") Then
        If (tgSchUsedSumEPE.sABCFormat <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC Format"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCFormat")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCFormat = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCPgmCode <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC Pgm Code"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCPgmCode")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCPgmCode = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCXDSMode <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC XDS Mode"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCXDSMODE")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCXDSMode = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCRecordItem <> "N") Then
            tgFilterFields(ilUpper).sFieldName = "ABC Recd Item"
            tgFilterFields(ilUpper).iFieldType = 2
            tgFilterFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCRecordItem")
            tgFilterFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCRecordItem = "Y") Then
                tgFilterFields(ilUpper).sMandatory = "Y"
            Else
                tgFilterFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgFilterFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
    End If
    
End Sub

Private Sub mInitReplaceInfo()
    Dim ilUpper As Integer
    ReDim tgReplaceFields(0 To 0) As FIELDSELECTION
    
    ilUpper = 0
    If ((tgSchUsedSumEPE.sAudioName <> "N") Or (tgSchUsedSumEPE.sProtAudioName <> "N") Or (tgSchUsedSumEPE.sBkupAudioName <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Name"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioName", 6)
        tgReplaceFields(ilUpper).sListFile = "ANE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sAudioName
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sBus <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusName", 6)
        tgReplaceFields(ilUpper).sListFile = "BDE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sBus
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sFollow <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Follow"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Follow", 6)
        tgReplaceFields(ilUpper).sListFile = "FNE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sFollow
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sMaterialType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Material"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Material", 6)
        tgReplaceFields(ilUpper).sListFile = "MTE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sMaterialType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sStartNetcue <> "N") Or (tgSchUsedSumEPE.sStopNetcue <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Netcue"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Netcue1", 6)
        tgReplaceFields(ilUpper).sListFile = "NNE"
        If (tgSchManSumEPE.sStartNetcue = "Y") Or (tgSchManSumEPE.sStopNetcue = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sRelay1 <> "N") Or (tgSchUsedSumEPE.sRelay2 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Relay"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Relay1", 6)
        tgReplaceFields(ilUpper).sListFile = "RNE"
        If (tgSchManSumEPE.sRelay1 = "Y") Or (tgSchManSumEPE.sRelay2 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sStartType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Start Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("StartType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTES"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sStartType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sEndType <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "End Type"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("EndType", 6)
        tgReplaceFields(ilUpper).sListFile = "TTEE"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sEndType
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sAudioControl <> "N") Or (tgSchUsedSumEPE.sProtAudioControl <> "N") Or (tgSchUsedSumEPE.sBkupAudioControl <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Audio Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("AudioCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEA"
        If (tgSchManSumEPE.sAudioControl = "Y") Or (tgSchManSumEPE.sProtAudioControl = "Y") Or (tgSchManSumEPE.sBkupAudioControl = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sBusControl <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Bus Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("BusCtrl", 6)
        tgReplaceFields(ilUpper).sListFile = "CCEB"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sBusControl
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sTitle2 <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Title 2"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title2", 6)
        tgReplaceFields(ilUpper).sListFile = "CTE2"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle2
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (tgSchUsedSumEPE.sTitle1 <> "N") Then
        tgReplaceFields(ilUpper).sFieldName = "Title 1"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Title1", 6)
        tgReplaceFields(ilUpper).sListFile = "CTE1"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sTitle1
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sSilence1 <> "N") Or (tgSchUsedSumEPE.sSilence2 <> "N") Or (tgSchUsedSumEPE.sSilence3 <> "N") Or (tgSchUsedSumEPE.sSilence4 <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Silence Control"
        tgReplaceFields(ilUpper).iFieldType = 5
        tgReplaceFields(ilUpper).iMaxNoChar = gSetMaxChars("Silence1", 6)
        tgReplaceFields(ilUpper).sListFile = "SCE"
        If (tgSchManSumEPE.sSilence1 = "Y") Or (tgSchManSumEPE.sSilence2 = "Y") Or (tgSchManSumEPE.sSilence3 = "Y") Or (tgSchManSumEPE.sSilence4 = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If tgSchUsedSumEPE.sFixedTime <> "N" Then
        tgReplaceFields(ilUpper).sFieldName = "Fixed Time"
        tgReplaceFields(ilUpper).iFieldType = 9
        tgReplaceFields(ilUpper).iMaxNoChar = 1
        tgReplaceFields(ilUpper).sListFile = "FTYN"
        tgReplaceFields(ilUpper).sMandatory = tgSchManSumEPE.sFixedTime
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sAudioItemID <> "N") Or (tgSchUsedSumEPE.sProtAudioItemID <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "Item ID"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = 0
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioItemID = "Y") Or (tgSchManSumEPE.sProtAudioItemID = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If ((tgSchUsedSumEPE.sAudioISCI <> "N") Or (tgSchUsedSumEPE.sProtAudioISCI <> "N")) Then
        tgReplaceFields(ilUpper).sFieldName = "ISCI"
        tgReplaceFields(ilUpper).iFieldType = 2
        tgReplaceFields(ilUpper).iMaxNoChar = 0
        tgReplaceFields(ilUpper).sListFile = ""
        If (tgSchManSumEPE.sAudioISCI = "Y") Or (tgSchManSumEPE.sProtAudioISCI = "Y") Then
            tgReplaceFields(ilUpper).sMandatory = "Y"
        Else
            tgReplaceFields(ilUpper).sMandatory = "N"
        End If
        ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
        ilUpper = ilUpper + 1
    End If
    If (sgClientFields = "A") Then
        If (tgSchUsedSumEPE.sABCFormat <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Format"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCFormat")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCFormat = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCPgmCode <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Pgm Code"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCPgmCode")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCPgmCode = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCXDSMode <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC XDS Mode"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCXDSMODE")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCXDSMode = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
        If (tgSchUsedSumEPE.sABCRecordItem <> "N") Then
            tgReplaceFields(ilUpper).sFieldName = "ABC Recd Item"
            tgReplaceFields(ilUpper).iFieldType = 2
            tgReplaceFields(ilUpper).iMaxNoChar = gGetMaxChars("ABCRecordItem")
            tgReplaceFields(ilUpper).sListFile = ""
            If (tgSchManSumEPE.sABCRecordItem = "Y") Then
                tgReplaceFields(ilUpper).sMandatory = "Y"
            Else
                tgReplaceFields(ilUpper).sMandatory = "N"
            End If
            ReDim Preserve tgReplaceFields(0 To ilUpper + 1) As FIELDSELECTION
            ilUpper = ilUpper + 1
        End If
    End If
    
End Sub

Private Function mCheckFilter(tlCurrSEE As SEE, slComment As String) As Integer
    Dim ilFilter As Integer
    Dim ilField As Integer
    Dim ilFilterType As Integer
    Dim slFileName As String
    Dim ilOrTest As Integer
    Dim ilMatch As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    
    mCheckFilter = True
    For ilFilter = LBound(tmFilterValues) To UBound(tmFilterValues) - 1 Step 1
        tmFilterValues(ilFilter).iUsed = False
    Next ilFilter
    For ilFilter = LBound(tmFilterValues) To UBound(tmFilterValues) - 1 Step 1
        If tmFilterValues(ilFilter).iUsed = False Then
            For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
                If tgFilterFields(ilField).sFieldName = tmFilterValues(ilFilter).sFieldName Then
                    ilFilterType = tgFilterFields(ilField).iFieldType
                    slFileName = tgFilterFields(ilField).sListFile
                    If ilFilterType = 5 Then
                        Select Case UCase$(Trim$(slFileName))
                            Case "ATE"
                                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                '    If tlCurrSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                    ilASE = gBinarySearchASE(tlCurrSEE.iAudioAseCode, tgCurrASE())
                                    If ilASE <> -1 Then
                                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                        '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                                            ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                                            If ilANE <> -1 Then
                                                ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrANE(ilANE).iAteCode))
                                        '        Exit For
                                            End If
                                        'Next ilANE
                                '        Exit For
                                    End If
                                'Next ilASE
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                    '    If tlCurrSEE.iProtAneCode = tgCurrANE(ilANE).iCode Then
                                        ilANE = gBinarySearchANE(tlCurrSEE.iProtAneCode, tgCurrANE())
                                        If ilANE <> -1 Then
                                            ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrANE(ilANE).iAteCode))
                                    '        Exit For
                                        End If
                                    'Next ilANE
                                End If
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                    '    If tlCurrSEE.iBkupAneCode = tgCurrANE(ilANE).iCode Then
                                        ilANE = gBinarySearchANE(tlCurrSEE.iBkupAneCode, tgCurrANE())
                                        If ilANE <> -1 Then
                                            ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrANE(ilANE).iAteCode))
                                    '        Exit For
                                        End If
                                    'Next ilANE
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                                
                            Case "ANE"
                                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                '    If tlCurrSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                    ilASE = gBinarySearchASE(tlCurrSEE.iAudioAseCode, tgCurrASE())
                                    If ilASE <> -1 Then
                                        ilMatch = mMatchTestFile(ilFilter, CLng(tgCurrASE(ilASE).iPriAneCode))
                                '        Exit For
                                    End If
                                'Next ilASE
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iProtAneCode))
                                End If
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBkupAneCode))
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "BDE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBdeCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "ETE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iEteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "FNE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iFneCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "MTE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iMteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "NNE"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iStartNneCode))
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iEndNneCode))
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "RNE"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i1RneCode))
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i2RneCode))
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "TTES"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iStartTteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "TTEE"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iEndTteCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "CCEA"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iAudioCceCode))
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBkupCceCode))
                                '    If Not ilMatch Then
                                    If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                        ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.iProtCceCode))
                                    End If
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "CCEB"
                                If Not mMatchTestFile(ilFilter, CLng(tlCurrSEE.iBusCceCode)) Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "SCE"
                                ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i1SceCode))
                                'If Not ilMatch Then
                                If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                    ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i2SceCode))
                                '    If Not ilMatch Then
                                    If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                        ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i3SceCode))
                                '        If Not ilMatch Then
                                        If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                            ilMatch = mMatchTestFile(ilFilter, CLng(tlCurrSEE.i4SceCode))
                                        End If
                                    End If
                                End If
                                If Not ilMatch Then
                                    mCheckFilter = False
                                    Exit Function
                                End If
                            Case "FTYN"
                                If Trim$(tlCurrSEE.sFixedTime) = "Y" Then
                                    If Not mMatchTestFile(ilFilter, 0) Then
                                        mCheckFilter = False
                                        Exit Function
                                    End If
                                ElseIf Trim$(tlCurrSEE.sFixedTime) = "N" Then
                                    If Not mMatchTestFile(ilFilter, 1) Then
                                        mCheckFilter = False
                                        Exit Function
                                    End If
                                End If
                            '7/8/11: Make T2 work like T1
                            'Case "CTE2"
                            '    If Not mMatchTestFile(ilFilter, tlCurrSEE.l2CteCode) Then
                            '        mCheckFilter = False
                            '        Exit Function
                            '    End If
                        End Select
                    ElseIf ilFilterType = 9 Then
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Fixed Time" Then
                            If Not mMatchTestList(ilFilter, tlCurrSEE.sFixedTime) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Title 1" Then
                            If Not mMatchTestList(ilFilter, slComment) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Title 2" Then
                            If Not mMatchTestList(ilFilter, slComment) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 1 Then    'Event ID
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Event ID" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lEventID) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 2 Then
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Item ID" Then
                            ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sAudioItemID)
                            'If Not ilMatch Then
                            If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sProtItemID)
                            End If
                            If Not ilMatch Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "ISCI" Then
                            ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sAudioISCI)
                            'If Not ilMatch Then
                            If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
                                ilMatch = mMatchTestString(ilFilter, tlCurrSEE.sProtISCI)
                            End If
                            If Not ilMatch Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 6 Then    'Time
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Time" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lTime) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    ElseIf ilFilterType = 8 Then    'Length
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Silence" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lSilenceTime) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                        If Trim$(tmFilterValues(ilFilter).sFieldName) = "Duration" Then
                            If Not mMatchTestValue(ilFilter, tlCurrSEE.lDuration) Then
                                mCheckFilter = False
                                Exit Function
                            End If
                        End If
                    End If
                    tmFilterValues(ilFilter).iUsed = True
                    Exit For
                End If
            Next ilField
        End If
    Next ilFilter
    
    
End Function

Private Function mMatchTestFile(ilFilter As Integer, llFileCode As Long) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    
    If tmFilterValues(ilFilter).iOperator = 1 Then
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If llFileCode = tmFilterValues(ilOrTest).lCode Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If llFileCode = tmFilterValues(ilAndTest).lCode Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    Else
        ilMatch = False
    End If
    For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
            tmFilterValues(ilOrTest).iUsed = True
        End If
    Next ilOrTest
    mMatchTestFile = ilMatch
End Function
Private Function mMatchTestList(ilFilter As Integer, slValue As String) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    
    If tmFilterValues(ilFilter).iOperator = 1 Then
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If StrComp(Trim$(slValue), Trim$(tmFilterValues(ilOrTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If StrComp(Trim$(slValue), Trim$(tmFilterValues(ilOrTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    Else
        ilMatch = False
    End If
    For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
            tmFilterValues(ilOrTest).iUsed = True
        End If
    Next ilOrTest
    mMatchTestList = ilMatch
End Function

Private Function mMatchTestString(ilFilter As Integer, slString As String) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    
    If tmFilterValues(ilFilter).iOperator = 1 Then
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If StrComp(Trim$(slString), Trim$(tmFilterValues(ilOrTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If StrComp(Trim$(slString), Trim$(tmFilterValues(ilAndTest).sValue), vbTextCompare) = 0 Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    Else
        ilMatch = False
    End If
    For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
            tmFilterValues(ilOrTest).iUsed = True
        End If
    Next ilOrTest
    mMatchTestString = ilMatch
End Function

Private Function mMatchTestValue(ilFilter As Integer, llValue As Long) As Integer
    Dim ilMatch As Integer
    Dim ilOrTest As Integer
    Dim ilAndTest As Integer
    Dim ilBetween As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilNameMatch As Integer
    
    ilMatch = False
    If tmFilterValues(ilFilter).iOperator = 1 Then   'Equal Match
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilOrTest).iOperator) Then
                If llValue = tmFilterValues(ilOrTest).lCode Then
                    ilMatch = True
                    Exit For
                End If
            End If
        Next ilOrTest
    ElseIf tmFilterValues(ilFilter).iOperator = 2 Then   'Not Equal Match
        ilMatch = True
        For ilAndTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilAndTest).sFieldName) And (tmFilterValues(ilFilter).iOperator = tmFilterValues(ilAndTest).iOperator) Then
                If llValue = tmFilterValues(ilAndTest).lCode Then
                    ilMatch = False
                    Exit For
                End If
            End If
        Next ilAndTest
    End If
    If tmFilterValues(ilFilter).iOperator <> 2 Then
        'Look for Greater Than
        If (Not ilMatch) Then
            For ilLoop = ilFilter To UBound(tmFilterValues) - 1 Step 1
                If (tmFilterValues(ilLoop).iOperator = 3) And (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilLoop).sFieldName) Then   'Greater Than
                    ilIndex = -1
                    For ilBetween = ilLoop + 1 To UBound(tmFilterValues) - 1 Step 1
                        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilBetween).sFieldName) And ((tmFilterValues(ilBetween).iOperator = 4) Or (tmFilterValues(ilBetween).iOperator = 6)) Then
                            If ilIndex = -1 Then
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode > 0 Then
                                    ilIndex = ilBetween
                                End If
                            Else
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilFilter).lCode > 0 Then
                                    If (tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode) < (tmFilterValues(ilIndex).lCode - tmFilterValues(ilLoop).lCode) Then
                                        ilIndex = ilBetween
                                    End If
                                End If
                            End If
                        End If
                    Next ilBetween
                    If ilIndex = -1 Then
'                        If llValue > tmFilterValues(ilLoop).lCode Then
'                            ilMatch = True
'                            Exit For
'                        End If
                    Else
                        tmFilterValues(ilLoop).iUsed = True
                        tmFilterValues(ilIndex).iUsed = True
                        If llValue > tmFilterValues(ilLoop).lCode Then
                            If tmFilterValues(ilIndex).iOperator = 4 Then
                                If llValue < tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                            If tmFilterValues(ilIndex).iOperator = 6 Then
                                If llValue <= tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilLoop
        End If
        'Look for Less than
        If (Not ilMatch) Then
            For ilLoop = ilFilter To UBound(tmFilterValues) - 1 Step 1
                If (tmFilterValues(ilLoop).iOperator = 5) And (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilLoop).sFieldName) Then   'Greater Than
                    ilIndex = -1
                    For ilBetween = ilLoop + 1 To UBound(tmFilterValues) - 1 Step 1
                        If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilBetween).sFieldName) And ((tmFilterValues(ilBetween).iOperator = 4) Or (tmFilterValues(ilBetween).iOperator = 6)) Then
                            If ilIndex = -1 Then
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode > 0 Then
                                    ilIndex = ilBetween
                                End If
                            Else
                                If tmFilterValues(ilBetween).lCode - tmFilterValues(ilFilter).lCode > 0 Then
                                    If (tmFilterValues(ilBetween).lCode - tmFilterValues(ilLoop).lCode) < (tmFilterValues(ilIndex).lCode - tmFilterValues(ilLoop).lCode) Then
                                        ilIndex = ilBetween
                                    End If
                                End If
                            End If
                        End If
                    Next ilBetween
                    If ilIndex = -1 Then
'                        If llValue >= tmFilterValues(ilLoop).lCode Then
'                            ilMatch = True
'                            Exit For
'                        End If
                    Else
                        tmFilterValues(ilLoop).iUsed = True
                        tmFilterValues(ilIndex).iUsed = True
                        If llValue >= tmFilterValues(ilLoop).lCode Then
                            If tmFilterValues(ilIndex).iOperator = 4 Then
                                If llValue < tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                            If tmFilterValues(ilIndex).iOperator = 6 Then
                                If llValue <= tmFilterValues(ilIndex).lCode Then
                                    ilMatch = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilLoop
        End If
    End If
    If ((Not ilMatch) And (tmFilterValues(ilFilter).iOperator <> 2)) Or ((ilMatch) And (tmFilterValues(ilFilter).iOperator = 2)) Then
        ilNameMatch = False
        ilMatch = False
        For ilOrTest = ilFilter To UBound(tmFilterValues) - 1 Step 1
            If (tmFilterValues(ilOrTest).iUsed = False) And ((tmFilterValues(ilOrTest).iOperator <> 1) And (tmFilterValues(ilOrTest).iOperator <> 2)) Then
                If (tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilOrTest).sFieldName) Then
                    If ilOrTest <> ilFilter Then
                        ilNameMatch = True
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 3 Then   'Greater Than
                        If llValue > tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 4 Then   'Less Than
                        If llValue < tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 5 Then   'Greater Than or Equal
                        If llValue >= tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                    If tmFilterValues(ilOrTest).iOperator = 6 Then   'Less Than or Equal
                        If llValue <= tmFilterValues(ilOrTest).lCode Then
                            ilMatch = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next ilOrTest
        If Not ilNameMatch Then
            If tmFilterValues(ilFilter).iOperator = 2 Then
                ilMatch = True
            End If
        End If
    End If
    For ilLoop = ilFilter To UBound(tmFilterValues) - 1 Step 1
        If tmFilterValues(ilFilter).sFieldName = tmFilterValues(ilLoop).sFieldName Then
            tmFilterValues(ilLoop).iUsed = True
        End If
    Next ilLoop
    mMatchTestValue = ilMatch
End Function

Private Sub mCreateUsedArrays()
    Dim llLoop As Long
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim ilETE As Integer
    Dim ilFNE As Integer
    Dim ilMTE As Integer
    Dim ilNNE As Integer
    Dim ilRNE As Integer
    Dim ilTTE As Integer
    Dim ilCCE As Integer
    Dim ilSCE As Integer
    Dim ilCTE As Integer
    
    ReDim tgYNMatchList(0 To 2) As MATCHLIST
    tgYNMatchList(0).sValue = "Y"
    tgYNMatchList(0).lValue = 0
    tgYNMatchList(1).sValue = "N"
    tgYNMatchList(1).lValue = 1
    If UBound(tmCurrSEE) <= 0 Then
        ReDim tgUsedBDE(0 To UBound(tgCurrBDE)) As BDE
        For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            LSet tgUsedBDE(ilBDE) = tgCurrBDE(ilBDE)
        Next ilBDE
        ReDim tgUsedANE(0 To UBound(tgCurrANE)) As ANE
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            LSet tgUsedANE(ilANE) = tgCurrANE(ilANE)
        Next ilANE
        ReDim tgUsedATE(0 To UBound(tgCurrATE)) As ATE
        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
            LSet tgUsedATE(ilATE) = tgCurrATE(ilATE)
        Next ilATE
        ReDim tgUsedETE(0 To UBound(tgCurrETE)) As ETE
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            LSet tgUsedETE(ilETE) = tgCurrETE(ilETE)
        Next ilETE
        ReDim tgUsedFNE(0 To UBound(tgCurrFNE)) As FNE
        For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
            LSet tgUsedFNE(ilFNE) = tgCurrFNE(ilFNE)
        Next ilFNE
        ReDim tgUsedMTE(0 To UBound(tgCurrMTE)) As MTE
        For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
            LSet tgUsedMTE(ilMTE) = tgCurrMTE(ilMTE)
        Next ilMTE
        ReDim tgUsedNNE(0 To UBound(tgCurrNNE)) As NNE
        For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
            LSet tgUsedNNE(ilNNE) = tgCurrNNE(ilNNE)
        Next ilNNE
        ReDim tgUsedRNE(0 To UBound(tgCurrRNE)) As RNE
        For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
            LSet tgUsedRNE(ilRNE) = tgCurrRNE(ilRNE)
        Next ilRNE
        ReDim tgUsedStartTTE(0 To UBound(tgCurrStartTTE)) As TTE
        For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            LSet tgUsedStartTTE(ilTTE) = tgCurrStartTTE(ilTTE)
        Next ilTTE
        ReDim tgUsedEndTTE(0 To UBound(tgCurrEndTTE)) As TTE
        For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            LSet tgUsedEndTTE(ilTTE) = tgCurrEndTTE(ilTTE)
        Next ilTTE
        ReDim tgUsedAudioCCE(0 To UBound(tgCurrAudioCCE)) As CCE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            LSet tgUsedAudioCCE(ilCCE) = tgCurrAudioCCE(ilCCE)
        Next ilCCE
        ReDim tgUsedBusCCE(0 To UBound(tgCurrBusCCE)) As CCE
        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            LSet tgUsedBusCCE(ilCCE) = tgCurrBusCCE(ilCCE)
        Next ilCCE
        ReDim tgUsedSCE(0 To UBound(tgCurrSCE)) As SCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            LSet tgUsedSCE(ilSCE) = tgCurrSCE(ilSCE)
        Next ilSCE
        '7/8/11: Make T2 work like T1
        'ReDim tgUsedT2CTE(0 To UBound(tgCurrCTE)) As CTE
        'For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
        '    LSet tgUsedT2CTE(ilCTE) = tgCurrCTE(ilCTE)
        'Next ilCTE
        ReDim tgT1MatchList(0 To 0) As MATCHLIST
        ReDim tgT2MatchList(0 To 0) As MATCHLIST
        Exit Sub
    End If
    ReDim tgUsedBDE(0 To 0) As BDE
    ReDim tgUsedANE(0 To 0) As ANE
    ReDim tgUsedATE(0 To 0) As ATE
    ReDim tgUsedETE(0 To 0) As ETE
    ReDim tgUsedFNE(0 To 0) As FNE
    ReDim tgUsedMTE(0 To 0) As MTE
    ReDim tgUsedNNE(0 To 0) As NNE
    ReDim tgUsedRNE(0 To 0) As RNE
    ReDim tgUsedStartTTE(0 To 0) As TTE
    ReDim tgUsedEndTTE(0 To 0) As TTE
    ReDim tgUsedAudioCCE(0 To 0) As CCE
    ReDim tgUsedBusCCE(0 To 0) As CCE
    ReDim tgUsedSCE(0 To 0) As SCE
    '7/8/11: Make T2 work like T1
    'ReDim tgUsedT2CTE(0 To 0) As CTE
    ReDim tgT1MatchList(0 To 0) As MATCHLIST
    '7/8/11: Make T2 work like T1
    ReDim tgT2MatchList(0 To 0) As MATCHLIST
    For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iBdeCode = tgCurrBDE(ilBDE).iCode Then
            ilBDE = gBinarySearchBDE(tmCurrSEE(llLoop).iBdeCode, tgCurrBDE())
            If ilBDE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedBDE) - 1 Step 1
                    If tgUsedBDE(ilTest).iCode = tgCurrBDE(ilBDE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedBDE(UBound(tgUsedBDE)) = tgCurrBDE(ilBDE)
                    ReDim Preserve tgUsedBDE(0 To UBound(tgUsedBDE) + 1) As BDE
                End If
        '        Exit For
            End If
        'Next ilBDE
        'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASE).iCode Then
            ilASE = gBinarySearchASE(tmCurrSEE(llLoop).iAudioAseCode, tgCurrASE())
            If ilASE <> -1 Then
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        ilFound = False
                        For ilTest = 0 To UBound(tgUsedANE) - 1 Step 1
                            If tgUsedANE(ilTest).iCode = tgCurrANE(ilANE).iCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilTest
                        If Not ilFound Then
                            LSet tgUsedANE(UBound(tgUsedANE)) = tgCurrANE(ilANE)
                            ReDim Preserve tgUsedANE(0 To UBound(tgUsedANE) + 1) As ANE
                        End If
                    End If
                'Next ilANE
        '        Exit For
            End If
        'Next ilASE
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iProtAneCode, tgCurrANE())
            If ilANE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedANE) - 1 Step 1
                    If tgUsedANE(ilTest).iCode = tgCurrANE(ilANE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedANE(UBound(tgUsedANE)) = tgCurrANE(ilANE)
                    ReDim Preserve tgUsedANE(0 To UBound(tgUsedANE) + 1) As ANE
                End If
            End If
        'Next ilANE
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iBkupAneCode = tgCurrANE(ilANE).iCode Then
            ilANE = gBinarySearchANE(tmCurrSEE(llLoop).iBkupAneCode, tgCurrANE())
            If ilANE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedANE) - 1 Step 1
                    If tgUsedANE(ilTest).iCode = tgCurrANE(ilANE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedANE(UBound(tgUsedANE)) = tgCurrANE(ilANE)
                    ReDim Preserve tgUsedANE(0 To UBound(tgUsedANE) + 1) As ANE
                End If
            End If
        'Next ilANE
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedETE) - 1 Step 1
                    If tgUsedETE(ilTest).iCode = tgCurrETE(ilETE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedETE(UBound(tgUsedETE)) = tgCurrETE(ilETE)
                    ReDim Preserve tgUsedETE(0 To UBound(tgUsedETE) + 1) As ETE
                End If
                Exit For
            End If
        Next ilETE
        For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
            If tmCurrSEE(llLoop).iFneCode = tgCurrFNE(ilFNE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedFNE) - 1 Step 1
                    If tgUsedFNE(ilTest).iCode = tgCurrFNE(ilFNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedFNE(UBound(tgUsedFNE)) = tgCurrFNE(ilFNE)
                    ReDim Preserve tgUsedFNE(0 To UBound(tgUsedFNE) + 1) As FNE
                End If
                Exit For
            End If
        Next ilFNE
        For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
            If tmCurrSEE(llLoop).iMteCode = tgCurrMTE(ilMTE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedMTE) - 1 Step 1
                    If tgUsedMTE(ilTest).iCode = tgCurrMTE(ilMTE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedMTE(UBound(tgUsedMTE)) = tgCurrMTE(ilMTE)
                    ReDim Preserve tgUsedMTE(0 To UBound(tgUsedMTE) + 1) As MTE
                End If
                Exit For
            End If
        Next ilMTE
        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iStartNneCode, tgCurrNNE())
            If ilNNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedNNE) - 1 Step 1
                    If tgUsedNNE(ilTest).iCode = tgCurrNNE(ilNNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedNNE(UBound(tgUsedNNE)) = tgCurrNNE(ilNNE)
                    ReDim Preserve tgUsedNNE(0 To UBound(tgUsedNNE) + 1) As NNE
                End If
        '        Exit For
            End If
        'Next ilNNE
        'For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            ilNNE = gBinarySearchNNE(tmCurrSEE(llLoop).iEndNneCode, tgCurrNNE())
            If ilNNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedNNE) - 1 Step 1
                    If tgUsedNNE(ilTest).iCode = tgCurrNNE(ilNNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedNNE(UBound(tgUsedNNE)) = tgCurrNNE(ilNNE)
                    ReDim Preserve tgUsedNNE(0 To UBound(tgUsedNNE) + 1) As NNE
                End If
        '        Exit For
            End If
        'Next ilNNE
        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).i1RneCode = tgCurrNNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i1RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedRNE) - 1 Step 1
                    If tgUsedRNE(ilTest).iCode = tgCurrRNE(ilRNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedRNE(UBound(tgUsedRNE)) = tgCurrRNE(ilRNE)
                    ReDim Preserve tgUsedRNE(0 To UBound(tgUsedRNE) + 1) As RNE
                End If
        '        Exit For
            End If
        'Next ilRNE
        'For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        '    If tmCurrSEE(llLoop).i2RneCode = tgCurrNNE(ilRNE).iCode Then
            ilRNE = gBinarySearchRNE(tmCurrSEE(llLoop).i2RneCode, tgCurrRNE())
            If ilRNE <> -1 Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedRNE) - 1 Step 1
                    If tgUsedRNE(ilTest).iCode = tgCurrRNE(ilRNE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedRNE(UBound(tgUsedRNE)) = tgCurrRNE(ilRNE)
                    ReDim Preserve tgUsedRNE(0 To UBound(tgUsedRNE) + 1) As RNE
                End If
        '        Exit For
            End If
        'Next ilRNE
        For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
            If tmCurrSEE(llLoop).iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedStartTTE) - 1 Step 1
                    If tgUsedStartTTE(ilTest).iCode = tgCurrStartTTE(ilTTE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedStartTTE(UBound(tgUsedStartTTE)) = tgCurrStartTTE(ilTTE)
                    ReDim Preserve tgUsedStartTTE(0 To UBound(tgUsedStartTTE) + 1) As TTE
                End If
                Exit For
            End If
        Next ilTTE
        For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
            If tmCurrSEE(llLoop).iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedEndTTE) - 1 Step 1
                    If tgUsedEndTTE(ilTest).iCode = tgCurrEndTTE(ilTTE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedEndTTE(UBound(tgUsedEndTTE)) = tgCurrEndTTE(ilTTE)
                    ReDim Preserve tgUsedEndTTE(0 To UBound(tgUsedEndTTE) + 1) As TTE
                End If
                Exit For
            End If
        Next ilTTE
        For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iBusCceCode = tgCurrBusCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedBusCCE) - 1 Step 1
                    If tgUsedBusCCE(ilTest).iCode = tgCurrBusCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedBusCCE(UBound(tgUsedBusCCE)) = tgCurrBusCCE(ilCCE)
                    ReDim Preserve tgUsedBusCCE(0 To UBound(tgUsedBusCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                    If tgUsedAudioCCE(ilTest).iCode = tgCurrAudioCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedAudioCCE(UBound(tgUsedAudioCCE)) = tgCurrAudioCCE(ilCCE)
                    ReDim Preserve tgUsedAudioCCE(0 To UBound(tgUsedAudioCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
         For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                    If tgUsedAudioCCE(ilTest).iCode = tgCurrAudioCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedAudioCCE(UBound(tgUsedAudioCCE)) = tgCurrAudioCCE(ilCCE)
                    ReDim Preserve tgUsedAudioCCE(0 To UBound(tgUsedAudioCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tmCurrSEE(llLoop).iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                    If tgUsedAudioCCE(ilTest).iCode = tgCurrAudioCCE(ilCCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedAudioCCE(UBound(tgUsedAudioCCE)) = tgCurrAudioCCE(ilCCE)
                    ReDim Preserve tgUsedAudioCCE(0 To UBound(tgUsedAudioCCE) + 1) As CCE
                End If
                Exit For
            End If
        Next ilCCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i1SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i2SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i3SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
            If tmCurrSEE(llLoop).i4SceCode = tgCurrSCE(ilSCE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedSCE) - 1 Step 1
                    If tgUsedSCE(ilTest).iCode = tgCurrSCE(ilSCE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedSCE(UBound(tgUsedSCE)) = tgCurrSCE(ilSCE)
                    ReDim Preserve tgUsedSCE(0 To UBound(tgUsedSCE) + 1) As SCE
                End If
                Exit For
            End If
        Next ilSCE
        '7/8/11: Make T2 work like T1
        'For ilCTE = 0 To UBound(tgCurrCTE) - 1 Step 1
        '    If tmCurrSEE(llLoop).l2CteCode = tgCurrCTE(ilCTE).lCode Then
        '        ilFound = False
        '        For ilTest = 0 To UBound(tgUsedT2CTE) - 1 Step 1
        '            If tgUsedT2CTE(ilTest).lCode = tgCurrCTE(ilCTE).lCode Then
        '                ilFound = True
        '                Exit For
        '            End If
        '        Next ilTest
        '        If Not ilFound Then
        '            LSet tgUsedT2CTE(UBound(tgUsedT2CTE)) = tgCurrCTE(ilCTE)
        '            ReDim Preserve tgUsedT2CTE(0 To UBound(tgUsedT2CTE) + 1) As CTE
        '        End If
        '        Exit For
        '    End If
        'Next ilCTE
        ilFound = False
        For ilTest = 0 To UBound(tgT2MatchList) - 1 Step 1
            If StrComp(Trim$(tgT2MatchList(ilTest).sValue), Trim$(smT2Comment(llLoop)), vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilTest
        If Not ilFound Then
            tgT2MatchList(UBound(tgT2MatchList)).sValue = smT2Comment(llLoop)
            tgT2MatchList(UBound(tgT2MatchList)).lValue = llLoop
            ReDim Preserve tgT2MatchList(0 To UBound(tgT2MatchList) + 1) As MATCHLIST
        End If
        
        ilFound = False
        For ilTest = 0 To UBound(tgT1MatchList) - 1 Step 1
            If StrComp(Trim$(tgT1MatchList(ilTest).sValue), Trim$(smT1Comment(llLoop)), vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilTest
        If Not ilFound Then
            tgT1MatchList(UBound(tgT1MatchList)).sValue = smT1Comment(llLoop)
            tgT1MatchList(UBound(tgT1MatchList)).lValue = llLoop
            ReDim Preserve tgT1MatchList(0 To UBound(tgT1MatchList) + 1) As MATCHLIST
        End If
   Next llLoop
    For ilANE = 0 To UBound(tgUsedANE) - 1 Step 1
        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
            If tgUsedANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                ilFound = False
                For ilTest = 0 To UBound(tgUsedATE) - 1 Step 1
                    If tgUsedATE(ilTest).iCode = tgCurrATE(ilATE).iCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    LSet tgUsedATE(UBound(tgUsedATE)) = tgCurrATE(ilATE)
                    ReDim Preserve tgUsedATE(0 To UBound(tgUsedATE) + 1) As ATE
                End If
                Exit For
            End If
        Next ilATE
    Next ilANE
End Sub

Private Sub mReplaceValues()
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilReplace As Integer
    Dim ilField As Integer
    Dim ilFieldType As Integer
    Dim slFileName As String
    Dim ilColumn As Integer
    Dim ilSet As Integer
    Dim slNewValue As String
    Dim slOldValue As String
    Dim ilETE As Integer
    Dim ilCol(0 To 3) As Integer
    
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        If slStr <> "" Then
            grdLibEvents.Row = llRow
            For ilColumn = EVENTTYPEINDEX To imMaxCols Step 1
                grdLibEvents.Col = ilColumn
                If (grdLibEvents.CellForeColor <> vbRed) And (grdLibEvents.CellForeColor <> vbMagenta) Then
                    If Not mExportCol(grdLibEvents.Row, grdLibEvents.Col) Then
                        grdLibEvents.CellForeColor = vbBlue
                    Else
                        grdLibEvents.CellForeColor = vbBlack
                    End If
                End If
            Next ilColumn
        End If
    Next llRow
    For ilReplace = LBound(tgSchdReplaceValues) To UBound(tgSchdReplaceValues) - 1 Step 1
        For ilField = LBound(tgReplaceFields) To UBound(tgReplaceFields) - 1 Step 1
            If tgReplaceFields(ilField).sFieldName = tgSchdReplaceValues(ilReplace).sFieldName Then
                ilFieldType = tgReplaceFields(ilField).iFieldType
                slFileName = tgReplaceFields(ilField).sListFile
                For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                    grdLibEvents.Row = llRow
                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
                    If slStr <> "" Then
                        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                                Select Case tgCurrETE(ilETE).sCategory
                                    Case "P"
                                        If Not bgApplyToEventType(0) Then
                                            slStr = ""
                                        End If
                                    Case "A"
                                        If Not bgApplyToEventType(1) Then
                                            slStr = ""
                                        End If
                                    Case "S"
                                        If Not bgApplyToEventType(2) Then
                                            slStr = ""
                                        End If
                                    Case Else
                                        slStr = ""
                                End Select
                                Exit For
                            End If
                        Next ilETE
                    End If
                    If slStr <> "" Then
                        'Disallow changes in the past
                        If mColOk(llRow, BUSNAMEINDEX, True) Then
                            ilCol(0) = -1
                            ilCol(1) = -1
                            ilCol(2) = -1
                            ilCol(3) = -1
                            If ilFieldType = 5 Then
                                Select Case UCase$(Trim$(slFileName))
                                    Case "ANE"
                                        ilCol(0) = AUDIONAMEINDEX
                                        ilCol(1) = PROTNAMEINDEX
                                        ilCol(2) = BACKUPNAMEINDEX
                                    Case "BDE"
                                        ilCol(0) = BUSNAMEINDEX
                                    Case "FNE"
                                        ilCol(0) = FOLLOWINDEX
                                    Case "MTE"
                                        ilCol(0) = MATERIALINDEX
                                    Case "NNE"
                                        ilCol(0) = NETCUE1INDEX
                                        ilCol(1) = NETCUE2INDEX
                                    Case "RNE"
                                        ilCol(0) = RELAY1INDEX
                                        ilCol(1) = RELAY2INDEX
                                    Case "TTES"
                                        ilCol(0) = STARTTYPEINDEX
                                    Case "TTEE"
                                        ilCol(0) = ENDTYPEINDEX
                                    Case "CCEA"
                                        ilCol(0) = AUDIOCTRLINDEX
                                        ilCol(1) = PROTCTRLINDEX
                                        ilCol(2) = BACKUPCTRLINDEX
                                    Case "CCEB"
                                        ilCol(0) = BUSCTRLINDEX
                                    Case "SCE"
                                        ilCol(0) = SILENCE1INDEX
                                        ilCol(1) = SILENCE2INDEX
                                        ilCol(2) = SILENCE3INDEX
                                        ilCol(3) = SILENCE4INDEX
                                    '7/8/11: Make T2 work like T1
                                    'Case "CTE2"
                                    '    ilCol(0) = TITLE2INDEX
                                End Select
                            ElseIf ilFieldType = 9 Then
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "Fixed Time" Then
                                    ilCol(0) = FIXEDINDEX
                                End If
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "Title 1" Then
                                    ilCol(0) = TITLE1INDEX
                                End If
                                '7/8/11: Make T2 work like T1
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "Title 2" Then
                                    ilCol(0) = TITLE2INDEX
                                End If
                            ElseIf ilFieldType = 2 Then
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "Item ID" Then
                                    ilCol(0) = AUDIOITEMIDINDEX
                                    ilCol(1) = PROTITEMIDINDEX
                                End If
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "ISCI" Then
                                    ilCol(0) = AUDIOISCIINDEX
                                    ilCol(1) = PROTISCIINDEX
                                End If
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC Format" Then
                                    ilCol(0) = ABCFORMATINDEX
                                End If
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC Pgm Code" Then
                                    ilCol(0) = ABCPGMCODEINDEX
                                End If
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC XDS Mode" Then
                                    ilCol(0) = ABCXDSMODEINDEX
                                End If
                                If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC Recd Item" Then
                                    ilCol(0) = ABCRECORDITEMINDEX
                                End If
                            End If
                            For ilSet = 0 To 3 Step 1
                                If ilCol(ilSet) >= 0 Then
                                    slStr = Trim$(grdLibEvents.TextMatrix(llRow, ilCol(ilSet)))
                                    slOldValue = Trim$(tgSchdReplaceValues(ilReplace).sOldValue)
                                    slNewValue = Trim$(tgSchdReplaceValues(ilReplace).sNewValue)
                                    If (StrComp(slOldValue, slStr, vbTextCompare) = 0) Or ((slStr = "") And (StrComp(slOldValue, "[None]", vbTextCompare) = 0)) Then
                                        imFieldChgd = True
                                        grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
                                        grdLibEvents.TextMatrix(llRow, ilCol(ilSet)) = slNewValue
                                        grdLibEvents.Col = ilCol(ilSet)
                                        grdLibEvents.CellForeColor = DARKGREEN
                                    End If
                                End If
                            Next ilSet
                        End If
                    End If
                Next llRow
            End If
        Next ilField
    Next ilReplace
    mSetCommands
    
End Sub



Private Sub mBuildItemIDChk()
    Dim llRow As Long
    Dim slItemID As String
    Dim slTitle As String
    Dim slDuration As String
    Dim llDuration As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    
    ReDim tgItemIDChk(0 To 0) As ITEMIDCHK
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
            If StrComp("Spot", Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)), vbTextCompare) = 0 Then
                slItemID = Trim$(grdLibEvents.TextMatrix(llRow, AUDIOITEMIDINDEX))
                slTitle = Trim$(grdLibEvents.TextMatrix(llRow, TITLE1INDEX))
                slDuration = grdLibEvents.TextMatrix(llRow, DURATIONINDEX)
                llDuration = gStrLengthInTenthToLong(slDuration)
                If (slItemID <> "") And (slTitle <> "") And (StrComp(slTitle, "[None]", vbTextCompare) <> 0) Then
                    ilFound = False
                    For ilLoop = 0 To UBound(tgItemIDChk) - 1 Step 1
                        If StrComp(slItemID, Trim$(tgItemIDChk(ilLoop).sItemID), vbTextCompare) = 0 Then
                            ilFound = True
                            tgItemIDChk(ilLoop).sAudioStatus = "U"
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        ilUpper = UBound(tgItemIDChk)
                        tgItemIDChk(ilUpper).sItemID = slItemID
                        tgItemIDChk(ilUpper).sAudioStatus = "U"
                        tgItemIDChk(ilUpper).sProtStatus = "U"
                        tgItemIDChk(ilUpper).sTitle = slTitle
                        tgItemIDChk(ilUpper).sPriResult = ""
                        tgItemIDChk(ilUpper).sProtResult = ""
                        tgItemIDChk(ilUpper).lLength = 100 * llDuration
                        tgItemIDChk(ilUpper).sPriLen = ""
                        tgItemIDChk(ilUpper).sProtLen = ""
                        tgItemIDChk(ilUpper).lSeeCode = 0
                        ReDim Preserve tgItemIDChk(0 To ilUpper + 1) As ITEMIDCHK
                    End If
                End If
                slItemID = Trim$(grdLibEvents.TextMatrix(llRow, PROTITEMIDINDEX))
                If (slItemID <> "") And (slTitle <> "") And (StrComp(slTitle, "[None]", vbTextCompare) <> 0) Then
                    ilFound = False
                    For ilLoop = 0 To UBound(tgItemIDChk) - 1 Step 1
                        If StrComp(slItemID, Trim$(tgItemIDChk(ilLoop).sItemID), vbTextCompare) = 0 Then
                            ilFound = True
                            tgItemIDChk(ilLoop).sProtStatus = "U"
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        ilUpper = UBound(tgItemIDChk)
                        tgItemIDChk(ilUpper).sItemID = slItemID
                        tgItemIDChk(ilUpper).sAudioStatus = ""
                        tgItemIDChk(ilUpper).sProtStatus = "U"
                        tgItemIDChk(ilUpper).sTitle = slTitle
                        tgItemIDChk(ilUpper).sPriResult = ""
                        tgItemIDChk(ilUpper).sProtResult = ""
                        tgItemIDChk(ilUpper).lLength = 100 * llDuration
                        tgItemIDChk(ilUpper).sPriLen = ""
                        tgItemIDChk(ilUpper).sProtLen = ""
                        tgItemIDChk(ilUpper).lSeeCode = 0
                        ReDim Preserve tgItemIDChk(0 To ilUpper + 1) As ITEMIDCHK
                    End If
                End If
            End If
        End If
    Next llRow

End Sub

Private Function mMerge() As Integer
'    Dim ilRet As Integer
'    Dim ilEof As Integer
'    Dim slLine As String
'    Dim slDate As String
'    Dim llAirDate As Long
'    Dim slTime As String
'    Dim llTime As Long
'    Dim slTitle As String
'    Dim slLen As String
'    Dim slBus As String
'    Dim slCopy As String
'    Dim llLoop As Long
'    Dim ilETE As Integer
'    Dim ilBDE As Integer
'    Dim ilBus As Integer
'    Dim llRow As Long
'    Dim llUpper As Long
'    Dim ilFound As Integer
'    Dim llPrevAvailLoop As Long
'    Dim slDateTime As String
'    Dim slNowDate As String
'    Dim slNowTime As String
'    Dim llNowDate As Long
'    Dim llNowTime As Long
'    Dim ilRemove As Integer
'    Dim ilFindMatch As Integer
'    Dim llAvailLength As Long
'    Dim llCheck As Long
'    Dim llCounter As Long
'    Dim slCounter As String
'    Dim slStr As String
'
'    mMerge = True
'    llAirDate = gDateValue(smAirDate)
'    ReDim tmSpotCurrSEE(0 To 0) As SEE
'    slDateTime = gNow()
'    slNowDate = Format(slDateTime, "ddddd")
'    slNowTime = Format(slDateTime, "ttttt")
'    llNowDate = gDateValue(slNowDate)
'    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
'    If llAirDate = llNowDate Then
'        Print #hmMsg, "Commercial Merge Spots Prior to " & gLongToTime(llNowTime) & " on " & smAirDate & " not checked"
'    End If
'    'Remove Spots
'    llLoop = LBound(tmCurrSEE)
'    Do While llLoop < UBound(tmCurrSEE)
'        ilFound = False
'        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'            If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
'                ilFound = True
'                If tgCurrETE(ilETE).sCategory = "S" Then
'                    ilRemove = True
'                    If llAirDate = llNowDate Then
'                        If llNowTime > tmCurrSEE(llLoop).lTime Then
'                            ilRemove = False
'                        End If
'                    End If
'                    If ilRemove Then
'                        LSet tmSpotCurrSEE(UBound(tmSpotCurrSEE)) = tmCurrSEE(llLoop)
'                        ReDim Preserve tmSpotCurrSEE(0 To UBound(tmSpotCurrSEE) + 1) As SEE
'                        For llRow = llLoop + 1 To UBound(tmCurrSEE) - 1 Step 1
'                            LSet tmCurrSEE(llRow - 1) = tmCurrSEE(llRow)
'                            smT1Comment(llRow - 1) = smT1Comment(llRow)
'                        Next llRow
'                        ReDim Preserve tmCurrSEE(0 To UBound(tmCurrSEE) - 1) As SEE
'                        ReDim Preserve smT1Comment(0 To UBound(smT1Comment) - 1) As String
'                    Else
'                        llLoop = llLoop + 1
'                    End If
'                Else
'                    llLoop = llLoop + 1
'                End If
'            End If
'        Next ilETE
'        If Not ilFound Then
'            llLoop = llLoop + 1
'        End If
'    Loop
'    lbcCommercialSort.Clear
'    llCounter = 0
'    Do
'        'Get Lines
'        ilRet = 0
'        On Error GoTo mMergeErr:
'        Line Input #hmMerge, slLine
'        On Error GoTo 0
'        If ilRet <> 0 Then
'            Exit Do
'        End If
'        If Trim$(slLine) <> "" Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            End If
'        End If
'        DoEvents
'        If Trim$(slLine) <> "" Then
'            slTime = Mid$(slLine, 11, 2) & ":" & Mid$(slLine, 13, 2) & ":" & Mid$(slLine, 15, 2)
'            llTime = 10 * gLengthToLong(slTime)
'            slStr = Trim$(Str$(llTime))
'            Do While Len(slStr) < 8
'                slStr = "0" & slStr
'            Loop
'            llCounter = llCounter + 1
'            slCounter = Trim$(Str$(llCounter))
'            Do While Len(slCounter) < 6
'                slCounter = "0" & slCounter
'            Loop
'            lbcCommercialSort.AddItem slStr & "|" & slCounter & "|" & slLine
'        End If
'    Loop Until ilEof
'    For llCounter = 0 To lbcCommercialSort.ListCount - 1 Step 1
'        slStr = lbcCommercialSort.List(llCounter)
'        slLine = Mid$(slStr, 17)
'
'            slDate = Mid$(slLine, 3, 2) & "/" & Mid$(slLine, 5, 2) & "/" & Mid$(slLine, 1, 2)
'            If gDateValue(slDate) <> llAirDate Then
'                mMerge = False
'                Print #hmMsg, "Commercial Merge Spot Date " & slDate & " does not Match Schedule Date " & smAirDate
'                Exit Function
'            End If
'            slTime = Mid$(slLine, 11, 2) & ":" & Mid$(slLine, 13, 2) & ":" & Mid$(slLine, 15, 2)
'            llTime = 10 * gLengthToLong(slTime)
'            slBus = Trim$(Mid$(slLine, 18, 5))
'            slCopy = Mid$(slLine, 24, 5)
'            slTitle = Trim$(Mid$(slLine, 30, 15))
'            slLen = "00:" & Mid$(slLine, 46, 2) & ":" & Mid$(slLine, 48, 2)
'            ilFound = False
'            llPrevAvailLoop = -1
'            ilFindMatch = True
'            If llAirDate = llNowDate Then
'                If llNowTime > llTime Then
'                    ilFindMatch = False
'                End If
'            End If
'            If ilFindMatch Then
'                For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
'                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'                        If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
'                            If tgCurrETE(ilETE).sCategory = "A" Then
'                                For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
'                                    If tmCurrSEE(llLoop).iBdeCode = tgCurrBDE(ilBDE).iCode Then
'                                        ilBus = StrComp(Trim$(tgCurrBDE(ilBDE).sName), slBus, vbTextCompare)
'                                        If ilBus = 0 Then
'                                            If (tmCurrSEE(llLoop).lTime = llTime) Then  'Or ((tmCurrSEE(llLoop).lTime > llTime) And (llPrevAvailLoop <> -1)) Then
'                                                ilFound = True
'                                                'Create event
'                                                llUpper = UBound(tmCurrSEE)
'                                                mInitSEE tmCurrSEE(llUpper)
'                                                smT1Comment(llUpper) = ""
'                                                If (tmCurrSEE(llLoop).lTime = llTime) Then
'                                                    LSet tmCurrSEE(llUpper) = tmCurrSEE(llLoop)
'                                                    llPrevAvailLoop = llLoop
'                                                Else
'                                                    LSet tmCurrSEE(llUpper) = tmCurrSEE(llPrevAvailLoop)
'                                                End If
'                                                tmCurrSEE(llUpper).lCode = 0
'                                                tmCurrSEE(llUpper).iEteCode = imSpotETECode
'                                                tmCurrSEE(llUpper).lDuration = 10 * gLengthToLong(slLen)
'                                                If tmCurrSEE(llUpper).iAudioAseCode > 0 Then
'                                                    tmCurrSEE(llUpper).sAudioItemID = slCopy
'                                                End If
'                                                If tmCurrSEE(llUpper).iProtAneCode > 0 Then
'                                                    tmCurrSEE(llUpper).sProtItemID = slCopy
'                                                End If
'                                                tmCurrSEE(llUpper).lSpotTime = llTime
'                                                tmARE.lCode = 0
'                                                tmARE.sName = slTitle
'                                                tmARE.sUnusued = ""
'                                                llAvailLength = tmCurrSEE(llLoop).lDuration
'                                                For llCheck = 0 To llUpper - 1 Step 1
'                                                    If (tmCurrSEE(llLoop).iBdeCode = tmCurrSEE(llCheck).iBdeCode) And (tmCurrSEE(llCheck).iEteCode = imSpotETECode) And (tmCurrSEE(llCheck).lTime = llTime) Then
'                                                        llAvailLength = llAvailLength - tmCurrSEE(llCheck).lDuration
'                                                    End If
'                                                Next llCheck
'                                                llAvailLength = llAvailLength - tmCurrSEE(llUpper).lDuration
'                                                If llAvailLength >= 0 Then
'                                                    ilRet = gPutInsert_ARE_AdvertiserRefer(tmARE, "EngrSchdDef-Merge Insert Advertiser Name")
'                                                    If ilRet Then
'                                                        tmCurrSEE(llUpper).lAreCode = tmARE.lCode
'                                                        mSpotMatch tmCurrSEE(llUpper)
'                                                        ReDim Preserve tmCurrSEE(0 To llUpper + 1) As SEE
'                                                        ReDim Preserve smT1Comment(0 To llUpper + 1) As String
'                                                    Else
'                                                        mMerge = False
'                                                        Print #hmMsg, "Unable to Add Advertiser/Product " & slDate & " " & slTime & " " & slTitle
'                                                        mInitSEE tmCurrSEE(llUpper)
'                                                    End If
'                                                Else
'                                                    mMerge = False
'                                                    Print #hmMsg, "Commercial Merge Spot Overbooked Avail " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
'                                                End If
'                                                Exit For
'                                            ElseIf tmCurrSEE(llLoop).lTime < llTime Then
'                                                If llPrevAvailLoop <> -1 Then
'                                                    If tmCurrSEE(llLoop).lTime > tmCurrSEE(llPrevAvailLoop).lTime Then
'                                                        llPrevAvailLoop = llLoop
'                                                    End If
'                                                Else
'                                                    llPrevAvailLoop = llLoop
'                                                End If
'                                            End If
'                                        End If
'                                    End If
'                                Next ilBDE
'                                If ilFound Then
'                                    Exit For
'                                End If
'                            End If
'                        End If
'                    Next ilETE
'                    If ilFound Then
'                        Exit For
'                    End If
'                Next llLoop
'                If (Not ilFound) And (llPrevAvailLoop >= 0) Then
'                    ilFound = True
'                    'Create event
'                    llUpper = UBound(tmCurrSEE)
'                    mInitSEE tmCurrSEE(llUpper)
'                    smT1Comment(llUpper) = ""
'                    LSet tmCurrSEE(llUpper) = tmCurrSEE(llPrevAvailLoop)
'                    tmCurrSEE(llUpper).lCode = 0
'                    tmCurrSEE(llUpper).iEteCode = imSpotETECode
'                    tmCurrSEE(llUpper).lDuration = 10 * gLengthToLong(slLen)
'                    If tmCurrSEE(llUpper).iAudioAseCode > 0 Then
'                        tmCurrSEE(llUpper).sAudioItemID = slCopy
'                    End If
'                    If tmCurrSEE(llUpper).iProtAneCode > 0 Then
'                        tmCurrSEE(llUpper).sProtItemID = slCopy
'                    End If
'                    tmCurrSEE(llUpper).lSpotTime = llTime
'                    tmARE.lCode = 0
'                    tmARE.sName = slTitle
'                    tmARE.sUnusued = ""
'                    llAvailLength = tmCurrSEE(llPrevAvailLoop).lDuration
'                    For llCheck = 0 To llUpper - 1 Step 1
'                        If (tmCurrSEE(llPrevAvailLoop).iBdeCode = tmCurrSEE(llCheck).iBdeCode) And (tmCurrSEE(llCheck).iEteCode = imSpotETECode) And (tmCurrSEE(llCheck).lTime = tmCurrSEE(llPrevAvailLoop).lTime) Then
'                            llAvailLength = llAvailLength - tmCurrSEE(llCheck).lDuration
'                        End If
'                    Next llCheck
'                    llAvailLength = llAvailLength - tmCurrSEE(llUpper).lDuration
'                    If llAvailLength >= 0 Then
'                        ilRet = gPutInsert_ARE_AdvertiserRefer(tmARE, "EngrSchdDef-Merge Insert Advertiser Name")
'                        If ilRet Then
'                            tmCurrSEE(llUpper).lAreCode = tmARE.lCode
'                            mSpotMatch tmCurrSEE(llUpper)
'                            ReDim Preserve tmCurrSEE(0 To llUpper + 1) As SEE
'                            ReDim Preserve smT1Comment(0 To llUpper + 1) As String
'                        Else
'                            mMerge = False
'                            Print #hmMsg, "Unable to Add Advertiser/Product " & slDate & " " & slTime & " " & slTitle
'                            mInitSEE tmCurrSEE(llUpper)
'                        End If
'                    Else
'                        mMerge = False
'                        Print #hmMsg, "Commercial Merge Spot Overbooked Avail " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
'                    End If
'                End If
'                If Not ilFound Then
'                    mMerge = False
'                    Print #hmMsg, "Commercial Merge Spot Avail Not Found " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
'                End If
'            End If
'    '    End If
'    'Loop Until ilEof
'    Next llCounter
'    Exit Function
'mMergeErr:
'    ilRet = Err.Number
'    Resume Next
End Function
Private Sub mTestItemID()
'    Dim llLoop As Long
'    Dim ilETE As Integer
'    Dim ilITE As Integer
'    Dim tlPriITE As ITE
'    Dim tlSecITE As ITE
'    Dim slCart As String
'    Dim slQuery As String
'    Dim slPriQuery As String
'    Dim slResult As String
'    Dim slTitle As String
'    Dim ilASE As Integer
'    Dim slTestItemID As String
'    Dim ilATE As Integer
'    Dim ilANE As Integer
'    Dim ilRet As Integer
'    Dim ilTestPort As Integer
'
'    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
'        If tgCurrITE(ilITE).sType = "P" Then
'            LSet tlPriITE = tgCurrITE(ilITE)
'            Exit For
'        End If
'    Next ilITE
'    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
'        If tgCurrITE(ilITE).sType = "S" Then
'            LSet tlSecITE = tgCurrITE(ilITE)
'            Exit For
'        End If
'    Next ilITE
'    ilTestPort = True
'    For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
'        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'            If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
'                If tgCurrETE(ilETE).sCategory = "S" Then
'                    slCart = Trim$(tmCurrSEE(llLoop).sAudioItemID)
'                    slTitle = ""
'                    If slCart <> "" Then
'                        slTestItemID = ""
'                        For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
'                            If tmCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASE).iCode Then
'                                For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
'                                    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
'                                        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
'                                            If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
'                                                slTestItemID = tgCurrATE(ilATE).sTestItemID
'                                                Exit For
'                                            End If
'                                        Next ilATE
'                                        If slTestItemID <> "" Then
'                                            Exit For
'                                        End If
'                                    End If
'                                Next ilANE
'                                If slTestItemID <> "" Then
'                                    Exit For
'                                End If
'                            End If
'                        Next ilASE
'                        If (slTestItemID = "Y") And (ilTestPort) Then
'                            ilRet = gGetRec_ARE_AdvertiserRefer(tmCurrSEE(llLoop).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
'                            If ilRet Then
'                                slTitle = Trim$(tmARE.sName)
'                            End If
'                        End If
'                        If (slTestItemID = "Y") And (slTitle <> "") And (ilTestPort) Then
'                            gBuildItemIDQuery slCart, tlPriITE, slQuery, slPriQuery
'                            ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slQuery, slPriQuery, slResult)
'                            If ilRet Then
'                                slResult = Mid$(slResult, Len(slPriQuery) + 1)
'                                If StrComp(Trim$(slTitle), slResult, vbTextCompare) = 0 Then
'                                    tmCurrSEE(llLoop).sAudioItemIDChk = "O"
'                                Else
'                                    tmCurrSEE(llLoop).sAudioItemIDChk = "F"
'                                End If
'                            Else
'                                If StrComp(slResult, "Failed", vbTextCompare) = 0 Then
'                                    ilTestPort = False
'                                End If
'                                tmCurrSEE(llLoop).sAudioItemIDChk = "N"
'                            End If
'                        End If
'                    End If
'                    slCart = Trim$(tmCurrSEE(llLoop).sProtItemID)
'                    If slCart <> "" Then
'                        slTestItemID = ""
'                        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
'                            If tmCurrSEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
'                                For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
'                                    If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
'                                        slTestItemID = tgCurrATE(ilATE).sTestItemID
'                                        Exit For
'                                    End If
'                                Next ilATE
'                                If slTestItemID <> "" Then
'                                    Exit For
'                                End If
'                            End If
'                        Next ilANE
'                        If (slTestItemID = "Y") And (slTitle = "") And (ilTestPort) Then
'                            ilRet = gGetRec_ARE_AdvertiserRefer(tmCurrSEE(llLoop).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
'                            If ilRet Then
'                                slTitle = Trim$(tmARE.sName)
'                            End If
'                        End If
'                        If (slTestItemID = "Y") And (slTitle <> "") And (ilTestPort) Then
'                            gBuildItemIDQuery slCart, tlPriITE, slQuery, slPriQuery
'                            ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slQuery, slPriQuery, slResult)
'                            If ilRet Then
'                                slResult = Mid$(slResult, Len(slPriQuery) + 1)
'                                If StrComp(Trim$(slTitle), slResult, vbTextCompare) = 0 Then
'                                    tmCurrSEE(llLoop).sProtItemIDChk = "O"
'                                Else
'                                    tmCurrSEE(llLoop).sProtItemIDChk = "F"
'                                End If
'                            Else
'                                If StrComp(slResult, "Failed", vbTextCompare) = 0 Then
'                                    ilTestPort = False
'                                End If
'                                tmCurrSEE(llLoop).sProtItemIDChk = "N"
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        Next ilETE
'    Next llLoop
End Sub

Private Sub mSpotMatch(tlSEE As SEE)
'    Dim ilLoop As Integer
'
'    tlSEE.lCode = 0
'    For ilLoop = 0 To UBound(tmSpotCurrSEE) - 1 Step 1
'        If tlSEE.iBdeCode = tmSpotCurrSEE(ilLoop).iBdeCode Then
'            If tlSEE.lTime = tmSpotCurrSEE(ilLoop).lTime Then
'                If tlSEE.lDuration = tmSpotCurrSEE(ilLoop).lDuration Then
'                    If tlSEE.lAreCode = tmSpotCurrSEE(ilLoop).lAreCode Then
'                        tlSEE.lCode = tmSpotCurrSEE(ilLoop).lCode
'                        tlSEE.sAudioItemIDChk = tmSpotCurrSEE(ilLoop).sAudioItemIDChk
'                        tlSEE.sProtItemIDChk = tmSpotCurrSEE(ilLoop).sProtItemIDChk
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'    Next ilLoop
End Sub

Private Sub mReplaceValuesAvails()
    Dim ilRowOk As Integer
    Dim slCategory As String
    Dim llAvailLength As Long
    Dim llAirDate As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llLoop As Long
    Dim ilETE As Integer
    Dim llTest As Long
    Dim ilField As Integer
    Dim ilReplace As Integer
    Dim slOldValue As String
    Dim slNewValue As String
    Dim llOldValue As Long
    Dim llNewValue As Long
    Dim ilASEOld As Integer
    Dim ilASENew As Integer
    Dim ilFieldType As Integer
    Dim slFileName As String
    
    If Not bgApplyToEventType(1) Then
        Exit Sub
    End If
    llAirDate = gDateValue(smAirDate)
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    For ilReplace = LBound(tgSchdReplaceValues) To UBound(tgSchdReplaceValues) - 1 Step 1
        For ilField = LBound(tgReplaceFields) To UBound(tgReplaceFields) - 1 Step 1
            If tgReplaceFields(ilField).sFieldName = tgSchdReplaceValues(ilReplace).sFieldName Then
                ilFieldType = tgReplaceFields(ilField).iFieldType
                slFileName = tgReplaceFields(ilField).sListFile
                For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
                    slCategory = ""
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If tmCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                            slCategory = tgCurrETE(ilETE).sCategory
                            Exit For
                        End If
                    Next ilETE
                    'If Avail, then check if any time left
                    ilRowOk = True
                    If slCategory = "A" Then
                        If Not bgApplyToEventType(1) Then
                            ilRowOk = False
                        Else
                            llAvailLength = tmCurrSEE(llLoop).lDuration
                            For llTest = 0 To UBound(tmCurrSEE) - 1 Step 1
                                If (tmCurrSEE(llLoop).iBdeCode = tmCurrSEE(llTest).iBdeCode) And (tmCurrSEE(llLoop).lTime = tmCurrSEE(llTest).lTime) And (tmCurrSEE(llTest).iEteCode = imSpotETECode) Then
                                    llAvailLength = llAvailLength - tmCurrSEE(llTest).lDuration
                                End If
                            Next llTest
                            If llAvailLength <= 0 Then
                                ilRowOk = False
                            End If
                        End If
                    Else
                        ilRowOk = False
                    End If
                    'Add Criteria test here
                    If ilRowOk Then
                        ilRowOk = mCheckFilter(tmCurrSEE(llLoop), smT1Comment(llLoop))
                    End If
                    If ilRowOk Then
                        If llAirDate = llNowDate Then
                            If llNowTime > tmCurrSEE(llLoop).lTime Then
                                ilRowOk = False
                            End If
                        End If
                    End If
                    If ilRowOk Then
                        slOldValue = Trim$(tgSchdReplaceValues(ilReplace).sOldValue)
                        slNewValue = Trim$(tgSchdReplaceValues(ilReplace).sNewValue)
                        llOldValue = tgSchdReplaceValues(ilReplace).lOldCode
                        llNewValue = tgSchdReplaceValues(ilReplace).lNewCode
                        If ilFieldType = 5 Then
                            Select Case UCase$(Trim$(slFileName))
                                Case "ANE"
                                    'For ilASEOld = 0 To UBound(tgCurrASE) - 1 Step 1
                                    '    If tmCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASEOld).iCode Then
                                        ilASEOld = gBinarySearchASE(tmCurrSEE(llLoop).iAudioAseCode, tgCurrASE())
                                        If ilASEOld <> -1 Then
                                            If llOldValue = tgCurrASE(ilASEOld).iPriAneCode Then
                                                For ilASENew = 0 To UBound(tgCurrASE) - 1 Step 1
                                                    If llNewValue = tgCurrASE(ilASENew).iPriAneCode Then
                                                        tmCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASENew).iCode
                                                        Exit For
                                                    End If
                                                Next ilASENew
                                    '            Exit For
                                            End If
                                        End If
                                    'Next ilASEOld
                                    If llOldValue = tmCurrSEE(llLoop).iProtAneCode Then
                                        tmCurrSEE(llLoop).iProtAneCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).iBkupAneCode Then
                                        tmCurrSEE(llLoop).iBkupAneCode = CInt(llNewValue)
                                    End If
                                Case "BDE"
                                    If llOldValue = tmCurrSEE(llLoop).iBdeCode Then
                                        tmCurrSEE(llLoop).iBdeCode = CInt(llNewValue)
                                    End If
                                Case "FNE"
                                    If llOldValue = tmCurrSEE(llLoop).iFneCode Then
                                        tmCurrSEE(llLoop).iFneCode = CInt(llNewValue)
                                    End If
                                Case "MTE"
                                    If llOldValue = tmCurrSEE(llLoop).iMteCode Then
                                        tmCurrSEE(llLoop).iMteCode = CInt(llNewValue)
                                    End If
                                Case "NNE"
                                    If llOldValue = tmCurrSEE(llLoop).iStartNneCode Then
                                        tmCurrSEE(llLoop).iStartNneCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).iEndNneCode Then
                                        tmCurrSEE(llLoop).iEndNneCode = CInt(llNewValue)
                                    End If
                                Case "RNE"
                                    If llOldValue = tmCurrSEE(llLoop).i1RneCode Then
                                        tmCurrSEE(llLoop).i1RneCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).i2RneCode Then
                                        tmCurrSEE(llLoop).i2RneCode = CInt(llNewValue)
                                    End If
                                Case "TTES"
                                    If llOldValue = tmCurrSEE(llLoop).iStartTteCode Then
                                        tmCurrSEE(llLoop).iStartTteCode = CInt(llNewValue)
                                    End If
                                Case "TTEE"
                                    If llOldValue = tmCurrSEE(llLoop).iEndTteCode Then
                                        tmCurrSEE(llLoop).iEndTteCode = CInt(llNewValue)
                                    End If
                                Case "CCEA"
                                    If llOldValue = tmCurrSEE(llLoop).iAudioCceCode Then
                                        tmCurrSEE(llLoop).iAudioCceCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).iProtCceCode Then
                                        tmCurrSEE(llLoop).iProtCceCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).iBkupCceCode Then
                                        tmCurrSEE(llLoop).iBkupCceCode = CInt(llNewValue)
                                    End If
                                Case "CCEB"
                                    If llOldValue = tmCurrSEE(llLoop).iBusCceCode Then
                                        tmCurrSEE(llLoop).iBusCceCode = CInt(llNewValue)
                                    End If
                                Case "SCE"
                                    If llOldValue = tmCurrSEE(llLoop).i1SceCode Then
                                        tmCurrSEE(llLoop).i1SceCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).i2SceCode Then
                                        tmCurrSEE(llLoop).i2SceCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).i3SceCode Then
                                        tmCurrSEE(llLoop).i3SceCode = CInt(llNewValue)
                                    End If
                                    If llOldValue = tmCurrSEE(llLoop).i4SceCode Then
                                        tmCurrSEE(llLoop).i4SceCode = CInt(llNewValue)
                                    End If
                                '7/8/11: Make T2 work like T1
                                'Case "CTE2"
                                '    If llOldValue = tmCurrSEE(llLoop).l2CteCode Then
                                '        tmCurrSEE(llLoop).l2CteCode = llNewValue
                                '    End If
                            End Select
                        ElseIf ilFieldType = 9 Then
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "Fixed Time" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sFixedTime), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sFixedTime = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "Title 1" Then
                                If StrComp(slOldValue, Trim$(smT1Comment(llLoop)), vbTextCompare) = 0 Then
                                    smT1Comment(llLoop) = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "Title 2" Then
                                If StrComp(slOldValue, Trim$(smT2Comment(llLoop)), vbTextCompare) = 0 Then
                                    smT2Comment(llLoop) = slNewValue
                                End If
                            End If
                        ElseIf ilFieldType = 2 Then
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "Item ID" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sAudioItemID), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sAudioItemID = slNewValue
                                End If
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sProtItemID), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sProtItemID = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "ISCI" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sAudioISCI), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sAudioISCI = slNewValue
                                End If
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sProtISCI), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sProtISCI = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC Format" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sABCFormat), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sABCFormat = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC Pgm Code" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sABCPgmCode), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sABCPgmCode = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC XDS Mode" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sABCXDSMode), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sABCXDSMode = slNewValue
                                End If
                            End If
                            If Trim$(tgReplaceFields(ilField).sFieldName) = "ABC Recd Item" Then
                                If StrComp(slOldValue, Trim$(tmCurrSEE(llLoop).sABCRecordItem), vbTextCompare) = 0 Then
                                    tmCurrSEE(llLoop).sABCRecordItem = slNewValue
                                End If
                            End If
                        End If
                    End If
                Next llLoop
            End If
        Next ilField
    Next ilReplace
End Sub

Private Sub mShowConflictGrid()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilColumn As Integer
    Dim ilPos As Integer
    Dim llRowFd As Long
    Dim llIndex As Integer
    Dim slCurrEBEStamp As String
    Dim tlCurrEBE() As EBE
    Dim ilDNE As Integer
    Dim ilDSE As Integer
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilFound As Integer
    Dim slHours As String
    Dim tlDEE As DEE
    Dim tlDHE As DHE
    Dim tlDNE As DNE
    Dim tlDSE As DSE
    Dim ilEBE As Integer
    
    If lmEEnableRow < grdLibEvents.FixedRows Then
        'grdLibEvents.Height = lmGridLibEventsHeight
        grdConflicts.Visible = False
        Exit Sub
    End If
    If lmConflictRow = lmEEnableRow Then
        If lmEEnableRow <> -1 Then
            'grdLibEvents.Height = lmGridLibEventsHeight
            grdConflicts.Visible = True
        End If
        Exit Sub
    End If
    gGrid_Clear grdConflicts, True
    slStr = grdLibEvents.TextMatrix(lmEEnableRow, ERRORCONFLICTINDEX)
    If slStr = "" Then
        lmConflictRow = -1
        'grdLibEvents.Height = lmGridLibEventsHeight
        grdConflicts.Visible = False
        Exit Sub
    End If
    If Val(slStr) <= 0 Then
        lmConflictRow = -1
        'grdLibEvents.Height = lmGridLibEventsHeight
        grdConflicts.Visible = False
        Exit Sub
    End If
    llRow = grdConflicts.FixedRows
    ilLoop = Val(slStr)
    If llRow + 1 > grdConflicts.Rows Then
        grdConflicts.AddItem ""
    End If
    grdConflicts.Row = llRow
    slStr = Trim$(grdLibEvents.TextMatrix(lmEEnableRow, LIBNAMEINDEX))
    grdLibEvents.TextMatrix(lmEEnableRow, LIBNAMEINDEX) = mGetLibName(slStr)
    slStr = grdLibEvents.TextMatrix(lmEEnableRow, LIBNAMEINDEX)
    ilPos = InStr(1, slStr, "/", vbTextCompare)
    If ilPos > 0 Then
        grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = Left$(slStr, ilPos - 1)
        grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = Mid$(slStr, ilPos + 1)
    Else
        grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, LIBNAMEINDEX)
        grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = ""
    End If
    grdConflicts.TextMatrix(llRow, CONFLICTSTARTDATEINDEX) = ""
    grdConflicts.TextMatrix(llRow, CONFLICTENDDATEINDEX) = ""
    grdConflicts.TextMatrix(llRow, CONFLICTDAYSINDEX) = ""
    grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = ""
    slStr = grdLibEvents.TextMatrix(lmEEnableRow, EVENTIDINDEX)
    If Val(slStr) > 0 Then
        grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = slStr
    End If
    grdConflicts.TextMatrix(llRow, CONFLICTHOURSINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, TIMEINDEX)
    grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, DURATIONINDEX)
    grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, BUSNAMEINDEX)
    grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, AUDIONAMEINDEX)
    grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, BACKUPNAMEINDEX)
    grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = grdLibEvents.TextMatrix(lmEEnableRow, PROTNAMEINDEX)
    llRow = llRow + 1
    Do
        If llRow + 1 > grdConflicts.Rows Then
            grdConflicts.AddItem ""
        End If
        grdConflicts.Row = llRow
        If (tmConflictList(ilLoop).sType = "L") Or (tmConflictList(ilLoop).sType = "T") Then
            ilRet = gGetRec_DEE_DayEvent(tmConflictList(ilLoop).lDeeCode, "EngrLibDef-gGetRec_DEE_DayEvent", tlDEE)
            ilRet = gGetRec_DHE_DayHeaderInfo(tmConflictList(ilLoop).lDheCode, "EngrLibDef-gGetRec_DHE_DayHeaderInfo", tlDHE)
            ilRet = gGetRec_DNE_DayName(tlDHE.lDneCode, "EngrLibDef-gGetRec_DNE_DayName", tlDNE)
            ilRet = gGetRec_DSE_DaySubName(tlDHE.lDseCode, "EngrLibDef-gGetRec_DSE_DaySubName", tlDSE)
            
            grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = Trim$(tlDNE.sName)
            grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = Trim$(tlDSE.sName)
            grdConflicts.TextMatrix(llRow, CONFLICTSTARTDATEINDEX) = tmConflictList(ilLoop).sStartDate
            If gDateValue(Trim$(tmConflictList(ilLoop).sEndDate)) <> gDateValue("12/31/2069") Then
                grdConflicts.TextMatrix(llRow, CONFLICTENDDATEINDEX) = tmConflictList(ilLoop).sEndDate
            Else
                grdConflicts.TextMatrix(llRow, CONFLICTENDDATEINDEX) = ""
            End If
            slStr = gDayMap(tlDEE.sDays)
            grdConflicts.TextMatrix(llRow, CONFLICTDAYSINDEX) = Trim$(slStr)
            grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = gLongToStrLengthInTenth(tlDEE.lTime, False)
            slHours = Trim$(tlDEE.sHours)
            slStr = gHourMap(slHours)
            grdConflicts.TextMatrix(llRow, CONFLICTHOURSINDEX) = slStr
            If (tlDEE.lDuration > 0) Then
                grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = gLongToStrLengthInTenth(tlDEE.lDuration, True)
            Else
                grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = gLongToStrLengthInTenth(tlDEE.lDuration, True)
            End If
            slStr = ""
            slCurrEBEStamp = ""
            Erase tlCurrEBE
            ilRet = gGetRecs_EBE_EventBusSel(slCurrEBEStamp, tlDEE.lCode, "Bus Definition-mShowConflictGrid", tlCurrEBE())
            For ilEBE = 0 To UBound(tlCurrEBE) - 1 Step 1
                'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                '    If tlCurrEBE(ilEBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                    ilBDE = gBinarySearchBDE(tlCurrEBE(ilEBE).iBdeCode, tgCurrBDE())
                    If ilBDE <> -1 Then
                        slStr = slStr & Trim$(tgCurrBDE(ilBDE).sName) & ","
                '        Exit For
                    End If
                'Next ilBDE
            Next ilEBE
            If slStr <> "" Then
                slStr = Left$(slStr, Len(slStr) - 1)
            End If
            grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = slStr
            grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = ""
            'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
            '    If tlDEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                ilASE = gBinarySearchASE(tlDEE.iAudioAseCode, tgCurrASE())
                If ilASE <> -1 Then
                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                    '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                        ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                        If ilANE <> -1 Then
                            grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                        End If
                    'Next ilANE
            '        Exit For
                End If
            'Next ilASE
            grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = ""
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tlDEE.iBkupAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tlDEE.iBkupAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = Trim$(tgCurrANE(ilANE).sName)
            '        Exit For
                End If
            'Next ilANE
            grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = ""
            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            '    If tlDEE.iProtAneCode = tgCurrANE(ilANE).iCode Then
                ilANE = gBinarySearchANE(tlDEE.iProtAneCode, tgCurrANE())
                If ilANE <> -1 Then
                    grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = Trim$(tgCurrANE(ilANE).sName)
            '        Exit For
                End If
            'Next ilANE
            llRow = llRow + 1
        ElseIf (tmConflictList(ilLoop).sType = "E") Then
            'Find Row Information
            ilFound = False
            For llRowFd = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
                If Trim$(grdLibEvents.TextMatrix(llRowFd, EVENTTYPEINDEX)) <> "" Then
                    If Trim$(grdLibEvents.TextMatrix(llRowFd, TMCURRSEEINDEX)) <> "" Then
                        llIndex = Val(grdLibEvents.TextMatrix(llRowFd, TMCURRSEEINDEX))
                        If tmConflictList(ilLoop).lIndex = llIndex Then
                            slStr = Trim$(grdLibEvents.TextMatrix(llRowFd, LIBNAMEINDEX))
                            grdLibEvents.TextMatrix(llRowFd, LIBNAMEINDEX) = mGetLibName(slStr)
                            slStr = grdLibEvents.TextMatrix(llRowFd, LIBNAMEINDEX)
                            ilPos = InStr(1, slStr, "/", vbTextCompare)
                            If ilPos > 0 Then
                                grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = Left$(slStr, ilPos - 1)
                                grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = Mid$(slStr, ilPos + 1)
                            Else
                                grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = grdLibEvents.TextMatrix(llRowFd, LIBNAMEINDEX)
                                grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = ""
                            End If
                            grdConflicts.TextMatrix(llRow, CONFLICTSTARTDATEINDEX) = ""
                            grdConflicts.TextMatrix(llRow, CONFLICTENDDATEINDEX) = ""
                            grdConflicts.TextMatrix(llRow, CONFLICTDAYSINDEX) = ""
                            grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = ""
                            slStr = grdLibEvents.TextMatrix(llRowFd, EVENTIDINDEX)
                            If Val(slStr) > 0 Then
                                grdConflicts.TextMatrix(llRow, CONFLICTOFFSETINDEX) = slStr
                            End If
                            grdConflicts.TextMatrix(llRow, CONFLICTHOURSINDEX) = grdLibEvents.TextMatrix(llRowFd, TIMEINDEX)
                            grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = grdLibEvents.TextMatrix(llRowFd, DURATIONINDEX)
                            grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = grdLibEvents.TextMatrix(llRowFd, BUSNAMEINDEX)
                            grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = grdLibEvents.TextMatrix(llRowFd, AUDIONAMEINDEX)
                            grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = grdLibEvents.TextMatrix(llRowFd, BACKUPNAMEINDEX)
                            grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = grdLibEvents.TextMatrix(llRowFd, PROTNAMEINDEX)
                            ilFound = True
                        End If
                    End If
                End If
            Next llRowFd
            If Not ilFound Then
                llIndex = tmConflictList(ilLoop).lIndex
                grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = ""
                grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = ""
                If tmCurrSEE(llIndex).lDeeCode > 0 Then
                    ilRet = gGetRec_DEE_DayEvent(tmCurrSEE(llIndex).lDeeCode, "EngrSchdDef-mMoveSEERecToCtrls: DEE", tmDee)
                    ilRet = gGetRec_DHE_DayHeaderInfo(tmDee.lDheCode, "EngrSchdDef-mMoveSEERecToCtrls: DHE", tmDHE)
                    
                    If tmDHE.sType <> "T" Then
                        For ilDNE = 0 To UBound(tgCurrLibDNE) - 1 Step 1
                            If tmDHE.lDneCode = tgCurrLibDNE(ilDNE).lCode Then
                                grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = Trim$(tgCurrLibDNE(ilDNE).sName)
                                Exit For
                            End If
                        Next ilDNE
                    Else
                        For ilDNE = 0 To UBound(tgCurrTempDNE) - 1 Step 1
                            If tmDHE.lDneCode = tgCurrTempDNE(ilDNE).lCode Then
                                grdConflicts.TextMatrix(llRow, CONFLICTNAMEINDEX) = Trim$(tgCurrTempDNE(ilDNE).sName)
                                Exit For
                            End If
                        Next ilDNE
                    End If
                    For ilDSE = 0 To UBound(tgCurrDSE) - 1 Step 1
                        If tmDHE.lDseCode = tgCurrDSE(ilDSE).lCode Then
                            grdConflicts.TextMatrix(llRow, CONFLICTSUBNAMEINDEX) = Trim$(tgCurrDSE(ilDSE).sName)
                            Exit For
                        End If
                    Next ilDSE
                End If
                grdConflicts.TextMatrix(llRow, CONFLICTSTARTDATEINDEX) = ""
                grdConflicts.TextMatrix(llRow, CONFLICTENDDATEINDEX) = ""
                grdConflicts.TextMatrix(llRow, CONFLICTDAYSINDEX) = ""
                grdLibEvents.TextMatrix(llRow, CONFLICTOFFSETINDEX) = ""
                If tmCurrSEE(llIndex).iEteCode <> imSpotETECode Then
                    grdLibEvents.TextMatrix(llRow, CONFLICTHOURSINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llIndex).lTime)
                Else
                    grdLibEvents.TextMatrix(llRow, CONFLICTHOURSINDEX) = gLongToStrTimeInTenth(tmCurrSEE(llIndex).lSpotTime)
                End If
                grdConflicts.TextMatrix(llRow, CONFLICTDURATIONINDEX) = gLongToStrLengthInTenth(tmCurrSEE(llIndex).lDuration, True)
                grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = ""
                'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                '    If tmCurrSEE(llIndex).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                    ilBDE = gBinarySearchBDE(tmCurrSEE(llIndex).iBdeCode, tgCurrBDE())
                    If ilBDE <> -1 Then
                        grdConflicts.TextMatrix(llRow, CONFLICTBUSESINDEX) = Trim$(tgCurrBDE(ilBDE).sName)
                '        Exit For
                    End If
                'Next ilBDE
                grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = ""
                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                '    If tmCurrSEE(llIndex).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                    ilASE = gBinarySearchASE(tmCurrSEE(llIndex).iAudioAseCode, tgCurrASE())
                    If ilASE <> -1 Then
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                            ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                            If ilANE <> -1 Then
                                grdConflicts.TextMatrix(llRow, CONFLICTAUDIOINDEX) = Trim$(tgCurrANE(ilANE).sName)
                            End If
                        'Next ilANE
                '        Exit For
                    End If
                'Next ilASE
                grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = ""
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tmCurrSEE(llIndex).iBkupAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tmCurrSEE(llIndex).iBkupAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        grdConflicts.TextMatrix(llRow, CONFLICTBACKUPINDEX) = Trim$(tgCurrANE(ilANE).sName)
                '        Exit For
                    End If
                'Next ilANE
                grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = ""
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tmCurrSEE(llIndex).iProtAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tmCurrSEE(llIndex).iProtAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        grdConflicts.TextMatrix(llRow, CONFLICTPROTINDEX) = Trim$(tgCurrANE(ilANE).sName)
                '        Exit For
                    End If
                'Next ilANE
            End If
            llRow = llRow + 1
        End If
        ilLoop = tmConflictList(ilLoop).iNextIndex
    Loop While ilLoop > 0

    For llRow = grdConflicts.FixedRows To grdConflicts.Rows - 1 Step 1
        grdConflicts.Row = llRow
        For ilColumn = CONFLICTNAMEINDEX To CONFLICTPROTINDEX Step 1
            grdConflicts.Col = ilColumn
            grdConflicts.CellBackColor = LIGHTYELLOW
        Next ilColumn
    Next llRow
    
    lmConflictRow = lmEEnableRow
    'grdLibEvents.Height = lmGridLibEventsHeight - grdConflicts.Height - 120
    'gGrid_IntegralHeight grdLibEvents
    'grdLibEvents.Height = grdLibEvents.Height - 15
    grdConflicts.Visible = True
    grdConflicts.Redraw = True
End Sub

Private Sub mHideConflictGrid()
    grdConflicts.Visible = False
    'grdLibEvents.Height = lmGridLibEventsHeight
End Sub

Private Function mAddConflict(ilStartIndex As Integer, tlConflictList() As CONFLICTLIST) As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    
    If ilStartIndex = 0 Then
        mAddConflict = UBound(tlConflictList)
        ReDim Preserve tlConflictList(1 To UBound(tlConflictList) + 1) As CONFLICTLIST
        tlConflictList(UBound(tlConflictList)).iNextIndex = -1
        Exit Function
    End If
    mAddConflict = ilStartIndex
    ilIndex = ilStartIndex
    ilUpper = UBound(tlConflictList)
    Do
        If (tlConflictList(ilIndex).sType = "E") And (tlConflictList(ilUpper).sType = "E") Then
            If (tlConflictList(ilIndex).lIndex = tlConflictList(ilUpper).lIndex) Then
                Exit Function
            End If
        End If
        ilIndex = tlConflictList(ilIndex).iNextIndex
    Loop While ilIndex > 0
    'Add to chain
    ilIndex = ilStartIndex
    Do
        If tlConflictList(ilIndex).iNextIndex = -1 Then
            tlConflictList(ilIndex).iNextIndex = UBound(tlConflictList)
            ReDim Preserve tlConflictList(1 To UBound(tlConflictList) + 1) As CONFLICTLIST
            tlConflictList(UBound(tlConflictList)).iNextIndex = -1
            Exit Function
        End If
        ilIndex = tlConflictList(ilIndex).iNextIndex
    Loop While ilIndex <> -1
End Function


Private Sub mAdjustAvailTime(llSpotStartIndex As Long)
    Dim llRow As Long
    Dim ilETE As Integer
    Dim slCategory As String
    Dim llAvailLength As Long
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    Dim llStart As Long
    Dim ilTest As Integer
    
    For llRow = 0 To UBound(tmCurrSEE) - 1 Step 1
        slCategory = ""
        ilETE = gBinarySearchETE(tmCurrSEE(llRow).iEteCode, tgCurrETE)
        If ilETE <> -1 Then
            slCategory = tgCurrETE(ilETE).sCategory
        End If
        If slCategory = "A" Then
            tmCurrSEE(llRow).lSpotTime = tmCurrSEE(llRow).lTime
            llAvailLength = tmCurrSEE(llRow).lDuration
            'If llSpotStartIndex <= 0 Then
            '    llStart = llRow + 1
            'Else
            '    llStart = llSpotStartIndex
            'End If
            'For llAvailTest = llStart To UBound(tmCurrSEE) - 1 Step 1
            For llAvailTest = 0 To UBound(tmCurrSEE) - 1 Step 1
                ilTest = False
                'If llSpotStartIndex <= 0 Then
                '    If tmCurrSEE(llAvailTest).lTime <> tmCurrSEE(llRow).lTime Then
                '        Exit For
                '    End If
                '    ilTest = True
                'Else
                '    'If tmCurrSEE(llAvailTest).lTime > tmCurrSEE(llRow).lTime Then
                '    '    Exit For
                '    'End If
                '    If tmCurrSEE(llAvailTest).lTime = tmCurrSEE(llRow).lTime Then
                '        ilTest = True
                '    End If
                'End If
                If (tmCurrSEE(llAvailTest).sAction <> "D") And (tmCurrSEE(llAvailTest).sAction <> "R") Then
                    If (tmCurrSEE(llAvailTest).lTime = tmCurrSEE(llRow).lTime) And (llRow <> llAvailTest) Then
                        ilTest = True
                    End If
                End If
                If ilTest Then
                    If tmCurrSEE(llAvailTest).iBdeCode = tmCurrSEE(llRow).iBdeCode Then
                        ilETE = gBinarySearchETE(tmCurrSEE(llAvailTest).iEteCode, tgCurrETE)
                        If ilETE <> -1 Then
                            If tgCurrETE(ilETE).sCategory = "S" Then
                                llAvailLength = llAvailLength - tmCurrSEE(llAvailTest).lDuration
                                llTimeTest = tmCurrSEE(llAvailTest).lSpotTime + tmCurrSEE(llAvailTest).lDuration
                                If llTimeTest > tmCurrSEE(llRow).lSpotTime Then
                                    tmCurrSEE(llRow).lSpotTime = llTimeTest
                                End If
                            End If
                        End If
                    End If
                End If
            Next llAvailTest
            tmCurrSEE(llRow).lAvailLength = llAvailLength
        End If
    Next llRow
End Sub

Private Function mExportCol(llRow As Long, llCol As Long) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    Dim ilUsed As Integer
    
    mExportCol = True
    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "U" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            Select Case llCol
                                Case BUSNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BUSCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case EVENTTYPEINDEX
                                    'Event Type exported if any other column exported and tgStartColAFE.iEventType >0
                                Case EVENTIDINDEX
                                    'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                                Case TIMEINDEX
                                    If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case STARTTYPEINDEX
                                    If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FIXEDINDEX
                                    If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ENDTYPEINDEX
                                    If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case DURATIONINDEX
                                    If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case MATERIALINDEX
                                    If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIONAMEINDEX
                                    If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOISCIINDEX
                                    If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOCTRLINDEX
                                    If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTNAMEINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTISCIINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTCTRLINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY1INDEX
                                    If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY2INDEX
                                    If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FOLLOWINDEX
                                    If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCETIMEINDEX
                                    If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE1INDEX
                                    If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE2INDEX
                                    If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE3INDEX
                                    If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE4INDEX
                                    If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE1INDEX
                                    If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE2INDEX
                                    If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE1INDEX
                                    If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE2INDEX
                                    If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCFORMATINDEX
                                    If (tgCurrEPE(ilEPE).sABCFormat <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCPGMCODEINDEX
                                    If (tgCurrEPE(ilEPE).sABCPgmCode <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCXDSMODEINDEX
                                    If (tgCurrEPE(ilEPE).sABCXDSMode <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCRECORDITEMINDEX
                                    If (tgCurrEPE(ilEPE).sABCRecordItem <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                            End Select
                            Exit For
                        End If
                    End If
                Next ilEPE
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "E" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            Select Case llCol
                                Case BUSNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BUSCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case EVENTTYPEINDEX
                                    'Always exported if any other col is exported
                                Case EVENTIDINDEX
                                    'Always exported if any other col is exported
                                Case TIMEINDEX
                                    If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case STARTTYPEINDEX
                                    If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FIXEDINDEX
                                    If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ENDTYPEINDEX
                                    If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case DURATIONINDEX
                                    If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case MATERIALINDEX
                                    If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIONAMEINDEX
                                    If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOISCIINDEX
                                    If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case AUDIOCTRLINDEX
                                    If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPNAMEINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case BACKUPCTRLINDEX
                                    If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTNAMEINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTITEMIDINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTISCIINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case PROTCTRLINDEX
                                    If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY1INDEX
                                    If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case RELAY2INDEX
                                    If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case FOLLOWINDEX
                                    If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCETIMEINDEX
                                    If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE1INDEX
                                    If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE2INDEX
                                    If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE3INDEX
                                    If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case SILENCE4INDEX
                                    If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE1INDEX
                                    If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case NETCUE2INDEX
                                    If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE1INDEX
                                    If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case TITLE2INDEX
                                    If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCFORMATINDEX
                                    If (tgCurrEPE(ilEPE).sABCFormat <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCPGMCODEINDEX
                                    If (tgCurrEPE(ilEPE).sABCPgmCode <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCXDSMODEINDEX
                                    If (tgCurrEPE(ilEPE).sABCXDSMode <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                                Case ABCRECORDITEMINDEX
                                    If (tgCurrEPE(ilEPE).sABCRecordItem <> "Y") Or (sgClientFields <> "A") Then
                                        mExportCol = False
                                        Exit Function
                                    End If
                            End Select
                            Exit For
                        End If
                    End If
                Next ilEPE
                Exit For
            End If
        Next ilETE
    End If
End Function

Private Function mExportRow(llRow As Long, slEventCategory As String, slEventAutoCode As String) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    
    mExportRow = False
    If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
        slStr = Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX))
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If StrComp(Trim$(tgCurrETE(ilETE).sName), slStr, vbTextCompare) = 0 Then
                slEventCategory = tgCurrETE(ilETE).sCategory
                slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
                If tgCurrETE(ilETE).sCategory = "A" Then
                    Exit Function
                End If
                For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                    If tgCurrEPE(ilEPE).sType = "E" Then
                        If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                            If tgCurrEPE(ilEPE).sBus = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sBusControl = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            'Event Type exported if any other column exported and tgStartColAFE.iEventType >0
                            'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                            If tgCurrEPE(ilEPE).sTime = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sStartType = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sFixedTime = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sEndType = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sDuration = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sMaterialType = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sAudioName = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sAudioItemID = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sAudioISCI = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sAudioControl = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sBkupAudioName = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sBkupAudioControl = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sProtAudioName = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sProtAudioItemID = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sProtAudioISCI = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sProtAudioControl = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sRelay1 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sRelay2 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sFollow = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sSilenceTime = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sSilence1 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sSilence2 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sSilence3 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sSilence4 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sStartNetcue = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sStopNetcue = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sTitle1 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sTitle2 = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If (tgCurrEPE(ilEPE).sABCFormat = "Y") And (sgClientFields = "A") Then
                                mExportRow = True
                                Exit Function
                            End If
                            If (tgCurrEPE(ilEPE).sABCPgmCode = "Y") And (sgClientFields = "A") Then
                                mExportRow = True
                                Exit Function
                            End If
                            If (tgCurrEPE(ilEPE).sABCXDSMode = "Y") And (sgClientFields = "A") Then
                                mExportRow = True
                                Exit Function
                            End If
                            If (tgCurrEPE(ilEPE).sABCRecordItem = "Y") And (sgClientFields = "A") Then
                                mExportRow = True
                                Exit Function
                            End If
                            Exit For
                        End If
                    End If
                Next ilEPE
                Exit For
            End If
        Next ilETE
    End If
End Function

Private Sub mSetColExportColor(llRow As Long)
    Dim ilRet As Integer
    Dim slEventCategory As String
    Dim slEventAutoCode As String
    Dim llCol As Long
    Dim llSvCol As Long
    Dim llSvRow As Long
    
    llSvRow = grdLibEvents.Row
    llSvCol = grdLibEvents.Col
    ilRet = mExportRow(llRow, slEventCategory, slEventAutoCode)
    If Not ilRet Then
        For llCol = EVENTTYPEINDEX To imMaxCols Step 1
            grdLibEvents.Row = llRow
            grdLibEvents.Col = llCol
            grdLibEvents.CellForeColor = vbBlue
        Next llCol
    Else
        For llCol = EVENTTYPEINDEX To imMaxCols Step 1
            grdLibEvents.Row = llRow
            grdLibEvents.Col = llCol
            If (grdLibEvents.CellForeColor <> vbRed) And (grdLibEvents.CellForeColor <> vbMagenta) Then
                If Not mExportCol(llRow, llCol) Then
                    grdLibEvents.CellForeColor = vbBlue
                Else
                    grdLibEvents.CellForeColor = vbBlack
                End If
            End If
        Next llCol
    End If
    grdLibEvents.Col = llSvCol
    grdLibEvents.Row = llSvRow
End Sub

Private Sub mCreateAudioRecs(llRow As Long, slType As String, slAudio As String, llEventStartTime As Long, llEventEndTime As Long, slDays As String, tlConflict() As CONFLICTTEST)
    Dim llUpper As Long
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim slCheckConflicts As String
    
    If slAudio = "" Then
        Exit Sub
    End If
    llUpper = UBound(tlConflict)
    tlConflict(llUpper).lRow = llRow
    tlConflict(llUpper).sType = slType
    tlConflict(llUpper).sDays = slDays
    gGetPreAndPostAudioTime slAudio, llPreTime, llPostTime, slCheckConflicts
    If slCheckConflicts = "N" Then
        Exit Sub
    End If
    If llEventEndTime <= 864000 Then
        If llEventStartTime - llPreTime >= 0 Then
            If llEventEndTime + llPostTime <= 864000 Then
                tlConflict(llUpper).lEventStartTime = llEventStartTime - llPreTime
                tlConflict(llUpper).lEventEndTime = llEventEndTime + llPostTime
                llUpper = llUpper + 1
                ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
            Else
                tlConflict(llUpper).lEventStartTime = llEventStartTime - llPreTime
                tlConflict(llUpper).lEventEndTime = 864000
                llUpper = llUpper + 1
                ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
            End If
        Else
            tlConflict(llUpper).lRow = llRow
            tlConflict(llUpper).sType = slType
            tlConflict(llUpper).sDays = slDays
            tlConflict(llUpper).lEventStartTime = 0
            tlConflict(llUpper).lEventEndTime = llEventEndTime + llPostTime
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
    Else
        tlConflict(llUpper).lEventStartTime = llEventStartTime - llPreTime
        tlConflict(llUpper).lEventEndTime = 864000
        llUpper = llUpper + 1
        ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
    End If
End Sub

Private Sub mCreateBusRecs(llRow As Long, slType As String, slIgnoreConflicts As String, llEventStartTime As Long, llEventEndTime As Long, slDays As String, tlConflict() As CONFLICTTEST)
    Dim llUpper As Long
    
    If (slIgnoreConflicts = "B") Or (slIgnoreConflicts = "I") Then
        Exit Sub
    End If
    llUpper = UBound(tlConflict)
    tlConflict(llUpper).lRow = llRow
    tlConflict(llUpper).sType = slType
    tlConflict(llUpper).sDays = slDays
    If llEventEndTime <= 864000 Then
        tlConflict(llUpper).lEventStartTime = llEventStartTime
        tlConflict(llUpper).lEventEndTime = llEventEndTime
        llUpper = llUpper + 1
        ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
    Else
        tlConflict(llUpper).lEventStartTime = llEventStartTime
        tlConflict(llUpper).lEventEndTime = 864000
        llUpper = llUpper + 1
        ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
    End If
End Sub

Private Sub mInitConflictTest()
    Dim llRow As Long
    
    lmConflictRow = -1
    imAnyEvtChgs = False
    For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
        If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
            If Trim$(grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX)) = "Y" Then
                imAnyEvtChgs = True
            End If
            If Trim$(grdLibEvents.TextMatrix(llRow, PCODEINDEX)) = "" Then
                grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0"
            End If
            'Only test events inserted by row or events modified by the user
            'Creation of new schedule events don't need to be checked
            'If grdLibEvents.TextMatrix(llRow, PCODEINDEX) = "0" Then
            '    grdLibEvents.TextMatrix(llRow, CHGSTATUSINDEX) = "Y"
            '    imAnyEvtChgs = True
            'End If
            grdLibEvents.TextMatrix(llRow, ERRORCONFLICTINDEX) = "0"
        End If
    Next llRow
End Sub

Private Function mCheckLibConflicts() As Integer
    Dim llRow As Long
    Dim llDheCode As Long
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slDays As String
    Dim slStr As String
    ReDim ilCols(0 To 15) As Integer
    
    mCheckLibConflicts = False
    llDheCode = 0
    
    slStartDate = smAirDate
    slEndDate = smAirDate
    If Not gIsDate(slStartDate) Then
        mCheckLibConflicts = True
        Exit Function
    End If
    If Not gIsDate(slEndDate) Then
        mCheckLibConflicts = True
        Exit Function
    End If
    If Not imAnyEvtChgs Then
        Exit Function
    End If
    ilCols(0) = ERRORCONFLICTINDEX
    ilCols(1) = EVENTTYPEINDEX
    ilCols(2) = -1  'AIRHOURSINDEX
    ilCols(3) = -1  'AIRDAYSINDEX
    ilCols(4) = TIMEINDEX
    ilCols(5) = DURATIONINDEX
    ilCols(6) = BUSNAMEINDEX
    ilCols(7) = AUDIONAMEINDEX
    ilCols(8) = PROTNAMEINDEX
    ilCols(9) = BACKUPNAMEINDEX
    ilCols(10) = AUDIOITEMIDINDEX
    ilCols(11) = PROTITEMIDINDEX
    ilCols(12) = AUDIOITEMIDINDEX
    ilCols(13) = CHGSTATUSINDEX
    ilCols(14) = EVTCONFLICTINDEX
    ilCols(15) = DEECODEINDEX
    'ilRet = gCheckConflicts("S", llDheCode, 0, slStartDate, slEndDate, "", grdLibEvents, ilCols(), tmConflictList())
    ilRet = gConflictTableCheck("S", llDheCode, 0, slStartDate, slEndDate, "", grdLibEvents, ilCols(), tmConflictList())
    If ilRet Then
        mCheckLibConflicts = True
        For llRow = grdLibEvents.FixedRows To grdLibEvents.Rows - 1 Step 1
            If Trim$(grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX)) <> "" Then
                If Trim$(grdLibEvents.TextMatrix(llRow, ERRORCONFLICTINDEX)) <> "" Then
                    If Val(grdLibEvents.TextMatrix(llRow, ERRORCONFLICTINDEX)) > 0 Then
                        grdLibEvents.TextMatrix(llRow, ERRORFIELDSORTINDEX) = "0"
                    End If
                End If
            End If
        Next llRow
    End If
End Function

Private Function mGetLibName(slLibNameOrDDECode As String) As String
    Dim slStr As String
    Dim llDNE As Long
    Dim llDSE As Long
    Dim ilRet As Integer
    Dim llDeeCode As Long
    Dim ilPos As Integer
    Dim slSpotTime As String
    Dim slDEECode As String
    
    slStr = slLibNameOrDDECode
    ilPos = InStr(1, slLibNameOrDDECode, "DEE=", vbTextCompare)
    If ilPos > 0 Then
        slSpotTime = ""
        slDEECode = Val(Mid$(slStr, ilPos + 4))
        ilPos = InStr(1, slDEECode, "/", vbTextCompare)
        If ilPos > 0 Then
            slSpotTime = Mid$(slDEECode, ilPos)
            slDEECode = Left$(slDEECode, ilPos - 1)
        End If
        llDeeCode = Val(slDEECode)
        If llDeeCode > 0 Then
            ilRet = gGetRec_DEE_DayEvent(llDeeCode, "EngrSchdDef-mMoveSEERecToCtrls: DEE", tmDee)
            ilRet = gGetRec_DHE_DayHeaderInfo(tmDee.lDheCode, "EngrSchdDef-mMoveSEERecToCtrls: DHE", tmDHE)
            
            If tmDHE.sType <> "T" Then
                llDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrLibDNE)
                If llDNE <> -1 Then
                    slStr = Trim$(tgCurrLibDNE(llDNE).sName)
                Else
                    slStr = "Name Missing"
                End If
            Else
                llDNE = gBinarySearchDNE(tmDHE.lDneCode, tgCurrTempDNE)
                If llDNE <> -1 Then
                    slStr = Trim$(tgCurrTempDNE(llDNE).sName)
                End If
            End If
            llDSE = gBinarySearchDSE(tmDHE.lDseCode, tgCurrDSE)
            If llDSE <> -1 Then
                slStr = slStr & "/" & Trim$(tgCurrDSE(llDSE).sName)
            End If
            slStr = slStr & slSpotTime
        End If
    End If
    mGetLibName = slStr
End Function

Private Sub mSortErrorsToTop()
    Dim ilCol As Integer
    
    gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
    grdLibEvents.Redraw = False
    If imLastColSorted >= 0 Then
        If imLastSort = flexSortStringNoCaseDescending Then
            imLastSort = flexSortStringNoCaseAscending
        Else
            imLastSort = flexSortStringNoCaseDescending
        End If
        ilCol = imLastColSorted
        mSortCol ilCol
    Else
        imLastSort = -1
        mSortCol TIMEINDEX
    End If
    grdLibEvents.Redraw = True
    gSetMousePointer grdLibEvents, grdConflicts, vbDefault
End Sub

Private Sub tmcStart_Timer()
    Dim ilRet As Integer
    Dim ilETE As Integer
    tmcStart.Enabled = False
    mPopANE
    mPopASE
    mPopBDE
    mPopCCE_Audio
    mPopCCE_Bus
    mPopCTE
    mPopDNE
    mPopDSE
    mPopETE
    mPopFNE
    mPopMTE
    mPopNNE
    mPopRNE
    mPopSCE
    mPopTTE_EndType
    mPopTTE_StartType
    mPopARE
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    imSpotETECode = 0
    smSpotEventTypeName = "Spot"
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).sCategory = "S" Then
            imSpotETECode = tgCurrETE(ilETE).iCode
            smSpotEventTypeName = Trim$(tgCurrETE(ilETE).sName)
            Exit For
        End If
    Next ilETE
    If imSpotETECode <= 0 Then
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
        MsgBox "Spot Event Type not defined", vbCritical + vbOKOnly, "Schedule"
    End If
    
    If igSchdCallType <> 0 Then
        gSetMousePointer grdLibEvents, grdConflicts, vbHourglass
        cccDate.text = sgSchdDate
        mReorderFilter
        cmcShowEvents_Click
        cmcCancel.SetFocus
        gSetMousePointer grdLibEvents, grdConflicts, vbDefault
    End If
    imFieldChgd = False
    imStartChgModeCompleted = True
    
End Sub

Public Sub mPopGrid()
    Dim llRow As Long
    Dim llLoop As Long
    Dim slCategory As String
    Dim ilETE As Integer
    Dim llAvailLength As Long
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    Dim ilRet As Integer
    Dim llCTE As Long
    
    ReDim lgLibDheUsed(0 To 0) As Long
    ReDim lmDeleteCodes(0 To 0) As Long
    ReDim tmCurrSEE(0 To UBound(tgCurrSEE)) As SEE
    ReDim smT1Comment(0 To UBound(tmCurrSEE)) As String
    ReDim tmCurr1CTE_Name(0 To 0) As DEECTE
    ReDim smT2Comment(0 To UBound(tmCurrSEE)) As String
    ReDim tmCurr2CTE_Name(0 To 0) As DEECTE
    llRow = 0
    For llLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
        If (tgCurrSEE(llLoop).sAction <> "D") And (tgCurrSEE(llLoop).sAction <> "R") Then
            LSet tmCurrSEE(llRow) = tgCurrSEE(llLoop)
            tmCurrSEE(llRow).sInsertFlag = "N"
            tmCurrSEE(llRow).lAvailLength = tmCurrSEE(llRow).lDuration
            'Adjust avail time to be after last spot
            slCategory = ""
            ilETE = gBinarySearchETE(tmCurrSEE(llRow).iEteCode, tgCurrETE)
            If ilETE <> -1 Then
                slCategory = tgCurrETE(ilETE).sCategory
            End If
            If slCategory = "A" Then
                tmCurrSEE(llRow).lSpotTime = tmCurrSEE(llRow).lTime
                llAvailLength = tmCurrSEE(llRow).lDuration
                For llAvailTest = 0 To UBound(tgCurrSEE) - 1 Step 1
                    If (tgCurrSEE(llAvailTest).sAction <> "D") And (tgCurrSEE(llAvailTest).sAction <> "R") Then
'                        If tgCurrSEE(llAvailTest).lTime <> tgCurrSEE(llLoop).lTime Then
'                            Exit For
'                        End If
                        If (tgCurrSEE(llAvailTest).lTime = tgCurrSEE(llLoop).lTime) And (llLoop <> llAvailTest) Then
                            If tgCurrSEE(llAvailTest).iBdeCode = tmCurrSEE(llRow).iBdeCode Then
                                ilETE = gBinarySearchETE(tgCurrSEE(llAvailTest).iEteCode, tgCurrETE)
                                If ilETE <> -1 Then
                                    If tgCurrETE(ilETE).sCategory = "S" Then
                                        llAvailLength = llAvailLength - tgCurrSEE(llAvailTest).lDuration
                                        llTimeTest = tgCurrSEE(llAvailTest).lSpotTime + tgCurrSEE(llAvailTest).lDuration
                                        If llTimeTest > tmCurrSEE(llRow).lSpotTime Then
                                            tmCurrSEE(llRow).lSpotTime = llTimeTest
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next llAvailTest
                tmCurrSEE(llRow).lAvailLength = llAvailLength
            End If
            smT1Comment(llRow) = ""
            If tgCurrSEE(llLoop).l1CteCode > 0 Then
                For llCTE = 0 To UBound(tmCurr1CTE_Name) - 1 Step 1
                    If (tmCurr1CTE_Name(llCTE).lDheCode = tgCurrSEE(llLoop).lDheCode) Then
                        If (tmCurr1CTE_Name(llCTE).lCteCode = tgCurrSEE(llLoop).l1CteCode) Then
                            smT1Comment(llRow) = Trim$(tmCurr1CTE_Name(llCTE).sComment)
                            Exit For
                        End If
                    End If
                Next llCTE
                If smT1Comment(llRow) = "" Then
                    ilRet = gGetRec_CTE_CommtsTitleAPI(hmCTE, tgCurrSEE(llLoop).l1CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
                    smT1Comment(llRow) = Trim$(tmCTE.sComment)
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).sComment = smT1Comment(llRow)
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lCteCode = tgCurrSEE(llLoop).l1CteCode
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lDeeCode = tgCurrSEE(llLoop).lDeeCode
                    tmCurr1CTE_Name(UBound(tmCurr1CTE_Name)).lDheCode = tgCurrSEE(llLoop).lDheCode
                    ReDim Preserve tmCurr1CTE_Name(0 To UBound(tmCurr1CTE_Name) + 1) As DEECTE
                End If
            End If
            '7/8/11: Make T2 work like T1
            smT2Comment(llRow) = ""
            If tgCurrSEE(llLoop).l2CteCode > 0 Then
                For llCTE = 0 To UBound(tmCurr2CTE_Name) - 1 Step 1
                    If (tmCurr2CTE_Name(llCTE).lDheCode = tgCurrSEE(llLoop).lDheCode) Then
                        If (tmCurr2CTE_Name(llCTE).lCteCode = tgCurrSEE(llLoop).l2CteCode) Then
                            smT2Comment(llRow) = Trim$(tmCurr2CTE_Name(llCTE).sComment)
                            Exit For
                        End If
                    End If
                Next llCTE
                If smT2Comment(llRow) = "" Then
                    ilRet = gGetRec_CTE_CommtsTitleAPI(hmCTE, tgCurrSEE(llLoop).l2CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
                    smT2Comment(llRow) = Trim$(tmCTE.sComment)
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).sComment = smT2Comment(llRow)
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lCteCode = tgCurrSEE(llLoop).l2CteCode
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lDeeCode = tgCurrSEE(llLoop).lDeeCode
                    tmCurr2CTE_Name(UBound(tmCurr2CTE_Name)).lDheCode = tgCurrSEE(llLoop).lDheCode
                    ReDim Preserve tmCurr2CTE_Name(0 To UBound(tmCurr2CTE_Name) + 1) As DEECTE
                End If
            End If
            llRow = llRow + 1
        Else
            If tgCurrSEE(llLoop).sSentStatus <> "S" Then
                lmDeleteCodes(UBound(lmDeleteCodes)) = tgCurrSEE(llLoop).lCode
                ReDim Preserve lmDeleteCodes(0 To UBound(lmDeleteCodes) + 1) As Long
            End If
        End If
    Next llLoop
    ReDim Preserve tmCurrSEE(0 To llRow) As SEE
    ReDim Preserve smT1Comment(0 To llRow) As String
    ReDim Preserve smT2Comment(0 To llRow) As String
    ReDim lmChgStatusSEECode(0 To 0) As Long
    If UBound(tgCurrSEE) > 0 Then
        ArraySortTyp fnAV(tgCurrSEE(), 0), UBound(tgCurrSEE), 0, LenB(tgCurrSEE(0)), 0, -2, 0
    End If
    grdLibEvents.Redraw = False
    grdLibEvents.Visible = False
    mMoveSEERecToCtrls
    imLastColSorted = -1
    mSortCol TIMEINDEX
    grdLibEvents.Redraw = True
    grdLibEvents.Visible = True

End Sub

Private Sub mGridErrorRows(llErrorColor As Long, llGridRow1 As Long, ilGridCol1 As Integer, llGridRow2 As Long, ilGridCol2 As Integer, ilError As Integer)
    Dim ilConflictIndex As Integer
    Dim ilRet As Integer
    Dim ilStartIndex As Integer
    
    grdLibEvents.TextMatrix(llGridRow1, ERRORFIELDSORTINDEX) = "0"
    grdLibEvents.TextMatrix(llGridRow2, ERRORFIELDSORTINDEX) = "0"
    grdLibEvents.Row = llGridRow1   'llRow2
    grdLibEvents.Col = ilGridCol1
    grdLibEvents.CellForeColor = llErrorColor
    grdLibEvents.Row = llGridRow2
    grdLibEvents.Col = ilGridCol2
    grdLibEvents.CellForeColor = llErrorColor
    If Not ilError Then
        ilStartIndex = Val(grdLibEvents.TextMatrix(llGridRow1, ERRORCONFLICTINDEX))
        If ilStartIndex = 0 Then
            grdLibEvents.TextMatrix(llGridRow1, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartIndex, tmConflictList())))
        Else
            ilRet = mAddConflict(ilStartIndex, tmConflictList())
        End If
        ilConflictIndex = UBound(tmConflictList)
        tmConflictList(ilConflictIndex).sType = "E"
        tmConflictList(ilConflictIndex).sStartDate = ""
        tmConflictList(ilConflictIndex).sEndDate = ""
        tmConflictList(ilConflictIndex).lIndex = Val(grdLibEvents.TextMatrix(llGridRow1, TMCURRSEEINDEX))
        tmConflictList(ilConflictIndex).iNextIndex = -1
        ilStartIndex = Val(grdLibEvents.TextMatrix(llGridRow2, ERRORCONFLICTINDEX))
        If ilStartIndex = 0 Then
            grdLibEvents.TextMatrix(llGridRow2, ERRORCONFLICTINDEX) = Trim$(Str$(mAddConflict(ilStartIndex, tmConflictList())))
        Else
            ilRet = mAddConflict(ilStartIndex, tmConflictList())
        End If
    End If
    ilError = True
End Sub

Private Sub mSetAvailTime()
    Dim llLoop As Long
    Dim slCategory As String
    Dim ilETE As Integer
    Dim llAvailLength As Long
    Dim llAvailTest As Long
    Dim llTimeTest As Long
    
    For llLoop = 0 To UBound(tmCurrSEE) - 1 Step 1
        If (tmCurrSEE(llLoop).sAction <> "D") And (tmCurrSEE(llLoop).sAction <> "R") Then
            'Adjust avail time to be after last spot
            slCategory = ""
            ilETE = gBinarySearchETE(tmCurrSEE(llLoop).iEteCode, tgCurrETE)
            If ilETE <> -1 Then
                slCategory = tgCurrETE(ilETE).sCategory
            End If
            If slCategory = "A" Then
                tmCurrSEE(llLoop).lSpotTime = tmCurrSEE(llLoop).lTime
                llAvailLength = tmCurrSEE(llLoop).lDuration
                For llAvailTest = 0 To UBound(tmCurrSEE) - 1 Step 1
                    If (tmCurrSEE(llAvailTest).sAction <> "D") And (tmCurrSEE(llAvailTest).sAction <> "R") Then
                        If (tmCurrSEE(llAvailTest).lTime = tmCurrSEE(llLoop).lTime) And (llLoop <> llAvailTest) Then
                            If tmCurrSEE(llAvailTest).iBdeCode = tmCurrSEE(llLoop).iBdeCode Then
                                ilETE = gBinarySearchETE(tmCurrSEE(llAvailTest).iEteCode, tgCurrETE)
                                If ilETE <> -1 Then
                                    If tgCurrETE(ilETE).sCategory = "S" Then
                                        llAvailLength = llAvailLength - tmCurrSEE(llAvailTest).lDuration
                                        llTimeTest = tmCurrSEE(llAvailTest).lSpotTime + tmCurrSEE(llAvailTest).lDuration
                                        If llTimeTest > tmCurrSEE(llLoop).lSpotTime Then
                                            tmCurrSEE(llLoop).lSpotTime = llTimeTest
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next llAvailTest
                tmCurrSEE(llLoop).lAvailLength = llAvailLength
            End If
        End If
    Next llLoop

End Sub

Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    Dim llSvCol As Long
    Dim llSvRow As Long
    Dim llSvTopRow As Long
    
    If (llRow >= grdLibEvents.FixedRows) And (llRow < grdLibEvents.Rows) Then
        grdLibEvents.Redraw = False
        llSvTopRow = grdLibEvents.TopRow
        llSvRow = grdLibEvents.Row
        llSvCol = grdLibEvents.Col
        If grdLibEvents.TextMatrix(llRow, EVENTTYPEINDEX) <> "" Then
            For llCol = EVENTTYPEINDEX To ABCRECORDITEMINDEX Step 1
                grdLibEvents.Row = llRow
                grdLibEvents.Col = llCol
                If grdLibEvents.CellBackColor <> LIGHTYELLOW Then
                    If lmEEnableRow <> llRow Then
                        grdLibEvents.CellBackColor = vbWhite
                    Else
                        grdLibEvents.CellBackColor = GRAY
                    End If
                End If
            Next llCol
        End If
        grdLibEvents.TopRow = llSvTopRow
        grdLibEvents.Row = llSvRow
        grdLibEvents.Col = llSvCol
        grdLibEvents.Redraw = True
    End If
End Sub

Public Function mBinarySearchOldSEE(llCode As Long) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If llCode <= 0 Then
        mBinarySearchOldSEE = -1
        Exit Function
    End If
    llMin = LBound(tgCurrSEE)
    llMax = UBound(tgCurrSEE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tgCurrSEE(llMiddle).lCode Then
            'found the match
            mBinarySearchOldSEE = llMiddle
            Exit Function
        ElseIf llCode < tgCurrSEE(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchOldSEE = -1
End Function


Private Sub mReorderFilter()
    Dim ilFilter As Integer
    Dim ilIndex As Integer
    
    'Reorder, Place Equal and Not Equal at Top
    ReDim tmFilterValues(0 To UBound(tgFilterValues)) As FILTERVALUES
    For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
        tgFilterValues(ilFilter).iUsed = False
    Next ilFilter
    ilIndex = LBound(tmFilterValues)
    For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
        If tgFilterValues(ilFilter).iUsed = False Then
            If (tgFilterValues(ilFilter).iOperator = 1) Or (tgFilterValues(ilFilter).iOperator = 2) Then
                LSet tmFilterValues(ilIndex) = tgFilterValues(ilFilter)
                ilIndex = ilIndex + 1
                tgFilterValues(ilFilter).iUsed = True
            End If
        End If
    Next ilFilter
    For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
        If tgFilterValues(ilFilter).iUsed = False Then
            LSet tmFilterValues(ilIndex) = tgFilterValues(ilFilter)
            ilIndex = ilIndex + 1
            tgFilterValues(ilFilter).iUsed = True
        End If
    Next ilFilter
End Sub

Private Function mSaveAndClearFilter(blSave As Boolean, blClearFilter As Boolean) As Integer
    Dim ilRet As Integer
    
    grdLibEvents.Redraw = False
    If (imFieldChgd = True) And (blSave) Then
        cmcTask.Caption = "Saving Changed Events...."
        cmcTask.Visible = True
        DoEvents
        ilRet = mSave()
        bmInSave = False
        If ilRet = False Then
            cmcTask.Visible = False
            grdLibEvents.Redraw = True
            gSetMousePointer grdLibEvents, grdConflicts, vbDefault
            mSaveAndClearFilter = False
            Exit Function
        End If
        cmcTask.Visible = False
    End If
    grdLibEvents.Redraw = False
    If (UBound(tgFilterValues) > LBound(tgFilterValues)) And (blClearFilter) Then
        cmcTask.Caption = "Removing Filter...."
        cmcTask.Visible = True
        imStartChgModeCompleted = False
        ReDim tgFilterValues(0 To 0) As FILTERVALUES
        mReorderFilter
        DoEvents
        cmcTask.Caption = "Loading Events...."
        cmcShowEvents_Click
        cmcTask.Visible = False
    End If
    mSaveAndClearFilter = True
End Function

Private Function mBusInFilter() As Boolean
    Dim ilFilter As Integer
    Dim ilField As Integer
    Dim ilFilterType As Integer
    Dim slFileName As String
    
    For ilFilter = LBound(tmFilterValues) To UBound(tmFilterValues) - 1 Step 1
        If tmFilterValues(ilFilter).iOperator = 1 Then
            For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
                If tgFilterFields(ilField).sFieldName = tmFilterValues(ilFilter).sFieldName Then
                    ilFilterType = tgFilterFields(ilField).iFieldType
                    slFileName = tgFilterFields(ilField).sListFile
                    If ilFilterType = 5 Then
                        If UCase$(Trim$(slFileName)) = "BDE" Then
                            mBusInFilter = True
                            Exit Function
                        End If
                    End If
                End If
            Next ilField
        End If
    Next ilFilter
    mBusInFilter = False
End Function

Private Sub mFindBrackets(llSEEOld As Long)
    Dim llRow As Long
    Dim llStartIndex As Long
    Dim llEndIndex As Long
    Dim llUpper As Long
    Dim llIndex As Long
    Dim llDel As Integer
    Dim blBypass As Boolean
    
    If tmSHE.sLoadedAutoStatus <> "L" Then
        Exit Sub
    End If
    llIndex = -1
    For llRow = 0 To UBound(tmCCurrSEE) - 1 Step 1
        If tmCCurrSEE(llRow).lCode = tgCurrSEE(llSEEOld).lCode Then
            llIndex = llRow
            Exit For
        End If
    Next llRow
    If llIndex = -1 Then
        Exit Sub
    End If
    'Find previous matching Bus
    llUpper = UBound(tmSeeBracket)
    llStartIndex = -1
    For llRow = llIndex - 1 To LBound(tmCCurrSEE) Step -1
        If tmCCurrSEE(llRow).iBdeCode = tmCCurrSEE(llIndex).iBdeCode Then
            '2/15/12: Bypass adjacent deleted items
            blBypass = False
            For llDel = LBound(lmDeleteCodes) To UBound(lmDeleteCodes) - 1 Step 1
                If lmDeleteCodes(llDel) = tmCCurrSEE(llRow).lCode Then
                    lmDeleteCodes(llDel) = -lmDeleteCodes(llDel)
                    blBypass = True
                    Exit For
                End If
            Next llDel
            If Not blBypass Then
                llStartIndex = llRow
                tmSeeBracket(llUpper).sSource = "C"
                tmSeeBracket(llUpper).lIndex = llStartIndex
                llUpper = llUpper + 1
                ReDim Preserve tmSeeBracket(0 To llUpper) As SEEBRACKET
                Exit For
            End If
        End If
    Next llRow
    If llStartIndex = -1 Then
        'Search previous day
        For llRow = UBound(tmPCurrSEE) - 1 To LBound(tmPCurrSEE) Step -1
            If tmPCurrSEE(llRow).iBdeCode = tmCCurrSEE(llIndex).iBdeCode Then
                llStartIndex = llRow
                tmSeeBracket(llUpper).sSource = "P"
                tmSeeBracket(llUpper).lIndex = llStartIndex
                llUpper = llUpper + 1
                ReDim Preserve tmSeeBracket(0 To llUpper) As SEEBRACKET
                Exit For
            End If
        Next llRow
    End If
    'Find Next matching bus
    llEndIndex = -1
    For llRow = llIndex + 1 To UBound(tmCCurrSEE) Step 1
        If tmCCurrSEE(llRow).iBdeCode = tmCCurrSEE(llIndex).iBdeCode Then
            '2/15/12: Bypass adjacent deleted items
            blBypass = False
            For llDel = LBound(lmDeleteCodes) To UBound(lmDeleteCodes) - 1 Step 1
                If lmDeleteCodes(llDel) = tmCCurrSEE(llRow).lCode Then
                    lmDeleteCodes(llDel) = -lmDeleteCodes(llDel)
                    blBypass = True
                    Exit For
                End If
            Next llDel
            If Not blBypass Then
                llEndIndex = llRow
                tmSeeBracket(llUpper).sSource = "C"
                tmSeeBracket(llUpper).lIndex = llEndIndex
                llUpper = llUpper + 1
                ReDim Preserve tmSeeBracket(0 To llUpper) As SEEBRACKET
                Exit For
            End If
        End If
    Next llRow
    If llEndIndex = -1 Then
        'Search Next day
        For llRow = 0 To UBound(tmNCurrSEE) Step 1
            If tmNCurrSEE(llRow).iBdeCode = tmCCurrSEE(llIndex).iBdeCode Then
                llEndIndex = llRow
                tmSeeBracket(llUpper).sSource = "N"
                tmSeeBracket(llUpper).lIndex = llEndIndex
                llUpper = llUpper + 1
                ReDim Preserve tmSeeBracket(0 To llUpper) As SEEBRACKET
                Exit For
            End If
        Next llRow
    End If
End Sub

Private Function mGenUPDFile() As Boolean
    Dim llRow As Long
    Dim tlSEE As SEE
    Dim ilLength As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilSend As Integer
    Dim slDate As String
    Dim slDateTime As String
    Dim ilEteCode As Integer
    Dim slEventCategory As String
    Dim slEventAutoCode As String
    Dim llSEECode As Long
    Dim ilRet As Integer
    Dim llOldSHECode As Long
    Dim slToFile As String
    

    If Not mBusInFilter() Then
        ilRet = mOpenAutoExportFile(slToFile)
        If Not ilRet Then
            mGenUPDFile = False
            Exit Function
        End If
    End If
    
    ilLength = gExportStrLength()
    
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = DateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    
    If tgNoCharAFE.iDate = 8 Then
        slDate = Format$(smAirDate, "yyyymmdd")
    ElseIf tgNoCharAFE.iDate = 6 Then
        slDate = Format$(smAirDate, "yymmdd")
    End If
    
    For llRow = 0 To UBound(tmSeeBracket) - 1 Step 1
        If tmSeeBracket(llRow).sSource = "P" Then
            tlSEE = tmPCurrSEE(tmSeeBracket(llRow).lIndex)
        ElseIf tmSeeBracket(llRow).sSource = "C" Then
            tlSEE = tmCCurrSEE(tmSeeBracket(llRow).lIndex)
        ElseIf tmSeeBracket(llRow).sSource = "N" Then
            tlSEE = tmNCurrSEE(tmSeeBracket(llRow).lIndex)
        End If
        ilEteCode = tlSEE.iEteCode
        If gAutoExportRow(ilEteCode, slEventCategory, slEventAutoCode) Then
            'Check If today and enough time
            ilSend = True
            If DateValue(smAirDate) = DateValue(slNowDate) Then
                If llNowTime > tlSEE.lTime Then
                    ilSend = False
                End If
            End If
            If ilSend Then
                If tlSEE.sAction <> "D" Then
                    gAutoSendSEE hmExport, slEventCategory, slEventAutoCode, slDate, ilEteCode, ilLength, tlSEE
                End If
                'Update SEE
                llSEECode = tlSEE.lCode
                If llSEECode > 0 Then
                    ilRet = gPutUpdate_SEE_SentFlag(llSEECode, "EngrSchd- Update SEE Sent Flag")
                End If
            End If
        End If
    Next llRow
    
    For llRow = 0 To UBound(lmChgSEE) - 1 Step 1
        tlSEE = tmCurrSEE(lmChgSEE(llRow))
        ilEteCode = tlSEE.iEteCode
        If gAutoExportRow(ilEteCode, slEventCategory, slEventAutoCode) Then
            'Check If today and enough time
            ilSend = True
            If DateValue(smAirDate) = DateValue(slNowDate) Then
                If llNowTime > tlSEE.lTime Then
                    ilSend = False
                End If
            End If
            If ilSend Then
                If tlSEE.sAction <> "D" Then
                    gAutoSendSEE hmExport, slEventCategory, slEventAutoCode, slDate, ilEteCode, ilLength, tlSEE
                End If
                'Update SEE
                llSEECode = tlSEE.lCode
                If llSEECode > 0 Then
                    ilRet = gPutUpdate_SEE_SentFlag(llSEECode, "EngrSchd- Update SEE Sent Flag")
                End If
            End If
        End If
    Next llRow
    On Error Resume Next
    Close hmMsg
    Close hmExport
    gRenameExportFile
    tmSHE.sLoadStatus = "N"
    ilRet = gPutUpdate_SHE_ScheduleHeader(7, tmSHE, "Schedule Definition-mCreateAuto: Update SHE", 0)
    ilRet = gPutUpdate_SHE_SentFlags(tmSHE.lCode, "EngrSchd- Update SHE Sent Flags")
    tmSHE.sLoadedAutoStatus = "L"
    tmSHE.iChgSeqNo = tmSHE.iChgSeqNo + 1
    tmSHE.sLoadedAutoDate = Format$(gNow(), sgShowDateForm)
    tmSHE.sCreateLoad = "N"
    ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE", llOldSHECode)
End Function

Private Function mOkToGenUPD() As Integer
    Dim ilRet As Integer
    
    If Not gOpenAutoMsgFile(smAirDate, smMsgFileName, hmMsg) Then
        MsgBox "Unable to Create Load Message file: " & smMsgFileName & " for " & smAirDate
        mOkToGenUPD = False
        Exit Function
    End If
    If Not gOpenAutoExportFile(tmSHE, smAirDate, smExportFileName, hmExport) Then
        Close #hmMsg
        MsgBox "Unable to Create Load file: " & smExportFileName & " for " & smAirDate
        mOkToGenUPD = False
        Exit Function
    End If
    mOkToGenUPD = True
End Function
